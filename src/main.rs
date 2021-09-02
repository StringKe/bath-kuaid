use std::fs;
use std::io::Write;
use std::ops::Add;
use std::path::Path;
use std::process;

use base64::encode;
use calamine::{open_workbook, Reader, Xlsx};
use chrono::{DateTime, Utc};
use clap::{App, Arg};
use hyper::{Body, Client, Method, Request};
use serde::{Deserialize, Serialize};
use serde_json::Value;
use simple_excel_writer;

#[derive(Serialize, Deserialize, Debug)]
struct Config {
    id: String,
    key: String,
    limit: i64,
    url: String,
    company_cn: Vec<String>,
    company_en: Vec<String>
}

#[derive(Serialize, Deserialize, Debug, Clone)]
struct ListInfo {
    company: String,
    company_code: String,
    oid: String,
    is_retry: bool,
}

#[derive(Default, Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ApiParams {
    #[serde(rename = "RequestData")]
    pub request_data: String,
    #[serde(rename = "EBusinessID")]
    pub ebusiness_id: String,
    #[serde(rename = "RequestType")]
    pub request_type: String,
    #[serde(rename = "DataSign")]
    pub data_sign: String,
    #[serde(rename = "DataType")]
    pub data_type: i64,
}


#[derive(Default, Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ApiData {
    #[serde(rename = "OrderCode")]
    pub order_code: String,
    #[serde(rename = "ShipperCode")]
    pub shipper_code: String,
    #[serde(rename = "LogisticCode")]
    pub logistic_code: String,
}

#[derive(Default, Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TraceItem {
    #[serde(rename = "Action")]
    pub action: String,
    #[serde(rename = "AcceptStation")]
    pub accept_station: String,
    #[serde(rename = "AcceptTime")]
    pub accept_time: String,
    #[serde(rename = "Location")]
    pub location: String,
}


#[tokio::main]
async fn main() {
    let execute_path = std::env::current_dir().unwrap();
    let execute_path = execute_path.to_str().unwrap();

    let config_name = "config.toml";
    let full_path = Path::new(execute_path).join(Path::new(config_name));

    let config = match fs::read_to_string(&full_path) {
        Ok(config) => {
            let c: Config = toml::from_str(&config).unwrap();
            c
        },
        Err(_) => {
            let mut company_cn = Vec::new();
            company_cn.push(String::from("京东"));
            let mut company_en = Vec::new();
            company_en.push(String::from("JD"));

            let default: Config = Config {
                id: String::from("未设置"),
                key: String::from("未设置"),
                limit: 60,
                url: String::from("未设置"),
                company_cn,
                company_en
            };

            let default_raw = toml::to_string(&default).unwrap();
            let file_path = *&full_path.to_str().unwrap();
            println!("配置文件已创建，请补充完成 {}", file_path);
            println!("{:?}", default_raw);
            let mut file = fs::File::create(file_path).unwrap();
            file.write_all(default_raw.as_bytes()).expect("写入配置文件错误");
            process::exit(1);
        }
    };

    let matches = App::new("快递鸟批量查询")
        .version("1.0")
        .author("青木 <atuehmail@qq.com>")
        .about("批量查询快递信息，注意！！！ 数据列表的第一行和第最后一行是不会进行查询\n\
        数据格式：\n\
第一行     物流公司|运单|是否重试
最后一行    要存在，随便输入什么
        ")
        .arg(Arg::new("file")
            .short('f')
            .long("file")
            .value_name("FILE")
            .about("xlsx 文件路径")
            .required(true)
            .takes_value(true))
        .arg(Arg::new("retry")
            .short('r')
            .about("是否重试文件中失败的"))
        .get_matches();


    println!("配置信息 {:?}", config);

    let file_path = matches.value_of("file").unwrap();
    let is_retry = matches.is_present("retry");
    let row_list = read_xls(&file_path, &config);

    let all_execute_list = match is_retry {
        true => {
            let mut list: Vec<ListInfo> = Vec::new();
            for row in row_list {
                if row.is_retry {
                    list.push(ListInfo {
                        oid: row.oid,
                        company: row.company,
                        company_code: row.company_code,
                        is_retry: false,
                    })
                }
            }
            list
        }
        false => {
            row_list
        }
    };

    let mut rows: Vec<Vec<String>> = Vec::new();
    let execute_start_time: DateTime<Utc> = Utc::now();

    let mut start = Vec::new();
    start.push(String::from("物流公司"));
    start.push(String::from("快递单号"));
    start.push(String::from("执行状态"));
    start.push(String::from("==="));
    start.push(String::from("代号"));
    start.push(String::from("物流状态"));
    start.push(String::from("增值物流状态"));
    start.push(String::from("城市"));
    start.push(String::from("==="));
    start.push(String::from("最后更新时间"));
    start.push(String::from("最后更新信息"));
    start.push(String::from("最后更新位置"));
    start.push(String::from("==="));
    start.push(String::from("发出时间"));
    start.push(String::from("发出信息"));
    start.push(String::from("发出位置"));
    start.push(String::from("派件时间"));
    start.push(String::from("派件信息"));
    start.push(String::from("派件位置"));
    start.push(String::from("签收时间"));
    start.push(String::from("签收信息"));
    start.push(String::from("签收位置"));
    start.push(String::from("==="));
    start.push(String::from("问题时间"));
    start.push(String::from("问题信息"));
    start.push(String::from("问题位置"));

    rows.push(start);


    println!("总执行数据 {}", all_execute_list.len());

    for (execute_index, execute) in all_execute_list.iter().enumerate() {
        println!("正在获取 {} {}#{} 的数据", execute_index + 1, &execute.company, &execute.oid);
        let data = request_data(&execute.oid, &execute.company, &config).await;
        match data {
            None => {
                println!("本条数据执行失败 {} {}", &execute.company, &execute.oid)
            }
            Some(execute_data) => {
                let id = (&execute.oid).parse().unwrap();
                let company = (&execute.company).parse().unwrap();
                let company_code = (&execute.company_code).parse().unwrap();
                let status = String::from("1");

                let mut line = Vec::new();
                line.push(company_code);
                line.push(id);
                line.push(status);
                line.push(String::from("==="));
                line.push(company);

                let state_raw = execute_data["State"].as_str().unwrap();
                line.push(transform_code(state_raw));

                let state_ex_raw = execute_data["StateEx"].as_str().unwrap();
                line.push(transform_code(state_ex_raw));

                let location_raw = execute_data["Location"].as_str();
                match location_raw {
                    None => {
                        line.push(String::from(""));
                    }
                    Some(val) => {
                        line.push(String::from(val));
                    }
                }

                line.push(String::from("==="));

                let traces = execute_data["Traces"].as_array();
                match traces {
                    None => {
                        line.push(String::from(""));
                        line.push(String::from(""));
                        line.push(String::from(""));
                        line.push(String::from(""));
                        line.push(String::from(""));
                        line.push(String::from(""));
                    }
                    Some(list) => {
                        let last = list.last();
                        match last {
                            None => {
                                line.push(String::from(""));
                                line.push(String::from(""));
                                line.push(String::from(""));
                            }
                            Some(info) => {
                                let trace_item: TraceItem = serde_json::from_value(info.clone()).unwrap();

                                line.push(trace_item.accept_time);
                                line.push(trace_item.accept_station);
                                line.push(trace_item.location);
                            }
                        }

                        let send_info = find_trace_by_code(list, String::from("1"), true);
                        match send_info {
                            None => {
                                line.push(String::from("==="));
                                line.push(String::from(""));
                                line.push(String::from(""));
                                line.push(String::from(""));
                            }
                            Some(info) => {
                                line.push(String::from("==="));
                                line.push(info.accept_time);
                                line.push(info.accept_station);
                                line.push(info.location);
                            }
                        }

                        let dispatch_info = find_trace_by_code(list, String::from("2"), false);
                        match dispatch_info {
                            None => {
                                line.push(String::from(""));
                                line.push(String::from(""));
                                line.push(String::from(""));
                            }
                            Some(info) => {
                                line.push(info.accept_time);
                                line.push(info.accept_station);
                                line.push(info.location);
                            }
                        }

                        let single_info = find_trace_by_code(list, String::from("3"), false);
                        match single_info {
                            None => {
                                line.push(String::from(""));
                                line.push(String::from(""));
                                line.push(String::from(""));
                            }
                            Some(info) => {
                                line.push(info.accept_time);
                                line.push(info.accept_station);
                                line.push(info.location);
                            }
                        }


                        let question_info = find_trace_by_code(list, String::from("4"), false);
                        match question_info {
                            None => {
                                line.push(String::from("==="));
                                line.push(String::from(""));
                                line.push(String::from(""));
                                line.push(String::from(""));
                            }
                            Some(info) => {
                                line.push(String::from("==="));
                                line.push(info.accept_time);
                                line.push(info.accept_station);
                                line.push(info.location);
                            }
                        }
                    }
                }

                rows.push(line);
            }
        }
    }

    let mut wb = simple_excel_writer::workbook::Workbook::create_in_memory();
    let mut sheet = wb.create_sheet("Sheet1");

    wb.write_sheet(&mut sheet, |sheet_writer| {
        let sw = sheet_writer;

        for row in rows {
            let row = make_row(row);
            sw.append_row(row).unwrap()
        }

        let execute_end_time: DateTime<Utc> = Utc::now();
        let mut end = Vec::new();

        let start_str = format!("开始时间 {:?}", execute_start_time.format("%Y-%m-%d %H:%M:%S").to_string());
        let end_str = format!("结束时间 {:?}", execute_end_time.format("%Y-%m-%d %H:%M:%S").to_string());

        end.push(String::from(start_str));
        end.push(String::from(end_str));
        sw.append_row(make_row(end))
    }).expect("写入失败!");

    let wb_buf = wb.close().expect("关闭文档失败!").unwrap();

    let old_path = Path::new(&file_path);
    fs::write(old_path, wb_buf).unwrap();
}

fn find_trace_by_code(list: &Vec<Value>, code: String, is_first: bool) -> Option<TraceItem> {
    let mut trace_item: Option<TraceItem> = None;
    for item in list.iter() {
        let action = item["Action"].as_str();
        match action {
            None => {
                continue
            }
            Some(val) => {
                if val.to_string().starts_with(&code) {
                    trace_item = serde_json::from_value(item.clone()).unwrap();
                    if is_first {
                        break;
                    }
                }
            }
        }
    }

    trace_item
}

fn transform_code(code: &str) -> String {
    match code {
        "1" => {
            String::from("已揽收")
        }
        "2" => {
            String::from("在途中")
        }
        "201" => {
            String::from("到达派件城市")
        }
        "202" => {
            String::from("派件中")
        }
        "211" => {
            String::from("已放入快递柜或驿站")
        }
        "3" => {
            String::from("已签收")
        }
        "301" => {
            String::from("正常签收")
        }
        "302" => {
            String::from("代收签收")
        }
        "304" => {
            String::from("快递柜或驿站签收")
        }
        "311" => {
            String::from("已取出快递柜或驿站")
        }
        "401" => {
            String::from("问题件")
        }
        "402" => {
            String::from("超时未签收")
        }
        "403" => {
            String::from("超时未更新")
        }
        "404" => {
            String::from("拒收（退件）")
        }
        "405" => {
            String::from("派件异常")
        }
        "406" => {
            String::from("退货签收")
        }
        "407" => {
            String::from("退货未签收")
        }
        "412" => {
            String::from("快递柜或驿站超时未取")
        }
        v => {
            String::from("未知 ").add(&v)
        }
    }
}

fn make_row(cells: Vec<String>) -> simple_excel_writer::Row {
    let mut row = simple_excel_writer::Row::new();
    for cell in cells {
        row.add_cell(cell)
    }
    row
}

fn make_sing(data: &str, key: &str) -> String {
    let digest = md5::compute(format!("{}{}", data, key));
    format!("{:x}", digest)
}

fn make_req(id: &str, company: &str, config: &Config) -> Request<Body> {
    let data: ApiData = ApiData {
        order_code: String::new(),
        logistic_code: (&id).parse().unwrap(),
        shipper_code: (&company).parse().unwrap(),
    };

    let data_json = serde_json::to_string(&data).unwrap();
    let data_sing = make_sing(&data_json, &config.key);

    let params: ApiParams = ApiParams {
        ebusiness_id: (&config.id).parse().unwrap(),
        request_data: (data_json),
        request_type: String::from("8001"),
        data_sign: encode(data_sing),
        data_type: 2,
    };

    let post_body = serde_urlencoded::to_string(params).unwrap();

    let req = Request::builder()
        .method(Method::POST)
        .uri(&config.url)
        .header("content-type", "application/x-www-form-urlencoded;charset=utf-8")
        .header("user-agent", "kdniao_api_client_bath_query_001")
        .body(Body::from(post_body)).unwrap();
    req
}

async fn request_data(id: &str, company: &str, config: &Config) -> Option<Value> {
    let client = Client::new();

    let req = make_req(&id, &company, &config);

    let resp = client.request(req).await.unwrap();

    if resp.status() != 200 {
        return None
    }

    let body = hyper::body::to_bytes(resp.into_body()).await.unwrap();
    let body_str = String::from_utf8(body.to_vec()).unwrap();

    let body_parser: Value = serde_json::from_str(&*body_str).unwrap();

    match body_parser["Success"] {
        Value::Bool(data) => {
            match data {
                true => {
                    Some(body_parser)
                }
                false => {
                    None
                }
            }
        }
        _ => {
            println!("ERROR {:?}", body_parser);
            None
        }
    }
}

fn convert_company(cn_name: &String, config: &Config) -> Option<String> {
    let index = config.company_cn.iter().position(|r| r == cn_name);
    match index {
        None => {
            None
        }
        Some(result) => {
            match config.company_en.get(result) {
                None => {
                    None
                }
                Some(val) => {
                    Some(val.to_string())
                }
            }
        }
    }
}

fn read_xls(path: &str, config: &Config) -> Vec<ListInfo> {
    match Path::new(&path).exists() {
        true => {
            let mut workbook: Xlsx<_> = open_workbook(path).expect(&*format!("无法读取文件 {}", &path));
            if let Some(Ok(range)) = workbook.worksheet_range("Sheet1") {
                let mut list: Vec<ListInfo> = Vec::new();
                let count = range.rows().count();
                for (index, row) in range.rows().enumerate() {
                    if index == 0 { continue; }
                    if index == count - 1 { continue }

                    let company_raw = row.get(0).unwrap().to_string();
                    let company = convert_company(&company_raw, &config);
                    match company {
                        None => {
                            println!("当前数据转换快递公司失败 {} {}", index + 1, company_raw);
                        }
                        Some(company_real) => {
                            let oid = row.get(1).unwrap().to_string();
                            let is_retry = !(row.get(2).unwrap().is_empty());

                            list.push(ListInfo {
                                company: company_real,
                                company_code: company_raw,
                                oid,
                                is_retry
                            });
                        }
                    }
                }
                list
            } else {
                println!("快递列表 表一 读取失败 {}", &path);
                println!("快递列表 表一 名字必须为 Sheet1");
                process::exit(1);
            }
        }
        false => {
            println!("快递列表文件不存在 {}", &path);
            process::exit(1);
        }
    }
}