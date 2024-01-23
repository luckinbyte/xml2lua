extern crate calamine;

use clap::{crate_description, crate_name, crate_version, Arg, Command};

use calamine::{open_workbook, Reader, Xlsx, Xls};
use calamine::Data;

use std::collections::HashMap;
use std::thread;

use std::fs;
use std::path::Path;
use std::io::Write;
use std::io::Cursor;
use std::{error::Error as StdError, fs::File};

extern crate piccolo;
//use piccolo::{Lua,};
use piccolo::{
    compiler::{self, interning::BasicInterner, CompiledPrototype},
    io,
};

fn is_skip_row(row:&[Data]) -> bool{
    let cell = row.get(0).unwrap().to_string();
    return cell.starts_with("//");
}

fn is_table(s:&String) -> bool{
    return s.starts_with("{")&s.ends_with("}");
}

fn str_trans2_bracenode(s:String) -> BraceNode{
    let mut stack:Vec<BraceNode> = Vec::new();
    let mut temp_node = BraceNode::new("".to_string(), Vec::new());
    for c in s.chars() {
        temp_node.value.push(c);
        match c{
            '}' =>{
                let s_len = stack.len() as usize;
                if s_len==1{
                    return temp_node
                }
                if let Some(mut last_s) = stack.pop(){
                    last_s.value = last_s.value + &temp_node.value;
                    last_s.children.push(temp_node);
                    temp_node = last_s
                }
            }
            '{' =>{
                temp_node.value.pop();
                stack.push(temp_node);
                temp_node = BraceNode::new("".to_string(), Vec::new());
                temp_node.value.push('{');
            }
            _ => ()
        }
    }
    return temp_node;
}

struct BraceNode{
    pub value:String,
    pub children:Vec<BraceNode>,
}

impl BraceNode{
    fn new(s:String, children:Vec<BraceNode>) -> Self{
        return Self{
            value:s,
            children:children
        }
    }
}

fn update_table_counter(node:&BraceNode, map:&mut HashMap<String, u32>, is_child:u32){
    let count = map.entry(node.value.to_string()).or_insert(0);
    *count += 1;
    if is_child==1{
        for child in &node.children{
            update_table_counter(&child, map, 1)
        }
    }
}

fn update_table_str(node:&BraceNode, cell:&String, t_map:&HashMap<&String, u32>) -> String{
    if cell.contains(&node.value){
        if let Some(n) = t_map.get(&node.value){
            let new_s:String = "t[".to_string()+&n.to_string()+&"]".to_string();
            cell.replace(&node.value, &new_s)
        }else{
            cell.to_string()
        }
    }else{
        let mut cell_n = cell.to_string();
        for child in &node.children{
            cell_n = update_table_str(&child, &cell_n, &t_map)
        }
        return cell_n
    }
}

fn create_files(files_dir:Vec<String>, out_dir:String){
    for file in files_dir{
        create_file(file, &out_dir)
    }
}

fn create_file(file_dir:String, out_dir:&String){
    //let path = format!("./ee.xls");
    let mut excel:Xls<_> = open_workbook(file_dir).unwrap();
    let mut out_string:String = "".to_string();

    let first_sheet_name = excel.sheet_names().first().unwrap().to_owned();
    let range = excel.worksheet_range(&first_sheet_name).unwrap();
    let mut i = 0;
    let mut init_head_flag = false;
    let mut heads:Vec<(u32,String)> = Vec::new();
    let mut node_cache:HashMap<String, BraceNode> = HashMap::new();
    let mut table_counter:HashMap<String, u32> = HashMap::new();
    for row in range.rows(){
        if is_skip_row(&row){
            i = i+1;
            continue;
        };
        if !init_head_flag{
            let mut j:u32 = 0;
            for cell in row{
                match excel.getcolor(i,j) {
                    Some(number) => {
                        heads.push((j, cell.to_string()));
                    }
                    _ => {
                        ()
                    }
                }
                j = j+1;
            }
            init_head_flag = true;
            continue;
        }
        i = i+1;
        if init_head_flag{
            for head in &heads{
                let cell = row.get(head.0 as usize).unwrap().to_string();
                if is_table(&cell){
                    if node_cache.contains_key(&cell){
                        if let Some(node) = node_cache.get(&cell){
                            update_table_counter(node, &mut table_counter, 1);
                        }
                    }else{
                        let node = str_trans2_bracenode(cell.to_string());
                        update_table_counter(&node, &mut table_counter, 0);
                        node_cache.insert(cell, node);
                    }
                }
            }
        }
    }
    // table_counter 按times排序， times小于2的忽略
    let mut t:Vec<(&String, &u32)> = table_counter.iter().filter(|(_, &value)| value>=2).collect();
    t.sort_by(|a,b| b.1.cmp(a.1));

    let mut t_map:HashMap<&String, u32> = HashMap::new();

    if t.len()>0{
        out_string.push_str("local t = {\n");
        let mut n:u32 = 1;
        for tt in &t{
            t_map.insert(&tt.0, n);
            out_string.push_str("   [");
            out_string.push_str(&n.to_string());
            out_string.push_str("] = ");
            out_string.push_str(tt.0);
            out_string.push_str(",\n");
            n = n+1;
        }
        out_string.push_str("}\n");
    }
    out_string.push_str("local table = {");
 
    init_head_flag = false;
    for row in range.rows(){
        if is_skip_row(&row){
            continue;
        };
        if !init_head_flag{
            init_head_flag = true;
            continue;
        }
        if init_head_flag{
            let mut init_key = false;
            for head in &heads{
                let mut cell = row.get(head.0 as usize).unwrap().to_string();
                if !init_key{
                    out_string.push_str("\n   [");
                    out_string.push_str(&cell);
                    out_string.push_str("] = {");
                    init_key = true
                }
                out_string.push_str(&head.1);
                out_string.push_str(" = ");
                if is_table(&cell){
                    if t_map.contains_key(&cell){
                        if let Some(n) = t_map.get(&cell){
                            out_string.push_str("t[");
                            out_string.push_str(&n.to_string());
                            out_string.push_str("], ");
                        }
                    }else{
                        if let Some(node) = node_cache.get(&cell){
                            cell = update_table_str(node, &cell, &t_map);
                        }
                        out_string.push_str(&cell);
                        out_string.push_str(", ");
                    }
                }else{
                    out_string.push_str(&cell);
                    out_string.push_str(", ");
                }
            }
            out_string.push_str("}, ");
        }
    }
    out_string.push_str("\n}};\nreturn table;");

    println!("{}", out_string);
    // fs::write("output.lua", out_string);

    // let ff = File::open("output.lua".to_string()).unwrap();
    // let file = io::buffered_read(ff).unwrap();

    let mut cursor = Cursor::new(&out_string);
    // let mut buffer = vec![];
    // cursor.read_exact(&mut buffer)?;

    let mut interner = BasicInterner::default();
    //let chunk = 
    match compiler::parse_chunk(cursor, &mut interner){
        Err(e) =>{
            println!("{:#?}", e);
            //fs::write("output.lua", out_string);
        }
        _ =>{
            println!("success");
        }
    }
    
    //println!("{}", out_string);
}

fn main() {
    let matches = Command::new(crate_name!())
    .arg(
        Arg::new("files")
            .short('f')
            .help("Parse files"),
    )
    .arg(
        Arg::new("path")
            .short('p')
            .help("Parse files in path"),
    )
    .arg(
        Arg::new("o")
            .required(true)
            .help("out path"),
    )
    .get_matches();

    let out_path = matches.get_one::<String>("o").unwrap().to_string();
    let mut files:Vec<String> = Vec::new();
    if matches.contains_id("files") {
        let file_s = matches.get_one::<String>("files").unwrap();
        files = file_s.split(' ').map(|t| t.to_string()).collect();
    } else {
        let path = matches.get_one::<String>("path").unwrap();
        let dir_path = Path::new(path);
        if let Ok(entries) = fs::read_dir(dir_path){
            for entry in entries{
                if let Ok(entry) = entry{
                    let file_path = entry.path();
                    if let Some(extension) = file_path.extension(){
                        if extension == "xlsx" || extension == "xls"{
                            files.push(file_path.to_str().unwrap().to_string())
                        }
                    }
                }
            }
        }
    }
    let mut file_sizes = Vec::new();
    for path in files {
        if let Ok(metadata) = fs::metadata(&path){
            let size = metadata.len();
            file_sizes.push((path, size));
        }
    }
    file_sizes.sort_by(|a, b| b.1.cmp(&a.1));

    let mut thread_num = 10;
    if file_sizes.len() < thread_num{
        thread_num = file_sizes.len()
    }
    let mut equalize:Vec<(Vec<String>, u64)> = Vec::new();
    for i in 1..=thread_num {
        let mut temp:Vec<String> = Vec::new();
        temp.push(file_sizes[i-1].0.to_string());
        equalize.push((temp, file_sizes[i-1].1))
    }
    for i in thread_num..=file_sizes.len(){
        let mut j = 1;
        let mut temp_min:u64 = std::u64::MAX;
        for k in 1..=thread_num{
            if equalize[k-1].1 < temp_min{
                temp_min = equalize[k-1].1;
                j = k
            }
        }
        equalize[j-1].1 = equalize[j-1].1+temp_min;
        equalize[j-1].0.push(file_sizes[i-1].0.to_string())
    }

    let mut handles = vec![];
    for s in equalize {
        let out_path_ = out_path.to_string();
        let handle = thread::spawn(|| {
            create_files(s.0, out_path_);
        });
        handles.push(handle);
    }

    for handle in handles {
        handle.join().unwrap();
    }
}
