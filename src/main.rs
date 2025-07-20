use csv::Writer;
extern crate umya_spreadsheet;
use std::fs::{self, File};
use std::io::Write;
use std::io;
use std::path::{Path, PathBuf};
use regex::Regex;
use umya_spreadsheet::Worksheet;

struct coordinates {
    row: u32,
    column: u32,
}

//checks if our row is in the same range as merged cells
fn check_range(merged: &String, selected: &str) -> bool {
    let re = Regex::new(r"^[A-Za-z](\d+):[A-Za-z](\d+)$").unwrap();
    let caps = re.captures(merged);

    let num2 = caps.as_ref().expect("error").get(2).expect("no group 2").as_str().parse::<u32>().ok().unwrap();
    let num1 = caps.as_ref().expect("error").get(1).expect("no group 1").as_str().parse::<u32>().ok().unwrap();
    let selected_row = selected.parse().unwrap();
    if num1 < selected_row && selected_row < num2 {
        return true
    }
    else {
        return false
    }
}
//creates a vector of everything in the row
 fn get_row(row: u32, sheet: &Worksheet) -> Vec<String> {    
    let mut row_values = Vec::new();
    let merged = sheet.get_merge_cells();
    let cell_row = row.to_string();
    
    for range in merged{
        let co
    let mut range_value = range.get_range();
    if check_range(&range_value, &cell_row) == true {
        let mut merge_coord = sheet.map_merged_cell(&*range_value);
        let mut value = sheet.get_value(merge_coord);
        row_values.push(value.to_string());
    }
   }
    let cell = sheet.get_collection_by_row(&row);
    for item in cell {
        let value = item.get_cell_value().get_value();
        row_values.push(value.to_string());
    }
    row_values
}

fn get_keyword_coord(query: &str, sheet: &Worksheet) -> Vec<coordinates>
{
    let mut coords = Vec::new();
    let cells = sheet.get_cell_collection();
    for item in cells {
        let mut value = item.get_cell_value().get_value();
        if query == value{
            coords.push(coordinates {
                row: *item.get_coordinate().get_row_num(),
                column: *item.get_coordinate().get_col_num(),
            });
        }
    }
    coords
}

fn main() {
    let path = std::path::Path::new("./smarts.xlsx");
    let mut book = umya_spreadsheet::reader::xlsx::read(path).unwrap();
    let sheet  = book.get_sheet_by_name("SMART Data").unwrap(); 
    let keyword = "Wear_Leveling_Count";
    let coords = get_keyword_coord(&keyword, &sheet);
    let mut wtr = Writer::from_path("output.csv").unwrap();
    for cord in coords {
        let row = get_row(cord.row, &sheet);
        wtr.write_record(&row);
        println!("Row is:");
        for cell in row{
            println!("cell is: {}", cell);
        }

}
wtr.flush();
}
