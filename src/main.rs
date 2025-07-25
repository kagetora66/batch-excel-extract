use csv::Writer;
extern crate umya_spreadsheet;
use std::fs::{self, File};
use std::io::Write;
use std::io;
use std::path::{Path, PathBuf};
use regex::Regex;
use umya_spreadsheet::Worksheet;
use walkdir::WalkDir;
use anyhow::{Context, Result};
use std::time::Duration;
use std::thread;

struct coordinates {
    row: u32,
    column: u32,
}

fn select_folder() -> Option<PathBuf> {
    rfd::FileDialog::new()
    .set_title("Select a folder containing XLSX files")
    .pick_folder()
}

fn find_xlsx_files(folder: &Path) -> Result<Vec<PathBuf>> {
    let mut xlsx_files = Vec::new();

    for entry in WalkDir::new(folder) {
        let entry = entry?;
        let path = entry.path();

        if path.is_file() {
            if let Some(ext) = path.extension() {
                if ext == "xlsx" {
                    xlsx_files.push(path.to_path_buf());
                }
            }
        }
    }

    Ok(xlsx_files)
}

//checks if our row is in the same range as merged cells
fn check_range(merged: &String, selected: &str) -> bool {
    let re = Regex::new(r"^[A-Za-z](\d{1,3}):[A-Za-z](\d{1,3})$").unwrap();
    let caps = match re.captures(merged) {
        Some(c) => c,
        None => {
            return false
        }
    };

    let num1 = caps.get(1)
    .unwrap_or_else(|| panic!("No group 1 in string '{}'", merged))
    .as_str()
    .parse::<u32>()
    .unwrap_or_else(|_| panic!("Failed to parse group 1 in string '{}'", merged));

    let num2 = caps.get(2)
    .unwrap_or_else(|| panic!("No group 2 in string '{}'", merged))
    .as_str()
    .parse::<u32>()
    .unwrap_or_else(|_| panic!("Failed to parse group 2 in string '{}'", merged));

    let selected_row = selected.parse().unwrap();
    if num1 < selected_row && selected_row < num2 {
        return true
    }
    else {
        return false
    }
}
//creates a vector of everything in the row
 fn get_row(row: u32, sheet: &Worksheet, filter: &str) -> Vec<String> {    
    let mut row_values = Vec::new();
    let merged = sheet.get_merge_cells();
    let cell_row = row.to_string();
    let mut is_filtered = false;
    //if our filter word is merged
    for range in merged{
       let mut range_value = range.get_range();
    if check_range(&range_value, &cell_row) == true {
        let mut merge_coord = sheet.map_merged_cell(&*range_value);
        let mut value = sheet.get_value(merge_coord);
        if value.to_string() == filter {
        row_values.push(value.to_string());
        is_filtered = true;
        }
    }
   }
    let cell = sheet.get_collection_by_row(&row);
    for item in cell {
        if is_filtered == true {
        let value = item.get_cell_value().get_value();
        row_values.push(value.to_string());
        }
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
fn prompt_input(prompt: &str) -> io::Result<String> {
    let mut input = String::new();
    print!("{}", prompt);
    io::stdout().flush()?; // Ensure prompt appears immediately
    io::stdin().read_line(&mut input)?;
    Ok(input.trim().to_string())
}

fn main() {
    println!("Please select a folder containing the excel files");
    let folder = select_folder().ok_or(anyhow::anyhow!("No folder selected")).unwrap();
    let xlsx_files = find_xlsx_files(&folder).unwrap();

    println!("Found xlsx files");
    // Get the query
    let keyword = prompt_input("Enter your search query: ").expect("Failed to read query");

    // Get optional filter
    let filter = match prompt_input("Filter rows/columns by keyword? (press Enter if no): ") {
        Ok(s) if !s.trim().is_empty() => s.trim().to_string(),
        _ => String::new(), // Empty string if no filter
    };

    let mut wtr = Writer::from_path("output.csv").unwrap();
    for file in xlsx_files {
        let book = umya_spreadsheet::reader::xlsx::read(&file).unwrap();
        let sheet = book.get_sheet_by_name("SMART Data").unwrap();
        let coords = get_keyword_coord(&keyword, &sheet);
        let filename = &file.file_name().unwrap().to_str().unwrap();
        let mut empty_row = Vec::new();
        for cord in coords {
            let mut row = get_row(cord.row, &sheet, &filter);
            if row.len() != 0 {
                row.insert(0, filename.to_string()); // Add filename as first column
                empty_row = vec!["".to_string(); row.len()];
            }
            wtr.write_record(&row);
            thread::sleep(Duration::from_millis(10));
        }
        wtr.write_record(&empty_row);

        wtr.flush();


    }

}
