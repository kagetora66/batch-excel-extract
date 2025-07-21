Batch Excel Extractor

A Rust tool for bulk searching and extracting data from multiple Excel files
Overview

Batch Excel Extractor processes a folder of Excel files (.xlsx), searches for rows containing a specified keyword, and exports the results to a CSV file. Optional filtering allows for more precise data extraction.
Key Features

✔ Bulk processing - Scan multiple Excel files in one operation
✔ Keyword search - Extract rows containing your search term
✔ Optional filtering - Narrow results with secondary filters
✔ Merged cell support - Properly handles merged rows/columns
✔ CSV export - Clean output in standard comma-separated format
Current Status

✅ Implemented:

    Row extraction with keyword matching

    Merged row support

    Optional filter conditions

🛠 In Development:

    Column extraction functionality

    Advanced filtering options

Usage

    Run the executable

    Enter your search keyword

    Specify the folder containing Excel files

    (Optional) Add filter criteria when prompted

    Results will be saved in output.csv
