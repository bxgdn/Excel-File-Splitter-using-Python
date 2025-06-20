# Excel File Splitter

A Python utility to split large Excel files into smaller, more manageable chunks without losing any data. This tool helps reduce file sizes and makes large datasets easier to work with.

## Features

- 🔄 Split Excel files into multiple parts (default: 3 parts)
- 📊 Preserve all data integrity during the split process
- 🎯 Customizable number of output files
- 📈 Progress tracking with row count information
- ⚡ Fast processing using pandas and openpyxl

## Requirements

- Python 3.6+
- pandas 2.2.2
- openpyxl 3.1.2

## Installation

1. Clone this repository or download the script
2. Install the required dependencies:

```bash
pip install pandas==2.2.2 openpyxl==3.1.2
```

## Usage

1. Update the input file path in the script to point to your Excel file
2. Run the script:

```bash
python excel_splitter.py
```

3. The script will create 3 output files by default:
   - `data_part_1.xlsx`
   - `data_part_2.xlsx`
   - `data_part_3.xlsx`

## Customization

### Change Number of Parts

To split into a different number of parts, modify this line in the script:
```python
rows_per_part = total_rows // 3  # Change 3 to your desired number
```

### Change Input File Path

Update this line with your file path:
```python
df = pd.read_excel("your_file_path_here.xlsx", engine="openpyxl")
```

### Change Output File Names

Modify the output file names in the `to_excel()` calls:
```python
df_part1.to_excel("custom_name_part_1.xlsx", index=False)
```

## How It Works

1. **Load**: Loads the entire Excel file into a pandas DataFrame
2. **Calculate**: Determines the optimal number of rows per part based on total rows
3. **Split**: Divides the DataFrame into equal chunks using pandas slicing
4. **Save**: Exports each chunk as a separate Excel file
5. **Report**: Displays the number of rows in each resulting file

## Example Output

```
Split into 3 parts complete:
  Part 1: 21812 rows
  Part 2: 21812 rows
  Part 3: 21813 rows
```

## Notes

- The script automatically handles remainder rows by adding them to the last part
- All original data formatting and structure is preserved
- No data is lost during the splitting process
- Uses openpyxl engine for full Excel compatibility

## Troubleshooting

### Common Issues

1. **Memory Error**: If working with very large files, consider splitting into more parts
2. **File Not Found**: Ensure the input file path is correct and accessible
3. **Permission Error**: Make sure you have write permissions in the output directory

### Performance Tips

- Close other applications when processing very large files
- Ensure sufficient disk space for output files
- Consider using SSD storage for faster I/O operations

## Contributing

Feel free to fork this project and submit pull requests for improvements or additional features.

## License

This project is open source and available under the MIT License.
