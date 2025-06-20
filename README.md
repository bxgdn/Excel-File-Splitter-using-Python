# Excel File Splitter

A Python utility to split large Excel files into smaller, more manageable parts. Perfect for handling large datasets that exceed memory limits or need to be processed in smaller chunks.

## 🚀 Features

- **Smart Splitting**: Automatically calculates optimal row distribution across parts
- **Remainder Handling**: Intelligently distributes extra rows to the last part
- **Data Integrity**: Verifies all rows are preserved during the split process
- **Progress Tracking**: Real-time feedback during processing
- **Error Handling**: Comprehensive error messages and validation
- **Flexible Configuration**: Easily customize number of parts and output naming

## 📋 Requirements

- Python 3.6+
- pandas
- openpyxl

## 🔧 Installation

1. Clone this repository:
```bash
git clone https://github.com/bxgdn/excel-file-splitter.git
cd excel-file-splitter
```

2. Install required dependencies:
```bash
pip install pandas openpyxl
```

## 📖 Usage

### Basic Usage

```python
# Simple split into 3 parts
python excel_splitter.py
```

### Advanced Usage

```python
from excel_splitter import split_excel_file

# Split into 5 parts with custom naming
parts_info = split_excel_file(
    input_file="large_dataset.xlsx",
    num_parts=5,
    output_prefix="dataset_chunk"
)
```

### Configuration

Modify these variables in the script for your needs:

```python
INPUT_FILE = "yourfile.xlsx"    # Path to your Excel file
NUM_PARTS = 3                   # Number of parts to create
```

## 📁 Input/Output

### Input
- **File Format**: Excel files (.xlsx, .xls)
- **Size**: Any size (limited by available memory)
- **Structure**: Any Excel structure with data in rows

### Output
- **Files**: Multiple Excel files with sequential naming
- **Naming Pattern**: `{prefix}_{number}.xlsx` (e.g., `data_part_01.xlsx`)
- **Format**: Same structure as input file
- **Index**: No index column added (clean data preservation)

## 📊 Example

**Input File**: `sales_data.xlsx` (65,437 rows)

**Command**:
```python
split_excel_file("sales_data.xlsx", num_parts=3, output_prefix="sales_chunk")
```

**Output**:
```
sales_chunk_01.xlsx - 21,812 rows
sales_chunk_02.xlsx - 21,812 rows  
sales_chunk_03.xlsx - 21,813 rows (includes remainder)
```

**Console Output**:
```
Loading Excel file: sales_data.xlsx
Total rows in file: 65,437
Rows per part (base): 21,812
Remainder rows (added to last part): 1
Split indices: [21812, 43624]
Saving sales_chunk_01.xlsx with 21,812 rows...
Saving sales_chunk_02.xlsx with 21,812 rows...
Saving sales_chunk_03.xlsx with 21,813 rows...

==================================================
EXCEL FILE SPLIT COMPLETE
==================================================
  Part 1: sales_chunk_01.xlsx - 21,812 rows
  Part 2: sales_chunk_02.xlsx - 21,812 rows
  Part 3: sales_chunk_03.xlsx - 21,813 rows
--------------------------------------------------
Total rows processed: 65,437
Files created: 3
✓ Data integrity verified: All rows preserved
```

## ⚙️ How It Works

1. **Load**: Reads the entire Excel file using pandas with openpyxl engine
2. **Calculate**: Determines optimal row distribution using integer division
3. **Split**: Creates data slices using pandas iloc for efficient processing
4. **Save**: Exports each part as a separate Excel file
5. **Verify**: Confirms all rows are preserved in the output files

### Split Logic

```
Total Rows: 65,437
Parts: 3
Base Rows per Part: 65,437 ÷ 3 = 21,812
Remainder: 65,437 % 3 = 1

Distribution:
- Part 1: Rows 0 to 21,811 (21,812 rows)
- Part 2: Rows 21,812 to 43,623 (21,812 rows)  
- Part 3: Rows 43,624 to 65,436 (21,813 rows) ← Gets the remainder
```

## 🛠️ Customization

### Function Parameters

```python
def split_excel_file(input_file, num_parts=3, output_prefix="data_part"):
    """
    Args:
        input_file (str): Path to input Excel file
        num_parts (int): Number of parts to create
        output_prefix (str): Prefix for output filenames
    
    Returns:
        list: [(filename, row_count), ...] for each part
    """
```

### Error Handling

The script handles common issues:
- **File Not Found**: Clear message if input file doesn't exist
- **Invalid Format**: Guidance for file format issues
- **Memory Issues**: Suggestions for very large files
- **Permission Errors**: Helpful hints for file access problems

## 🐛 Troubleshooting

### Common Issues

**"FileNotFoundError"**
- Check that your input file path is correct
- Ensure the file exists in the specified location

**"Memory Error"**
- File too large for available RAM
- Consider processing on a machine with more memory
- Try splitting into more parts to reduce memory usage per operation

**"Permission Denied"**
- Check file permissions
- Ensure the file isn't open in Excel
- Verify write permissions in the output directory

### Performance Tips

- **Large Files**: Increase `num_parts` for better memory management
- **SSD Storage**: Use SSD for faster I/O operations
- **Memory**: Close other applications to free up RAM

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/improvement`)
3. Commit your changes (`git commit -am 'Add new feature'`)
4. Push to the branch (`git push origin feature/improvement`)
5. Create a Pull Request

## 📧 Support

If you encounter any issues or have questions:
- Open an issue on GitHub
- Check the troubleshooting section above
- Review the example usage

## 🏷️ Version History

- **v1.0.0** (2025-06-20): Initial release
  - Basic Excel file splitting functionality
  - Data integrity verification
  - Comprehensive error handling

---

**Created by**: [@bxgdn](https://github.com/bxgdn)  
**Last Updated**: June 20, 2025
