import pandas as pd

def split_excel_file(input_file, num_parts=3, output_prefix="data_part"):
    """
    Split a large Excel file into smaller, more manageable parts.
    
    Args:
        input_file (str): Path to the input Excel file
        num_parts (int): Number of parts to split the file into (default: 3)
        output_prefix (str): Prefix for output filenames (default: "data_part")
    
    Returns:
        list: List of tuples containing (filename, row_count) for each part
    """
    
    # Load the full Excel file using openpyxl engine for .xlsx files
    # openpyxl is recommended for reading .xlsx files as it handles formatting better
    print(f"Loading Excel file: {input_file}")
    df = pd.read_excel(input_file, engine="openpyxl")
    
    # Calculate total rows and determine optimal split strategy
    total_rows = len(df)
    print(f"Total rows in file: {total_rows:,}")
    
    # Calculate base number of rows per part using integer division
    # This ensures we get the maximum equal distribution
    rows_per_part = total_rows // num_parts
    
    # Calculate remainder rows that will be distributed to the last part
    # This handles cases where total_rows is not evenly divisible by num_parts
    remainder = total_rows % num_parts
    
    print(f"Rows per part (base): {rows_per_part:,}")
    print(f"Remainder rows (added to last part): {remainder}")
    
    # Define split indices for creating data slices
    # These indices mark where each part begins and ends
    split_indices = []
    for i in range(1, num_parts):
        split_indices.append(rows_per_part * i)
    
    print(f"Split indices: {split_indices}")
    
    # Create data parts using pandas iloc for efficient row slicing
    parts_info = []
    
    for i in range(num_parts):
        # Determine start and end indices for current part
        if i == 0:
            # First part: from beginning to first split
            start_idx = 0
            end_idx = split_indices[0] if split_indices else total_rows
        elif i == num_parts - 1:
            # Last part: from last split to end (includes remainder rows)
            start_idx = split_indices[i - 1]
            end_idx = total_rows
        else:
            # Middle parts: between consecutive splits
            start_idx = split_indices[i - 1]
            end_idx = split_indices[i]
        
        # Extract the data slice for current part
        df_part = df.iloc[start_idx:end_idx]
        
        # Generate output filename with zero-padded numbering for proper sorting
        output_filename = f"{output_prefix}_{i + 1:02d}.xlsx"
        
        # Save to Excel file without index column to maintain clean data structure
        # index=False prevents pandas from adding an extra index column
        print(f"Saving {output_filename} with {len(df_part):,} rows...")
        df_part.to_excel(output_filename, index=False, engine="openpyxl")
        
        # Store information about this part for summary reporting
        parts_info.append((output_filename, len(df_part)))
    
    return parts_info

# Main execution block
if __name__ == "__main__":
    # Configuration - modify these variables as needed
    INPUT_FILE = "yourfile.xlsx"  # Replace with your actual Excel file path
    NUM_PARTS = 3                 # Number of parts to split into
    
    try:
        # Execute the split operation
        parts_info = split_excel_file(INPUT_FILE, NUM_PARTS)
        
        # Display comprehensive summary of the split operation
        print("\n" + "="*50)
        print("EXCEL FILE SPLIT COMPLETE")
        print("="*50)
        
        total_output_rows = 0
        for i, (filename, row_count) in enumerate(parts_info, 1):
            print(f"  Part {i}: {filename} - {row_count:,} rows")
            total_output_rows += row_count
        
        print("-"*50)
        print(f"Total rows processed: {total_output_rows:,}")
        print(f"Files created: {len(parts_info)}")
        
        # Verify data integrity
        original_df = pd.read_excel(INPUT_FILE, engine="openpyxl")
        if len(original_df) == total_output_rows:
            print("✓ Data integrity verified: All rows preserved")
        else:
            print("⚠ Warning: Row count mismatch detected")
            
    except FileNotFoundError:
        print(f"Error: Could not find input file '{INPUT_FILE}'")
        print("Please check the file path and ensure the file exists.")
    except Exception as e:
        print(f"Error occurred during processing: {str(e)}")
        print("Please check your input file format and try again.")
