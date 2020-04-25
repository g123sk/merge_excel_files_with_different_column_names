# merge_excels_with_different_column_names

## Overview
Perl utility to merge all excel files (.xlsx, .xls .csv ignores other file types) in a directory including sub-directories. If the excel files have different headers (column names), they can be mapped accordingly. Works fine with different header column order as well. 

Supports excel formats like .xls, .xlsx and .csv and multiple worksheets in the same excel.

## Typical Use Case
Suited for merging excels that are similar in content but compiled by or from different sources in different excel formats (.xls, .xlsx, .csv). There might be slight differences in the column headers or order. Like `First Name` might be called `first_name` in one excel, `first` in another. It might be 1st column in one excel but 3rd in another. These differences can be mapped with header_map_file. If all excels have the same header, it still works. Just create a mapping file with same mapping (same values in both columns). 

### Upcoming Features:
  * ~~Support for reading all sheets in an excel and not just the first one~~ Done
  * Auto-generate the header mapping file if all excels have the same header.
  * Support to merge duplicate entries (if different excels have the same entry (same email id or same SSN), merge them)

## System Requirements:
Need Perl installation with excel packages mentioned in the perl module merge_excel_pkg.pm

## Usage
```
./merge_excel.pl <-source_dir excel_source_directory> <-header_map_file header_mapping_excel> [-output_excel_name output_filename] [-debug] [-full_debug] [-help]
```
## Command Line Options
```
  -help                   Print this message [Optional]
  -source_dir=s           Top parent directory containing all the excels to be merged
  -header_map_file=s       Excel file with header mapping and header column order
  -output_excel_name=s    Output merged excel filename  [Optional]. Default: output_merged_excel.xlsx
  -debug                  Debug the script. Lower   verbosity [Optional]
  -full_debug             Debug the script. Highest verbosity [Optional]
```

## Example 1
Merge all excels under directory "all_excels_dir" with default output excel name
```
./merge_excel.pl -source_directory all_excels_dir -header_map_file input_header_mapping.xlsx
```

## Example 2
Merge all excels under directory "all_excels_dir" and merge into excel file "custom_merged_excel.xlsx"
```
./merge_excel.pl -source_directory all_excels_dir -output_excel_name custom_merged_excel -header_map_file input_header_mapping.xlsx
```

## File Structure
```
merge_excel.pl              Main perl code
merge_excel_pkg.pm          Package file with all local functions needed
input_header_mapping.xlsx    Sample mapping file
all_excels_dir/             Sample input directory with 3 excel files
                                Each excel has a list of user properties which will be merged
output_merged_excel.xlsx    Output merged excel file for the sample inputs
```

## Demo Video
[Demo](https://www.youtube.com/watch?v=jY3ZrWaHpfs)
