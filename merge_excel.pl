#!/usr/bin/perl
use strict;
use warnings;

# Required libraries
use Getopt::Long;
use merge_excel_pkg;  # local pkg

## --- Handle Command Line Arguments--- ##
our %options = ();
GetOptions(
           "help"                 => \$options{help},               ## Describe usage
           "source_directory=s"   => \$options{source_dir},         ## Top parent directory containing all the excels to be merged
           "header_map_file=s"    => \$options{header_map_file},    ## Excel file with header mapping and header column order
           "output_excel_name=s"  => \$options{output_excel_name},  ## Output merged excel filename  [Optional]
           "debug"                => \$options{debug},              ## Debug the script. Lower   verbosity [Optional]
           "full_debug"           => \$options{full_debug},         ## Debug the script. Highest verbosity [Optional]
           );

## --- Main --- ##
usage() if ($options{help} || !defined $options{source_dir} || !defined $options{header_map_file});

# variable mapping
$src_dir              = $options{source_dir};
$header_map_xl_file   = $options{header_map_file};
$outname              = (defined $options{output_excel_name}) ? "$options{output_excel_name}.xlsx" : "output_merged_excel.xlsx";
$debug                = $options{debug};
$full_debug           = $options{full_debug};

# Main function call
die "Source directory \"$src_dir\" does not exists\n" if (! -d $src_dir);
&main();
exit(0);


#################################################################################
#########    LOCAL FUNCTIONS        #############################################
#################################################################################

## ------------------------------------------------------------------------------
## usage()
##
## Explains the functionality of the script
## ------------------------------------------------------------------------------
sub usage {

print <<USAGE;

Usage: merge_excel.pl <-source_directory excel_source_directory> <-header_map_file header_mapping_excel> [-output_excel_name output_filename] [-debug] [-full_debug] [-help]

Merge all excels (.xlsx, .xls, .csv) in a directory including all sub-directories

---------
Use Case:
---------
Suited for merging excels that are similar in content but compiled by or from different sources
There might be slight differences in the column headers or order. Like "First Name" might be called "first_name" in one excel, "first" in another
It might be 1st column in one excel but 3rd in another. These differences can be mapped with header_map_file
If all excels have the same header, it still works. Just create a mapping file with same mapping (same values in both columns)
TODO: Next version might auto-generate this mapping file for this usecase.

You may provide the output file name. Else it will be "output_merged_excel.xlsx".

Options:
  -help                   Print this message [Optional]
  -source_dir=s           Top parent directory containing all the excels to be merged
  -header_map_file=s      Excel file with header mapping and header column order
  -output_excel_name=s    Output merged excel filename  [Optional]. Default: autogen_merge_excel.xlsx
  -debug                  Debug the script. Lower   verbosity [Optional]
  -full_debug             Debug the script. Highest verbosity [Optional]

Example 1: Merge all excels under directory "all_excels_dir" with default output excel name
---------
./merge_excel.pl -source_directory all_excels_dir -header_map_file input_header_mapping.xlsx

Example 2: Merge all excels under directory "all_excels_dir" and merge into excel file "custom_merged_excel.xlsx"
---------
./merge_excel.pl -source_directory all_excels_dir -output_excel_name custom_merged_excel -header_map_file input_header_mapping.xlsx

USAGE

exit(1);
}

## ------------------------------------------------------------------------------
## main()
##
## Main function that parses all excels in the source directory and merges them
##   mapping file is used to merge entries appropriately and order them as different files
##   might have different header heading and different order
##   Like "First Name" might be called "first_name" in one excel, "first" in another
##   It might be 1st column in one excel but 3rd in another
## ------------------------------------------------------------------------------
sub main {

  # build excel header mapping
  my $header_map_xl = &readExcelFile($header_map_xl_file);
  &print_xl($header_map_xl->[0]->{data}) if ($full_debug);
  &build_header_row_mapping($header_map_xl);

  # Start from parent directory, read each excel, and then merge them
  &processItem($src_dir);

  # write out the final excel output
  &write_out_all_rows($outname);
  &printStat();

}

## ------------------------------------------------------------------------------
## printStat()
##
## Print statistics about the run
## ------------------------------------------------------------------------------
sub printStat {

  &print_divider("=", 40);
  printf ("Final Stats:\n");
  &print_divider("=", 40);
  printf ("Number of Directories     = %5d\n", $dir_count);
  printf ("Number of Files           = %5d\n", $file_count);
  printf ("Number of Excel Files     = %5d\n", $xl_count);
  printf ("Number of Non-Excel Files = %5d\n", $non_xl_count);
  &print_divider("=", 40);
  printf ("Number of Merged Rows     = %5d\n", scalar @all_merged_rows_arr);
  &print_divider("=", 40);

}
