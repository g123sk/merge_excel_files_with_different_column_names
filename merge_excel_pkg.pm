#!/usr/bin/perl
use strict;
use warnings;

package merge_excel_pkg;

# Required libraries
use Data::Dumper;
use Excel::Writer::XLSX;
use Spreadsheet::ParseExcel;
use Spreadsheet::XLSX;
use Spreadsheet::Read;

require Exporter;
our @ISA = qw(Exporter);
our @EXPORT = qw(create_new_output_row add_output_row_to_list
                 $debug $full_debug
                 $src_dir $outname $header_map_xl_file
                 $dir_count $file_count $xl_count $non_xl_count
                 %all_header_mappings_hash @header_row_names @all_merged_rows_arr
                 create_mapped_header print_xl build_header_row_mapping
                 handle_xl_content print_divider write_out_all_rows
                 read_xlsx read_xls read_csv
                 processItem handleDir handleFile
                 massage_cell_value);

# global vars
our $debug;                       ## Debug the script. Lower   verbosity [Optional]
our $full_debug;                  ## Debug the script. Highest verbosity [Optional]
our $src_dir;                     ## Top parent directory containing all the excels to be merged
our $outname;                     ## Output merged excel filename [Optional]
our $header_map_xl_file;          ## Excel file with header mapping and header column order
our $dir_count      = 0;          ## Number of directories  processed
our $file_count     = 0;          ## Number of native files processed
our $xl_count       = 0;          ## Number of directories processed
our $non_xl_count   = 0;          ## Number of directories processed
our @header_row_names;            ## Array with the final header contents
our %all_header_mappings_hash;    ## Hash holding mapping for header array from each excel to final header values
our @all_merged_rows_arr;         ## Array to hold all rows in the final merged excel

##################################################
##########    Local Functions    #################
##################################################

# Function to create a new output excel row entry with contents from current row in current excel
# New entry will use headers from mapping file
sub create_new_output_row_with_info {
  my ($curr_row, $header_arr, $my_fname) = @_;
  my %new_output_row = ();

  my $arr_length = scalar @$header_arr;

  # populate the new output row with current excel row data
  for (my $i=0; $i<$arr_length; $i++) {
    $new_output_row{$header_arr->[$i]} = $curr_row->[$i];
  }

  $new_output_row{Source_Excel_FileName} = $my_fname; # Save the source file info in output excel for downstream processing

  return(\%new_output_row);
}

# Add the new output row to final merged excel
sub add_output_row_to_list {
  my ($output_row) = @_;
  push @all_merged_rows_arr, $output_row;
}

# Function to create header mapping for current excel
# Takes in the unmapped header (first row in current excel) and
# generates the mapped header array to be used in new output row
sub create_mapped_header {
  my ($unmapped_header_arr, $my_fname) = @_;

  my @mapped_headers = ();

  foreach my $header ( @{$unmapped_header_arr} ) {

    die "Empty header value not allowed \"$header\"\n" if ($header =~ /^\s*$/);
    if (! exists $all_header_mappings_hash{$header}) {
      #print Dumper \%all_header_mappings_hash;
      die "Unknown header value \"$header\" in excel file \"$my_fname\"\n";
    }
    push @mapped_headers, $all_header_mappings_hash{$header};
  }

  return (\@mapped_headers);
}

# Function to print excel for debug purpose
sub print_xl {
  my ($xl_ref) = @_;

  &print_divider("=", 15);
  print "Printing Excel:\n";
  &print_divider("=", 15);

  foreach my $row (@$xl_ref) {
    foreach my $col (@$row) {
      print "\"$col\" ";
    }
    print "\n";
  }
}

# Function to print dividing line for formatting output
sub print_divider {
  my ($char, $rep) = @_;
  printf ("%s\n", $char x $rep);
}

# creates mapping for all column headers that are used across different excels
# Like "First Name" might be called "first_name" in one excel, "first" in another
# It might be 1st column in one excel but 3rd in another. These differences can be mapped with header_map_file.
# so we need this mapping. Final output excel will use values in the column one of mapping file and in that order
sub build_header_row_mapping {
  my ($header_mapping_xl_ref) = @_;

  shift @$header_mapping_xl_ref; # Throw away the header row

  foreach my $row (@$header_mapping_xl_ref) {
    my $header = shift @$row;
    push @header_row_names, $header; # add new header name

    foreach my $col (@$row) {
      $all_header_mappings_hash{$col} = $header;
    }
  }

  push @header_row_names, "Source_Excel_FileName"; # Final Column will be for saving the original excel file
}

# handle contents of a single excel
# go through each row, create the merged row entry and add to final excel output row list
sub handle_xl_content {
  my ($xl_ref, $my_fname) = @_;

  # handle header (first row)
  # create a header mapping between current excel and output excel
  my $header_row_ref = shift @$xl_ref;
  my $mapped_headers = &create_mapped_header($header_row_ref, $my_fname);

  if (scalar @$mapped_headers == 0) {
    &print_divider("=", 40);
    print "Empty Excel File \"$my_fname\"?\n";
    &print_divider("=", 40);
    return 0;
  }

  # handle the rows in this excel
  foreach my $row (@$xl_ref) {
    my $new_row_ref = &create_new_output_row_with_info($row, $mapped_headers, $my_fname);
    &add_output_row_to_list($new_row_ref);
  } # each row

}

# write final all row into excel output with formatting
sub write_out_all_rows {
  my ($my_outname) = @_;
  my $Excel_book1  = Excel::Writer::XLSX->new($my_outname);
  my $Excel_sheet1 = $Excel_book1->add_worksheet("Merged_rows");

  # add the header row
  my $header_format = $Excel_book1->add_format();
  $header_format->set_bold();
  $header_format->set_size(14);
  $header_format->set_color('blue');

  $Excel_sheet1->write(0, 0, \@header_row_names, $header_format);

  my $merged_row_cnt = scalar @all_merged_rows_arr;
  for (my $i=0; $i < $merged_row_cnt; $i++) {
    my $curr_output_row_info_arr = &rebuild_merged_row_info($all_merged_rows_arr[$i]);

    # write out each row
    $Excel_sheet1->write( $i+1, 0, $curr_output_row_info_arr);
  }

}

# Function to create one output excel entry
# Takes in hash for one row and builds the array in expected header order
sub rebuild_merged_row_info {
  my ($curr_merged_row_info) = @_;
  my @rebuilt_merged_row = ();

  foreach my $header_name (@header_row_names) {
    push @rebuilt_merged_row, $curr_merged_row_info->{$header_name};
  }

  return \@rebuilt_merged_row;
}

# Add any text manipulation you need
# called after reading each cell in the excel
sub massage_cell_value {
  $_[0] =~ s/&amp;/&/g;         # excel cell reads of '&' comes as '&amp;' so change it back to '&'
  $_[0] =~ s/[^[:ascii:]]+//g;  # get rid of non-ASCII characters
}

# Read Excel with xls extension
sub read_xls {
  #my (@args) = @_;
  my ($filename) = @_;

  my @curr_xl = ();

  die "read_xls() can only read xls file extension. But called for file \"$filename\"" if ($filename !~ /\.xls$/);
  die "Cannot find file named \"$filename\" to read in read_xls()\n" if (! -e $filename);

  print "read_xls() on file \"$filename\"\n" if ($debug);

  my $parser   = Spreadsheet::ParseExcel->new();
  my $workbook = $parser->parse($filename);

  if ( !defined $workbook ) {
           die $parser->error(), ".\n";
  }

  for my $worksheet ( $workbook->worksheets() ) {

    my ( $row_min, $row_max ) = $worksheet->row_range();
    my ( $col_min, $col_max ) = $worksheet->col_range();

    for my $row ( $row_min .. $row_max ) {
      my @curr_row = ();

      for my $col ( $col_min .. $col_max ) {

		     my $cell = $worksheet->get_cell( $row, $col );
         my $val  = $cell ? $cell->value() : "";
         &massage_cell_value($val);

         push @curr_row, $val; # populate current cell into current row

		     print "Row, Col    = ($row, $col)\n" if ($full_debug);
	       print "Value       = ", $val,                 "\n" if ($full_debug);
         print "Unformatted = ", $cell->unformatted(), "\n" if ($full_debug && $cell);
         print "\n" if ($full_debug);
      } # each col

      push @curr_xl, \@curr_row; # add reference of current row
    } # each row
  } # each worksheet

  return (\@curr_xl);
}

# Read Excel with xlsx extension
sub read_xlsx {
  #my (@args) = @_;
  my ($filename) = @_;
  my @unmapped_header = ();

  my @curr_xl = ();

  die "read_xlsx() can only read xlsx file extension. But called for file \"$filename\"" if ($filename !~ /\.xlsx$/);
  die "Cannot find file named \"$filename\" to read in read_xlsx()\n" if (! -e $filename);

  print "read_xlsx() on file \"$filename\"\n" if ($debug);

  # my $excel = Spreadsheet::XLSX -> new ('test.xlsx', $converter);
  my $excel = Spreadsheet::XLSX -> new ($filename);

  foreach my $sheet (@{$excel -> {Worksheet}}) {

    printf("Sheet: %s\n", $sheet->{Name}) if ($debug);
    printf("maxrow = %0d min row = %0d\n",  $sheet -> {MaxRow},  $sheet -> {MinRow}) if ($full_debug);
    $sheet -> {MaxRow} ||= $sheet -> {MinRow};
    printf("maxrow = %0d min row = %0d\n",  $sheet -> {MaxRow},  $sheet -> {MinRow}) if ($full_debug);

    foreach my $row ($sheet -> {MinRow} .. $sheet -> {MaxRow}) {
      my @curr_row = ();
      $sheet -> {MaxCol} ||= $sheet -> {MinCol};

      foreach my $col ($sheet -> {MinCol} ..  $sheet -> {MaxCol}) {
        my $cell = $sheet -> {Cells} [$row] [$col];
        my $val  = $cell ? $cell -> {Val} : "";
        &massage_cell_value($val);

        push @curr_row, $val; # populate current cell into current row
        printf("( %s , %s ) => %s\n", $row, $col, $val) if ($full_debug);
      } # each col

      push @curr_xl, \@curr_row; # add reference of current row
    } # each row
  }

  return (\@curr_xl);
}


# Read Excel with csv extension
sub read_csv {
  #my (@args) = @_;
  my ($filename) = @_;
  my @unmapped_header = ();

  my @curr_xl = ();

  die "read_csv() can only read csv file extension. But called for file \"$filename\"" if ($filename !~ /\.csv$/);
  die "Cannot find file named \"$filename\" to read in read_csv()\n" if (! -e $filename);

  print "read_csv() on file \"$filename\"\n" if ($debug);

  my $workbook = ReadData ($filename);
  print Dumper($workbook) if ($full_debug);

  my $info     = $workbook->[0];
  print "Parsed $filename with $info->{parser} $info->{version}\n" if ($full_debug);

  my $data     = $workbook->[1];
  print Dumper($data) if ($full_debug);

  foreach my $row (1 .. $data->{maxrow}) {
    my @curr_row = ();

    foreach my $col (1 .. $data->{maxcol}) {
      my $cell = cr2cell ($col, $row);
      my $val  = $cell ? $data->{$cell} : "";
      &massage_cell_value($val);

      push @curr_row, $val; # populate current cell into current row
      printf "%-3s ", $val if ($full_debug);
    } # each col

    push @curr_xl, \@curr_row; # add reference of current row
  } # each row

  return (\@curr_xl);
}

# Function to handle each item in the source directory
# calls handling routines based on type (file or dir)
# will be recursively called to handle all excels under
#   the source directory including sub-directories
sub processItem {
  my ($name) = @_;
	if (-d "$name") {
		print "$name is a Directory \n" if ($debug);
    ++$dir_count;
		&handleDir($name);
	} elsif (-f $name) {
		print "$name is a File \n" if ($debug);
    ++$file_count;
		&handleFile($name);
	}
	return 0;
}

# Function to handle a single directory
# action on a dir item in the tree
# reads the directory contents
# removes . and ..
# process every other item in the directory
sub handleDir() {
	my ($dname) = @_;
	opendir (DIR, $dname);

	my @files;
	chomp(@files = readdir (DIR));
	print "Dir = $dname \n" if ($full_debug);
	print "Contents = @files \n" if ($full_debug);
	shift(@files);
	shift(@files);
	foreach my $i (@files) {
		&processItem("$dname/$i");
	}
	return 0;
}

# Function to handle a single file
# actions on file items in the dir tree
# based in excel file type, calls the appropriate excel reading function
sub handleFile() {
	my ($fname) = @_;
  my $curr_xl_ref;

  #return 0 if ( ($fname !~ /\.xls/) && ($fname !~ /\.xlsx/) && ($fname !~ /\.csv/) );

     if ( $fname =~ /\.xls$/  )  { $curr_xl_ref = &read_xls  ($fname); }
  elsif ( $fname =~ /\.xlsx$/ )  { $curr_xl_ref = &read_xlsx ($fname); }
  elsif ( $fname =~ /\.csv$/  )  { $curr_xl_ref = &read_csv  ($fname); }
  else {
         print "Non Excel File \"$fname\"\n";
         ++$non_xl_count;
         return 0; # nothing to do for non-excel files
       }

  &handle_xl_content($curr_xl_ref, $fname);

  ++$xl_count;
  return 0;
}

1;
