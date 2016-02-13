#!/usr/bin/perl -w

use utf8;
use strict;
use Excel::Writer::XLSX;
use Encode qw(decode encode);

my $usage = "Usage: <script.pl> <input> <filename>";

my ($inputFile, $outFile) = @ARGV;

$inputFile && $outFile || die ( "$usage" );

#Remove .xls or .xlsx
$outFile =~ s/\.xls(x?)//;

my ($sec,$min,$hour,$mday,$mon,$year) = localtime;
$year += 1900;
my $padString = sprintf("%04d%02d%02d_%02d%02d%02d", $year, $mon, $mday, $hour, $min, $sec).".xlsx";

$outFile .= "_".$padString;

binmode(STDOUT, ":encoding(utf8)");
print "Filename $outFile\n";

#open(my $inputFh, "<:encoding(utf8)", $inputFile);
open(my $inputFh, "<:encoding(euc-cn)", $inputFile);

# Create a new Excel workbook
my $workbook = Excel::Writer::XLSX->new($outFile);

# Create format for first subdirectory line
my $topFormat = $workbook->add_format();
$topFormat->set_bold();
$topFormat->set_bg_color( 'yellow' );
$topFormat->set_align("left");
$topFormat->set_font("Times New Roman");

# Create format for normal cell
my $itemFormat = $workbook->add_format();
$itemFormat->set_align("left");
$itemFormat->set_font("Times New Roman");

my $firstFormat = $workbook->add_format();
$firstFormat->set_align("left");
$firstFormat->set_font("Times New Roman");
$firstFormat->set_size(16);
$firstFormat->set_italic(1);

# Enable verbose output
my $debug = 1;

my $tmpVar = 0; 

my $rowNum = 0;
my $worksheet;

while(!eof($inputFh)) {
    my $tmpLine = readline($inputFh);  
    print $tmpLine if ($tmpVar==0);
    if($tmpLine =~ /^(├─|└─)(.*)\n/) {
        #die if $tmpVar;
        print "SheetName: $2\n" if $debug;
        #Create new worksheet
        my $subsheet = $2;
        $worksheet = $workbook->add_worksheet($subsheet);
        $worksheet->keep_leading_zeros();
        $worksheet->set_column('A:A',100,$itemFormat);
        $rowNum = 0;

        $worksheet->write($rowNum++, 0, $subsheet, $firstFormat);
        $tmpVar++;
    } else {
        if(defined($worksheet)) {
            $tmpLine = substr $tmpLine, 3;
            #print $tmpLine;
            if($tmpLine =~ /^\s?(├─|└─)/) {
                $worksheet->write($rowNum++, 0, $tmpLine, $topFormat);
            } else {
                $worksheet->write($rowNum++, 0, $tmpLine, $itemFormat);
            }
        }
    }
}

