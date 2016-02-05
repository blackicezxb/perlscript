#!/usr/bin/perl -w

use utf8;
use strict;
use Excel::Writer::XLSX;
use Encode qw(decode encode);

my $usage = "Usage: <script.pl> <directory> <filename>";

my ($directory, $outFile) = @ARGV;

$directory && $outFile || die ( "$usage" );

#Remove .xls or .xlsx
$outFile =~ s/\.xls(x?)//;

my ($sec,$min,$hour,$mday,$mon,$year) = localtime;
$year += 1900;
my $padString = sprintf("%04d%02d%02d_%02d%02d%02d", $year, $mon, $mday, $hour, $min, $sec).".xlsx";

$outFile .= "_".$padString;

binmode(STDOUT, ":encoding(utf8)");
print "Filename $outFile\n";

sub findFiles{
    my $parentDir = shift;
    my @dirs;
    if(opendir(my $dh, $parentDir)) {
        @dirs = grep { !/^\./ } readdir($dh);
        closedir $dh;
        @dirs = sort @dirs;
    }
    return @dirs;
}

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

# Enable verbose output
my $debug = 1;

for my $subsheet (findFiles($directory)) {
    #Create a saparate worksheet
    my $rowNum = 0;
    my $sheetPath = $directory."/".$subsheet; 
    if(opendir(my $dh, $sheetPath)) {
        closedir $dh;
        
        #Covert UTF8 to perl internal bytes
        #TODO: I guess Excel module only accept this kind of data 
        #TODO: However, output of Excel is UTF8 too. confusing.
        $subsheet = decode("utf8",$subsheet);
        print "Sheet: $subsheet\n" if $debug;

        #Create new worksheet
        my $worksheet = $workbook->add_worksheet($subsheet);
        $worksheet->keep_leading_zeros();
        $worksheet->set_column('A:A',100,$itemFormat);

        #Search subsheet
        for my $subDir (findFiles($sheetPath)) {
            my @treeOutputs = split /\n/, qx(tree $sheetPath/$subDir);
            my $errorCount = grep { /error opening dir/ } @treeOutputs;

            #Covert UTF8 to perl internal bytes
            $subDir = decode("utf8",$subDir);
            print "$subDir\n" if $debug;
            print "ErrorCounts: $errorCount\n" if $debug;

            #We don't need to highlight file node
            if($errorCount) {
                $worksheet->write($rowNum++, 0, $subDir, $itemFormat);

                #Clear arrary for file node
                @treeOutputs = ();
            } else {
                $worksheet->write($rowNum++, 0, $subDir, $topFormat);

                #Get arrary content
                @treeOutputs = grep { !/^\w/ } @treeOutputs;
            }

            for my $subline (@treeOutputs) {
                $subline = decode("utf8", $subline);
                print "$subline\n" if $debug;
                $worksheet->write($rowNum++, 0, $subline, $itemFormat);
            }
        }
    }
}

