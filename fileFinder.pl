#!/usr/bin/perl -w

#use feature 'unicode_strings';
use utf8;
use strict;
use Excel::Writer::XLSX;
use Encode qw(decode encode);

my $usage = "Usage: <script.pl> <directory> <filename>";

my ($directory, $outFile) = @ARGV;

$directory && $outFile || die ( "$usage" );

#Remove .xls or .xlsx
#$outFile =~ s/\.xls(x?)//;

my ($sec,$min,$hour,$mday,$mon,$year) = localtime;
$year += 1900;
my $padString = sprintf("%04d%02d%02d_%02d%02d%02d", $year, $mon, $mday, $hour, $min, $sec).".xlsx";

#$outFile .= "_".$padString;

#binmode(STDOUT, ":encoding(utf8)");
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
my $format = $workbook->add_format();
$format->set_bold();
$format->set_bg_color( 'yellow' );

my $debug = 1;
for my $subsheet (findFiles($directory)) {
    #Create a saparate worksheet
    my $rowNum = 0;
    my $sheetPath = $directory."/".$subsheet; 
    if(opendir(my $dh, $sheetPath)) {
        closedir $dh;

        my $worksheet = $workbook->add_worksheet($subsheet);
        $worksheet->keep_leading_zeros();
        print "Sheet: $subsheet\n" if $debug;

        #Search subsheet
        for my $subDir (findFiles($sheetPath)) {
            my @treeOutputs = split /\n/, qx(tree $sheetPath/$subDir);
            @treeOutputs = grep { !/^(\w|\s)/ } @treeOutputs;

            #We need to highlight this line
            print "$subDir\n" if $debug;
            $worksheet->write($rowNum, 0, $subDir, $format);
            $rowNum++;

            for my $subline (@treeOutputs) {
                #$subline =~ s/\W/ /g;
                #$subline = encode("utf8", $subline);
                print "$subline\n" if $debug;
                $worksheet->write($rowNum, 0, $subline);
                $rowNum++;
            }
        }
    }
}

