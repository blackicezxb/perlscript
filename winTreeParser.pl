#!/usr/bin/perl -w

use utf8;
use strict;
use Excel::Writer::XLSX;
use Encode qw(decode encode);

my $usage = "Usage: <script.pl> <input> <oldInput> <filename> <targetPath>";

my ($inputFile, $oldFile, $outFile, $targetPath) = @ARGV;

$inputFile && $outFile || die ( "$usage" );

#Remove .xls or .xlsx
$outFile =~ s/\.xls(x?)//;

my ($sec,$min,$hour,$mday,$mon,$year) = localtime;
$year += 1900;
$mon  += 1;
my $padString = sprintf("%04d%02d%02d_%02d%02d%02d", $year, $mon, $mday, $hour, $min, $sec).".xlsx";
#my $padString = sprintf("%04d%02d%02d", $year, $mon, $mday).".xlsx";

$outFile .= "_".$padString;

print "Filename $outFile\n";
binmode(STDOUT, ":encoding(utf8)");



my $diffOutput = qx(diff $oldFile $inputFile);
$diffOutput = decode("euc-cn", $diffOutput);
my @diffStrings= split /\n/, $diffOutput;

#Parse two file differences
my $newStart;
my (@oldArray, @diffArray); 
my $newIdx;
@diffArray = ();

for my $tmpLine (@diffStrings) { 
    #print "$tmpLine\n";
    
    if( $tmpLine =~ /^\d.*c(\d+)(,(\d+))?/ ) {
        $newStart = $1;
        @oldArray = (); 
        $newIdx=0;
        #print "NewStart: $newStart\n";
    } elsif ($tmpLine =~ /d(\d+)$/) { 
        $newStart = $1;
        @oldArray = (); 
        $newIdx=0;
        #print "NewStart: $newStart\n";
    } elsif ($tmpLine =~ /a(\d+),?/) { 
        $newStart = $1;
        @oldArray = (); 
        $newIdx=0;
        #print "NewStart: $newStart\n";
    } elsif( $tmpLine =~ /^<\W*(\w.*)/) {
        #print "OLD: $1\n";
        push @oldArray, $1;
    } elsif ( $tmpLine =~ /^>\W*(\w.*)/) {
        my $tmpNum = $newIdx + $newStart; 
        if($newIdx <= $#oldArray) {
            if($1 ne $oldArray[$newIdx]) {
                #print "$tmpNum\n";
                push @diffArray, $tmpNum; 
            }
        } else {
            #print "$tmpNum\n";
            push @diffArray, $tmpNum;
        }
    } 

    if ( $tmpLine =~ /^>/){
        $newIdx++;
    }
}

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

my $highlightTop = $workbook->add_format();
$highlightTop->set_bold();
$highlightTop->set_align("left");
$highlightTop->set_bg_color( 'cyan' );
$highlightTop->set_font("Times New Roman");

# Create format for normal cell
my $itemFormat = $workbook->add_format();
$itemFormat->set_align("left");
$itemFormat->set_font("Times New Roman");

my $highlightItem = $workbook->add_format();
$highlightItem->set_align("left");
$highlightItem->set_bg_color( 'cyan' );
$highlightItem->set_font("Times New Roman");

my $firstFormat = $workbook->add_format();
$firstFormat->set_align("left");
$firstFormat->set_font("Times New Roman");
$firstFormat->set_size(16);
$firstFormat->set_italic(1);

my %directHash;
sub createDirect{
    my $inputKey = shift;
    my $current = shift;

    foreach my $key (keys %directHash) {
        if($key >= ($inputKey - 1)) {
            delete $directHash{$key};
        }
    }
    $directHash{$inputKey} = $current;
}

sub getDirectory{
    my $direct;

    if(defined($targetPath)){
        $direct = "\\\\lmssha-ns01/Common/LMS/LMS LSR/".decode('utf8',$targetPath)."/";
    }

    foreach my $key (sort keys %directHash) {
        $direct .= $directHash{$key}."/";
    }
    return $direct;
}

# Enable verbose output
my $debug = 1;

my $tmpVar = 0; 

my $rowNum = 0;
my $worksheet;

my $lineNum = 0;
my $diffIdx = 0;

while(!eof($inputFh)) {
    my $tmpLine = readline($inputFh);  
    $lineNum++;

    #print $tmpLine if ($tmpVar==0);
    if($tmpLine =~ /^(├─|└─)(.*)\r\n/) {
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

        $tmpLine = $2;
        $tmpLine =~ s/(\w*)\r\n/$1/;
        $directHash{0} = $tmpLine;
    } else {
        if(defined($worksheet)) {
            $tmpLine = substr $tmpLine, 3;
            #print $tmpLine;
            if($tmpLine =~ /^\s?(├─|└─)/g) {
                #print "KEY:", pos $tmpLine, "\n";
                my $fileName = $tmpLine; 
                $fileName =~ s/\W*//;
                $fileName =~ s/(\w*)\r\n/$1/;
                createDirect(pos $tmpLine, $fileName);
                my $cellFormat = $topFormat;
                if($#diffArray >= $diffIdx) { if($lineNum == $diffArray[$diffIdx]) {
                    $cellFormat = $highlightTop;
                    #print "LINENUM: $lineNum\n";
                    $diffIdx++;
                }}

                if(defined($targetPath)) {
                    my $urlStr = "external:".getDirectory();
                    $worksheet->write($rowNum++, 0, $urlStr, $tmpLine, $cellFormat);
                } else {
                    $worksheet->write($rowNum++, 0, $tmpLine, $cellFormat);
                }
            } else {
                my $cellFormat = $itemFormat;
                if($#diffArray >= $diffIdx) { if($lineNum == $diffArray[$diffIdx]) {
                    $cellFormat = $highlightItem;
                    #print "LINENUM: $lineNum\n";
                    $diffIdx++;
                }}

                if( $tmpLine =~ /(├─|└─)/g ) {
                    #print "KEY:", pos $tmpLine, "\n";
                    my $fileName = $tmpLine; 
                    $fileName =~ s/\W*//;
                    $fileName =~ s/(\w*)\r\n/$1/;
                    createDirect(pos $tmpLine, $fileName);
                    if(defined($targetPath)) {
                        my $urlStr = "external:".getDirectory();
                        $worksheet->write($rowNum++, 0, $urlStr, $tmpLine, $cellFormat);
                    } else {
                        $worksheet->write($rowNum++, 0, $tmpLine, $cellFormat);
                    }
                } else {
                    if(defined($targetPath)) {
                        my $fileName = $tmpLine; 
                        $fileName =~ s/\W*//;
                        $fileName =~ s/(\w*)\r\n/$1/;
                        my $urlStr = "external:".getDirectory().$fileName;
                        if((length $urlStr) > 215 ) {
                            $urlStr = "external:".getDirectory();
                        }
                        $worksheet->write($rowNum++, 0, $urlStr, $tmpLine, $cellFormat); 
                    } else {
                        $worksheet->write($rowNum++, 0, $tmpLine, $cellFormat);
                    }
                }
            }
        }
    }
}

