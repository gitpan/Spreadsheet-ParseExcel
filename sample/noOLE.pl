use strict;
use Spreadsheet::ParseExcel;

my $oExcel = new Spreadsheet::ParseExcel(NotUseOLE=>1);
my $oBook = $oExcel->Parse('Excel/Test97.xls');

my($iR, $iC, $oWkS, $oWkC);
  
print "FILE  :", $oBook->{File} , "\n";
print "COUNT :", $oBook->{SheetCount} , "\n";
print "AUTHOR:", $oBook->{Author} , "\n";
for(my $iSheet=0; $iSheet < $oBook->{SheetCount} ; $iSheet++) {
    $oWkS = $oBook->{Worksheet}[$iSheet];
    print "--------- SHEET:", $oWkS->{Name}, "\n";
    for(my $iR = $oWkS->{MinRow} ; 
            defined $oWkS->{MaxRow} && $iR <= $oWkS->{MaxRow} ; $iR++) {
        for(my $iC = $oWkS->{MinCol} ;
               defined $oWkS->{MaxCol} && $iC <= $oWkS->{MaxCol} ; $iC++) {
            $oWkC = $oWkS->{Cells}[$iR][$iC];
            print "( $iR , $iC ) =>", $oWkC->Value, "\n" if($oWkC);
        }
    }
}

