use strict;
if(!(defined $ARGV[0])) {
    print<<EOF;
Usage: $0 Excel_File
EOF
    exit;
}
use Spreadsheet::ParseExcel;
my $oExcel = new Spreadsheet::ParseExcel;
my $oBook = $oExcel->Parse($ARGV[0]);

my($iR, $iC, $oWkS, $oWkC);
print "<html>\n";
print "<!-- ========================================= -->\n";
print "<!-- FILE  :", $oBook->{File} , " -->\n";
print "<!-- Sheet Count :", $oBook->{SheetCount} , " -->\n";
print "<!-- AUTHOR:", $oBook->{Author} , " -->\n";
for(my $iSheet=0; $iSheet < $oBook->{SheetCount} ; $iSheet++) {
    $oWkS = $oBook->{Worksheet}[$iSheet];
    print "<!-- --------- SHEET:", $oWkS->{Name}, " -->\n";
    print "<table>\n";
    for(my $iR = $oWkS->{MinRow} ; 
            defined $oWkS->{MaxRow} && $iR <= $oWkS->{MaxRow} ; $iR++) {
        print "<tr><!-- L. ", $iR, " -->";
        for(my $iC = $oWkS->{MinCol} ;
                        defined $oWkS->{MaxCol} && $iC <= $oWkS->{MaxCol} ; $iC++) {
            $oWkC = $oWkS->{Cells}[$iR][$iC];
            print "<td>";
            # print "( $iR , $iC ) =>", $oWkC->Value, "\n" if($oWkC);
            print $oWkC->Value if($oWkC);
            print "</td>";
        }
        print "</tr>\n";
    }
    print "</table>\n";
}
print "</html>\n";
