#!/usr/bin/perl
use strict;
use warnings;

use Test::More;
my $tests;
plan tests => $tests;
use Data::Dumper;

use_ok('Spreadsheet::ParseExcel');
BEGIN { $tests += 1; }


# these tests were created based on the values received using 0.27
# to create regressions tests

{
    # historically Parse returned and unblessed reference on missing file 
    my $excel = Spreadsheet::ParseExcel::Workbook->Parse('no_such_file.xls');
    is(ref($excel), 'HASH', 'failed Parse method returns HASH ref');
    isa_ok($excel->{_Excel}, 'Spreadsheet::ParseExcel', 
            'failed Parse method creates _Excel object');
    BEGIN { $tests += 2; }
}
{
    # historically Parse returned and unblessed reference on failure (e.g.
    # input file is not an Excel file)
    my $excel = Spreadsheet::ParseExcel::Workbook->Parse($0);
    is(ref($excel), 'HASH', 'failed Parse method returns HASH ref');
    isa_ok($excel->{_Excel}, 'Spreadsheet::ParseExcel', 
            'failed Parse method creates _Excel object');
    BEGIN { $tests += 2; }
}

{
    my $excel = Spreadsheet::ParseExcel::Workbook->Parse('sample/Excel/Test95.xls');
    is(ref($excel), 'Spreadsheet::ParseExcel::Workbook',
            'Spreadsheet::ParseExcel::Workbook created');
    isa_ok($excel->{_Excel}, 'Spreadsheet::ParseExcel', 
            'Parse method creates _Excel object');

    my @sheets = @{$excel->{Worksheet}};
    is (@sheets, 2, "two sheets");
    is($sheets[0]->{Name}, 'Sheet1-ASC');     # Open Office shows: 'Sheet1_ASC'
    is($sheets[1]->{Name}, 'Sheet1-ASC (2)'); # OO shows 'Sheet1_ASC_2_' 

    is($sheets[0]->{MinRow}, 0);
    is($sheets[0]->{MaxRow}, 7);
    #diag Dumper $sheets[0]->{Cells};
    #qw(ASC Date INTEGER Float Double Formula)
    is($sheets[0]->{Cells}[0][0]->{Val}, 'ASC');
    #diag Dumper $sheets[0]->{Cells}[0][0];


    is($sheets[1]->{MinRow}, 0);
    is($sheets[1]->{MaxRow}, 5);

    BEGIN { $tests += 10; }
}

