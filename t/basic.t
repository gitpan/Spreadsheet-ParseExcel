#!/usr/bin/perl
use strict;
use warnings;

use Test::More tests => 8;


use_ok('Spreadsheet::ParseExcel');
use_ok('Spreadsheet::ParseExcel::Dump');
use_ok('Spreadsheet::ParseExcel::FmtDefault');
use_ok('Spreadsheet::ParseExcel::Utility');

SKIP: {
    eval "use  Jcode";
    skip "Need Jcode for additional tests", 2 if $@;
    use_ok('Spreadsheet::ParseExcel::FmtJapan');
    use_ok('Spreadsheet::ParseExcel::FmtJapan2');
}

SKIP: {
    eval "use Unicode::Map";
    skip "Need Unicode::Map for additional tests", 1 if $@;
    use_ok('Spreadsheet::ParseExcel::FmtUnicode');
}

SKIP: {
    eval "use Spreadsheet::WriteExcel";
    skip "Need Spreadsheet::WriteExcel for additional tests", 1 if $@;
    use_ok('Spreadsheet::ParseExcel::SaveParser');
}



