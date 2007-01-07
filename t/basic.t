#!/usr/bin/perl
use strict;
use warnings;

use Test::More tests => 8;

use_ok('Spreadsheet::ParseExcel');
use_ok('Spreadsheet::ParseExcel::Dump');
use_ok('Spreadsheet::ParseExcel::FmtDefault');
use_ok('Spreadsheet::ParseExcel::FmtJapan');
use_ok('Spreadsheet::ParseExcel::FmtJapan2');
use_ok('Spreadsheet::ParseExcel::FmtUnicode');
use_ok('Spreadsheet::ParseExcel::SaveParser');
use_ok('Spreadsheet::ParseExcel::Utility');

