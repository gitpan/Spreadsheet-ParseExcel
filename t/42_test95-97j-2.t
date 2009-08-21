#!perl -w

use strict;
use Test::More tests => 66;

use utf8;
use Encode qw(encode);

use Spreadsheet::ParseExcel;
use Spreadsheet::ParseExcel::FmtJapan;

my $xl   = Spreadsheet::ParseExcel->new();
my $fmtj = Spreadsheet::ParseExcel::FmtJapan->new(encoding => 'sjis');

foreach my $xls(qw(Test95J.xls Test97J.xls)){

	my $book = $xl->Parse("t/excel_files/$xls", $fmtj);
	ok $book, "load $xls";

	my $sheet = $book->worksheet(0);

	is $sheet->{Name}, 'Sheet1-ASC', '1. ASCII name';

	my @expected = (
		[ASC     => q{This Data is 'ASC Only'}],
		[Date    => q{1964/3/23}],
		[INTEGER => 12345],
		[Float   => 1.29],
		[Double  => 1234567.89012345],
		[Formula => 1246912.89012345],
		[Data    => 1234567.89],

		['BIG INTEGER'  => 123456789012],
	);

	#binmode STDOUT, ':encoding(cp932)';

	my($rmin, $rmax) = $sheet->row_range();
	my($cmin, $cmax) = $sheet->col_range();

	for my $i($rmin .. $rmax){
		for my $j($cmin .. $cmax){
			#print $sheet->get_cell($i, $j)->value, "\n";
			is $sheet->get_cell($i, $j)->value, $expected[$i][$j], "[$i, $j]";
		}
	}

	$sheet = $book->worksheet(1);

	is $sheet->{Name}, encode(cp932 => '漢字名'), '2. Kanji name';

	@expected = (
		[ASC     => q{This Data is 'ASC Only'}],
		[encode(cp932 => '漢字も入る') => encode(cp932 => '漢字のデータ')],
		[Date    => q{1964/3/23}],
		[INTEGER => 12345],
		[Float   => 1.29],
		[Double  => 1234567.89012345],
		[Formula => 1246912.89012345],
		[Float   => undef],
	);

	($rmin, $rmax) = $sheet->row_range();
	($cmin, $cmax) = $sheet->col_range();

	for my $i($rmin .. $rmax){
		for my $j($cmin .. $cmax){
			#print $sheet->get_cell($i, $j)->value, "\n";
			my $cell = $sheet->get_cell($i, $j);
			is ref($cell) ? $cell->value : $cell, $expected[$i][$j], "[$i, $j]";
		}
	}
}
