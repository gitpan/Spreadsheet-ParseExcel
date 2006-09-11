#!/usr/bin/perl
use strict;
use warnings;

use Test::More;
my $tests;
plan tests => $tests;
my @tests;
my $no_jcode;

BEGIN {

    eval {
        require Jcode;
    };
    $no_jcode = $@ ? 1 : 0;

    
    # jcode => 1 means the given piece of code needs Jcode installed
    # skip  => 1 is used to skip long running tests during developmen
    @tests = (
        {
            in    => 'chkFmt.pl',
            out   => 'res_fmt',
        },
        {
            in    => 'chkInfo.pl',
            out   => 'res_info',
            skip  => 0,
        },
        {
            in    => 'sample.pl',
            out   => 'res_sample',
            skip  => 0,
        },
        {
            in    => 'sample_j.pl euc',
            out   => 'res_samplej',  
            jcode => 1, # needs Jcode
            skip  => 0,
        },
        {
            in    => 'sampleOEM.pl',
            out   => 'res_oem',
            unicode_map => 1,
            skip  => 1,
            #needs Unicode::Map and CP932Excel.map has to be installed somehow
        },
        {
            in    => 'dmpExR.pl Excel/Rich.xls',
            out   => 'res_rich',
        },
    );
    # More files from sample/README that have not been used yet in tests
    # dmpEx.pl 
    # dmpExJ.pl
    # dmpExU.pl
    # dmpExH.pl
    # Ilya.pl
    # smpFile.pl
    # xls2csv.pl
    # iftest.pl
    # iftestj.pl
    
    my $jcode_cnt = 0;
    my $skip = 0;
    foreach my $t (@tests) {
        if ($t->{skip}) {
            $skip++;
        }
        elsif ($t->{jcode}) {
            $jcode_cnt++;
        }
    }
    $tests += 2 * (scalar(@tests) - $jcode_cnt*$no_jcode - $skip);
    diag "Skipping " . (2 * $jcode_cnt) . " test as they need Jcode to be installed"
        if $no_jcode;
}

diag "Some of the tests can take a long time, please be patient";
chdir "sample";
foreach my $t (@tests) {
    next if $t->{skip};
    next if $t->{jcode} and $no_jcode;

    system "$^X $t->{in} > out 2>err";

    my $err = slurp('err');
    is($err, '', "stderr when running $t->{in} is empty");
    my @expected_out = slurp($t->{out});
    my @out = slurp('out');
    is_deeply(\@out, \@expected_out, "stdout when running $t->{in}");
}
unlink 'out', 'err';




sub slurp {
    my ($file) = @_;
    open my $fh, '<', $file or die "Could not open $file: $!";
    if (wantarray) {
        return <$fh>;
    } else {
        local $/ = undef;
        return <$fh>;
    }
}


