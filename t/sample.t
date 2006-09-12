#!/usr/bin/perl
use strict;
use warnings;

use Test::More;
my $tests;
#plan tests => $tests;
plan skip_all => "Tests moved to individual files in t/examples/";
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
            skip  => 1,
        },
        {
            in    => 'sample.pl',
            out   => 'res_sample',
            skip  => 1,
        },
        {
            in    => 'sample_j.pl',
            out   => 'res_samplej',  
            argv  => ['euc'],
            jcode => 1, # needs Jcode
            skip  => 1,
        },
        {
            in    => 'sampleOEM.pl',
            out   => 'res_oem',
            unicode_map => 1,
            skip  => 11,
            #needs Unicode::Map and CP932Excel.map has to be installed somehow
        },
        {
            in    => 'dmpExR.pl',
            argv  => ['Excel/Rich.xls'],
            out   => 'res_rich',
            skip  => 1,
        },
        {
            in    => 'warning.pl',
            out   => 'warning_output',
            skip  => 1,
        }
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
close STDERR;
close STDOUT;
unshift @INC, '../blib/lib';
foreach my $t (@tests) {
    next if $t->{skip};
    next if $t->{jcode} and $no_jcode;

    my $err = '';
    open STDERR, '>', \$err or die;
    my $out = '';
    open STDOUT, '>', \$out or die;
    #print "23\n";
    #system "$^X -I../blib/lib $t->{in}";
    if ($t->{argv}) {
        @ARGV = @{ $t->{argv} };
        diag "@ARGV";
    }
    do $t->{in};
    @ARGV = ();
    close STDERR;
    close STDOUT;
    #diag $out;
    #exit;
    #my $err = slurp('err');
    is($err, '', "stderr when running $t->{in} is empty");
    my @expected_out = slurp($t->{out});
    #my @out = slurp('out');
    my @out = $out =~ /^.*\n/mg;
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


