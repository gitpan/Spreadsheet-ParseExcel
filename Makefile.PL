use strict;
use warnings;
use ExtUtils::MakeMaker;
use Config;

my %deps = (
    'OLE::Storage_Lite'   => 0.08,
    'IO::File'            => 0,
    'Test::More'          => '0.47',
);
exists $Config{useperlio} && $Config{useperlio} eq "define" and
    $deps{'IO::Scalar'}   =  0;

WriteMakefile(
          'PL_FILES'     => {},
          'INSTALLDIRS'  => 'site',
          'NAME'         => 'Spreadsheet::ParseExcel',
          'EXE_FILES'    => [],
          'VERSION_FROM' => 'lib/Spreadsheet/ParseExcel.pm',
          'PREREQ_PM'    => \%deps, 
);
