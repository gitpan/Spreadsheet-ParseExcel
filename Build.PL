use strict;
use warnings;
use Module::Build;
use Config;

my %deps = (
    'OLE::Storage_Lite'   => 0.08,
    'IO::File'            => 0,
);
exists $Config{useperlio} && $Config{useperlio} eq "define" and
    $deps{'IO::Scalar'}   =  0;


my $builder = Module::Build->new(
    module_name           => 'Spreadsheet::ParseExcel',
    license               => 'perl',
    dist_author           => 'Kawai Takanori',
    #create_makefile_pl    => 'traditional',
    create_readme         => 0,
    recursive_test_files  => 1,
    requires              => \%deps,
    build_requires        => {
        'Test::More'                => '0.47',
    },
);

$builder->create_build_script();
