use strict;
use warnings;

use Test::More;
use lib 't/lib';
use Test::Example qw(test_example_do);

plan skip_all => "We need Unicode::Map and CP932Excel.map has to be installed somehow...  This test is not in use currently";

eval {
    require Unicode::Map;
};
if ($@) {
    plan skip_all => "Unicode::Map needed for this test";
}
else {
    plan tests => 2;
}

test_example_do(
    dir     => 'sample',
    script  => 'sampleOEM.pl',
    stdout  => 'res_oem', 
);

