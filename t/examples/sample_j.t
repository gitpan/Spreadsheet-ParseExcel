use strict;
use warnings;

use Test::More;
use lib 't/lib';
use Test::Example qw(test_example_do);

eval {
    require Jcode;
};
if ($@) {
    plan skip_all => "Jcode needed for this test";
}
else {
    plan tests => 2;
}

test_example_do(
    dir     => 'sample',
    script  => 'sample_j.pl',
    stdout  => 'res_samplej', 
    argv    => ['euc'],
);

