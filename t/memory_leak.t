#!/usr/bin/perl
use strict;
use warnings;

# USe it either as a regular test file or by typing something like this:
# perl -Ilib t/memory_leak.t --file sample/Excel/AuthorK.xls --count 100

use Test::More;
use Getopt::Long qw(GetOptions);
my $tests;

eval {
    require Proc::ProcessTable;
};
if ($@) {
    plan skip_all  => "Proc::ProcessTable is needed for this test";
}
else {
    plan tests => $tests;
}

my $count = 10;
my $file  = 'sample/Excel/Test95.xls';
GetOptions(
    "count" => \$count,
    "file"  => \$file,
) or die;


sub get_process_size {
    my ($pid) = @_;
    my $pt = Proc::ProcessTable->new;
    foreach my $p ( @{$pt->table} ) {
        return $p->size if $pid == $p->pid;
    }
    return;
}


use_ok('Spreadsheet::ParseExcel');
BEGIN { $tests += 1; }

diag "using version $Proc::ProcessTable::VERSION of Proc::ProcessTable";
diag "Testing version $Spreadsheet::ParseExcel::VERSION of Spreadsheet::ParseExcel";

# 1:  131072 - 135168
# 5:  270336
# 8:  405504
# 10: 540672
# 13:  675840
# 15:  811008
# 18:  946176
# 21:  1077248
# ...
# 100:  5218304

my $begin_size = get_process_size($$);
do_something();
my $start_size = get_process_size($$);
my $memory_consumption = $start_size - $begin_size;
diag "Memory consumption by one call is $memory_consumption"; 
# This value was about 286720 - 290816 when using version 0.29

diag "Testing memory leak, running $count times";
diag "Start size: $start_size";
foreach (2..$count) {
    do_something();
    my $size = get_process_size($$);
    #diag ("$_:  " . ($size - $start_size));
}
my $end_size = get_process_size($$);
my $size_change =  $end_size - $start_size;
diag "End size: $end_size";
diag "Size change was: $size_change";
cmp_ok($size_change, '<', 100, 
    'normally it should be 0 but we are not that picky');
BEGIN { $tests += 1; }

sub do_something {
    my $workbook = Spreadsheet::ParseExcel::Workbook->Parse($file);
}

