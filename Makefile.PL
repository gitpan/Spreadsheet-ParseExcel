use ExtUtils::MakeMaker;
# See lib/ExtUtils/MakeMaker.pm for details of how to influence
# the contents of the Makefile that is written.
WriteMakefile(
    'NAME'  => 'Spreadsheet::ParseExcel',
    'VERSION_FROM' => 'ParseExcel.pm', # finds $VERSION
    'PREREQ_PM' => { OLE::Storage_Lite => 0.08 },
);
