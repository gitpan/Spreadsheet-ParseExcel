# Spreadsheet::ParseExcel::SaveParser
#  by Kawai, Takanori (Hippo2000) 2001.5.1
# This Program is ALPHA version.
#==============================================================================
package Spreadsheet::ParseExcel::SaveParser::Workbook;
require Exporter;
use strict;
use vars qw($VERSION @ISA);
@ISA = qw(Spreadsheet::ParseExcel::Workbook Exporter);
$VERSION = '0.01'; # 
#==============================================================================
# Spreadsheet::ParseExcel::SaveParser::Workbook
#==============================================================================
package Spreadsheet::ParseExcel::SaveParser::Workbook;
require Exporter;
use strict;
use vars qw($VERSION @ISA);
@ISA = qw(Spreadsheet::ParseExcel::Workbook Exporter);
sub new($$) {
    my($sPkg, $oBook) = @_;
    return undef unless(defined $oBook);
    my %oThis = %$oBook;
    bless \%oThis, $sPkg;
    return \%oThis;
}
#------------------------------------------------------------------------------
# AddWorksheet (for Spreadsheet::ParseExcel::SaveParser)
#------------------------------------------------------------------------------
sub AddWorksheet($$%) {
    my($oThis, $sName, %hAttr) = @_;
    $hAttr{Name} ||= $sName;
    my $oWkS = Spreadsheet::ParseExcel::Worksheet->new(%hAttr);
    $oThis->{Worksheet}[$oThis->{SheetCount}] = $oWkS;
    $oThis->{SheetCount}++;
    return $oThis->{SheetCount} - 1;
}
#------------------------------------------------------------------------------
# AddFormat (for Spreadsheet::ParseExcel::SaveParser)
#------------------------------------------------------------------------------
sub AddFormat($%){
    my ($oBook, %hAttr) = @_;
    push @{$oBook->{Format}}, 
        Spreadsheet::ParseExcel::Format->new(%hAttr);
    return $#{$oBook->{Format}};
}
#------------------------------------------------------------------------------
# AddFont (for Spreadsheet::ParseExcel::SaveParser)
#------------------------------------------------------------------------------
sub AddFont($%){
    my ($oBook, %hAttr) = @_;
    push @{$oBook->{Font}}, 
        Spreadsheet::ParseExcel::Font->new(%hAttr);
    return $#{$oBook->{Font}};
}
#------------------------------------------------------------------------------
# AddCell (for Spreadsheet::ParseExcel::SaveParser)
#------------------------------------------------------------------------------
sub AddCell($$$$$$) {
    my($oBook, $iSheet, $iR, $iC, $sVal, $oCell)=@_;
    my %rhKey;

    my $iFmt = (UNIVERSAL::isa($oCell, 'Spreadsheet::ParseExcel::Cell'))?
                $oCell->{FormatNo} : $oCell;
    $rhKey{FormatNo} = $iFmt;
    $rhKey{Val}      = $sVal;
    $oBook->{_CurSheet} = $iSheet;
    Spreadsheet::ParseExcel::_NewCell($oBook, $iR, $iC, %rhKey);
    Spreadsheet::ParseExcel::_SetDimension($oBook, $iR, $iC, $iC);
}
1;
#==============================================================================
# Spreadsheet::ParseExcel::SaveParser
#==============================================================================
package Spreadsheet::ParseExcel::SaveParser;
require Exporter;
use strict;
use Spreadsheet::WriteExcel;
use Spreadsheet::ParseExcel;
use vars qw($VERSION @ISA);
@ISA = qw(Spreadsheet::ParseExcel Exporter);
$VERSION = '0.01'; # 
use constant MagicCol => 1.14;
#------------------------------------------------------------------------------
# new (for Spreadsheet::ParseExcel::SaveParser)
#------------------------------------------------------------------------------
sub new($%) {
    my($sPkg, %hKey) = @_;
    my $oThis = new Spreadsheet::ParseExcel(%hKey);
    bless $oThis, $sPkg;
    return $oThis;
}
#------------------------------------------------------------------------------
# Parse (for Spreadsheet::ParseExcel::SaveParser)
#------------------------------------------------------------------------------
sub Parse($$;$) {
    my($oThis, $sFile, $oWkFmt)=@_;
    my $oBook = $oThis->SUPER::Parse($sFile, $oWkFmt);
    return undef unless(defined $oBook);
    return Spreadsheet::ParseExcel::SaveParser::Workbook->new($oBook);
}
#------------------------------------------------------------------------------
# SaveAs (for Spreadsheet::ParseExcel::FmtJapan2)
#------------------------------------------------------------------------------
sub SaveAs($$$){
    my ($oThis, $oBook, $sName)=@_;
    # Create a new Excel workbook
    my $oWrEx = Spreadsheet::WriteExcel->new($sName);
    my %hFmt;

    my $iNo = 0;
    my @aAlH = (undef, 'left', 'center', 'right', 'fill', 'justify', 'merge', 'equal_space');
    my @aAlV = ('top' , 'vcenter', 'bottom', 'vjustify', 'vequal_space');

    foreach my $pFmt (@{$oBook->{Format}}) {
        my $oFmt = $oWrEx->addformat();    # Add Formats
        unless($pFmt->{Style}) {
            $hFmt{$iNo} = $oFmt;
            my $rFont = $pFmt->{Font};

            $oFmt->set_font($rFont->{Name});
            $oFmt->set_size($rFont->{Height});
            $oFmt->set_color($rFont->{Color});
            $oFmt->set_bold($rFont->{Bold});
            $oFmt->set_italic($rFont->{Italic});
            $oFmt->set_underline($rFont->{Underline});
            $oFmt->set_font_strikeout($rFont->{Strikeout});
            $oFmt->set_font_script($rFont->{Super});

            $oFmt->set_align($aAlH[$pFmt->{AlignH}]);
            $oFmt->set_align($aAlV[$pFmt->{AlignV}]);
            if($pFmt->{Rotate}==0) {
                #‰ñ“]–³‚µ
                $oFmt->set_rotation(0);
            }
            elsif($pFmt->{Rotate}> 0) {  # Mainly ==90
                $oFmt->set_rotation(3);
            }
            elsif($pFmt->{Rotate} < 0) {  # Mainly == -90
                $oFmt->set_rotation(2);
            }
            $oFmt->set_num_format($oBook->{FmtClass}->FmtStringDef($pFmt->{FmtIdx}, $oBook));

            $oFmt->set_text_wrap($pFmt->{Wrap});

            $oFmt->set_pattern($pFmt->{Fill}->[0]);
            $oFmt->set_fg_color($pFmt->{Fill}->[1]) 
                        if(($pFmt->{Fill}->[1] >= 8) && ($pFmt->{Fill}->[1] <= 63));
            $oFmt->set_bg_color($pFmt->{Fill}->[2])
                        if(($pFmt->{Fill}->[2] >= 8) && ($pFmt->{Fill}->[2] <= 63));

            $oFmt->set_left  (($pFmt->{BdrStyle}->[0]>7)? 3: $pFmt->{BdrStyle}->[0]);
            $oFmt->set_right (($pFmt->{BdrStyle}->[1]>7)? 3: $pFmt->{BdrStyle}->[1]);
            $oFmt->set_top   (($pFmt->{BdrStyle}->[2]>7)? 3: $pFmt->{BdrStyle}->[2]);
            $oFmt->set_bottom(($pFmt->{BdrStyle}->[3]>7)? 3: $pFmt->{BdrStyle}->[3]);

            $oFmt->set_left_color  ($pFmt->{BdrColor}->[0])
                        if(($pFmt->{BdrColor}->[0] >= 8) && ($pFmt->{BdrColor}->[0] <= 63));
            $oFmt->set_right_color ($pFmt->{BdrColor}->[1])
                        if(($pFmt->{BdrColor}->[1] >= 8) && ($pFmt->{BdrColor}->[1] <= 63));
            $oFmt->set_top_color   ($pFmt->{BdrColor}->[2])
                        if(($pFmt->{BdrColor}->[2] >= 8) && ($pFmt->{BdrColor}->[2] <= 63));
            $oFmt->set_bottom_color($pFmt->{BdrColor}->[3])
                        if(($pFmt->{BdrColor}->[3] >= 8) && ($pFmt->{BdrColor}->[3] <= 63));
        }
        $iNo++;
    }
    for(my $iSheet=0; $iSheet < $oBook->{SheetCount} ; $iSheet++) {
        my $oWkS = $oBook->{Worksheet}[$iSheet];
        my $oWrS = $oWrEx->addworksheet($oWkS->{Name});
=cmmt
        $oWrS->set_printinfo( 
            Landscape    => $oWkS->{Landscape},            # Landscape (0:Horizontal, 1:Vertical)
            Scale        => $oWkS->{Scale},                # Scale
            FitWidth     => $oWkS->{FitWidth},             # Pages on fit with width
            FitHeight    => $oWkS->{FitHeight},            # Pages on fit with height
            PageFit      => $oWkS->{PageFit},              # Pages on fit
            PaperSize    => $oWkS->{PaperSize},            # Paper size (8=A3)
            PageStart    => $oWkS->{PageStart},            # Page number for start
            UsePage      => $oWkS->{UsePage},              # Use own start page number

            Mergin       => [$oWkS->{LeftMergin},
                             $oWkS->{RightMergin},
                             $oWkS->{TopMergin},
                             $oWkS->{BottomMergin},
                             $oWkS->{HeaderMergin},
                             $oWkS->{FooterMergin}, ],      # Mergins(Left,Right,Top,Bottom,Header,Footer)
            HCenter      => $oWkS->{HCenter},               # Horizontal Center
            VCenter      => $oWkS->{VCenter},               # Vertical Center

            Header       => $oWkS->{Header},                # Header
            Footer       => $oWkS->{Footer},                # Footer

            PrintArea    => $oBook->{PrintArea}[$iSheet],
            PrintTitle   => {Row    => $oBook->{PrintTitle}[$iSheet]->{Row}[0],
                             Column => $oBook->{PrintTitle}[$iSheet]->{Column}[0]}, 
                                                            # PrintTitles(Row, Column)
            PrintGrid    => $oWkS->{PrintGrid},             # Print Gridlines
            PrintHeaders => $oWkS->{PrintHeaders},          # Print Headings
            NoColor      => $oWkS->{NoColor},               # Print in blcak-white
            Draft        => $oWkS->{Draft},                 # Print in draft mode
            Notes        => $oWkS->{Notes},                 # Print notes
            LeftToRight  => $oWkS->{LeftToRight},           # Left to Right

            HPageBreak   => $oWkS->{HPageBreak},            # Horizontal Page Breaks
            VPageBreak   => $oWkS->{VPageBreak},            # Veritical Page Breaks
        );
=cut
        for(my $iC = $oWkS->{MinCol} ;
                            defined $oWkS->{MaxCol} && $iC <= $oWkS->{MaxCol} ; $iC++) {            
            if(defined $oWkS->{ColWidth}[$iC]) {
                $oWrS->set_column($iC, $iC, $oWkS->{ColWidth}[$iC]) ;
                #$oWrS->set_column($iC, $iC, $oWkS->{ColWidth}[$iC] * MagicCol) ;
            }
        }
        for(my $iR = $oWkS->{MinRow} ; 
                defined $oWkS->{MaxRow} && $iR <= $oWkS->{MaxRow} ; $iR++) {
            $oWrS->set_row($iR, $oWkS->{RowHeight}[$iR]);
            for(my $iC = $oWkS->{MinCol} ;
                            defined $oWkS->{MaxCol} && $iC <= $oWkS->{MaxCol} ; $iC++) {

                my $oWkC = $oWkS->{Cells}[$iR][$iC];
                if($oWkC) {
                    if($oWkC->{Merged}) {
                        my $oFmtN = $oWrEx->addformat();
                        $oFmtN->copy($hFmt{$oWkC->{FormatNo}});
                        $oFmtN->set_merge(1);
                        $oWrS->write($iR , $iC, $oBook->{FmtClass}->TextFmt($oWkC->{Val}, $oWkC->{Code}), 
                            $oFmtN);
                    }
                    else {
                        $oWrS->write($iR , $iC, $oBook->{FmtClass}->TextFmt($oWkC->{Val}, $oWkC->{Code}), 
                            $hFmt{$oWkC->{FormatNo}});
                    }
                }
            }
        }
    }
}
1;

__END__

=head1 NAME

Spreadsheet::ParseExcel::SaveParser - Expand of Spreadsheet::ParseExcel with Spreadsheet::WriteExcel

=head1 SYNOPSIS

    use strict;
    use Spreadsheet::ParseExcel::SaveParser;
    my $oExcel = new Spreadsheet::ParseExcel::SaveParser;
    my $oBook = $oExcel->Parse('some.xls');

    my $oBWr = Spreadsheet::ParseExcel::SaveParser::Workbook->new($oBook);
    my $oExcel = new Spreadsheet::ParseExcel;

    #1.1 Normal Excel97
    my $oBook = $oExcel->Parse('Excel/Test97.xls');
    my($iR, $iC, $oWkS, $oWkC);
    print "FILE  :", $oBook->{File} , "\n";
    print "COUNT :", $oBook->{SheetCount} , "\n";
    print "AUTHOR:", $oBook->{Author} , "\n";
    for(my $iSheet=0; $iSheet < $oBook->{SheetCount} ; $iSheet++) {
        $oWkS = $oBook->{Worksheet}[$iSheet];
        print "--------- SHEET:", $oWkS->{Name}, "\n";
        for(my $iR = $oWkS->{MinRow} ; 
                defined $oWkS->{MaxRow} && $iR <= $oWkS->{MaxRow} ; $iR++) {
            for(my $iC = $oWkS->{MinCol} ;
                            defined $oWkS->{MaxCol} && $iC <= $oWkS->{MaxCol} ; $iC++) {
                $oWkC = $oWkS->{Cells}[$iR][$iC];
                print "( $iR , $iC ) =>", $oWkC->Value, "\n" if($oWkC);
            }
        }
    }

=head1 DESCRIPTION

Spreadsheet::ParseExcel makes you to get information from Excel95, Excel97, Excel2000 file.

=head2 Functions

=over 4

=item new

I<$oExcel> = new Spreadsheet::ParseExcel(
                    [ I<CellHandler> => \&subCellHandler, 
                      I<NotSetCell> => undef | 1,
                    ]);

Constructor.


=over 4

=item CellHandler I<(experimental)>

specify callback function when a cell is detected.

I<subCellHandler> gets arguments like below:

sub subCellHandler (I<$oBook>, I<$iSheet>, I<$iRow>, I<$iCol>, I<$oCell>);

B<CAUTION> : The atributes of Workbook may not be complete.
This function will be called almost order by rows and columns.
Take care B<almost>, I<not perfectly>.

=item NotSetCell I<(experimental)>

specify set or not cell values to Workbook object.

=back

=item Parse

I<$oWorkbook> = $oParse->Parse(I<$sFileName> [, I<$oFmt>]);

return L<"Workbook"> object.
if error occurs, returns undef.

=over 4

=item I<$sFileName>

name of the file to parse

From 0.12 (with OLE::Storage_Lite v.0.06), 
scalar reference of file contents (ex. \$sBuff) or 
IO::Handle object (inclucdng IO::File etc.) are also available.

=item I<$oFmt>

L<"Formatter Class"> to format the value of cells.

=back

=item ColorIdxToRGB

I<$sRGB> = $oParse->ColorIdxToRGB(I<$iColorIdx>);

I<ColorIdxToRGB> returns RGB string corresponding to specified color index.
RGB string has 6 charcters, representing RGB hex value. (ex. red = 'FF0000')

=back

=head2 Workbook

I<Spreadsheet::ParseExcel::Workbook>

Workbook class has these properties :

=over 4

=item File

Name of the file

=item Author

Author of the file

=item Flag1904

If this flag is on, date of the file count from 1904.

=item Version

Version of the file

=item SheetCount

Numbers of L<"Worksheet"> s in that Workbook

=item Worksheet[SheetNo]

Array of L<"Worksheet">s class

=item PrintArea[SheetNo]

Array of PrintArea array refs.

Each PrintArea is : [ I<StartRow>, I<StartColumn>, I<EndRow>, I<EndColumn>]

=item PrintTitle[SheetNo]

Array of PrintTitle hash refs.

Each PrintTitle is : 
        { Row => [I<StartRow>, I<EndRow>], 
          Column => [I<StartColumn>, I<EndColumn>]}

=back

=head2 Worksheet

I<Spreadsheet::ParseExcel::Worksheet>

Worksheet class has these properties:

=over 4

=item Name

Name of that Worksheet

=item DefRowHeight

Default height of rows

=item DefColWidth

Default width of columns

=item RowHeight[Row]

Array of row height

=item ColWidth[Col]

Array of column width (undef means DefColWidth)

=item Cells[Row][Col]

Array of L<"Cell">s infomation in the worksheet

=item Landscape

Print in horizontal(0) or vertical (1).

=item Scale

Print scale.

=item FitWidth

Number of pages with fit in width. 

=item FitHeight

Number of pages with fit in height.

=item PageFit

Print with fit (or not).

=item PaperSize

Papar size. The value is like below:

  Letter               1, LetterSmall          2, Tabloid              3 ,
  Ledger               4, Legal                5, Statement            6 ,
  Executive            7, A3                   8, A4                   9 ,
  A4Small             10, A5                  11, B4                  12 ,
  B5                  13, Folio               14, Quarto              15 ,
  10x14               16, 11x17               17, Note                18 ,
  Envelope9           19, Envelope10          20, Envelope11          21 ,
  Envelope12          22, Envelope14          23, Csheet              24 ,
  Dsheet              25, Esheet              26, EnvelopeDL          27 ,
  EnvelopeC5          28, EnvelopeC3          29, EnvelopeC4          30 ,
  EnvelopeC6          31, EnvelopeC65         32, EnvelopeB4          33 ,
  EnvelopeB5          34, EnvelopeB6          35, EnvelopeItaly       36 ,
  EnvelopeMonarch     37, EnvelopePersonal    38, FanfoldUS           39 ,
  FanfoldStdGerman    40, FanfoldLegalGerman  41, User                256

=item PageStart

Start page number.

=item UsePage

Use own start page number (or not).

=item LeftMergin, RightMergin, TopMergin, BottomMergin, HeaderMergin, FooterMergin

Mergins for left, right, top, bottom, header and footer.

=item HCenter

Print in horizontal center (or not)

=item VCenter

Print in vertical center  (or not)

=item Header

Content of print header.
Please refer Excel Help.

=item Footer

Content of print footer.
Please refer Excel Help.

=item PrintGrid

Print with Gridlines (or not)

=item PrintHeaders

Print with headings (or not)

=item NoColor

Print in black-white (or not).

=item Draft

Print in draft mode (or not).

=item Notes

Print with notes (or not).

=item LeftToRight

Print left to right(0) or top to down(1).

=item HPageBreak

Array ref of horizontal page breaks.

=item VPageBreak

Array ref of vertical page breaks.

=item MergedArea

Array ref of merged areas.
Each merged area is : [ I<StartRow>, I<StartColumn>, I<EndRow>, I<EndColumn>]

=back

=head2 Cell

I<Spreadsheet::ParseExcel::Cell>

Cell class has these properties:

=over 4

=item Value

I<Method>
Formatted value of that cell

=item Val

Original Value of that cell

=item Type

Kind of that cell ('Text', 'Numeric', 'Date')

=item Code

Character code of that cell (undef, 'ucs2', '_native_')
undef tells that cell seems to be ascii.
'_native_' tells that cell seems to be 'sjis' or something like that.

=item Format

L<"Format"> for that cell.

=item Merged

That cells is merged (or not).

=item Rich

Array ref of font informations about each characters.

Each entry has : [ I<Start Position>, I<Font Object>]

For more information please refer sample/dmpExR.pl

=back

=head2 Format

I<Spreadsheet::ParseExcel::Format>

Format class has these properties:

=over 4

=item Font

L<"Font"> object for that Format.

=item AlignH

Horizontal Alignment.

  0: (standard), 1: left,       2: center,     3: right,      
  4: fill ,      5: justify,    7:equal_space  

B<Notice:> 6 may be I<merge> but it seems not to work.

=item AlignV

Vertical Alignment.

    0: top,  1: vcenter, 2: bottom, 3: vjustify, 4: vequal_space

=item Indent

Number of indent

=item Wrap

Wrap (or not).

=item Shrink

Display in shrinking (or not)

=item Rotate

In Excel97, 2000      : degrees of string rotation.
In Excel95 or earlier : 0: No rotation, 1: Top down, 2: 90 degrees anti-clockwise, 
                        3: 90 clockwise

=item JustLast

JustLast (or not).
I<I have never seen this attribute.>

=item ReadDir

Direction for read.

=item BdrStyle

Array ref of boder styles : [I<Left>, I<Right>, I<Top>, I<Bottom>]

=item BdrColor

Array ref of boder color indexes : [I<Left>, I<Right>, I<Top>, I<Bottom>]

=item BdrDiag

Array ref of diag boder kind, style and color index : [I<Kind>, I<Style>, I<Color>]
  Kind : 0: None, 1: Right-Down, 2:Right-Up, 3:Both

=item Fill

Array ref of fill pattern and color indexes : [I<Pattern>, I<Front Color>, I<Back Color>]

=item Lock

Locked (or not).

=item Hidden

Hiddedn (or not).

=item Style

Style format (or Cell format)

=back

=head2 Font

I<Spreadsheet::ParseExcel::Font>

Format class has these properties:

=over 4

=item Name

Name of that font.

=item Bold

Bold (or not).

=item Italic

Italic (or not).

=item Height

Size (height) of that font.

=item Underline

Underline (or not).

=item UnderlineStyle

0: None, 1: Single, 2: Double, 0x21: Single(Account), 0x22: Double(Account)

=item Color

Color index for that font.

=item Strikeout

Strikeout (or not).

=item Super

0: None, 1: Upper, 2: Lower

=back

=head1 Formatter class

I<Spreadsheet::ParseExcel::Fmt*>

Formatter class will convert cell data.

Spreadsheet::ParseExcel includes 2 formatter classes: FmtDefault and FmtJapanese. 
You can create your own FmtClass as you like.

Formatter class(Spreadsheet::ParseExcel::Fmt*) should provide these functions:

=over 4

=item ChkType($oSelf, $iNumeric, $iFmtIdx)

tells type of the cell that has specified value.

=over 8

=item $oSelf

Formatter itself

=item $iNumeric

If on, the value seems to be number

=item $iFmtIdx

Format index number of that cell

=back

=item TextFmt($oSelf, $sText, $sCode)

converts original text into applicatable for Value.

=over 8

=item $oSelf

Formatter itself

=item $sText

Original text

=item $sCode

Character code of Original text

=back

=item ValFmt($oSelf, $oCell, $oBook) 

converts original value into applicatable for Value.

=over 8

=item $oSelf

Formatter itself

=item $oCell

Cell object

=item $oBook

Workbook object

=back

=item FmtString($oSelf, $oCell, $oBook)

get format string for the I<$oCell>.

=over 8

=item $oSelf

Formatter itself

=item $oCell

Cell object

=item $oBook

WorkBook object contains that cell

=back

=back

=head1 KNOWN PROBLEM

This module can not get the values of fomulas in 
Excel files made with Spreadsheet::WriteExcel.
Normaly (ie. By Excel application), formula has the result with it.
But Spreadsheet::WriteExcel writes formula with no result.
If you set your Excel application "Auto Calculation" off.
(maybe [Tool]-[Option]-[Calculation] or something)
You will see the same result.

=head1 AUTHOR

Kawai Takanori (Hippo2000) kwitknr@cpan.org

    http://member.nifty.ne.jp/hippo2000/            (Japanese)
    http://member.nifty.ne.jp/hippo2000/index_e.htm (English)

=head1 SEE ALSO

XLHTML, OLE::Storage, Spreadsheet::WriteExcel, OLE::Storage_Lite

This module is based on herbert within OLE::Storage and XLHTML.

=head1 COPYRIGHT

Copyright (c) 2000-2001 Kawai Takanori and Nippon-RAD Co. OP Division
All rights reserved.

You may distribute under the terms of either the GNU General Public
License or the Artistic License, as specified in the Perl README file.

=head1 ACKNOWLEDGEMENTS

First of all, I would like to acknowledge valuable program and modules :
XHTML, OLE::Storage and Spreadsheet::WriteExcel.

In no particular order: Yamaji Haruna, Simamoto Takesi, Noguchi Harumi, 
Ikezawa Kazuhiro, Suwazono Shugo, Hirofumi Morisada, Michael Edwards, Kim Namusk 
and many many people + Kawai Mikako.

=cut
