# Spreadsheet::ParseExcel
#  by Kawai, Takanori (Hippo2000) 2000.10.2
#                                 2001. 2.2 (Ver. 0.15)
# This Program is ALPHA version.
#//////////////////////////////////////////////////////////////////////////////
# Spreadsheet::ParseExcel Objects
#//////////////////////////////////////////////////////////////////////////////
#==============================================================================
# Spreadsheet::ParseExcel::Workbook
#==============================================================================
package Spreadsheet::ParseExcel::Workbook;
require Exporter;
use strict;
use vars qw($VERSION @ISA);
@ISA = qw(Exporter);
sub new($) {
  my $oThis = {};
  bless $oThis;
  return $oThis;
}
#==============================================================================
# Spreadsheet::ParseExcel::Worksheet
#==============================================================================
package Spreadsheet::ParseExcel::Worksheet;
require Exporter;
use strict;
use vars qw($VERSION @ISA);
@ISA = qw(Exporter);
sub new($%) {
  my ($sClass, %rhIni) = @_;
  my $oThis = \%rhIni;
=cmmnt
        {   Name        => $rhIni{Name},
            Kind        => $rhIni{Kind},
            _Pos        => $rhIni{_Pos},
            Cells       => undef,
        };
=cut
  bless $oThis;
  $oThis->{Cells}=undef;
  $oThis->{DefColWidth}=8.38;
  return $oThis;
}
#==============================================================================
# Spreadsheet::ParseExcel::Font
#==============================================================================
package Spreadsheet::ParseExcel::Font;
require Exporter;
use strict;
use vars qw($VERSION @ISA);
@ISA = qw(Exporter);
sub new($%) {
  my($sClass, %rhIni) = @_;
  my $oThis = { 
        Height    => $rhIni{Height},
        Attr      => $rhIni{Attr},
        CIdx      => $rhIni{CIdx},
        Bold      => $rhIni{Bold},
        Super     => $rhIni{Super},
        UnderLine => $rhIni{UnderLine},
        Name      => $rhIni{Name},
        NameCode  => $rhIni{NameCode},
    };
  bless $oThis;
  return $oThis;
}
#==============================================================================
# Spreadsheet::ParseExcel::Format
#==============================================================================
package Spreadsheet::ParseExcel::Format;
require Exporter;
use strict;
use vars qw($VERSION @ISA);
@ISA = qw(Exporter);
sub new($%) {
  my($sClass, %rhIni) = @_;
  my $oThis = { 
        FontNo   => $rhIni{FontNo},
        Font     => $rhIni{Font},
        FmtIdx   => $rhIni{FmtIdx},
        Gen      => $rhIni{Gen},
        Align    => $rhIni{Align},
        BdrStyle => $rhIni{BdrStyle},
        BdrLClr  => $rhIni{BdrLClr},
        BdrRClr  => $rhIni{BdrRClr},
        BdrTClr  => $rhIni{BdrTClr},
        BdrBClr  => $rhIni{BdrBClr},
    };
  bless $oThis;
  return $oThis;
}
#==============================================================================
# Spreadsheet::ParseExcel::Cell
#==============================================================================
package Spreadsheet::ParseExcel::Cell;
require Exporter;
use strict;
use vars qw($VERSION @ISA);
@ISA = qw(Exporter);

sub new($%) {
    my($sPkg, %rhKey)=@_;
    my($sWk, $iLen);
    my $oThis = {
        Val     => $rhKey{Val},
        Format  => $rhKey{Format},
        Code    => $rhKey{Code},
        Type    => $rhKey{Type},
    };
    bless $oThis;
    return $oThis;
}
sub Value($){
    my($oThis)=@_;
    return $oThis->{_Value};
}
#==============================================================================
# Spreadsheet::ParseExcel
#==============================================================================
package Spreadsheet::ParseExcel;
require Exporter;
use strict;
use OLE::Storage_Lite;
use vars qw($VERSION @ISA );
@ISA = qw(Exporter);
$VERSION = '0.16'; # 
my $oFmtClass;
my @aColor =
(
    '000000',   # 0x00
    'FFFFFF', 'FFFFFF', 'FFFFFF', 'FFFFFF',
    'FFFFFF', 'FFFFFF', 'FFFFFF', 'FFFFFF', #0x08 - This one's Black, too ???
    'FFFFFF', 'FF0000', '00FF00', '0000FF',
    'FFFF00', 'FF00FF', '00FFFF', '800000', # 0x10
    '008000', '000080', '808000', '800080',
    '008080', 'C0C0C0', '808080', '9999FF', # 0x18
    '993366', 'FFFFCC', 'CCFFFF', '660066',
    'FF8080', '0066CC', 'CCCCFF', '000080', # 0x20
    'FF00FF', 'FFFF00', '00FFFF', '800080',
    '800000', '008080', '0000FF', '00CCFF', # 0x28
    'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF',
    'FF99CC', 'CC99FF', 'FFCC99', '3366FF', # 0x30
    '33CCCC', '99CC00', 'FFCC00', 'FF9900',
    'FF6600', '666699', '969696', '003366', # 0x38
    '339966', '003300', '333300', '993300',
    '993366', '333399', '333333', 'FFFFFF'  # 0x40
);
use constant verExcel95 => 0x500;
use constant verExcel97 =>0x600;
use constant verBIFF2 =>0x00;
use constant verBIFF3 =>0x02;
use constant verBIFF4 =>0x04;
use constant verBIFF5 =>0x08;
use constant verBIFF8 =>0x18;   #Added (Not in BOOK)

my %ProcTbl =(
#Develpers' Kit P292
    0x22    => \&_subFlg1904,           # 1904 Flag
    0x3C => \&_subContinue,             # Continue
    0x43    => \&_subXF,                # ExTended Format(?)
#Develpers' Kit P292
    0x55   =>\&_subDefColWidth,         # Consider
    0x5C    => \&_subWriteAccess,          # WRITEACCESS
    0x7D    => \&_subColInfo,           # Colinfo
    0x7E    => \&_subRK,                # RK
    0x85    => \&_subBoundSheet,        # BoundSheet

    0x99    => \&_subStandardWidth,     # Standard Col
#Develpers' Kit P293
    0xBD    => \&_subMulRK,             # MULRK
    0xBE    => \&_subMulBlank,          # MULBLANK
    0xD6    => \&_subRString,           # RString
#Develpers' Kit P294
    0xE0    => \&_subXF,                # ExTended Format
    0xFC    => \&_subSST,               # Shared String Table
    0xFD    => \&_subLabelSST,          # Label SST
#Develpers' Kit P295
    0x201   => \&_subBlank,             # Blank

    0x202   => \&_subInteger,           # Integer(Not Documented)
    0x203   => \&_subNumber,            # Number
    0x204   => \&_subLabel ,            # Label
    0x205   => \&_subBoolErr,           # BoolErr
    0x207   => \&_subString,            # STRING
    0x208   => \&_subRow,               # RowData
    0x221   => \&_subArray,             #Array (Consider)
    0x225   => \&_subDefaultRowHeight,  # Consider

    0x31    => \&_subFont,              # Font
    0x231   => \&_subFont,              # Font

    0x27E   => \&_subRK,                # RK
    0x41E   => \&_subFormat,            # Format

    0x06    => \&_subFormula,           # Formula
    0x406   => \&_subFormula,           # Formula

    0x09    => \&_subBOF,               # BOF(BIFF2)
    0x209   => \&_subBOF,               # BOF(BIFF3)
    0x409   => \&_subBOF,               # BOF(BIFF4)
    0x809   => \&_subBOF,               # BOF(BIFF5-8)
    );

my $BIGENDIAN;
my $PREFUNC;
#------------------------------------------------------------------------------
# Spreadsheet::ParseExcel->new
#------------------------------------------------------------------------------
sub new($;%) {
    my ($sPkg, %hParam) =@_;

#0. Check ENDIAN(Little: Interl etc. BIG: Sparc etc)
    $BIGENDIAN = (defined $hParam{Endian})? $hParam{Endian} :
                    (unpack("H08", pack("L", 2)) eq '02000000')? 0: 1;
    my $oThis = { };
    bless $oThis;

#1. Set Parameter
#1.1 Get Content
    $oThis->{GetContent} = \&_subGetContent;

#1.2 Set Event Handler
    if($hParam{EventHandlers}) {
        $oThis->SetEventHandlers($hParam{EventHandlers});
    }
    else {
        $oThis->SetEventHandlers(\%ProcTbl);
    }
    if($hParam{AddHandlers}) {
        foreach my $sKey (keys(%{$hParam{AddHandlers}})) {
            $oThis->SetEventHandler($sKey, $hParam{AddHandlers}->{$sKey});
        }
    }
    return $oThis;
}
#------------------------------------------------------------------------------
# Spreadsheet::ParseExcel->SetEventHandler
#------------------------------------------------------------------------------
sub SetEventHandler($$\&) {
    my($oThis, $sKey, $oFunc) = @_;
    $oThis->{FuncTbl}->{$sKey} = $oFunc;
}
#------------------------------------------------------------------------------
# Spreadsheet::ParseExcel->SetEventHandlers
#------------------------------------------------------------------------------
sub SetEventHandlers($$) {
    my($oThis, $rhTbl) = @_;
    $oThis->{FuncTbl} = undef;
    foreach my $sKey (keys %$rhTbl) {
        $oThis->{FuncTbl}->{$sKey} = $rhTbl->{$sKey};
    }
}
#------------------------------------------------------------------------------
# Spreadsheet::ParseExcel->Parse
#------------------------------------------------------------------------------
sub Parse($$;$) {
    my($oThis, $sFile, $oWkFmt)=@_;
    my($sWk, $bLen);

#0. New $oBook
    my $oBook = Spreadsheet::ParseExcel::Workbook->new;
    $oBook->{SheetCount} = 0;

#1.Get content
    my($sBIFF, $iLen);
    if(ref($sFile) eq "SCALAR") {
#1.1 Specified by Buffer
        $sBIFF = $$sFile;
        $iLen  = length($sBIFF);
    }
    elsif(ref($sFile)) {
#1.2 Specified by Other Things(HASH reference etc)
        return undef;
    }
    else {
#1.3 Specified by File name
        $oBook->{File} = $sFile;
        return undef unless (-e $sFile);
        ($sBIFF, $iLen) = $oThis->{GetContent}->($sFile);
        return undef unless($sBIFF);
    }

#2. Ready for format
    if ($oWkFmt) {
        $oFmtClass = $oWkFmt;
    }
    else {
        require Spreadsheet::ParseExcel::FmtDefault;
        $oFmtClass = new Spreadsheet::ParseExcel::FmtDefault;
    }

#3. Parse content
    my $lPos = 0;
    $sWk = substr($sBIFF, $lPos, 4);
    $lPos += 4;
    my $iEfFlg = 0;
    while($lPos<=$iLen) {
        my($bOp, $bLen) = unpack("v2", $sWk);
       if($bLen) {
            $sWk = substr($sBIFF, $lPos, $bLen);
            $lPos += $bLen;
        }
#printf STDERR "%4X:%s\n", $bOp, 'UNDEFIND---:' . unpack("H*", $sWk) unless($NameTbl{$bOp});
    #Check EF, EOF
    if($bOp == 0xEF) {    #EF
            $iEfFlg = $bOp;
    }
    elsif($bOp == 0x0A) { #EOF
            undef $iEfFlg;
    }
    unless($iEfFlg) {
        if(defined $oThis->{FuncTbl}->{$bOp}) {
                $oThis->{FuncTbl}->{$bOp}->($oBook, $bOp, $bLen, $sWk);
            }
            $PREFUNC = $bOp if ($bOp != 0x3C); #Not Continue 
    }
        $sWk = substr($sBIFF, $lPos, 4) if(($lPos+4) <= $iLen);
        $lPos += 4;
    }
#4.return $oBook
    return $oBook;
}
#------------------------------------------------------------------------------
# _subGetContent (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub _subGetContent($)
{
    my($sFile)=@_;
    my $oOl = OLE::Storage_Lite->new($sFile);
    return (undef, undef) unless($oOl);
    my @aRes = $oOl->getPpsSearch(
            [OLE::Storage_Lite::Asc2Ucs('Book'), 
             OLE::Storage_Lite::Asc2Ucs('Workbook')], 1, 1);
    return (undef, undef) if($#aRes < 0);
#Hack from Herbert
    unless($aRes[0]->{Data}) {
        #Same as OLE::Storage_Lite
        my $oIo;
        #1. $sFile is Ref of scalar
        if(ref($sFile) eq 'SCALAR') {
            $oIo = new IO::Scalar;
            $oIo->open($sFile);
        }
        #2. $sFile is a IO::Handle object
        elsif(UNIVERSAL::isa($sFile, 'IO::Handle')) {
            $oIo = $sFile;
            binmode($oIo);
        }
        #3. $sFile is a simple filename string
        elsif(!ref($sFile)) {
            $oIo = new IO::File;
            $oIo->open("<$sFile") || return undef;
            binmode($oIo);
        }
        my $sWk;
        my $sBuff ='';

        while($oIo->read($sWk, 4096)) { #4_096 has no special meanings
            $sBuff .= $sWk;
        }
        $oIo->close();
        #Not Excel file (simple method)
        return (undef, undef) if (substr($sBuff, 0, 1) ne "\x09");
        return ($sBuff, length($sBuff));
    }
    else {
        return ($aRes[0]->{Data}, length($aRes[0]->{Data}));
    }
}
#------------------------------------------------------------------------------
# _subBOF (for Spreadsheet::ParseExcel) Developers' Kit : P303
#------------------------------------------------------------------------------
sub _subBOF($$$$)
{
    my($oBook, $bOp, $bLen, $sWk) = @_;

    if(defined $oBook->{_CurSheet}) {
        $oBook->{_CurSheet}++; 
        ($oBook->{Worksheet}[$oBook->{_CurSheet}]->{SheetVersion},
         $oBook->{Worksheet}[$oBook->{_CurSheet}]->{SheetType},) 
                = unpack("v2", $sWk) if(length($sWk) > 4);
    }
    else {
        $oBook->{BIFFVersion} = int($bOp / 0x100);
        if (($oBook->{BIFFVersion} == verBIFF2) ||
            ($oBook->{BIFFVersion} == verBIFF3) ||
            ($oBook->{BIFFVersion} == verBIFF4)) {
            $oBook->{Version} = $oBook->{BIFFVersion};
        }
        else {
            $oBook->{Version} = unpack("v", $sWk);
            $oBook->{BIFFVersion} = 
                ($oBook->{Version}==verExcel95)? verBIFF5:verBIFF8;
        }
        $oBook->{_CurSheet} = -1;
    }
}
#------------------------------------------------------------------------------
# _subBlank (for Spreadsheet::ParseExcel) DK:P303
#------------------------------------------------------------------------------
sub _subBlank($$$$)
{
    my($oBook, $bOp, $bLen, $sWk) = @_;
    my ($iR, $iC, $iF) = unpack("v3", $sWk);
    $oBook->{Worksheet}[$oBook->{_CurSheet}]->{Cells}[$iR][$iC] = 
        _NewCell(
            Kind    => 'BLANK',
            Val     => '',
            Format  => $oBook->{Format}[$iF],
            Numeric => 0,
            Code    => undef,
            Book    => $oBook,
        );
#2.MaxRow, MaxCol, MinRow, MinCol
    _SetDimension($oBook, $iR, $iC, $iC);
}
#------------------------------------------------------------------------------
# _subInteger (for Spreadsheet::ParseExcel) Not in DK
#------------------------------------------------------------------------------
sub _subInteger($$$$)
{
    my($oBook, $bOp, $bLen, $sWk) = @_;
    my($iR, $iC, $iF, $sTxt, $sDum);

    ($iR, $iC, $iF, $sDum, $sTxt) = unpack("v3cv", $sWk);
    $oBook->{Worksheet}[$oBook->{_CurSheet}]->{Cells}[$iR][$iC] = 
            _NewCell (
                Kind    => 'INTEGER',
                Val     => $sTxt,
                Format  => $oBook->{Format}[$iF],
                Numeric => 0,
                Code    => undef,
                Book    => $oBook,
            );
#2.MaxRow, MaxCol, MinRow, MinCol
    _SetDimension($oBook, $iR, $iC, $iC);
}
#------------------------------------------------------------------------------
# _subNumber (for Spreadsheet::ParseExcel)  : DK: P354
#------------------------------------------------------------------------------
sub _subNumber($$$$)
{
    my($oBook, $bOp, $bLen, $sWk) = @_;

    my ($iR, $iC, $iF) = unpack("v3", $sWk);
    my $dVal = _convDval(substr($sWk, 6, 8));
    $oBook->{Worksheet}[$oBook->{_CurSheet}]->{Cells}[$iR][$iC] = 
            _NewCell (
                Kind    => 'Number',
                Val     => $dVal,
                Format  => $oBook->{Format}[$iF],
                Numeric => 1,
                Code    => undef,
                Book    => $oBook,
            );
#2.MaxRow, MaxCol, MinRow, MinCol
    _SetDimension($oBook, $iR, $iC, $iC);
}
#------------------------------------------------------------------------------
# _convDval (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub _convDval($) {
    my($sWk)=@_;
    return  unpack("d", ($BIGENDIAN)? 
                    pack("c8", reverse(unpack("c8", $sWk))) : $sWk);
}
#------------------------------------------------------------------------------
# _subRString (for Spreadsheet::ParseExcel) DK:P405
#------------------------------------------------------------------------------
sub _subRString($$$$)
{
    my($oBook, $bOp, $bLen, $sWk) = @_;
    my($iR, $iC, $iF, $iL, $sTxt);
    ($iR, $iC, $iF, $iL) = unpack("v4", $sWk);
    $sTxt = substr($sWk, 8, $iL);
    $oBook->{Worksheet}[$oBook->{_CurSheet}]->{Cells}[$iR][$iC] = 
        _NewCell (
            Kind    => 'RString',
            Val     => $sTxt,
            Format  => $oBook->{Format}[$iF],
            Numeric => 0,
            Code    => undef,
            Book    => $oBook,
        );
    #Has STRUN
    if(length($sWk) > (8+$iL)) {
        $oBook->{Worksheet}[$oBook->{_CurSheet}]->{Cells}[$iR][$iC]->{STRUN} = 
            substr($sWk, (8+$iL)+1);
    }
#2.MaxRow, MaxCol, MinRow, MinCol
    _SetDimension($oBook, $iR, $iC, $iC);
}
#------------------------------------------------------------------------------
# _subBoolErr (for Spreadsheet::ParseExcel) DK:P306
#------------------------------------------------------------------------------
sub _subBoolErr($$$$)
{
    my($oBook, $bOp, $bLen, $sWk) = @_;
    my ($iR, $iC, $iF) = unpack("v3", $sWk);
    my ($iVal, $iFlg) = unpack("cc", substr($sWk, 6, 2));
    my $sTxt = DecodeBoolErr($iVal, $iFlg);

    $oBook->{Worksheet}[$oBook->{_CurSheet}]->{Cells}[$iR][$iC] = 
        _NewCell (
            Kind    => 'BoolError',
            Val     => $sTxt,
            Format  => $oBook->{Format}[$iF],
            Numeric => 0,
            Code    => undef,
            Book    => $oBook,
        );
#2.MaxRow, MaxCol, MinRow, MinCol
    _SetDimension($oBook, $iR, $iC, $iC);
}
#------------------------------------------------------------------------------
# _subRK (for Spreadsheet::ParseExcel)  DK:P401
#------------------------------------------------------------------------------
sub _subRK($$$$)
{
    my($oBook, $bOp, $bLen, $sWk) = @_;
    my ($iR, $iC) = unpack("v3", $sWk);

    my($iF, $sTxt)= _UnpackRKRec(substr($sWk, 4, 6));
    $oBook->{Worksheet}[$oBook->{_CurSheet}]->{Cells}[$iR][$iC] = 
        _NewCell (
            Kind    => 'RK',
            Val     => $sTxt,
            Format  => $oBook->{Format}[$iF],
            Numeric => 1,
            Code    => undef,
            Book    => $oBook,
        );
#2.MaxRow, MaxCol, MinRow, MinCol
    _SetDimension($oBook, $iR, $iC, $iC);
}
#------------------------------------------------------------------------------
# _subArray (for Spreadsheet::ParseExcel)   DK:P297
#------------------------------------------------------------------------------
sub _subArray($$$$)
{
    my($oBook, $bOp, $bLen, $sWk) = @_;
    my ($iBR, $iER, $iBC, $iEC) = unpack("v2c2", $sWk);
    
}
#------------------------------------------------------------------------------
# _subFormula (for Spreadsheet::ParseExcel) DK:P336
#------------------------------------------------------------------------------
sub _subFormula($$$$)
{
    my($oBook, $bOp, $bLen, $sWk) = @_;
    my($iR, $iC, $iF) = unpack("v3", $sWk);

    my ($iFlg) = unpack("v", substr($sWk,12,2));
    if($iFlg == 0xFFFF) {
        my($iKind) = unpack("c", substr($sWk, 6, 1));
        my($iVal)  = unpack("c", substr($sWk, 8, 1));

        if(($iKind==1) or ($iKind==2)) {
            my $sTxt = ($iKind == 1)? DecodeBoolErr($iVal, 0):DecodeBoolErr($iVal, 1);
            $oBook->{Worksheet}[$oBook->{_CurSheet}]->{Cells}[$iR][$iC] = 
                _NewCell (
                    Kind    => 'Formulra Bool',
                    Val     => $sTxt,
                    Format  => $oBook->{Format}[$iF],
                    Numeric => 0,
                    Code    => undef,
                    Book    => $oBook,
                );
        }
        else { # Result
            $oBook->{Worksheet}[$oBook->{_CurSheet}]->{Cells}[$iR][$iC] = 
                _NewCell (
                    Kind    => 'Formula String',
                    Val     => '',
                    Format  => $oBook->{Format}[$iF],
                    Numeric => 0,
                    Code    => undef,
                    Book    => $oBook,
                );
            $oBook->{_PrevPos} = [$iR, $iC, $iF];
        }
    }
    else {
        my $dVal = _convDval(substr($sWk, 6, 8));
        $oBook->{Worksheet}[$oBook->{_CurSheet}]->{Cells}[$iR][$iC] = 
            _NewCell (
                Kind    => 'Formula Number',
                Val     => $dVal,
                Format  => $oBook->{Format}[$iF],
                Numeric => 1,
                Code    => undef,
                Book    => $oBook,
            );
    }
#2.MaxRow, MaxCol, MinRow, MinCol
    _SetDimension($oBook, $iR, $iC, $iC);
}
#------------------------------------------------------------------------------
# _subString (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub _subString($$$$)
{
    my($oBook, $bOp, $bLen, $sWk) = @_;
#Position (not enough for ARRAY)

    my $iPos = $oBook->{_PrevPos};
    return undef unless($iPos);
    $oBook->{_PrevPos} = undef;
    my ($iR, $iC, $iF) = @$iPos;

    my ($iLen, $sTxt, $iCode);
    if($oBook->{BIFFVersion} == verBIFF8) {
        my( $raBuff, $iLen) = _convBIFF8String($sWk);
        $sTxt  = $raBuff->[0];
        $iCode = $raBuff->[1];
    }
    elsif($oBook->{BIFFVersion} == verBIFF5) {
        $iCode = 0;
        $iLen = unpack("v", $sWk);
        $sTxt = substr($sWk, 2, $iLen);
    }
    else {
        $iCode = 0;
        $iLen = unpack("c", $sWk);
        $sTxt = substr($sWk, 1, $iLen);
    }
    $oBook->{Worksheet}[$oBook->{_CurSheet}]->{Cells}[$iR][$iC] = 
        _NewCell (
            Kind    => 'String',
            Val     => $sTxt,
            Format  => $oBook->{Format}[$iF],
            Numeric => 0,
            Code    => ($iCode)? 'ucs2': '_native_',
            Book    => $oBook,
        );
#2.MaxRow, MaxCol, MinRow, MinCol
    _SetDimension($oBook, $iR, $iC, $iC);
}
#------------------------------------------------------------------------------
# _subLabel (for Spreadsheet::ParseExcel)   DK:P344
#------------------------------------------------------------------------------
sub _subLabel($$$$)
{
    my($oBook, $bOp, $bLen, $sWk) = @_;
    my($iR, $iC, $iF) = unpack("v3", $sWk);
    $oBook->{Worksheet}[$oBook->{_CurSheet}]->{Cells}[$iR][$iC] = 
        _NewCell ( 
            Kind    => 'Label',
            Val     => substr($sWk, 8),
            Format  => $oBook->{Format}[$iF],
            Numeric => 0,
            Code    => '_native_',
            Book    => $oBook,
        );
#2.MaxRow, MaxCol, MinRow, MinCol
    _SetDimension($oBook, $iR, $iC, $iC);
}
#------------------------------------------------------------------------------
# _subMulRK (for Spreadsheet::ParseExcel)   DK:P349
#------------------------------------------------------------------------------
sub _subMulRK($$$$)
{
    my($oBook, $bOp, $bLen, $sWk) = @_;
    return if ($oBook->{SheetCount}<=0);

    my ($iR, $iSc) = unpack("v2", $sWk);
    my $iEc = unpack("v", substr($sWk, length($sWk) -2, 2));

    my $iPos = 4;
    for(my $iC=$iSc; $iC<=$iEc; $iC++) {
        my($iF, $lVal) = _UnpackRKRec(substr($sWk, $iPos, 6), $iR, $iC);
        $oBook->{Worksheet}[$oBook->{_CurSheet}]->{Cells}[$iR][$iC] = 
            _NewCell (
                Kind    => 'MulRK',
                Val     => $lVal,
                Format  => $oBook->{Format}[$iF],
                Numeric => 1,
                Code => undef,
                Book    => $oBook,
                );
        $iPos += 6;
    }
#2.MaxRow, MaxCol, MinRow, MinCol
    _SetDimension($oBook, $iR, $iSc, $iEc);
}
#------------------------------------------------------------------------------
# _subMulBlank (for Spreadsheet::ParseExcel) DK:P349
#------------------------------------------------------------------------------
sub _subMulBlank($$$$)
{
    my($oBook, $bOp, $bLen, $sWk) = @_;
    my ($iR, $iSc) = unpack("v2", $sWk);
    my $iEc = unpack("v", substr($sWk, length($sWk)-2, 2));
    my $iPos = 4;
    for(my $iC=$iSc; $iC<=$iEc; $iC++) {
        my $iF = unpack('v', substr($sWk, $iPos, 2));
        $oBook->{Worksheet}[$oBook->{_CurSheet}]->{Cells}[$iR][$iC] = 
            _NewCell (
                Kind    => 'MulBlank',
                Val     => '',
                Format  => $oBook->{Format}[$iF],
                Numeric => 0,
                Code    => undef,
                Book    => $oBook,
            );
        $iPos+=2;
    }
#2.MaxRow, MaxCol, MinRow, MinCol
    _SetDimension($oBook, $iR, $iSc, $iEc);
}
#------------------------------------------------------------------------------
# _subLabelSST (for Spreadsheet::ParseExcel) DK: P345
#------------------------------------------------------------------------------
sub _subLabelSST($$$$)
{
    my($oBook, $bOp, $bLen, $sWk) = @_;
    my ($iR, $iC, $iF, $iIdx) = unpack('v3V', $sWk);

    $oBook->{Worksheet}[$oBook->{_CurSheet}]->{Cells}[$iR][$iC] = 
        _NewCell (
            Kind    => 'PackedIdx',
            Val     => $oBook->{PkgStr}[$iIdx]->{Text},
            Format  => $oBook->{Format}[$iF],
            Numeric => 0,
            Code    => ($oBook->{PkgStr}[$iIdx]->{Unicode})? 'ucs2': undef,
            Book    => $oBook,
        );

#2.MaxRow, MaxCol, MinRow, MinCol
    _SetDimension($oBook, $iR, $iC, $iC);
}
#------------------------------------------------------------------------------
# _subFlg1904 (for Spreadsheet::ParseExcel) DK:P296
#------------------------------------------------------------------------------
sub _subFlg1904($$$$)
{
    my($oBook, $bOp, $bLen, $sWk) = @_;
    $oBook->{Flg1904} = unpack("v", $sWk);
}
#------------------------------------------------------------------------------
# _subRow (for Spreadsheet::ParseExcel) DK:P403
#------------------------------------------------------------------------------
sub _subRow($$$$)
{
    my($oBook, $bOp, $bLen, $sWk) = @_;
#0. Get Worksheet info (MaxRow, MaxCol, MinRow, MinCol)
    my($iR, $iSc, $iEc, $iHght, $iXf) = unpack("v5", $sWk);
    $iEc--;

#1. RowHeight
    $oBook->{Worksheet}[$oBook->{_CurSheet}]->{RowHeight}[$iR] = $iHght/20.0;

#2.MaxRow, MaxCol, MinRow, MinCol
    _SetDimension($oBook, $iR, $iSc, $iEc);
}
#------------------------------------------------------------------------------
# _SetDimension (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub _SetDimension(($$$$))
{
    my($oBook, $iR, $iSc, $iEc)=@_;
#2.MaxRow, MaxCol, MinRow, MinCol
#2.1 MinRow
    $oBook->{Worksheet}[$oBook->{_CurSheet}]->{MinRow} = $iR 
        unless (defined $oBook->{Worksheet}[$oBook->{_CurSheet}]->{MinRow}) and 
               ($oBook->{Worksheet}[$oBook->{_CurSheet}]->{MinRow} <= $iR);

#2.2 MaxRow
    $oBook->{Worksheet}[$oBook->{_CurSheet}]->{MaxRow} = $iR 
        unless (defined $oBook->{Worksheet}[$oBook->{_CurSheet}]->{MaxRow}) and
               ($oBook->{Worksheet}[$oBook->{_CurSheet}]->{MaxRow} > $iR);
#2.3 MinCol
    $oBook->{Worksheet}[$oBook->{_CurSheet}]->{MinCol} = $iSc
            unless (defined $oBook->{Worksheet}[$oBook->{_CurSheet}]->{MinCol}) and
               ($oBook->{Worksheet}[$oBook->{_CurSheet}]->{MinCol} <= $iSc);
#2.4 MaxCol
    $oBook->{Worksheet}[$oBook->{_CurSheet}]->{MaxCol} = $iEc 
            unless (defined $oBook->{Worksheet}[$oBook->{_CurSheet}]->{MaxCol}) and
               ($oBook->{Worksheet}[$oBook->{_CurSheet}]->{MaxCol} > $iEc);

}
#------------------------------------------------------------------------------
# _subDefaultRowHeight (for Spreadsheet::ParseExcel)    DK: P318
#------------------------------------------------------------------------------
sub _subDefaultRowHeight($$$$)
{
    my($oBook, $bOp, $bLen, $sWk) = @_;
#1. RowHeight
    my($iDum, $iHght) = unpack("v2", $sWk);
    $oBook->{Worksheet}[$oBook->{_CurSheet}]->{DefRowHeight} = $iHght/20;

}
#------------------------------------------------------------------------------
# _subStandardWidth(for Spreadsheet::ParseExcel)    DK:P413
#------------------------------------------------------------------------------
sub _subStandardWidth($$$$)
{
    my($oBook, $bOp, $bLen, $sWk) = @_;
    my $iW = unpack("v", $sWk);
    $oBook->{StandardWidth}= _adjustColWidth($iW);
}
#------------------------------------------------------------------------------
# _subDefColWidth(for Spreadsheet::ParseExcel)      DK:P319
#------------------------------------------------------------------------------
sub _subDefColWidth($$$$)
{
    my($oBook, $bOp, $bLen, $sWk) = @_;
    my $iW = unpack("v", $sWk);
    $oBook->{Worksheet}[$oBook->{_CurSheet}]->{DefColWidth}= _adjustColWidth($iW);
}
#------------------------------------------------------------------------------
# _subColInfo (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub _adjustColWidth($) {
    my($iW)=@_;
    return ($iW -0xA0)/256;
}
#------------------------------------------------------------------------------
# _subColInfo (for Spreadsheet::ParseExcel) DK:P309
#------------------------------------------------------------------------------
sub _subColInfo($$$$)
{
    my($oBook, $bOp, $bLen, $sWk) = @_;
    my($iSc, $iEc, $iW, $iXF, $iGr) = unpack("v5", $sWk);
    for(my $i= $iSc; $i<=$iEc; $i++) {
        $oBook->{Worksheet}[$oBook->{_CurSheet}]->{ColWidth}[$i] = _adjustColWidth($iW);
        $oBook->{Worksheet}[$oBook->{_CurSheet}]->{ColFmtNo}[$i] = $iXF;
        # $oBook->{Worksheet}[$oBook->{_CurSheet}]->{ColCr}[$i]    = $iGr; #Not Implemented
    }
}
#------------------------------------------------------------------------------
# _subSST (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub _subSST($$$$)
{
    my($oBook, $bOp, $bLen, $sWk) = @_;
    _subStrWk($oBook, substr($sWk, 8));
}
#------------------------------------------------------------------------------
# _subContinue (for Spreadsheet::ParseExcel)    DK:P311
#------------------------------------------------------------------------------
sub _subContinue($$$$)
{
    my($oBook, $bOp, $bLen, $sWk) = @_;
=cmmt
    if(defined $oThis->{FuncTbl}->{$bOp}) {
        $oThis->{FuncTbl}->{$PREFUNC}->($oBook, $bOp, $bLen, $sWk);
    }
=cut
    _subStrWk($oBook, $sWk, 1) if($PREFUNC == 0xFC);
}
#------------------------------------------------------------------------------
# _subWriteAccess (for Spreadsheet::ParseExcel) DK:P451
#------------------------------------------------------------------------------
sub _subWriteAccess($$$$)
{
    my($oBook, $bOp, $bLen, $sWk) = @_;
    return if (defined $oBook->{_Author});

    #BIFF8
    if($oBook->{BIFFVersion} >= verBIFF8) {
        $oBook->{Author} = _convBIFF8String($sWk);
    }
    #Before BIFF8
    else {
        my($iLen) = unpack("c", $sWk);
        $oBook->{Author} = $oFmtClass->TextFmt(substr($sWk, 1, $iLen), '_native_');
    }
}
#------------------------------------------------------------------------------
# _convBIFF8String (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub _convBIFF8String($;$){
    my($sWk, $iCnvFlg) = @_;
    my($iLen, $iFlg) = unpack("vc", $sWk);
    my($iHigh, $iExt, $iRich) = ($iFlg & 0x01, $iFlg & 0x04, $iFlg & 0x08);
    my($iStPos, $iExtCnt, $iRichCnt, $sStr);
#2. Rich and Ext
    if($iRich && $iExt) {
        $iStPos   = 9;
        ($iRichCnt, $iExtCnt) = unpack('vV', substr($sWk, 3, 6));
    }
    elsif($iRich) { #Only Rich
        $iStPos   = 5;
        $iRichCnt = unpack('v', substr($sWk, 3, 2));
        $iExtCnt  = 0;
    }
    elsif($iExt)  { #Only Ext
        $iStPos   = 7;
        $iRichCnt = 0;
        $iExtCnt  = unpack('V', substr($sWk, 3, 4));
    }
    else {          #Nothing Special
        $iStPos   = 3;
        $iExtCnt  = 0;
        $iRichCnt = 0;
    }
#3.Get String
    if($iHigh) {    #Compressed
        $iLen *= 2;
        $sStr = substr($sWk,    $iStPos, $iLen);
        _SwapForUnicode(\$sStr);
        $sStr = $oFmtClass->TextFmt($sStr, 'ucs2') unless($iCnvFlg);
    }
    else {              #Not Compressed
        $sStr = substr($sWk, $iStPos, $iLen);
    }

#4. return 
    if(wantarray) {
        #4.1 Get Rich and Ext
        if(length($sWk) < $iStPos + $iLen+ $iRichCnt*4+$iExtCnt) {
            return ([undef, $iHigh, undef, undef], 
                $iStPos + $iLen+ $iRichCnt*4+$iExtCnt, $iStPos, $iLen);
        }
        else {
            return ([$sStr, $iHigh,
                    substr($sWk, $iStPos + $iLen, $iRichCnt*4),
                    substr($sWk, $iStPos + $iLen+ $iRichCnt*4, $iExtCnt)], 
                $iStPos + $iLen+ $iRichCnt*4+$iExtCnt,  $iStPos, $iLen);
        }
    }
    else {
        return $sStr;
    }
}
#------------------------------------------------------------------------------
# _subXF (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub _subXF($$$$)
{
    my($oBook, $bOp, $bLen, $sWk) = @_;
    my($iFnt, $iIdx, $iGen, $iAlign, $iIndent, $iDum);
    my($iBdrStyle, $iBdrLClr, $iBdrTClr, $iCellColor);

    ($iFnt, $iIdx, $iGen, $iAlign, $iIndent) = unpack("v5", $sWk);
    my $sBdrL;
    my $sBdrT;
    my $sBitC;
 
    if($oBook->{BIFFVersion} == verBIFF8) {
        ($iBdrStyle, $iBdrLClr, $iBdrTClr, $iCellColor)
                = unpack("vvVv", substr($sWk, 10, 10));
        $sBdrL = ($iBdrLClr==0)? '0'x16 : unpack("B16", $iBdrLClr);
        $sBdrT = ($iBdrTClr==0)? '0'x16 : unpack("B16", $iBdrTClr);
        $sBitC = ($iCellColor==0)? '0'x16: unpack("B16", $iCellColor);
    }
    elsif($oBook->{BIFFVersion} == verBIFF5) {
        $iBdrLClr = 0;
        $iBdrTClr = 0;
        ($iCellColor, $iBdrTClr, $iBdrStyle) = unpack("vvv", substr($sWk, 12, 6));
#madamada
        $sBdrL = ($iBdrLClr==0)? '0'x16 : unpack("B16", $iBdrLClr);
        $sBdrT = ($iBdrTClr==0)? '0'x16 : unpack("B16", $iBdrTClr);
        $sBitC = ($iCellColor==0)? '0'x16: unpack("B16", $iCellColor);
    }

   push @{$oBook->{Format}} , 
         Spreadsheet::ParseExcel::Format->new (
            FontNo   => $iFnt,
            Font     => ($iFnt == 0)? $oBook->{Font}[0] : $oBook->{Font}[$iFnt-1],
            FmtIdx   => $iIdx,
            Gen      => $iGen,
            Align    => $iAlign,
            Indent   => $iIndent,
            BackColor => ord(pack("B8", '0'. substr($sBitC, 9, 7))),
            ForeColor => ord(pack("B8", '0'. substr($sBitC, 2, 7))),
            BdrStyle => $iBdrStyle,
            BdrLClr  => $aColor[ord(pack("B8", '0'. substr($sBdrL, 9, 7)))],
            BdrRClr  => $aColor[ord(pack("B8", '0'. substr($sBdrL, 2, 7)))],
            BdrTClr  => $aColor[ord(pack("B8", '0'. substr($sBdrT, 9, 7)))],
            BdrBClr  => $aColor[ord(pack("B8", '0'. substr($sBdrT, 2, 7)))],
        );
}
#------------------------------------------------------------------------------
# _subFormat (for Spreadsheet::ParseExcel)  DK: P336
#------------------------------------------------------------------------------
sub _subFormat($$$$)
{
    my($oBook, $bOp, $bLen, $sWk) = @_;
    my $sFmt;
    if (($oBook->{BIFFVersion} == verBIFF2) ||
        ($oBook->{BIFFVersion} == verBIFF3) ||
        ($oBook->{BIFFVersion} == verBIFF4) ||
        ($oBook->{BIFFVersion} == verBIFF5) ) {
        $sFmt = substr($sWk, 3, unpack('c', substr($sWk, 2, 1)));
    }
    else {
        $sFmt = _convBIFF8String(substr($sWk, 2));
    }
    $oBook->{FormatStr}->{unpack('v', substr($sWk, 0, 2))} = $sFmt;
}
#------------------------------------------------------------------------------
# _subPalette (for Spreadsheet::ParseExcel) DK: P393
#------------------------------------------------------------------------------
sub _subPalette($$$$)
{
    my($oBook, $bOp, $bLen, $sWk) = @_;
    for(my $i=0;$i<unpack('v', $sWk);$i++) {
        push @aColor, unpack('H6', substr($sWk, $i*4+2));
    }
}
#------------------------------------------------------------------------------
# _subFont (for Spreadsheet::ParseExcel) DK:P333
#------------------------------------------------------------------------------
sub _subFont($$$$)
{
    my($oBook, $bOp, $bLen, $sWk) = @_;
    my($iHeight, $iAttr, $iCIdx, $iBold, $iSuper, $iUnderLine, $iUnicode, $sFntName);
    my($bBold, $bItalic, $bUnderLine, $bStrikeout);
    if($oBook->{BIFFVersion} == verBIFF8) {
        ($iHeight, $iAttr, $iCIdx, $iBold, $iSuper, $iUnderLine) = 
            unpack("v6", $sWk);
        $iUnicode = 1;
        $sFntName = substr($sWk, 16);
        _SwapForUnicode(\$sFntName);
        $sFntName = $oFmtClass->TextFmt($sFntName, 'ucs2');

        $bBold       = ($iBold >= 0x2BC)? 1: 0;
        $bItalic     = ($iAttr & 0x02)? 1: 0;
        $bStrikeout  = ($iAttr & 0x08)? 1: 0;
        $bUnderLine  = ($iUnderLine)? 1: 0;
    }
    elsif($oBook->{BIFFVersion} == verBIFF5) {
        ($iHeight, $iAttr, $iCIdx, $iBold, $iSuper, $iUnderLine) = 
            unpack("v6", $sWk);
        $sFntName = $oFmtClass->TextFmt(
                    substr($sWk, 15, unpack("c", substr($sWk, 14, 1))), 
                    '_native_');
        $bBold       = ($iBold >= 0x2BC)? 1: 0;
        $bItalic     = ($iAttr & 0x02)? 1: 0;
        $bStrikeout  = ($iAttr & 0x08)? 1: 0;
        $bUnderLine  = ($iUnderLine)? 1: 0;
    }
    else {
        ($iHeight, $iAttr) = unpack("v2", $sWk);
        $iCIdx       = undef;
        $iSuper      = 0;

        $bBold       = ($iAttr & 0x01)? 1: 0;
        $bItalic     = ($iAttr & 0x02)? 1: 0;
        $bUnderLine  = ($iAttr & 0x04)? 1: 0;
        $bStrikeout  = ($iAttr & 0x08)? 1: 0;

        $sFntName = substr($sWk, 5, unpack("c", substr($sWk, 4, 1)));
    }
    push @{$oBook->{Font}}, 
        Spreadsheet::ParseExcel::Font->new(
            Height  => $iHeight / 20.0,
            Attr => $iAttr,
            Color=> $aColor[$iCIdx],
            Super       => $iSuper,
            BoldStyle       => $iBold,
            UnderLineStyle  => $iUnderLine,
            Name        => $sFntName,
            _CIdx       => $iCIdx,

            Bold        => $bBold,
            Italic      => $bItalic,
            UnderLine   => $bUnderLine,
            Strikeout   => $bStrikeout,
    );
    #Skip Font[4]
    push @{$oBook->{Font}}, {} if(scalar(@{$oBook->{Font}}) == 4);

}
#------------------------------------------------------------------------------
# _subBoundSheet (for Spreadsheet::ParseExcel): DK: P307
#------------------------------------------------------------------------------
sub _subBoundSheet($$$$)
{
    my($oBook, $bOp, $bLen, $sWk) = @_;
    my($iPos, $iGr) = unpack("v2", $sWk);
    my $iKind = $iGr & 0x0F;
    if($oBook->{BIFFVersion} >= verBIFF8) {
        my($iSize, $iUni) = unpack("cc", substr($sWk, 6, 2));
        my $sWsName = substr($sWk, 8);
        if($iUni & 0x01) {
            _SwapForUnicode(\$sWsName);
            $sWsName = $oFmtClass->TextFmt($sWsName, 'ucs2');
        }
        $oBook->{Worksheet}[$oBook->{SheetCount}] = 
            new Spreadsheet::ParseExcel::Worksheet(
                    Name => $sWsName,
                    Kind => $iKind,
                    _Pos => $iPos,
                );
    }
    else {
        $oBook->{Worksheet}[$oBook->{SheetCount}] = 
            new Spreadsheet::ParseExcel::Worksheet(
                    Name => $oFmtClass->TextFmt(substr($sWk, 7), '_native_'),
                    Kind => $iKind,
                    _Pos => $iPos,
                );
    }
    $oBook->{SheetCount}++;
}
#------------------------------------------------------------------------------
# DecodeBoolErr (for Spreadsheet::ParseExcel) DK: P306
#------------------------------------------------------------------------------
sub DecodeBoolErr($$)
{
    my($iVal, $iFlg) = @_;
    if($iFlg) {     # ERROR
        if($iVal == 0x00) {
            return "#NULL!";
        }
        elsif($iVal == 0x07) {
            return "#DIV/0!";
        }
        elsif($iVal == 0x0F) {
            return "#VALUE!";
        }
        elsif($iVal == 0x17) {
            return "#REF!";
        }
        elsif($iVal == 0x1D) {
            return "#NAME?";
        }
        elsif($iVal == 0x24) {
            return "#NUM!";
        }
        elsif($iVal == 0x2A) {
            return "#N/A!";
        }
        else {
            return "#ERR";
        }
    }
    else {
        return ($iVal)? "TRUE" : "FALSE";
    }
}
#------------------------------------------------------------------------------
# _UnpackRKRec (for Spreadsheet::ParseExcel)    DK:P 401
#------------------------------------------------------------------------------
sub _UnpackRKRec($) {
    my($sArg) = @_;

    my $iF  = unpack('v', substr($sArg, 0, 2));

    my $lWk = substr($sArg, 2, 4);
    my $sWk = pack("c4", reverse(unpack("c4", $lWk)));
    my $iPtn = unpack("c",substr($sWk, 3, 1)) & 0x03;
    if($iPtn == 0) {
        return ($iF, unpack("d", ($BIGENDIAN)? $sWk . "\0\0\0\0": "\0\0\0\0". $lWk));
    }
    elsif($iPtn == 1) {
        substr($sWk, 3, 1) &=  pack('c', unpack("c",substr($sWk, 3, 1)) & 0xFC);
        substr($lWk, 0, 1) &=  pack('c', unpack("c",substr($lWk, 0, 1)) & 0xFC);
        return ($iF, unpack("d", ($BIGENDIAN)? $sWk . "\0\0\0\0": "\0\0\0\0". $lWk)/ 100);
    }
    elsif($iPtn == 2) {
        my $sWkLB = pack("B32", "00" . substr(unpack("B32", $sWk), 0, 30));
        my $sWkL  = ($BIGENDIAN)? $sWkLB: pack("c4", reverse(unpack("c4", $sWkLB)));
        return ($iF, unpack("i", $sWkL));
    }
    else {
        my $sWkLB = pack("B32", "00" . substr(unpack("B32", $sWk), 0, 30));
        my $sWkL  = ($BIGENDIAN)? $sWkLB: pack("c4", reverse(unpack("c4", $sWkLB)));
        return ($iF, unpack("i", $sWkL) / 100);
    }
}
#------------------------------------------------------------------------------
# _subStrWk (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub _subStrWk($$;$)
{
    my($oBook, $sWk, $fCnt) = @_;

    #1. Continue
    if(defined($fCnt)) {
    #1.1 Before No Data No
        if(($oBook->{StrBuff} eq '') || (!(defined($oBook->{_PrevCond})))){
            $oBook->{StrBuff} .= $sWk;
        }
        else {
#print "CONT\n";
            my $iCnt1st = ord($sWk); # 1st byte of Continue may be a GR byte
            my($iStP, $iLenS) = @{$oBook->{_PrevInfo}};
            my $iLenB = length($oBook->{StrBuff});

        #1.1 Not in String
            if($iLenB >= ($iStP + $iLenS)) {
#print "NOT STR\n";
                $oBook->{StrBuff} .= $sWk;
            }
        #1.2 Same code (Unicode or ASCII)
            elsif(($oBook->{_PrevCond} & 0x01) == ($iCnt1st & 0x01)) {
#print "SAME\n";
                $oBook->{StrBuff} .= substr($sWk, 1);
            }
            else {
        #1.3 Diff code (Unicode or ASCII)
                my $iDiff = ($iStP + $iLenS) - $iLenB;
                if($iCnt1st & 0x01) {
#print "DIFF ASC $iStP $iLenS $iLenB DIFF:$iDiff\n";
#print "BEF:", unpack("H6", $oBook->{StrBuff}), "\n";
                  my ($iDum, $iGr) =unpack('vc', $oBook->{StrBuff});
                  substr($oBook->{StrBuff}, 2, 1) = pack('c', $iGr | 0x01);
#print "AFT:", unpack("H6", $oBook->{StrBuff}), "\n";
                    for(my $i = ($iLenB-$iStP); $i >=1; $i--) {
                        substr($oBook->{StrBuff}, $iStP+$i, 0) =  "\x00"; 
                    }
                }
                else {
#print "DIFF UNI:", $oBook->{_PrevCond}, ":", $iCnt1st, " DIFF:$iDiff\n";
                    for(my $i = ($iDiff/2); $i>=1;$i--) {
                        substr($sWk, $i+1, 0) =  "\x00";
                    }
                }
                $oBook->{StrBuff} .= substr($sWk, 1);
           }
        }
    }
    else {
    #2. Saisho
        $oBook->{StrBuff} .= $sWk;
    }
#print " AFT:", unpack("H60", $oBook->{StrBuff}), "\n";

    $oBook->{_PrevCond} = undef;
    $oBook->{_PrevInfo} = undef;

    while(length($oBook->{StrBuff}) >= 4) {
        my ( $raBuff, $iLen, $iStPos, $iLenS) = _convBIFF8String($oBook->{StrBuff}, 1);
                                                    #No Code Convert
        if(defined($raBuff->[0])) {
            push @{$oBook->{PkgStr}}, 
                {
                    Text    => $raBuff->[0],
                    Unicode => $raBuff->[1],
                    Rich    => $raBuff->[2],
                    Ext     => $raBuff->[3],
            };
            $oBook->{StrBuff} = substr($oBook->{StrBuff}, $iLen);
        }
        else {
            $oBook->{_PrevCond} = $raBuff->[1];
            $oBook->{_PrevInfo} = [$iStPos, $iLenS];
            last;
        }
    }
}
#------------------------------------------------------------------------------
# _SwapForUnicode (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub _SwapForUnicode(\$) 
{
    my($sObj) = @_;

    for(my $i = 0; $i<length($$sObj); $i+=2){
            my $sIt = substr($$sObj, $i, 1);
            substr($$sObj, $i, 1) = substr($$sObj, $i+1, 1);
            substr($$sObj, $i+1, 1) = $sIt;
    }
}
#------------------------------------------------------------------------------
# _NewCell (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub _NewCell(%) 
{
    my(%rhKey)=@_;
    my($sWk, $iLen);
    my $oCell = 
        Spreadsheet::ParseExcel::Cell->new(
            Val     => $rhKey{Val},
            Format  => $rhKey{Format},
            Code    => $rhKey{Code},
            Type    => $oFmtClass->ChkType(
                            $rhKey{Numeric}, 
                            $rhKey{Format}->{FmtIdx}),
        );
        $oCell->{_Kind} = $rhKey{Kind};
        $oCell->{_Value} = $oFmtClass->ValFmt($oCell, $rhKey{Book});
    return $oCell;
}
1;
__END__

=head1 NAME

Spreadsheet::ParseExcel - Get information from Excel file

=head1 SYNOPSIS

    use strict;
    use Spreadsheet::ParseExcel;
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

I<$oExcel> = new Spreadsheet::ParseExcel;

Constructor.

=item Parse

I<$oWorkbook> = $oParse->Parse(I<$sFileName> [, I<$oFmt>]);

return L<Workbook> object.
if error occurs, returns undef.

=over 4

=item I<$sFileName>

name of the file to parse

From 0.12 (with OLE::Storage_Lite v.0.06), 
scalar reference of file contents (ex. \$sBuff) or 
IO::Handle object (inclucdng IO::File etc.) are also available.

=item I<$oFmt>

L<Formatter Class> to format the value of cells.

=back

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

Numbers of L<Worksheet> s in that Workbook

=item Worksheet[SheetNo]

Array of L<Worksheet>s class

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

=item ColHeight[Col]

Array of column width (undef means DefColWidth)

=item Cells[Row][Col]

Array of L<Cell>s infomation in the worksheet

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

=back

=head1 AUTHOR

Kawai Takanori (Hippo2000) kwitknr@cpan.org

    http://member.nifty.ne.jp/hippo2000/            (Japanese)
    http://member.nifty.ne.jp/hippo2000/index_e.htm (English)

=head1 SEE ALSO

XLHTML, OLE::Storage, Spreadsheet::WriteExcel, OLE::Storage_Lite

This module is based on herbert within OLE::Storage and XLHTML.

=head1 COPYRIGHT

The Spreadsheet::ParseExcel module is Copyright (c) 2000,2001 Kawai Takanori. Japan.
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
