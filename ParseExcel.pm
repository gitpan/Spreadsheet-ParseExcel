# Spreadsheet::ParseExcel
#  by Kawai, Takanori (Hippo2000) 2000.10.2
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
  my $oThis = { Name        => $rhIni{Name},
                NameUnicode => $rhIni{NameUnicode},
                Cells       => undef,
                };
  bless $oThis;
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
        Size      => $rhIni{Size},
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
    my($sWk, $bVer, $iLen);
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
$VERSION = '0.13'; # 

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
sub verExcel95 {0x500};
sub verExcel97 {0x600};

my %ProcTbl =(
#    0x00 => undef,          # Dimensions (SKIP)
    0x01 => \&_subBlank,     # Blank
    0x02 => \&_subInteger,   # Integer
    0x03 => \&_subNumFloat,  # Number
    0x04 => \&_subLabelUni , # Label
    0x05 => \&_subBoolErr,   # BoolErr
    0x06 => \&_subFormula,   # Formula
    0x07 => \&_subString,    # STRING
    0x08 => \&_subRowData,   # RowData
    0x09 => \&_subBOF,       # BOF
#    0x0A => undef,          # EOF
#    0x0B => undef,          # INDEX
#    0x0C => undef,          # CALCCOUNT
#    0x0D => undef,          # CALCMODE
#    0x0E => undef,          # PRECISION
#    0x0F => undef,          # REFMODE
#    0x10 => undef,          # DELTA
#    0x11 => undef,          # ITERATION
#    0x12 => undef,          # PROTECT
#    0x13 => undef,          # PASSWORD
#    0x14 => undef,          # Header(Skip)
#    0x15 => undef,          # Footer(Skip)
#    0x16 => undef,          # EXTERNCOUNT                
#    0x17 => undef,          # EXTERNSHEET                
    0x18 => \&_subNameUNI,   # Name UNI
#    0x19 => undef,          # WINDOW PROTECT             
#    0x1A => undef,          # VERTICAL PAGE BREAKS       
#    0x1B => undef,          # HORIZONTAL PAGE BREAKS     
#    0x1C => undef,          # NOTE                       
#    0x1D => undef,          # SELECTION                  
#    0x1E => undef,          # FORMAT                     
#    0x1F => undef,          # FORMATCOUNT                
#    0x20 => undef,          # COLUMN DEFAULT             
#    0x21 => undef,          # Arrays (Skip)
    0x22 => \&_subFlg1904,   # 1904 Flag
#    0x23 => undef,          # EXTERNNAME                 
#    0x24 => undef,          # COLWIDTH                   
    0x25 => undef,          # DEFAULT ROW HEIGHT         
#    0x26 => undef,          # LEFT MARGIN                
#    0x27 => undef,          # RIGHT MARGIN               
#    0x28 => undef,          # TOP MARGIN                 
#    0x29 => undef,          # BOTTOM MARGIN              
#    0x2A => undef,          # PRINT ROW HEADERS          
#    0x2B => undef,          # PRINT GRIDLINES            
#    0x2F => undef,          # FILEPASS                   
    0x31 => \&_subFont,      # Font
#    0x32 => undef,          # FONT2                      
#    0x36 => undef,          # TABLE                      
#    0x37 => undef,          # TABLE2                     
    0x3C => \&_subContinue,  # Continue
#    0x3D => undef,          # WINDOW1                    
#    0x3E => undef,          # WINDOW2                    
#    0x40 => undef,          # BACKUP                     
#    0x41 => undef,          # PANE                       
    0x5C => \&_subAuthors,   # Author's
    0x7D => \&_subColW,      # Col Width (?)
    0x7E => \&_subRKNumber,  # RK Number
    0x85 => \&_subBoundSheet,# BoundSheet
    0x99 => \&_subColWDef,   # Default Col
#    0xBC => undef,          # Shared Fomula (Skip)
    0xBD => \&_subMulRK,     # MULRK
    0xBE => \&_subMulBlank,  # MULBLANK
    0xD6 => \&_subRString,   # RString
    0xE0 => \&_subExFmt,     # ExTended Format
#    0xE5 => undef,          # Cell Merge Instructions (skip)
    0xFC => \&_subPackedStr, # Packed String Array
    0xFD => \&_subPackedIdx, # String Index
);
    my %NameTbl = (
        0x00=>'DIMENSIONS',             0x01=>'BLANK',
        0x02=>'INTEGER',                0x03=>'NUMBER',
        0x04=>'LABEL',                  0x05=>'BOOLERR',        
        0x06=>'FORMULA',                0x07=>'STRING',
        0x08=>'ROW',                    0x09=>'BOF',            
        0x0A=>'EOF',                    0x0B=>'INDEX',
        0x0C=>'CALCCOUNT',              0x0D=>'CALCMODE',       
        0x0E=>'PRECISION',              0x0F=>'REFMODE',
        0x10=>'DELTA',                  0x11=>'ITERATION',      
        0x12=>'PROTECT',                0x13=>'PASSWORD',
        0x14=>'HEADER',                 0x15=>'FOOTER',         
        0x16=>'EXTERNCOUNT',            0x17=>'EXTERNSHEET',
        0x18=>'NAME',                   0x19=>'WINDOW PROTECT', 
        0x1A=>'VERTICAL PAGE BREAKS',   0x1B=>'HORIZONTAL PAGE BREAKS',
        0x1C=>'NOTE',                   0x1D=>'SELECTION',      
        0x1E=>'FORMAT',                 0x1F=>'FORMATCOUNT',
        0x20=>'COLUMN DEFAULT',         0x21=>'ARRAY',          
        0x22=>'1904',                   0x23=>'EXTERNNAME',
        0x24=>'COLWIDTH',               0x25=>'DEF ROW HEIGHT', 
        0x26=>'LEFT MARGIN',            0x27=>'RIGHT MARGIN',
        0x28=>'TOP MARGIN',             0x29=>'BOTTOM MARGIN',  
        0x2A=>'PRINT ROW HEADERS',      0x2B=>'PRINT GRIDLINES',
        0x2F=>'FILEPASS',               0x31=>'FONT',           
        0x32=>'FONT2',                  0x36=>'TABLE',
        0x37=>'TABLE2',                 0x3C=>'CONTINUE',       
        0x3D=>'WINDOW1',                0x3E=>'WINDOW2',
        0x40=>'BACKUP',                 0x41=>'PANE',           
        0x5C=>'Author',                 
        0x7D=>'COLUMN WIDTH?',
        0x7E=>'RK Number',
        0x85=>'BoundSheet',             
        0xBC=>'Shared Fomula',          0xBD=>'MUL RK',
        0xBE=>'MUL BLANK',
        0xD6=>'RString',                
        0xE0=>'Extended Format',
        0xE5=>'Cell Merge',             
        0xFC=>'Packed String Array',    0xFD=>'String Index',
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
    my($sWk, $bVer, $bLen);

#0. New $oBook
    my $oBook = Spreadsheet::ParseExcel::Workbook->new;
    $oBook->{SheetCount} = 0;

#1.Get content
    $oBook->{File} = $sFile;
    my($sBiff, $iLen) = $oThis->{GetContent}->($sFile);
    return undef unless($sBiff);

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
    $sWk = substr($sBiff, $lPos, 4);
    $lPos += 4;
    my $iEfFlg = 0;
    while($lPos<=$iLen) {
        my($bOp, $bVer, $bLen) = unpack("C2v", $sWk);
       if($bLen) {
            $sWk = substr($sBiff, $lPos, $bLen);
            $lPos += $bLen;
        }
	#Check EF, EOF
	if($bOp == 0xEF) {    #EF
            $iEfFlg = $bOp;
	}
	elsif($bOp == 0x0A) { #EOF
            undef $iEfFlg;
	}
	unless($iEfFlg) {
	    if(defined $oThis->{FuncTbl}->{$bOp}) {
            	$oThis->{FuncTbl}->{$bOp}->($oBook, $bOp, $bVer, $bLen, $sWk);
            }
            $PREFUNC = $bOp if ($bOp != 0x3C); #Not Continue 
	}
        $sWk = substr($sBiff, $lPos, 4) if(($lPos+4) <= $iLen);
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
# _subBOF (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub _subBOF($$$$$)
{
    my($oBook, $bOp, $bVer, $bLen, $sWk) = @_;

    if(defined $oBook->{_CurSheet}) {
        $oBook->{_CurSheet}++; 
    }
    else {
        if($bVer ==8) {
            $oBook->{Version} = unpack("v", $sWk);
            $oBook->{_CurSheet} = -1;
        }
        else {
            $oBook->{Version} = $bVer;
            $oBook->{_CurSheet} = 0;
            $oBook->{Worksheet}[$oBook->{SheetCount}] =
                    new Spreadsheet::ParseExcel::Worksheet(
                         _Name => '',
                          Name => '',
            );
    	    $oBook->{SheetCount}++;
        }
    }
}
#------------------------------------------------------------------------------
# _subBlank (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub _subBlank($$$$$)
{
    my($oBook, $bOp, $bVer, $bLen, $sWk) = @_;
    my($iR, $iC, $iF);

    if($bVer == 2) {
        ($iR, $iC, $iF) = unpack("v3", $sWk);
    }
    else {
        $iF = 0;
        ($iR, $iC) = unpack("v2", $sWk);
    }
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
# _subInteger (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub _subInteger($$$$$)
{
    my($oBook, $bOp, $bVer, $bLen, $sWk) = @_;
    my($iR, $iC, $iF, $sTxt, $sDum);
    if($bVer == 2) {
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
    }
    else {
        ($iR, $iC) = unpack("v2", $sWk);
        $iF = 0;
        $sTxt = "****INT";
        $oBook->{Worksheet}[$oBook->{_CurSheet}]->{Cells}[$iR][$iC] = 
            _NewCell (
                Kind => 'INTEGER',
                Val     => $sTxt,
                Format  => $oBook->{Format}[$iF],
                Numeric => 0,
                Code    => undef,
                Book    => $oBook,
            );
    }
#2.MaxRow, MaxCol, MinRow, MinCol
    _SetDimension($oBook, $iR, $iC, $iC);
}
#------------------------------------------------------------------------------
# _subNumFloat (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub _subNumFloat($$$$$)
{
    my($oBook, $bOp, $bVer, $bLen, $sWk) = @_;
    my($iR, $iC, $iF, $sTxt, $sDum);
    if($bVer == 2) {
        ($iR, $iC, $iF) = unpack("v3", $sWk);
        $sTxt = unpack("d", ($BIGENDIAN)? 
                    pack("c8", reverse(unpack("c8", substr($sWk, 6, 8)))) :
                    substr($sWk, 6, 8));

        $oBook->{Worksheet}[$oBook->{_CurSheet}]->{Cells}[$iR][$iC] = 
            _NewCell (
                Kind    => 'Float',
                Val     => $sTxt,
                Format  => $oBook->{Format}[$iF],
                Numeric => 1,
                Code    => undef,
                Book    => $oBook,
            );
    }
    else {
        ($iR, $iC) = unpack("v2", $sWk);
        $iF = 0;
        $sTxt = "****FPv";
        $oBook->{Worksheet}[$oBook->{_CurSheet}]->{Cells}[$iR][$iC] = 
            _NewCell (
                Kind    => 'Float',
                Val     => $sTxt,
                Format  => $oBook->{Format}[$iF],
                Numeric => 0,
                Code    => undef,
                Book    => $oBook,
            );
    }
#2.MaxRow, MaxCol, MinRow, MinCol
    _SetDimension($oBook, $iR, $iC, $iC);
}
#------------------------------------------------------------------------------
# _subRString (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub _subRString($$$$$)
{
    my($oBook, $bOp, $bVer, $bLen, $sWk) = @_;
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
#2.MaxRow, MaxCol, MinRow, MinCol
    _SetDimension($oBook, $iR, $iC, $iC);
}
#------------------------------------------------------------------------------
# _subBoolErr (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub _subBoolErr($$$$$)
{
    my($oBook, $bOp, $bVer, $bLen, $sWk) = @_;
    my($iR, $iC, $iF, $sTxt, $sDum);
    if($bVer == 2) {
        ($iR, $iC, $iF) = unpack("v3", $sWk);
        my ($iVal, $iFlg) = unpack("cc", substr($sWk, 6, 2));
        $sTxt = DecodeBoolErr($iVal, $iFlg);
    }
    else {
        ($iR, $iC) = unpack("v2", $sWk);
        $iF = 0;
        $sTxt = "****Bool";
    }
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
# _subRKNumber (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub _subRKNumber($$$$$)
{
    my($oBook, $bOp, $bVer, $bLen, $sWk) = @_;
    my($iR, $iC, $iF, $lWk, $sTxt, $sDum);
    ($iR, $iC, $iF) = unpack("v3", $sWk);
    my ($iPtn, $iTxt);
    $sTxt = _LongToDt(substr($sWk, 6, 4));
    $oBook->{Worksheet}[$oBook->{_CurSheet}]->{Cells}[$iR][$iC] = 
    _NewCell (
            Kind    => 'RKNumber',
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
# _subFormula (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub _subFormula($$$$$)
{
    my($oBook, $bOp, $bVer, $bLen, $sWk) = @_;
    my($iR, $iC, $iF, $sTxt);
    ($iR, $iC, $iF) = unpack("v3", $sWk);
    my ($iFlg) = unpack("v", substr($sWk,12,2));
    if($iFlg == 0xFFFF) {
        my($iKind) = unpack("c", substr($sWk, 6, 1));
        my($iVal)  = unpack("c", substr($sWk, 8, 1));
        if($iKind == 1) { # Boolean Value
            $sTxt = DecodeBoolErr($iVal, 0);
        }
        elsif($iKind == 2) { #Error Code
            $sTxt = DecodeBoolErr($iVal, 1);
        }
        else {
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
            return;
            #$sTxt = substr($sWk, 8);
            #$sTxt = "NOT IMPLEMENT:$iKind:" . unpack("H34",substr($sWk, 13));
        }
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
    else {
        my ($dVal) = unpack("d", ($BIGENDIAN)? 
                        pack("c8", reverse(unpack("c8", substr($sWk, 6, 8)))) :
                        substr($sWk, 6, 8));
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
sub _subString($$$$$)
{
    my($oBook, $bOp, $bVer, $bLen, $sWk) = @_;
    my $iPos = $oBook->{_PrevPos};
    return undef unless($iPos);
    $oBook->{_PrevPos} = undef;
    my ($iR, $iC, $iF) = @$iPos;

    my ($iLen, $sTxt, $iCode);
    if($oBook->{Version} == verExcel95) {
        $iCode = 0;
        $iLen = unpack("v", $sWk);
        $sTxt = substr($sWk, 2, $iLen);
    }
    elsif($oBook->{Version} == verExcel97) {
        ($iLen, $iCode) = unpack("vc", $sWk);
        if($iCode) {
            $iLen *= 2 if($iCode);
            $sTxt = substr($sWk, 3, $iLen);
            _SwapForUnicode(\$sTxt);
        }
        else {
            $sTxt = substr($sWk, 3, $iLen);
        }
    }
    else {
        $iCode = 0;
        $iLen = unpack("c", $sWk);
        $sTxt = substr($sWk, 1, $iLen);
    }
    $oBook->{Worksheet}[$oBook->{_CurSheet}]->{Cells}[$iR][$iC] = 
        _NewCell (
            Kind    => 'Formula String',
            Val     => $sTxt,
            Format  => $oBook->{Format}[$iF],
            Numeric => 0,
            Code    => ($iCode)? 'ucs2': undef,
            Book    => $oBook,
        );
#2.MaxRow, MaxCol, MinRow, MinCol
    _SetDimension($oBook, $iR, $iC, $iC);
}
#------------------------------------------------------------------------------
# _subLabelUni (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub _subLabelUni($$$$$)
{
    my($oBook, $bOp, $bVer, $bLen, $sWk) = @_;
    my($iR, $iC, $iF, $sTxt);
    ($iR, $iC, $iF) = unpack("v3", $sWk);
    $sTxt = substr($sWk, 8);
    $oBook->{Worksheet}[$oBook->{_CurSheet}]->{Cells}[$iR][$iC] = 
    _NewCell ( 
            Kind    => 'LabelUni',
            Val     => $sTxt,
            Format  => $oBook->{Format}[$iF],
            Numeric => 0,
            Code    => '_native_',
            Book    => $oBook,
        );
#2.MaxRow, MaxCol, MinRow, MinCol
    _SetDimension($oBook, $iR, $iC, $iC);
}
#------------------------------------------------------------------------------
# _subMulRK (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub _subMulRK($$$$$)
{
    my($oBook, $bOp, $bVer, $bLen, $sWk) = @_;
    return if ($oBook->{SheetCount}<=0);

    my($iR, $iSc, $iEc,$sTxt);
    ($iR, $iSc) = unpack("v2", $sWk);
    $iEc = unpack("v", substr($sWk, length($sWk) -2, 2));
    my $iPos = 4;
    for(my $iC=$iSc; $iC<=$iEc; $iC++) {
        my($iF) = unpack("v", substr($sWk, $iPos, 2));
        $sTxt = _LongToDt(substr($sWk, $iPos+2, 4), $iR, $iC);
        $oBook->{Worksheet}[$oBook->{_CurSheet}]->{Cells}[$iR][$iC] = 
        _NewCell (
                Kind    => 'MulRK',
                Val     => $sTxt,
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
# _subMulBlank (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub _subMulBlank($$$$$)
{
    my($oBook, $bOp, $bVer, $bLen, $sWk) = @_;
    my($iR, $iSc, $iEc,$sTxt);
    ($iR, $iSc) = unpack("v2", $sWk);
    $iEc = unpack("v", substr($sWk, length($sWk)-2, 2));

    for(my $iC=$iSc; $iC<=$iEc; $iC++) {

    $oBook->{Worksheet}[$oBook->{_CurSheet}]->{Cells}[$iR][$iC] = 
    _NewCell (
            Kind    => 'MulBlank',
            Val     => '',
            Format  => $oBook->{Format}[0],
            Numeric => 0,
            Code    => undef,
            Book    => $oBook,
        );
    }
#2.MaxRow, MaxCol, MinRow, MinCol
    _SetDimension($oBook, $iR, $iSc, $iEc);
}
#------------------------------------------------------------------------------
# _subPackedIdx (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub _subPackedIdx($$$$$)
{
    my($oBook, $bOp, $bVer, $bLen, $sWk) = @_;
    my($iR, $iC, $iF, $iIdx);
    ($iR, $iC, $iF) = unpack("v3", $sWk);
    $iIdx = unpack("L", ($BIGENDIAN)?
                substr($sWk, 9, 1) . substr($sWk, 8, 1) . substr($sWk, 7, 1) . substr($sWk, 6, 1) :
                substr($sWk, 6, 4));
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
# _subNameUNI (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub _subNameUNI($$$$$)
{
    my($oBook, $bOp, $bVer, $bLen, $sWk) = @_;
    if (substr($sWk, 15) =~ /^(.*)\x17(.)\x00(.*)$/) {
        $oBook->{$1} = $3;
    }
}
#------------------------------------------------------------------------------
# _subFlg1904 (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub _subFlg1904($$$$$)
{
    my($oBook, $bOp, $bVer, $bLen, $sWk) = @_;
    $oBook->{Flg1904} = unpack("v", $sWk);
}
#------------------------------------------------------------------------------
# _subRowData (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub _subRowData($$$$$)
{
    my($oBook, $bOp, $bVer, $bLen, $sWk) = @_;

#0. Get Worksheet info (MaxRow, MaxCol, MinRow, MinCol)
    my($iR, $iSc, $iEc, $iHght, $iXf) = unpack("v5", $sWk);
    $iEc--;

#1. RowHeight
    $oBook->{Worksheet}[$oBook->{_CurSheet}]->{RowHeight}[$iR] = $iHght/20;

#2.MaxRow, MaxCol, MinRow, MinCol
    _SetDimension($oBook, $iR, $iSc, $iEc);
}
#------------------------------------------------------------------------------
# _SetDimension (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub _SetDimension($$$$)
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
# _subDefRowHeight (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub _subDefRowHeight($$$$$)
{
    my($oBook, $bOp, $bVer, $bLen, $sWk) = @_;
#1. RowHeight
    my($iDum, $iHght) = unpack("v2", $sWk);
    $oBook->{Worksheet}[$oBook->{_CurSheet}]->{DefRowHeight} = $iHght/20;

}
#------------------------------------------------------------------------------
# _subColWDef(for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub _subColWDef($$$$$)
{
    my($oBook, $bOp, $bVer, $bLen, $sWk) = @_;
    my $iW = unpack("v", $sWk);
    $oBook->{Worksheet}[$oBook->{_CurSheet}]->{DefColWidth}= ($iW -0xA0)/256;
}
#------------------------------------------------------------------------------
# _subColW (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub _subColW($$$$$)
{
    my($oBook, $bOp, $bVer, $bLen, $sWk) = @_;
    my($iSc, $iEc, $iW) = unpack("v3", $sWk);
    for(my $i= $iSc; $i<=$iEc; $i++) {
        $oBook->{Worksheet}[$oBook->{_CurSheet}]->{ColWidth}[$i] = ($iW -0xA0)/256;
    }
}
#------------------------------------------------------------------------------
# SubPackedStr (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub _subPackedStr($$$$$)
{
    my($oBook, $bOp, $bVer, $bLen, $sWk) = @_;
    _subStrWk($oBook, substr($sWk, 8));
}
#------------------------------------------------------------------------------
# _subContinue (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub _subContinue($$$$$)
{
    my($oBook, $bOp, $bVer, $bLen, $sWk) = @_;
    _subStrWk($oBook, $sWk, 1) if($PREFUNC == 0xFC);
}
#------------------------------------------------------------------------------
# _subAuthors (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub _subAuthors($$$$$)
{
    my($oBook, $bOp, $bVer, $bLen, $sWk) = @_;
    return if (defined $oBook->{_Author});
    if($oBook->{Version} == verExcel97) {
        my($bLen, $iFlg) = unpack("vc", $sWk);
        if($iFlg & 0x01) {
            my $sUni = substr($sWk, 3, $bLen * 2);
            _SwapForUnicode(\$sUni);
            $oBook->{_Author} = $sUni;
            $oBook->{Author} = $oFmtClass->TextFmt($oBook->{_Author}, 'ucs2');
        }
        else {
            $oBook->{_Author} = substr($sWk, 3, $bLen);
            $oBook->{Author} = substr($sWk, 3, $bLen);
        }
    }
    elsif($oBook->{Version} == verExcel95) {
        my($iLen) = unpack("c", $sWk);
        $oBook->{_Author} = substr($sWk, 1, $iLen);
        $oBook->{Author} = $oFmtClass->TextFmt($oBook->{_Author}, '_native_');
    }
}
#------------------------------------------------------------------------------
# _subExFmt (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub _subExFmt($$$$$)
{
    my($oBook, $bOp, $bVer, $bLen, $sWk) = @_;
    my($iFnt, $iIdx, $iGen, $iAlign, $iIndent, $iDum);
    my($iBdrStyle, $iBdrLClr, $iBdrTClr, $iCellColor);
    ($iFnt, $iIdx, $iGen, $iAlign, $iIndent) = unpack("v5", $sWk);

    if($oBook->{Version} == verExcel95) {
        $iBdrLClr = 0;
        $iBdrTClr = 0;
        $iCellColor = unpack("v", substr($sWk, 12, 2));
    }
    elsif($oBook->{Version} == verExcel97) {
        ($iBdrLClr, $iBdrTClr, $iCellColor)
                = unpack("vlv", substr($sWk, 12, 8));
    }
    my $sBdrL = ($iBdrLClr==0)? '0'x16 : unpack("B16", $iBdrLClr);
    my $sBdrT = ($iBdrTClr==0)? '0'x16 : unpack("B16", $iBdrTClr);
    my $sBitC = ($iCellColor==0)? '0'x16: unpack("B16", $iCellColor);

    push @{$oBook->{Format}} , 
        Spreadsheet::ParseExcel::Format->new (
        FontNo   => $iFnt,
        Font     => ($iFnt == 0)? $oBook->{Font}[0] : $oBook->{Font}[$iFnt-1],
        FmtIdx   => $iIdx,
        Gen      => $iGen,
        Align    => $iAlign,
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
# _subFont (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub _subFont($$$$$)
{
    my($oBook, $bOp, $bVer, $bLen, $sWk) = @_;
    my($iSize, $iAttr, $iCIdx, $iBold, $iSuper, $iUnderLine, $iUnicode, $sFntName);
    if($oBook->{Version} == verExcel95) {
        ($iSize, $iAttr, $iCIdx, $iBold, $iSuper, $iUnderLine) = 
            unpack("v6", $sWk);
        $iUnicode = 0;
        $sFntName = substr($sWk, 15, unpack("c", substr($sWk, 14, 1)));
    }
    elsif($oBook->{Version} == verExcel97) {
        ($iSize, $iAttr, $iCIdx, $iBold, $iSuper, $iUnderLine) = 
            unpack("v6", $sWk);
        $sFntName = substr($sWk, 16);
        for(my $i = 0; $i<length($sFntName); $i+=2){
            my $sIt = substr($sFntName, $i, 1);
            substr($sFntName, $i, 1) = substr($sFntName, $i+1, 1);
            substr($sFntName, $i+1, 1) = $sIt;
        }
    }
    push @{$oBook->{Font}}, 
        Spreadsheet::ParseExcel::Font->new(
        Size => $iSize / 20,
        Attr => $iAttr,
        Color=> (defined $aColor[$iCIdx])? $aColor[$iCIdx]: $aColor[0],
        Bold => $iBold,
        Super     => $iSuper,
        UnderLine => $iUnderLine,
        Name     => $oFmtClass->TextFmt($sFntName, 'ucs2'),
        _CIdx => $iCIdx,
        _Name     => $sFntName,
    );
}
#------------------------------------------------------------------------------
# _subBoundSheet (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub _subBoundSheet($$$$$)
{
    my($oBook, $bOp, $bVer, $bLen, $sWk) = @_;
    my $iKind  = unpack("c", substr($sWk, 4, 1));
    my ($iSize, $iUni, $sWsName);

    if(!($iKind & 0x0F)) {
        if($oBook->{Version} == verExcel97) {
            ($iSize, $iUni) = unpack("cc", substr($sWk, 6, 2));
            $sWsName = substr($sWk, 8);
            if($iUni & 0x01) {
                _SwapForUnicode(\$sWsName);
                $oBook->{Worksheet}[$oBook->{SheetCount}] = 
                    new Spreadsheet::ParseExcel::Worksheet(
                         _Name => $sWsName, 
                          Name => $oFmtClass->TextFmt($sWsName, 'ucs2'),
                    );
            }
            else {
                $oBook->{Worksheet}[$oBook->{SheetCount}] = 
                    new Spreadsheet::ParseExcel::Worksheet(
                                _Name => $sWsName, 
                                Name => $sWsName, 
                            );
            }
        }
        else {
                $oBook->{Worksheet}[$oBook->{SheetCount}] = 
                    new Spreadsheet::ParseExcel::Worksheet(
                            _Name => substr($sWk, 7), 
                            Name => $oFmtClass->TextFmt(substr($sWk, 7), '_native_'),
                            );
        }
    }
    $oBook->{SheetCount}++;
}
#------------------------------------------------------------------------------
# subDUMP (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub subDUMP($$$$$)
{
    my($oBook, $bOp, $bVer, $bLen, $sWk) = @_;
    printf "%02X:%-22s (Len:%3d) : %s\n", 
            $bOp, OpName($bOp), $bLen, unpack("H40",$sWk);
}
#------------------------------------------------------------------------------
# DecodeBoolErr (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub DecodeBoolErr($$)
{
    my($iVal, $iFlg) = @_;
    if($iFlg) {     # ERROR
        if($iVal == 0x00) {
            return "#NULL!";
        }
        elsif($iVal == 0x07) {
            return "#NULL!";
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
# _LongToDt (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub _LongToDt($) {
    my($lWk, $iR, $iC) = @_;
    my $sWk = pack("c4", reverse(unpack("c4", $lWk)));
    my $iPtn = unpack("c",substr($sWk, 3, 1)) & 0x03;
    if($iPtn == 0) {
        return unpack("d", ($BIGENDIAN)? $sWk . "\0\0\0\0": "\0\0\0\0". $lWk );
    }
    elsif($iPtn == 1) {
        return unpack("d", ($BIGENDIAN)? $sWk . "\0\0\0\0": "\0\0\0\0". $lWk ) / 100.0;
    }
    elsif($iPtn == 2) {
        my $sWkLB = pack("B32", "00" . substr(unpack("B32", $sWk), 0, 30));
        my $sWkL  = ($BIGENDIAN)? $sWkLB: pack("c4", reverse(unpack("c4", $sWkLB)));
        return unpack("i", $sWkL);
    }
    else {
        my $sWkLB = pack("B32", "00" . substr(unpack("B32", $sWk), 0, 30));
        my $sWkL  = ($BIGENDIAN)? $sWkLB: pack("c4", reverse(unpack("c4", $sWkLB)));
        return unpack("i", $sWkL) / 100.0;
    }
}
#------------------------------------------------------------------------------
# _subStrWk (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub _subStrWk($$;$)
{
    my($oBook, $sWk, $fCnt) = @_;
#print "NEW LEN:", length($oBook->{StrBuff}), "  CONT:", defined($fCnt);
#print " STR:", unpack("H60", $oBook->{StrBuff}), "\n";
#print " SWK:", unpack("H60", $sWk), "\n";
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
    my($iLen, $iLenS);
    my($iCrun, $iExrst);
    my($iStP);

    while(length($oBook->{StrBuff}) >= 4) {
        my($iChrs, $iGr) = unpack("vc", $oBook->{StrBuff});
        $iLenS = $iChrs;
        $iLenS *= 2 if($iGr & 0x01);
        $iLen = $iLenS;
        $iStP = 3;
        if(($iGr & 0x0C) == 0x0C) { #FarEast + RichText
            ($iCrun, $iExrst) = unpack("vV", substr($oBook->{StrBuff}, 3, 6));
            $iLen += $iCrun * 4 + $iExrst;
            $iStP = 9;
        }
        elsif(($iGr & 0x08) == 0x08) { #RichText
            $iCrun = unpack("v", substr($oBook->{StrBuff}, 3, 2));
            $iLen += $iCrun * 4;
            $iStP = 5;
        }
        elsif(($iGr & 0x04) == 0x04) { #FarEast
            $iExrst = unpack("V", substr($oBook->{StrBuff}, 3, 4));
            $iLen += $iExrst;
            $iStP = 7;
        }
        if(length($oBook->{StrBuff}) >= $iLen + $iStP) {
            my $sTxt = substr($oBook->{StrBuff}, $iStP, $iLenS);
            if($iGr & 0x01) {
                _SwapForUnicode(\$sTxt);
            }
#print "ADD ARRY:", $#{$oBook->{PkgStr}}, " LEN:", length($sTxt), " ST:", $iStP, , ":", substr($sTxt, length($sTxt)-10, 10), "\n";
            push @{$oBook->{PkgStr}}, {
                Text => $sTxt,
                Unicode => $iGr & 0x01,
            };
            $oBook->{StrBuff} = substr($oBook->{StrBuff}, $iStP+$iLen);
        }
        else {
            $oBook->{_PrevCond} = $iGr;
            $oBook->{_PrevInfo} = [$iStP, $iLenS];
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
    my($sWk, $bVer, $iLen);
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
#------------------------------------------------------------------------------
# Spreadsheet::ParseExcel->OpName
#------------------------------------------------------------------------------
sub OpName($) {
    my($bOp)=@_;
    return (defined $NameTbl{$bOp})? $NameTbl{$bOp}: 'undef';
}
#------------------------------------------------------------------------------
# ExcelLocaltime (for Spreadsheet::ParseExcel)
#------------------------------------------------------------------------------
sub ExcelLocaltime($$)
{
  my($dObj, $flg1904) = @_;

  my($iSec, $iMin, $iHour, $iDay, $iMon, $iYear, $iwDay, $iMSec);
  my($iDt, $iTime, $iYDays);

  $iDt  = int($dObj);
  $iTime = $dObj - $iDt;

  if($flg1904) {
    $iYear = 1904;
    $iDt++;         #Start from Jan 1st
    $iYDays = 366;
    $iwDay = (($iDt+4) % 7);
  }
  else {
    $iYear = 1900;
    $iYDays = 366;  #In Excel 1900 is leap year (That's not TRUE!)
    $iwDay = (($iDt+6) % 7);
  }

  while($iDt > $iYDays) {
    $iDt -= $iYDays;
    $iYear++;
    $iYDays = ((($iYear % 4)==0) && 
        (($iYear % 100) || ($iYear % 400)==0))? 366: 365;
  }
  $iYear -= 1900;
  for($iMon=1;$iMon < 12; $iMon++){
    my $iMD;
    if($iMon == 1 || $iMon == 3 || $iMon == 5 || $iMon == 7 || $iMon == 8
        || $iMon == 10 || $iMon == 12) {
        $iMD = 31;
    }
    elsif($iMon == 4 || $iMon == 6 || $iMon == 9 || $iMon == 11) {
        $iMD = 30;
    }
    elsif($iMon == 2) {
        $iMD = (($iYear % 4) == 0)? 29: 28;
    }
    last if($iDt <= $iMD);
    $iDt -= $iMD;
  }
  $iDay = $iDt;
  $iTime += (0.05 / 86400.0);
  $iTime*=24.0;
  $iHour = int($iTime);
  $iTime -= $iHour;
  $iTime *= 60.0;
  $iMin  = int($iTime);
  $iTime -= $iMin;
  $iTime *= 60.0;
  $iSec  = int($iTime);
  $iTime -= $iSec;
  $iTime *= 10.0;
  $iMSec = int($iTime);
  return ($iSec, $iMin, $iHour, $iDay, $iMon-1, $iYear, $iwDay, $iMSec);
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

    http://member.nifty.ne.jp/hippo2000/ (Sorry Only in Japanese)

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

In no particular order: Simamoto Takesi, Noguchi Harumi, Ikezawa Kazuhiro, 
Suwazono Shugo, Hirofumi Morisada, Michael Edwards, Kim Namusk 
and many many people + Kawai Mikako.

=cut
