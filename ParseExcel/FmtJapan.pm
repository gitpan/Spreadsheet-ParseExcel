# Spreadsheet::ParseExcel::FmtJapan
#  by Kawai, Takanori (Hippo2000) 2001.2.2
# This Program is ALPHA version.
#==============================================================================
package Spreadsheet::ParseExcel::FmtJapan;
require Exporter;
use strict;
use Spreadsheet::ParseExcel::FmtDefault;
use Jcode;
use vars qw($VERSION @ISA);
@ISA = qw(Spreadsheet::ParseExcel::FmtDefault Exporter);

$VERSION = '0.02'; # 
my %hFmtJapan = (
    0x00 => '@',
    0x01 => '0',
    0x02 => '0.00',
    0x03 => '#,##0',
    0x04 => '#,##0.00',
    0x05 => '(\\#,##0_);(\\#,##0)',
    0x06 => '(\\#,##0_);[RED](\\#,##0)',
    0x07 => '(\\#,##0.00_);(\\#,##0.00_)',
    0x08 => '(\\#,##0.00_);[RED](\\#,##0.00_)',
    0x09 => '0%',
    0x0A => '0.00%',
    0x0B => '0.00E+00',
    0x0C => '# ?/?',
    0x0D => '# ??/??',
#    0x0E => 'm/d/yy',
    0x0E => 'yyyy/m/d',
    0x0F => 'd-mmm-yy',
    0x10 => 'd-mmm',
    0x11 => 'mmm-yy',
    0x12 => 'h:mm AM/PM',
    0x13 => 'h:mm:ss AM/PM',
    0x14 => 'h:mm',
    0x15 => 'h:mm:ss',
    0x16 => 'm/d/yy h:mm',
#0x17-0x24 -- Differs in Natinal
    0x25 => '(#,##0_);(#,##0)',
    0x26 => '(#,##0_);[RED](#,##0)',
    0x27 => '(#,##0.00);(#,##0.00)',
    0x28 => '(#,##0.00);[RED](#,##0.00)',
    0x29 => '_(*#,##0_);_(*(#,##0);_(*"-"_);_(@_)',
    0x2A => '_(\\*#,##0_);_(\\*(#,##0);_(*"-"_);_(@_)',
    0x2B => '_(*#,##0.00_);_(*(#,##0.00);_(*"-"??_);_(@_)',
    0x2C => '_(\\*#,##0.00_);_(\\*(#,##0.00);_(*"-"??_);_(@_)',
    0x2D => 'mm:ss',
    0x2E => '[h]:mm:ss',
    0x2F => 'mm:ss.0',
    0x30 => '##0.0E+0',
    0x31 => '@',
);

#------------------------------------------------------------------------------
# new (for Spreadsheet::ParseExcel::FmtJapan)
#------------------------------------------------------------------------------
sub new($%) {
    my($sPkg, %hKey) = @_;
    my $oThis={ 
        Code => $hKey{Code},
    };
    bless $oThis;
    return $oThis;
}
#------------------------------------------------------------------------------
# TextFmt (for Spreadsheet::ParseExcel::FmtJapan)
#------------------------------------------------------------------------------
sub TextFmt($$;$) {
    my($oThis, $sTxt, $sCode) =@_;
    $sCode = 'sjis' if(defined($sCode) && ($sCode eq '_native_'));
    if($oThis->{Code}) {
        return Jcode::convert($sTxt, $oThis->{Code}, $sCode);
    }
    else {
        return $sTxt;
    }
}
#------------------------------------------------------------------------------
# ValFmt (for Spreadsheet::ParseExcel::FmtJapan)
#------------------------------------------------------------------------------
sub ValFmt($$$) {
    my($oThis, $oCell, $oBook) =@_;
    return $oThis->SUPER::ValFmt($oCell, $oBook, \%hFmtJapan);
}
#------------------------------------------------------------------------------
# ChkType (for Spreadsheet::ParseExcel::FmtJapan)
#------------------------------------------------------------------------------
sub ChkType($$$) {
    my($oPkg, $iNumeric, $iFmtIdx) =@_;
# Is there something special for Japan?
    return $oPkg->SUPER::ChkType($iNumeric, $iFmtIdx);
}
1;
