# Spreadsheet::ParseExcel::FmtUnicode
#  by Kawai, Takanori (Hippo2000) 2000.12.20
#                                 2001.2.2
# This Program is ALPHA version.
#==============================================================================
package Spreadsheet::ParseExcel::FmtUnicode;
require Exporter;
use strict;
use Spreadsheet::ParseExcel::FmtDefault;
use vars qw($VERSION @ISA);
@ISA = qw(Spreadsheet::ParseExcel::FmtDefault Exporter);
$VERSION = '0.01'; # 
use Unicode::Map;
#------------------------------------------------------------------------------
# new (for Spreadsheet::ParseExcel::FmtJapan2)
#------------------------------------------------------------------------------
sub new($%) {
    my($sPkg, %hKey) = @_;
    my $sMap = $hKey{Unicode_Map};
    my $oMap;
    $oMap = Unicode::Map->new($sMap) if $sMap;
    my $oThis={ 
        Unicode_Map => $sMap,
        _UniMap => $oMap,
    };
    bless $oThis;
    return $oThis;
}
#------------------------------------------------------------------------------
# TextFmt (for Spreadsheet::ParseExcel::FmtJapan2)
#------------------------------------------------------------------------------
sub TextFmt($$;$) {
    my($oThis, $sTxt, $sCode) =@_;
    if($oThis->{_UniMap}) {
        $sTxt = $oThis->{_UniMap}->from_unicode($sTxt)
         if($sCode eq 'ucs2');
        return $sTxt;
    }
    else {
        return $sTxt;
    }
}
1;
