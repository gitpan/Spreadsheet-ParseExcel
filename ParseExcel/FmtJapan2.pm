# Spreadsheet::ParseExcel::FmtJapan2
#  by Kawai, Takanori (Hippo2000) 2001.2.2
# This Program is ALPHA version.
#==============================================================================
package Spreadsheet::ParseExcel::FmtJapan2;
require Exporter;
use strict;
use Jcode;
use Unicode::Map;
use Spreadsheet::ParseExcel::FmtJapan;
use vars qw($VERSION @ISA);
@ISA = qw(Spreadsheet::ParseExcel::FmtJapan Exporter);
$VERSION = '0.02'; # 

#------------------------------------------------------------------------------
# new (for Spreadsheet::ParseExcel::FmtJapan2)
#------------------------------------------------------------------------------
sub new($%) {
    my($sPkg, %hKey) = @_;
    my $oMap = Unicode::Map->new('CP932Excel');
    my $oThis={ 
        Code => $hKey{Code},
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
    $sCode = 'sjis' if(defined($sCode) && ($sCode eq '_native_'));

    if($oThis->{Code}) {
    if($sCode eq 'ucs2') {
        $sCode = 'sjis';
            $sTxt = $oThis->{_UniMap}->from_unicode($sTxt);
        }
        return Jcode::convert($sTxt, $oThis->{Code}, $sCode);
    }
    else {
        return $sTxt;
    }
}
1;
