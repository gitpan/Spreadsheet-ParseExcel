package Spreadsheet::ParseExcel::FmtJapan;
#===================================
# Spreadsheet::ParseExcel::FmtJapan
#  by Kawai, Takanori (Hippo2000) 2000.9.20
# This Program is ALPHA version.
#===================================
require Exporter;
use strict;
use vars qw($VERSION @ISA);
@ISA = qw(Exporter);
$VERSION = '0.01'; # 
use Jcode;

sub new($%) {
    my($sPkg, %hKey) = @_;
    my $oThis={ 
        Code => $hKey{Code},
    };
    bless $oThis;
    return $oThis;
}

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

sub ValFmt($$$) {
    my($oThis, $oCell, $oBook) =@_;

    my($Dt, $iFmtIdx, $iNumeric, $Flg1904);

    $Dt       = $oCell->{Val};
    $iFmtIdx  = $oCell->{Format}->{FmtIdx};
    $Flg1904  = $oBook->{Flg1904};

    if ($oCell->{Type} eq 'Numeric') {
        if($iFmtIdx == 0x00) {      #General
            return sprintf "%.15g", $Dt;
        }
        elsif($iFmtIdx == 0x01) { # Number 0
            return sprintf "%0.0f", $Dt;
        }
        elsif($iFmtIdx == 0x02) { # Number 0.00
            return sprintf "%0.2f", $Dt;
        }
        elsif($iFmtIdx == 0x03) { # Number w/comma 0,000.0
            return sprintf "%0.0f", $Dt;
        }
        elsif($iFmtIdx == 0x04) { # Number w/comma  0,000.00
            return sprintf "%0.2f", $Dt;
        }
        elsif($iFmtIdx == 0x09) { # Percent 0%
            return sprintf("%.0f%%", $Dt * 100.0);
        }
        elsif($iFmtIdx == 0x0A) { # Percent 0.00%
            return sprintf("%0.2f%%", $Dt*100.0);
        }
        elsif($iFmtIdx == 0x0B) { # Scientific 0.00+E00
            return sprintf("%0.2E", $Dt);
        }
        elsif($iFmtIdx == 0x0C) { #Fraction 1 number  e.g. 1/2, 1/3
            return sprintf "%0.1f", $Dt;
        }
        elsif($iFmtIdx == 0x0D) { # Fraction 2 numbers  e.g. 1/50, 25/33
            return sprintf "%0.2f", $Dt;
        }
        elsif($iFmtIdx == 0x31) { # Text - if we are here...its a number
            return sprintf "%g", $Dt;
        }
        else { #// Unsupported...but, if we are here, its a number
            return sprintf "%g", $Dt;
        }
    }
    elsif($oCell->{Type} eq 'Date') {
        my($iSec, $iMin, $iHour, $iDay, $iMon, $iYear, $iwDay, $iMSec) = 
            Spreadsheet::ParseExcel::ExcelLocaltime($Dt, $Flg1904);
        $iMon++;
        $iYear+=1900;

        if(($iFmtIdx == 0x0E)||  # Date: m-d-y
          ($iFmtIdx == 0x0F) ||  # Date: d-mmm-yy
          ($iFmtIdx == 0x10) ||  # Date: d-mmm
          ($iFmtIdx == 0x11)     # Date: mmm-yy
        ){
            return sprintf("%4d/%2d/%2d", $iYear, $iMon, $iDay);
        }
        elsif(
          ($iFmtIdx == 0x12) ||  # Time: h:mm AM/PM
          ($iFmtIdx == 0x13) ||  # Time: h:mm:ss AM/PM
          ($iFmtIdx == 0x14) ||  # Time: h:mm
          ($iFmtIdx == 0x15) ||  # Time: h:mm:ss
          ($iFmtIdx == 0x2D) ||  # Time: mm:ss
          ($iFmtIdx == 0x2E) ||  # Time: [h]:mm:ss
          ($iFmtIdx == 0x2F)     # Time: mm:ss.0
          ) {
            return sprintf("%2d:%02d:%02d", $iHour, $iMin, $iSec);
        }
        else {
            return sprintf("%2d:%02d:%02d", $iHour, $iMin, $iSec);
        }
    }
    else {
        return $oThis->TextFmt($oCell->{Val}, $oCell->{Code});
    }
}
sub ChkType($$$) {
    my($oPkg, $iNumeric, $iFmtIdx) =@_;

    if ($iNumeric) {
        if(
          ($iFmtIdx == 0x0E) ||  # Date: m-d-y
          ($iFmtIdx == 0x0F) ||  # Date: d-mmm-yy
          ($iFmtIdx == 0x10) ||  # Date: d-mmm
          ($iFmtIdx == 0x11) ||  # Date: mmm-yy
          ($iFmtIdx == 0x12) ||  # Time: h:mm AM/PM
          ($iFmtIdx == 0x13) ||  # Time: h:mm:ss AM/PM
          ($iFmtIdx == 0x14) ||  # Time: h:mm
          ($iFmtIdx == 0x15) ||  # Time: h:mm:ss
          ($iFmtIdx == 0x2D) ||  # Time: mm:ss
          ($iFmtIdx == 0x2E) ||  # Time: [h]:mm:ss
          ($iFmtIdx == 0x2F)  # Time: mm:ss.0
          ) {
            return "Date";
        }
        else {
#         ($iFmtIdx == 0x00) or   #General
#         ($iFmtIdx == 0x01) or # Number 0
#         ($iFmtIdx == 0x02) or # Number 0.00
#         ($iFmtIdx == 0x03) or # Number w/comma 0,000.0
#         ($iFmtIdx == 0x04) or # Number w/comma    0,000.00
#         ($iFmtIdx == 0x09) or # Percent 0%
#         ($iFmtIdx == 0x0A) or # Percent 0.00%
#         ($iFmtIdx == 0x0B) or # Scientific 0.00+E00
#         ($iFmtIdx == 0x0C) or #Fraction 1 number  e.g. 1/2, 1/3
#         ($iFmtIdx == 0x0D)    # Fraction 2 numbers  e.g. 1/50, 25/33
#          )
            return "Numeric";
        }
    }
    else {
        return "Text";
    }
}
1;
