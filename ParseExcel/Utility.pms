# Spreadsheet::ParseExcel::Utility
#  by Kawai, Takanori (Hippo2000) 2001.2.2
# This Program is ALPHA version.
#==============================================================================
# Spreadsheet::ParseExcel::Utility;
#==============================================================================
package Spreadsheet::ParseExcel::Utility;
require Exporter;
use strict;
use vars qw($VERSION @ISA @EXPORT_OK);
@ISA = qw(Exporter);
@EXPORT_OK = qw(ExcelFmt LocaltimeExcel ExcelLocaltime);

#ProtoTypes
sub ExcelFmt($$;$);
sub LocaltimeExcel($$$$$$;$$);
sub ExcelLocaltime($;$);
sub AddComma($);
sub MakeBun($$);
sub MakeE($$);
sub LeapYear($);

#------------------------------------------------------------------------------
# ExcelFmt (for Spreadsheet::ParseExcel::Utility)
#------------------------------------------------------------------------------
sub ExcelFmt($$;$) {
    my($sFmt, $iData, $i1904) =@_;

    my $sCond;
    my $sWkF ='';
    my $sRes='';

#1. Get Condition
    if($sFmt=~/^\[([<>=][^\]]+)\](.*)$/) {
        $sCond = $1;
        $sFmt = $2;
    }
    $sFmt =~ s/_/ /g;

    my @sFmtWk;
    my $sFmtObj;
    my $iFmtPos=0;
    my $iDblQ=0;
    my $iQ = 0;
    foreach my $sWk (split //, $sFmt) {
        if($iDblQ or $iQ) {
            $sFmtWk[$iFmtPos] .=$sWk;
            $iDblQ = 0 if($sWk eq '"');
            $iQ = 0;
            next;
        }

        if($sWk eq ';') {
            $iFmtPos++;
            next;
        }
        elsif($sWk eq '"') {
            $iDblQ = 1;
        }
        elsif($sWk eq '!') {
            $iQ = 1;
        }
        $sFmtWk[$iFmtPos] .=$sWk;
    }
#Get FmtString
    if(scalar(@sFmtWk)>1) {
        if($sCond) {
            $sFmtObj = $sFmtWk[((eval("$iData $sCond"))? 0: 1)];
        }
        else {
            if(scalar(@sFmtWk)==2) {
                $sFmtObj = $sFmtWk[(($iData>=0)? 0: 1)];
            }
            else {
                $sFmtObj = $sFmtWk[(($iData>0)? 0: (($iData<0)? 1:  2))];
            }
        }
    }
    else {
        $sFmtObj = $sFmtWk[0];
    }

    my $sColor;
    if($sFmtObj =~ /^(\[[^hm\[\]]*\])/) {
        $sColor = $1;
        $sFmtObj = substr($sFmtObj, length($sColor));
        chop($sColor);
        $sColor = substr($sColor, 1);
    }
#print "FMT:$sFmtObj Co:$sColor\n";

#3.Build Data
    my $iFmtMode=0;   #1:Number, 2:Date
    my $i=0;
    my $ir=0;
    my $sFmtWk;
    my @aRep = ();
    my $sFmtRes='';

    my $iFflg = -1;
    my $iRpos = -1;
    my $iCmmCnt = 0;
    my $iBunFlg = 0;
    my $iFugouFlg = 0;
    my $iPer = 0;
    my $iAm=0;
    my $iSt;

    while($i<length($sFmtObj)) {
        $iSt = $i;
        my $sWk = substr($sFmtObj, $i, 1);

        if($sWk !~ /[#0\+\-\.\?eE\,\%]/) {
            if($iFflg != -1) {
                push @aRep, [substr($sFmtObj, $iFflg, $i-$iFflg),  
                                    $iRpos, $i-$iFflg];
                $iFflg= -1;
            }
        }

        if($sWk eq '"') {
            $iDblQ = $iDblQ? 0: 1;
            $i++;
            next;
        }
        elsif($sWk eq '!') {
            $iQ = 1;
            $i++;
            next;
        }
#print STDERR "DEF1: $iDblQ DEF2: $iQ\n";
        if((defined($iDblQ) and ($iDblQ)) or (defined($iQ) and ($iQ))) {
            $iQ = 0;
            $i++;
        }
        elsif(($sWk =~ /[#0\+\.\?eE\,\%]/) || 
              (($iFmtMode != 2) and ($sWk eq '-'))) {
            $iFmtMode = 1 unless($iFmtMode);
            if(substr($sFmtObj, $i, 1) =~ /[#0]/) {
                if(substr($sFmtObj, $i) =~ /^([#0]+)([\.]?)([0#]*)([eE])([\+\-])([0#]+)/){
                    push @aRep, [substr($sFmtObj, $i, length($&)), $i, length($&)];
                    $i +=length($&);
                }
                else{
                    if($iFflg==-1) {
                        $iFflg = $i;
                        $iRpos = length($sFmtRes);
                    }
                }
            }
            elsif(substr($sFmtObj, $i, 1) eq '?') {
                if($iFflg != -1) {
                    push @aRep, [substr($sFmtObj, $iFflg, $i-$iFflg+1),  
                                        $iRpos, $i-$iFflg+1];
                }
                $iFflg = $i;
                while($i<length($sFmtObj)) {
                    if (substr($sFmtObj, $i, 1) eq '/'){
                        $iBunFlg = 1;
                    }
                    elsif (substr($sFmtObj, $i, 1) eq '?'){
                        ;
                    }
                    else {
                        if(($iBunFlg) && (substr($sFmtObj, $i, 1) =~ /[0-9]/)) {
                            ;
                        }
                        else {
                            last;
                        }
                    }
                    $i++;
                }
                $i--;
                push @aRep, [substr($sFmtObj, $iFflg, $i-$iFflg+1),  
                                        length($sFmtRes), $i-$iFflg+1];
                $iFflg = -1;
            }
            elsif(substr($sFmtObj, $i, 3) =~ /^[eE][\+\-][0#]$/) {
                if(substr($sFmtObj, $i) =~ /([eE])([\+\-])([0#]+)/){
                    push @aRep, [substr($sFmtObj, $i, length($&)), $i, length($&)];
                    $i +=length($&);
                }
                $iFflg = -1;
            }
            else {
                if($iFflg != -1) {
                    push @aRep, [substr($sFmtObj, $iFflg, $i-$iFflg),  
                                        $iRpos, $i-$iFflg];
                    $iFflg= -1;
                }
                if(substr($sFmtObj, $i, 1) =~ /[\+\-]/) {
                    push @aRep, [substr($sFmtObj, $i, 1),  
                                        length($sFmtRes), 1];
                    $iFugouFlg = 1;
                }
                elsif(substr($sFmtObj, $i, 1) eq '.') {
                    push @aRep, [substr($sFmtObj, $i, 1),  
                                        length($sFmtRes), 1];
                }
                elsif(substr($sFmtObj, $i, 1) eq ',') {
                    $iCmmCnt++;
                    push @aRep, [substr($sFmtObj, $i, 1),  
                                        length($sFmtRes), 1];
                }
                elsif(substr($sFmtObj, $i, 1) eq '%') {
                    $iPer = 1;
                }
            }
            $i++;
        }
        elsif($sWk =~ /[ymdhsap]/) {
            $iFmtMode = 2 unless($iFmtMode);
            if(substr($sFmtObj, $i, 5) =~ /am\/pm/i) {
                push @aRep, ['am/pm', length($sFmtRes), 5];
                $iAm=1;
                $i+=5;
            }
            elsif(substr($sFmtObj, $i, 3) =~ /a\/p/i) {
                push @aRep, ['a/p', length($sFmtRes), 3];
                $iAm=1;
                $i+=3;
            }
            elsif(substr($sFmtObj, $i, 5) eq 'mmmmm') {
                push @aRep, ['mmmmm', length($sFmtRes), 5];
                $i+=5;
            }
            elsif((substr($sFmtObj, $i, 4) eq 'mmmm')  ||
                  (substr($sFmtObj, $i, 4) eq 'dddd')  ||
                  (substr($sFmtObj, $i, 4) eq 'yyyy')) {
                push @aRep, [substr($sFmtObj, $i, 4), length($sFmtRes), 4];
                $i+=4;
            }
            elsif((substr($sFmtObj, $i, 3) eq 'mmm')  ||
                  (substr($sFmtObj, $i, 3) eq 'yyy')) {
                push @aRep, [substr($sFmtObj, $i, 3), length($sFmtRes), 3];
                $i+=3;
            }
            elsif((substr($sFmtObj, $i, 2) eq 'yy')  ||
              (substr($sFmtObj, $i, 2) eq 'mm')  ||
              (substr($sFmtObj, $i, 2) eq 'dd')  ||
              (substr($sFmtObj, $i, 2) eq 'hh')  ||
              (substr($sFmtObj, $i, 2) eq 'ss')) {
                if((substr($sFmtObj, $i, 2) eq 'mm') &&
                   ($#aRep>=0) && 
                    (($aRep[$#aRep]->[0] eq 'h') or ($aRep[$#aRep]->[0] eq 'hh'))) {
                        push @aRep, ['mm', length($sFmtRes), 2, 'min'];
                }
                else {
                        push @aRep, [substr($sFmtObj, $i, 2), length($sFmtRes), 2];
                }
                if((substr($sFmtObj, $i, 2) eq 'ss') && ($#aRep>0)) {
                    if(($aRep[$#aRep-1]->[0] eq 'm') ||
                       ($aRep[$#aRep-1]->[0] eq 'mm')) {
                        push(@{$aRep[$#aRep-1]}, 'min');
                    }
                }
                $i+=2;
            }
            elsif((substr($sFmtObj, $i, 1) eq 'm')  ||
                  (substr($sFmtObj, $i, 1) eq 'd')  ||
                  (substr($sFmtObj, $i, 1) eq 'h')  ||
                  (substr($sFmtObj, $i, 1) eq 's')){
                if((substr($sFmtObj, $i, 1) eq 'm') &&
                   ($#aRep>=0) && 
                    (($aRep[$#aRep]->[0] eq 'h') or ($aRep[$#aRep]->[0]  eq 'hh'))) {
                        push @aRep, ['m', length($sFmtRes), 1, 'min'];
                }
                else {
                        push @aRep, [substr($sFmtObj, $i, 1), length($sFmtRes), 1];
                }
                if((substr($sFmtObj, $i, 1) eq 's') && ($#aRep>0)) {
                    if(($aRep[$#aRep-1]->[0] eq 'm') ||
                       ($aRep[$#aRep-1]->[0] eq 'mm')) {
                        push(@{$aRep[$#aRep-1]}, 'min');
                    }
                }
                $i+=1;
            }
        }
        elsif((substr($sFmtObj, $i, 3) eq '[h]')) {
            push @aRep, ['[h]', length($sFmtRes), 3];
            $i+=3;
        }
        elsif((substr($sFmtObj, $i, 4) eq '[mm]')) {
            push @aRep, ['[mm]', length($sFmtRes), 4];
            $i+=4;
        }
        elsif($sWk eq '@') {
            push @aRep, ['@', length($sFmtRes), 1];
            $i++;
        }
        else{
            $i++;
        }
        $i++ if($i == $iSt);        #No Format match
        $sFmtRes .= substr($sFmtObj, $iSt, $i-$iSt);
    }
#print "FMT: $iRpos ",$sFmtRes, "\n";
    if($iFflg != -1) {
        push @aRep, [substr($sFmtObj, $iFflg, $i-$iFflg+1),
                    $iRpos,, $i-$iFflg+1];
        $iFflg= 0;
    }
#For Date format
    if($iFmtMode==2) {
        my @aTime = ExcelLocaltime($iData, $i1904);
        $aTime[4]++;
        $aTime[5] += 1900;

        my @aMonL = 
            qw (dum January February March April May June July 
                August September October November December );
        my @aMonNm =
            qw (dum Jan Feb Mar May Jun Jul Aug Sep Oct Nov Dec);
        my @aWeekNm = 
            qw (Mon Tue Wed Thu Fri Sat Sun);
        my @aWeekL = 
            qw (Monday Tuesday Wednesday Thursday Friday Saturday Sunday);
        my $sRep;
        for(my $iIt=$#aRep; $iIt>=0;$iIt--) {
            my $rItem = $aRep[$iIt];
            if((scalar @$rItem) >=4) {
    #Min
                if($rItem->[0] eq 'mm') {
                    $sRep = sprintf("%02d", $aTime[1]);
                }
                else {
                    $sRep = sprintf("%d", $aTime[1]);
                }
            }
    #Year
            elsif($rItem->[0] eq 'yyyy') {
                $sRep = sprintf('%04d', $aTime[5]);
            }
            elsif($rItem->[0] eq 'yy') {
                $sRep = sprintf('%02d', $aTime[5] % 100);
            }
    #Mon
            elsif($rItem->[0] eq 'mmmmm') {
                $sRep = substr($aMonNm[$aTime[4]], 0, 1);
            }
            elsif($rItem->[0] eq 'mmmm') {
                $sRep = $aMonL[$aTime[4]];
            }
            elsif($rItem->[0] eq 'mmm') {
                $sRep = $aMonNm[$aTime[4]];
            }
            elsif($rItem->[0] eq 'mm') {
                $sRep = sprintf('%02d', $aTime[4]);
            }
            elsif($rItem->[0] eq 'm') {
                $sRep = sprintf('%d', $aTime[4]);
            }
    #Day
            elsif($rItem->[0] eq 'dddd') {
                $sRep = $aWeekL[$aTime[7]];
            }
            elsif($rItem->[0] eq 'ddd') {
                $sRep = $aWeekNm[$aTime[7]];
            }
            elsif($rItem->[0] eq 'dd') {
                $sRep = sprintf('%02d', $aTime[3]);
            }
            elsif($rItem->[0] eq 'd') {
                $sRep = sprintf('%d', $aTime[3]);
            }
    #Hour
            elsif($rItem->[0] eq 'hh') {
                if($iAm) {
                    $sRep = sprintf('%02d', $aTime[2]%12);
                }
                else {
                    $sRep = sprintf('%02d', $aTime[2]);
                }
            }
            elsif($rItem->[0] eq 'h') {
                if($iAm) {
                    $sRep = sprintf('%d', $aTime[2]%12);
                }
                else {
                    $sRep = sprintf('%d', $aTime[2]);
                }
            }
    #SS
            elsif($rItem->[0] eq 'ss') {
                $sRep = sprintf('%02d', $aTime[0]);
            }
            elsif($rItem->[0] eq 'S') {
                $sRep = sprintf('%d', $aTime[0]);
            }
    #am/pm
            elsif($rItem->[0] eq 'am/pm') {
                $sRep = ($aTime[4]>12)? 'pm':'am';
            }
            elsif($rItem->[0] eq 'a/p') {
                $sRep = ($aTime[4]>12)? 'p':'a';
            }
            elsif($rItem->[0] eq '.') {
                $sRep = '.';
            }
            elsif($rItem->[0] =~ /^0+$/) {
                my $i0Len = length($&);
#print "SEC:", $aTime[7], "\n";
                $sRep = substr(sprintf("%.${i0Len}f", $aTime[7]/1000.0), 2, $i0Len);
            }
            elsif($rItem->[0] eq '[h]') {
                $sRep = sprintf('%d', int($iData) * 24 + $aTime[2]);
            }
            elsif($rItem->[0] eq '[mm]') {
                $sRep = sprintf('%d', (int($iData) * 24 + $aTime[2])*60 + $aTime[1]);
            }
            elsif($rItem->[0] eq '@') {
                $sRep = $iData;
            }

#print "REP:$sRep ",$rItem->[0], ":", $rItem->[1], ":" ,$rItem->[2], "\n";
            substr($sFmtRes, $rItem->[1], $rItem->[2]) = $sRep;
        }
    }
    elsif($iFmtMode==1) {
        if($#aRep>=0) {
            while($aRep[$#aRep]->[0] eq ',') {
                $iCmmCnt--;
                substr($sFmtRes, $aRep[$#aRep]->[1], $aRep[$#aRep]->[2]) = '';
                $iData /= 1000;
                pop @aRep;
            }

            my $sNumFmt = join('', map {$_->[0]} @aRep);
            my $sNumRes;
            my $iTtl=0;
            my $iE=0;
            my $iP=0;
            my $iAftP=undef;
            foreach my $sItem (split //, $sNumFmt) {
                if($sItem eq '.') {
                    $iTtl++;
                    $iP = 1;
                }
                elsif(($sItem eq 'E') || ($sItem eq 'e')){
                    $iE = 1;
                }
                elsif($sItem eq '0') {
                    $iTtl++;
                    $iAftP++ if($iP);
                }
                elsif($sItem eq '#') {
                    #$iTtl++;
                    $iAftP++ if($iP);
                }
                elsif($sItem eq '?') {
                    #$iTtl++;
                    $iAftP++ if($iP);
                }
            }
    #print "DATA:$iData\n";
            $iData *= 100.0 if($iPer);
            my $iDData = ($iFugouFlg)? abs($iData) : $iData+0;
            if($iBunFlg) {
                $sNumRes = sprintf("%0${iTtl}d", int($iDData));
            }
            else {
                if($iP) {
                    $sNumRes = sprintf("%0${iTtl}.${iAftP}f", $iDData);
                }
                else {
    #print "DATA:", $iDData, "\n";
                    $sNumRes = sprintf("%0${iTtl}.0f", $iDData);
                }
            }
            $sNumRes = AddComma($sNumRes) if($iCmmCnt > 0);
    #print "RES:$sNumRes\n";
            my $iLen = length($sNumRes);
            my $iPPos = -1;
            my $sRep;

            for(my $iIt=$#aRep; $iIt>=0;$iIt--) {
                my $rItem = $aRep[$iIt];
    #print "Rep:", $rItem->[0], "\n";
                if($rItem->[0] =~/([#0]*)([\.]?)([0#]*)([eE])([\+\-])([0#]+)/) {
                    substr($sFmtRes, $rItem->[1], $rItem->[2]) = 
                            MakeE($rItem->[0], $iData);
                }
                elsif($rItem->[0] =~ /\//) {
                    substr($sFmtRes, $rItem->[1], $rItem->[2]) = 
                        MakeBun($rItem->[0], $iData);
                }
                elsif($rItem->[0] eq '.') {
                    $iLen--;
                    $iPPos=$iLen;
                }
                elsif($rItem->[0] eq '+') {
                    substr($sFmtRes, $rItem->[1], $rItem->[2]) = 
                        ($iData > 0)? '+': (($iData==0)? '+':'-');
                }
                elsif($rItem->[0] eq '-') {
                    substr($sFmtRes, $rItem->[1], $rItem->[2]) = 
                        ($iData > 0)? '': (($iData==0)? '':'-');
                }
                elsif($rItem->[0] eq '@') {
                    substr($sFmtRes, $rItem->[1], $rItem->[2]) = $iData;
                }
                else {
                    if($iLen>0) {
                        if($iIt <= 0) {
                            $sRep = substr($sNumRes, 0, $iLen);
                        }
                        else {
                            my $iReal = length($rItem->[0]);
                            if($iPPos >= 0) {
                                my $sWkF = $rItem->[0];
                                $sWkF=~s/^#+//;
                                $iReal = length($sWkF);
                                $iReal = ($iLen <=$iReal)? $iLen:$iReal;
                            }
                            $sRep = substr($sNumRes, $iLen - $iReal, $iReal);
                            $iLen -=$iReal;
                        }
                    }
                    else {
                            $sRep = '';
                    }
    #print "REP:$sRep ", $rItem->[1], ":" ,$rItem->[2], "\n";
                    substr($sFmtRes, $rItem->[1], $rItem->[2]) = $sRep;
                }
            }
        }
    }
    else {
        for(my $iIt=$#aRep; $iIt>=0;$iIt--) {
            my $rItem = $aRep[$iIt];
            if($rItem->[0] eq '@') {
                substr($sFmtRes, $rItem->[1], $rItem->[2]) = $iData;
            }
            else {
                substr($sFmtRes, $rItem->[1], $rItem->[2]) = '';
            }
        }
    }
    return wantarray()? ($sFmtRes, $sColor) : $sFmtRes;
}
#------------------------------------------------------------------------------
# AddComma (for Spreadsheet::ParseExcel::Utility)
#------------------------------------------------------------------------------
sub AddComma($) {
    my($sNum) = @_;

    if($sNum=~ /^([^\d]*)(\d\d\d\d+)(\.*.*)$/) {
        my($sPre, $sObj, $sAft) =($1, $2, $3);
        for(my $i=length($sObj)-3;$i>0; $i-=3) {
            substr($sObj, $i, 0) = ',';
        }
        return $sPre . $sObj . $sAft;
    }
    else {
        return $sNum;
    }
}
#------------------------------------------------------------------------------
# MakeBun (for Spreadsheet::ParseExcel::Utility)
#------------------------------------------------------------------------------
sub MakeBun($$) {
    my($sFmt, $iData) = @_;
    my $iBunbo;
    my $iShou;

#1. Init
    $iShou = $iData - int($iData);
    return '' if($iShou == 0);
    my $sSWk;

#2.Calc BUNBO
#2.1 BUNBO defined
    if($sFmt =~ /\/(\d+)$/) {
        $iBunbo = $1;
        return sprintf("%d/%d", $iShou*$iBunbo, $iBunbo);
    }
    else {
#2.2 Calc BUNBO
        $sFmt =~ /\/(\?+)$/;
        my $iKeta = length($1);
        my $iSWk = 1;
        my $sSWk = '';
        my $iBunsi;
        for(my $iBunbo = 2;$iBunbo<10**$iKeta;$iBunbo++) {
            $iBunsi = int($iShou*$iBunbo + 0.5);
            my $iCmp = abs($iShou - ($iBunsi/$iBunbo));
            if($iCmp < $iSWk) {
                $iSWk =$iCmp;
                $sSWk = sprintf("%d/%d", $iBunsi, $iBunbo);
                last if($iSWk==0);
            }
        }
        return $sSWk;
    }
}
#------------------------------------------------------------------------------
# MakeE (for Spreadsheet::ParseExcel::Utility)
#------------------------------------------------------------------------------
sub MakeE($$) {
    my($sFmt, $iData) = @_;

    $sFmt=~/(([#0]*)[\.]?[#0]*)([eE])([\+\-][0#]+)/;
    my($sKari, $iKeta, $sE, $sSisu) = ($1, length($2), $3, $4);
    $iKeta = 1 if($iKeta<=0);

    my $iLog10 = 0;
    $iLog10 = ($iData == 0)? 0 : (log(abs($iData))/ log(10));
    $iLog10 = (int($iLog10 / $iKeta) + 
            ((($iLog10 - int($iLog10 / $iKeta))<0)? -1: 0)) *$iKeta;

    my $sUe = ExcelFmt($sKari, $iData*(10**($iLog10*-1)),0);
    my $sShita = ExcelFmt($sSisu, $iLog10, 0);
    return $sUe . $sE . $sShita;
}
#------------------------------------------------------------------------------
# LeapYear (for Spreadsheet::ParseExcel::Utility)
#------------------------------------------------------------------------------
sub LeapYear($) {
    my($iYear)=@_;
    return 1 if($iYear==1900); #Special for Excel
    return ((($iYear % 4)==0) && (($iYear % 100) || ($iYear % 400)==0))? 1: 0;
}
#------------------------------------------------------------------------------
# LocaltimeExcel (for Spreadsheet::ParseExcel::Utility)
#------------------------------------------------------------------------------
sub LocaltimeExcel($$$$$$;$$) {
    my($iSec, $iMin, $iHour, $iDay, $iMon, $iYear, $iMSec, $flg1904) = @_;

#0. Init
    $iMon++;
    $iYear+=1900;

#1. Calc Time
    my $iTime;
    $iTime =$iHour;
    $iTime *=60;
    $iTime +=$iMin;
    $iTime *=60;
    $iTime +=$iSec;
    $iTime += $iMSec/1000.0 if(defined($iMSec)) ;
    $iTime /= 86400.0;      #3600*24(1day in seconds)
    my $iY;
    my $iYDays;

#2. Calc Days
    if($flg1904) {
        $iY = 1904;
        $iTime--;         #Start from Jan 1st
        $iYDays = 366;
    }
    else {
        $iY = 1900;
        $iYDays = 366;  #In Excel 1900 is leap year (That's not TRUE!)
    }
    while($iY<$iYear) {
        $iTime += $iYDays;
        $iY++;
        $iYDays = (LeapYear($iY))? 366: 365;
    }
    for(my $iM=1;$iM < $iMon; $iM++){
        if($iM == 1 || $iM == 3 || $iM == 5 || $iM == 7 || $iM == 8
            || $iM == 10 || $iM == 12) {
            $iTime += 31;
        }
        elsif($iM == 4 || $iM == 6 || $iM == 9 || $iM == 11) {
            $iTime += 30;
        }
        elsif($iM == 2) {
            $iTime += (LeapYear($iYear))? 29: 28;
        }
    }
    $iTime+=$iDay;
    return $iTime;
}
#------------------------------------------------------------------------------
# ExcelLocaltime (for Spreadsheet::ParseExcel::Utility)
#------------------------------------------------------------------------------
sub ExcelLocaltime($;$)
{
  my($dObj, $flg1904) = @_;
  my($iSec, $iMin, $iHour, $iDay, $iMon, $iYear, $iwDay, $iMSec);
  my($iDt, $iTime, $iYDays);

  $iDt  = int($dObj);
  $iTime = $dObj - $iDt;

#1. Calc Days
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

#2. Calc Time
  $iDay = $iDt;
  $iTime += (0.0005 / 86400.0);
  $iTime*=24.0;
  $iHour = int($iTime);
  $iTime -= $iHour;
  $iTime *= 60.0;
  $iMin  = int($iTime);
  $iTime -= $iMin;
  $iTime *= 60.0;
  $iSec  = int($iTime);
  $iTime -= $iSec;
  $iTime *= 1000.0;
  $iMSec = int($iTime);
    
  return ($iSec, $iMin, $iHour, $iDay, $iMon-1, $iYear, $iwDay, $iMSec);
}
1;
__END__

=head1 NAME

Spreadsheet::ParseExcel::Utility - Utility function for Spreadsheet::ParseExcel

=head1 SYNOPSIS

    use strict;
    #Declare
    use Spreadsheet::ParseExcel::Utility qw(ExcelFmt ExcelLocaltime LocaltimeExcel);
    
    #Convert localtime ->Excel Time
    my $iBirth = LocaltimeExcel(11, 10, 12, 23, 2, 64);
                               # = 1964-3-23 12:10:11
    print $iBirth, "\n";       # 23459.5070717593
    
    #Convert Excel Time -> localtime
    my @aBirth = ExcelLocaltime($iBirth, undef);
    print join(":", @aBirth), "\n";   # 11:10:12:23:2:64:1:0
    
    #Formatting
    print ExcelFmt('yyyy-mm-dd', $iBirth), "\n"; #1964-3-23
    print ExcelFmt('m-d-yy', $iBirth), "\n";     # 3-23-64
    print ExcelFmt('#,##0', $iBirth), "\n";      # 23,460
    print ExcelFmt('#,##0.00', $iBirth), "\n";   # 23,459.51
    print ExcelFmt('"My Birthday is (m/d):" m/d', $iBirth), "\n";
                                      # My Birthday is (m/d): 3/23

=head1 DESCRIPTION

Spreadsheet::ParseExcel::Utility exports utility functions concerned with Excel format setting.

=head1 Functions

This module can export 3 functions: ExcelFmt, ExcelLocaltime and LocaltimeExcel.

=head2 ExcelFmt

$sTxt = ExcelFmt($sFmt, $iData [, $i1904]);

I<$sFmt> is a format string for Excel. I<$iData> is the target value.
If I<$flg1904> is true, this functions assumes that epoch is 1904.
I<$sTxt> is the result.

For more detail and examples, please refer sample/chkFmt.pl in this distribution.

ex.
  
=head2 ExcelLocaltime

($iSec, $iMin, $iHour, $iDay, $iMon, $iYear, $iwDay, $iMSec) = 
            ExcelLocaltime($iExTime [, $flg1904]);

I<ExcelLocaltime> converts time information in Excel format into Perl localtime format.
I<$iExTime> is a time of Excel. If I<$flg1904> is true, this functions assumes that
epoch is 1904.
I<$iSec>, I<$iMin>, I<$iHour>, I<$iDay>, I<$iMon>, I<$iYear>, I<$iwDay> are same as localtime.
I<$iMSec> means 1/1,000,000 seconds(ms).


=head2 LocaltimeExcel

I<$iExTime> = LocaltimeExcel($iSec, $iMin, $iHour, $iDay, $iMon, $iYear [,$iMSec] [,$flg1904])

I<LocaltimeExcel> converts time information in Perl localtime format into Excel format .
I<$iSec>, I<$iMin>, I<$iHour>, I<$iDay>, I<$iMon>, I<$iYear> are same as localtime.

If I<$flg1904> is true, this functions assumes that epoch is 1904.
I<$iExTime> is a time of Excel. 

=head2 AUTHOR

Kawai Takanori (Hippo2000) kwitknr@cpan.org

    http://member.nifty.ne.jp/hippo2000/            (Japanese)
    http://member.nifty.ne.jp/hippo2000/index_e.htm (English)

=head1 SEE ALSO

Spreadsheet::ParseExcel, Spreadsheet::WriteExcel

=head1 COPYRIGHT

This module is part of the Spreadsheet::ParseExcel distribution.

=cut
