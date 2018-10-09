#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use Win32::OLE;
use File::Copy;


#先删除ERROR.txt
unlink "ERROR.txt", if(-f "ERROR.txt");


# 美国站xlsl 转换成加拿大xlsl
my $filename = "us.xlsx";


#初始化excel文件
my $Excel;
# use existing instance if Excel is already running
eval {$Excel = Win32::OLE->GetActiveObject('Excel.Application')};
die "Excel not installed" if $@;
unless (defined $Excel) {
    $Excel = Win32::OLE->new('Excel.Application', sub {$_[0]->Quit;})
          or die "Oops, cannot start Excel";
}
# 关掉Excel的提示，比如是否保存之类的。
$Excel->{DisplayAlerts} = 'False';

# 打开us.xlsx文件
my $Book11 = $Excel->Workbooks->Open("E:\\tool\\amazon_us_to_ca_clothes\\us.xlsx") //  die "Can not open $filename book\n" ;	
my $Sheet11 = $Book11->Worksheets(4);


#读取新文件名
my $psku	= $Sheet11->range("A4")->{Value};
my $ptitle 	= $Sheet11->range("B4")->{Value};
my @huohaos = split '-', $psku;
my $huohao 	= $huohaos[-1];
my @ptitles = split ' ', $ptitle;
my $brand   = $ptitles[0];

#拷贝ca excel文件
my $newxlsx = "amazon-ca-${brand}-${huohao}.xlsx";
copy("NOT_DEL_AMAZON_CA_CLOTHES.xlsx", "$newxlsx") or die "Copy NOT_DEL_AMAZON_CA_CLOTHES.xlsx failed: $!";

# 打开新文档 写数据
my $Book22 = $Excel->Workbooks->Open("E:\\tool\\amazon_us_to_ca_clothes\\$newxlsx") //  die "Can not open $newxlsx book\n" ;	
my $Sheet22 = $Book22->Worksheets(4);


my $line = 0;
for(4..200) {
	say "No Defined A$_" and last unless(defined $Sheet11->range("A$_")->{Value});
	$line = $_;
}
say "line is $line";

sub copy_paste_multi {
	my $t1 = shift;
	my $t2 = shift;
	
	$Sheet11->range("${t1}4:${t1}${line}")->copy();
	$Sheet22->range("${t2}4:${t2}${line}")->Select();
	$Sheet22->paste();
}

my %cphash = qw(
A  A
B  H
C  J
D  I
E  K
F  D
H  F
I  B
J  M
K  Z
L  R
M  AC
N  S
O  X
P  Q
Q  L
R  N
S  O
T  P
U  Y
V  AA
W  V
X  U
Y  T
Z  AB
AB W
AR AE
AS AD
AT AG
AU AH
AW AI
AX AF
AY AJ
AZ AN
BA AO
BB AP
BC AQ
BD AR
BE AM
BF AZ
BG AT
BH AU
BI AV
BJ AW
BK AX
BL AY
BM AS
BN BG
BO BE
BP BA
BQ BD
BS BC
BT BB
BU BK
BV BJ
BW BI
BX BH
BY BM
BZ BN
CA CA
CC BL
CF BP
CG BQ
CM BR
CN BS
CQ BT
CS BU
CU BV
CV BW
CY BX
CZ BY
DA BZ
DC CB
DD CC
DG CD
DH CE
DI CF
DL CG
DM CH
DN CI
DQ CJ
DR CK
DT CL
DW CN
DX CO
DY CP
DZ CQ
); 


while(my($key, $value) = each %cphash) {
	&copy_paste_multi("$key", "$value");
}


$Book11->Close;
undef $Book11;

$Book22->Save;
$Book22->Close;
undef $Book22;

undef $Excel;


=pod
US CA clothes
A  A
B  H
C  J
D  I
E  K
F  D
G
H  F
I  B
J  M
K  Z
L  R
M  AC
N  S
O  X
P  Q
Q  L
R  N
S  O
T  P
U  Y
V  AA
W  V
X  U
Y  T
Z  AB
AA 
AB W
AC 
AD 
AE
AF
AG
AH
AI
AJ
AK
AL
AM
AN
AO
AP
AQ
AR AE
AS AD
AT AG
AU AH
AV 
AW AI
AX AF
AY AJ
AZ AN
BA AO
BB AP
BC AQ
BD AR
BE AM
BF AZ
BG AT
BH AU
BI AV
BJ AW
BK AX
BL AY
BM AS
BN BG
BO BE
BP BA
BQ BD
BR
BS BC
BT BB
BU BK
BV BJ
BW BI
BX BH
BY BM
BZ BN
CA CA
CB 
CC BL
CD 
CE 
CF BP
CG BQ
CH 
CI
CJ 
CK
CL
CM BR
CN BS
CO
CP
CQ BT
CR
CS BU
CT
CU BV
CV BW
CW 
CX 
CY BX
CZ BY
DA BZ
DB 
DC CB
DD CC
DE 
DF
DG CD
DH CE
DI CF
DJ 
DK
DL CG
DM CH
DN CI
DO 
DP 
DQ CJ
DR CK
DS 
DT CL
DU 
DV
DW CN
DX CO
DY CP
DZ CQ
EA
EB
EC
ED
EE
=cut