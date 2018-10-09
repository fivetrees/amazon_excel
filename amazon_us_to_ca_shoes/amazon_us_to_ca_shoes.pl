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
my $Book11 = $Excel->Workbooks->Open("E:\\tool\\amazon_us_to_ca_shoes\\us.xlsx") //  die "Can not open $filename book\n" ;	
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
copy("NOT_DEL_AMAZON_CA_SHOES.xlsx", "$newxlsx") or die "Copy NOT_DEL_AMAZON_CA_SHOES.xlsx failed: $!";

# 打开新文档 写数据
my $Book22 = $Excel->Workbooks->Open("E:\\tool\\amazon_us_to_ca_shoes\\$newxlsx") //  die "Can not open $newxlsx book\n" ;	
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
A A
B B
C C
D D
E E
F F
H H
I I
J J
K K
L L
M M
N N
O O
P P
Q Q
R R
S S
T T
U U
V V
W W
X X
Y Y
Z Z
AP AA
AQ AB
AR AD
AS AC
AT AF
AU AG
AV AH
AX AI
AY AJ
AZ AK
BA AL
BB AM
BC AO
BD AP
BE AQ
BF AR
BG AS
BH AT
BI AU
BJ AV
BK AW
BL AX
BM AY
BN AZ
BO BA
BP BB
BQ BC
BR BD
BS BE
BT BF
BU BG
BV BH
BW BI
BX BJ
BY BK
BZ BL
CA BM
CC BN
CD BO
CF BP
CG BQ
CH BR
CI BS
CJ BT
CK BU
CL BV
CM BW
CN BX
CQ BY
CR BZ
CS CA
CT CB
DP CD
DQ CE
DR CF
DS CG
); 

#DM CF
#DZ CE

while(my($key, $value) = each %cphash) {
	&copy_paste_multi("$key", "$value");
}

#excel多行赋值
#for(4..$line) {
#	$Sheet22->Cells($_,7)->{Value} = "Shoes";
#}
$Sheet22->Range("G4:G$line")->{Value} = "Shoes";
								   
								   
$Book11->Close;
undef $Book11;

$Book22->Save;
$Book22->Close;
undef $Book22;

undef $Excel;


=pod
US	CA Shoes	
A	A	
B	B	
C	C	
D	D	
E	E	
F	F	
G	G   Shoes 	内容不一样，CA 的G列直接设置填Shoes
H	H	
I	I	
J	J	
K	K	
L	L	
M	M	
N	N	
O	O	
P	P	
Q	Q	
R	R	
S	S	
T	T	
U	U	
V	V	
W	W	
X	X	
Y	Y	
Z	Z	
AA		
AB		
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
AP	AA	
AQ	AB	
AR	AD	
AS	AC	
AT	AF	
AU	AG	
AV	AH	
AW		
AX	AI	
AY	AJ	
AZ	AK	
BA	AL	
BB	AM	
	AN	
BC	AO	
BD	AP	
BE	AQ	
BF	AR	
BG	AS	
BH	AT	
BI	AU	
BJ	AV	
BK	AW	
BL	AX	
BM	AY	
BN	AZ	
BO	BA	
BP	BB	
BQ	BC	
BR	BD	
BS	BE	
BT	BF	
BU	BG	
BV	BH	
BW	BI	
BX	BJ	
BY		
BZ	BK	
CA	BL	
CB		
	BM	
CC	BN	
CD	BZ	
CE	BO	
CF	BP	
CG	BQ	
CH	BR	
CI	BS	
CJ	BT	
CK	BU	
CL		
CM		
CN	BV	
CO	BW	
CP	BX	
CQ	BY	
CR		
CS		
CT		
CU		
CV		
CW		
CX		
CY		
CZ		
DA		
DB		
DC		
DD		
DE		
DF		
DG		
DH		
DI		
DJ	CF	
DK		
DL		
DM	CA	
DN	CB	
DO	CC	
DP	CD	
DQ		
DR		
DS		
DT		
DU		
DV		
DW	CE	
DX		
DY		
DZ		
EA		
EB		
EC		
ED		
EF		
EG		
EH		
=cut