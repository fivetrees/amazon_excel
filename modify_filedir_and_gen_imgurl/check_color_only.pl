#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use Encode; 
use File::Find;
use File::Basename;


#定义当前目录
my $mydir = "E:/tool/modify_filedir_and_gen_imgurl/imgdir";


chdir "$mydir";
my @dirs = <*>;
my @imgsizes;

#重命名所有图片文件 数字1开始命名
for my $dir (@dirs) {
	chdir "$mydir";
	chdir "$dir";
	#say "$mydir/$dir";
	my @color_dirs = <*>;
	my $num = 1;
	for my $color (@color_dirs) {
		#chdir "$mydir";
		#chdir "$dir";
		#chdir $color;
		#say "$color";
		
		unless ( grep { $_ eq $color } @imgsizes ){
			push @imgsizes, $color;
		}

	}
	
}


for(@imgsizes) {
	say;
}