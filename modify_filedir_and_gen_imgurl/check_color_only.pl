#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use Encode; 
use File::Find;
use File::Basename;


#���嵱ǰĿ¼
my $mydir = "E:/tool/modify_filedir_and_gen_imgurl/imgdir";


chdir "$mydir";
my @dirs = <*>;
my @imgsizes;

#����������ͼƬ�ļ� ����1��ʼ����
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