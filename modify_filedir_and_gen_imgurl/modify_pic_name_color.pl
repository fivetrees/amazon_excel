#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use Encode; 
use File::Find;
use File::Basename;

# 修改递归目录下的图片名字 从01.jpg开始

#定义当前目录
my $mydir = "E:/tool/modify_filedir_and_gen_imgurl/imgdir";


chdir "$mydir";
my @dirs = <*>;


#重命名所有图片文件 数字1开始命名
for my $dir (@dirs) {
	chdir "$mydir";
	chdir "$dir";
	say "$mydir/$dir";
	my @color_dirs = <*>;
	my $num = 1;
	for(@color_dirs) {
		chdir "$mydir";
		chdir "$dir";
		chdir $_;
		say "$mydir/$dir/$_";
		for(glob '*'){
			unless(-d) {
				if($num < 10) {
					say "modify $_ ==> 0$num.jpg";
					rename "$_", "fskj0$num.jpg";
				} elsif($num >= 10) {
					say "modify $_ ==> $num.jpg";
					rename "$_", "fskj$num.jpg";
				}
				$num++;
			}
		}
	}
	
}


for my $dir (@dirs) {
	chdir "$mydir";
	chdir "$dir";
	say "$mydir/$dir";
	my @color_dirs = <*>;
	my $num = 1;
	for(@color_dirs) {
		chdir "$mydir";
		chdir "$dir";
		chdir $_;
		say "$mydir/$dir/$_";
		for(glob '*'){
			unless(-d) {
				if($num < 10) {
					say "modify $_ ==> 0$num.jpg";
					rename "$_", "0$num.jpg";
				} elsif($num >= 10) {
					say "modify $_ ==> $num.jpg";
					rename "$_", "$num.jpg";
				}
				$num++;
			}
		}
	}
	
}


