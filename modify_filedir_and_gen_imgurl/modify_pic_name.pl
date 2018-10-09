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


#查找Thumbs.db并删除
for(@dirs) {
	chdir "$mydir";
	chdir $_;
	for(glob '*'){
		unless(-d) {
			if("$_" eq "Thumbs.db") {
				unlink $_;
				say "Thumbs.db ...";
			}
		}
	}
}

#重命名所有图片文件 数字1开始命名
	
for(@dirs) {
	chdir "$mydir";
	chdir "$_";
	my $num = 1;
	say "$mydir/$_";
	for(glob '*'){
		unless(-d) {
			next if(/txt/i);
			if($num < 10) {
				say "modify $_ ==> fs0$num.jpg";
				rename "$_", "fs0$num.jpg";
			} elsif($num >= 10) {
				say "modify $_ ==> fs$num.jpg";
				rename "$_", "fs$num.jpg";
			}
			$num++;
		}
	}
}

for(@dirs) {
	chdir "$mydir";
	chdir "$_";
	my $num = 1;
	say "$mydir/$_";
	for(glob '*'){
		unless(-d) {
			next if(/txt/i);
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

