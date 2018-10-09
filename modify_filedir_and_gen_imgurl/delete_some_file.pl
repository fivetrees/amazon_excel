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

#读取当面目录中所有文件名和目录名
my (@files, @dirs, @all);
finddepth sub {(push @all, $File::Find::name), $File::Find::name , '/', $_ }, qq($mydir);

#对提取的所有文件名和目录名进行文件和目录分类
for(@all) {
	if(-d) {
		push @dirs, $_;
	} elsif(-f) {
		push @files, $_;
	}
}

#查找Thumbs.db并删除
for(@dirs) {
	chdir $_;
	my $num = 1;
	for(glob '*'){
		unless(-d) {
			if("$_" eq "Thumbs.db") {
				unlink $_;
				say "Thumbs.db ...";
			}
		}
	}
}


