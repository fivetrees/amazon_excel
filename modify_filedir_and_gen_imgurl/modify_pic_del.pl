#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use Encode; 
use File::Find;
use File::Basename;

# 删除图片保留txt文件

#定义当前目录
my $mydir = "E:/tool/modify_filedir_and_gen_imgurl/imgdir";


chdir "$mydir";
my @dirs = <*>;


#重命名所有图片文件 数字1开始命名
	
for(@dirs) {
	chdir "$mydir";
	chdir "$_";
	my $num = 1;
	say "$mydir/$_";
	for(glob '*'){
		unless(-d) {
			if($_ =~ /jpg|JPG/) {
				unlink $_;
				say "del $_ ......";
			}
		}
	}
}



