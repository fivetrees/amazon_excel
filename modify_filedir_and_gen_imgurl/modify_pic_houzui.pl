#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use Encode; 
use File::Find;
use File::Basename;

# 修改图片的后缀名

#定义当前目录
my $mydir = "E:/tool/modify_filedir_and_gen_imgurl/imgdir";


chdir "$mydir";
my @dirs = <*>;


#修改图片的后缀名
for(@dirs) {
	chdir "$mydir";
	chdir "$_";
	say "$mydir/$_";
	for(glob '*'){
		unless(-d) {
			my $old_name = $_;
			if($_ =~ /JPG/) {
				$_ =~ s/JPG/jpg/g;
				rename "$old_name", "$_";
				say "rename $old_name $_";
			}
		}
	}
}
=pod
#重命名所有图片文件 数字1开始命名
for(@dirs) {
	chdir "$mydir";
	chdir "$_";
	my $num = 1;
	say "$mydir/$_";
	for(glob '*'){
		unless(-d) {
			my $old_name = $_;
			if($_ =~ /JPG/) {
				$_ =~ s/JPG/jpg/g;
				rename "$old_name", "$_";
				say "rename $old_name $_";
			}
		}
	}
}
=cut

