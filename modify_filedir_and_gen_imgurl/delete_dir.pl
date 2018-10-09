#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use Encode; 
use File::Find;
use File::Basename;

# 删除目录

#定义当前目录
my $mydir = "E:/tool/modify_filedir_and_gen_imgurl/imgdir";

chdir "$mydir";

#需要删除的目录
my @del_dir = qw(LC410077-10 LC410077-3 LC410077-7 LC410095-1 LC410095-14 LC410095-2 LC410095-3 LC410095-4 LC410095-5 LC410095-6 LC410095-7 LC410095-8 LC41041-1 LC41041-1 LC41068-1 LC41068-2 LC41071 LC41221 LC41251 LC41291 LC41430 LC41430-1 LC41633 LC41633-14 LC41650-1 LC41650-2 LC41651 LC41651-1 LC41654 LC41661 LC41661-2 LC41820 LC40902);

#删除目录里面的文件
for(@del_dir) {
	chdir "$mydir";
	chdir "$_";
	say "$mydir/$_";
	for(glob '*'){
		unless(-d) {
			unlink $_;
			say "del file $_ ......";
		}
	}
}

#删除目录
for(@del_dir) {
	chdir "$mydir";
	rmdir $_;
	say "del dir $_ ......";
}


