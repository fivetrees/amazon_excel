#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use File::Find;
use File::Basename;

unlink "imgfilenum.txt", if(-f "imgfilenum.txt");
open IMGNUM_FH,">imgfilenum.txt";

# 修改递归目录下的图片名字 从01.jpg开始

#定义当前目录
my $mydir = "E:/tool/modify_filedir_and_gen_imgurl/imgdir";

#读取当面目录中所有文件名和目录名
my (@files, @dirs, @all, @zero_file, @one_file);
finddepth sub {(push @all, $File::Find::name), $File::Find::name , '/', $_ }, qq($mydir);

#对提取的所有文件名和目录名进行文件和目录分类
for(@all) {
	if(-d) {
		push @dirs, $_;
	} elsif(-f) {
		push @files, $_;
	}
}

#查找出零图片文件的目录
for my $dir (@dirs) {
	chdir $mydir;
	chdir $dir;
	my $file_num = 0;

	for (glob '*'){
		unless(-d) {
			#去除txt
			next if($_ =~ /.txt/);
		}
		$file_num++;
	}
	push @zero_file, $dir if($file_num == 0);
	push @one_file, $dir if($file_num == 1);
	say IMGNUM_FH "file_num is $dir :: $file_num";
}

for my $zerodir (@zero_file) {
	say "$zerodir";
	my $base_file = basename($zerodir);
	#删除零图片的非图片文件
	chdir $mydir;
	chdir $zerodir;
	for(glob '*') {
		unlink "$_";
		say "delete $_ ......";
	}
	
	#删除零图片的目录
	chdir $mydir;
	rmdir $base_file;
	say "delete $base_file ......";
}


#say "one picture file $_" for(@one_file);

close IMGNUM_FH;
