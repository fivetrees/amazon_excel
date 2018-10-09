#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use Encode; 
use File::Find;
use File::Basename;

# 修改目录名  目录名为中文的，去除中文

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


for(@dirs) {
	my $dirname = dirname($_);
	my $basename = basename($_);
	#say "$dirname :: $basename";
	
	#修改#为空
	if($basename =~ /#/) {
		my $jingname = "$basename";
		$basename =~ s/#.*$//g;
		rename "$jingname", "$basename";
	}
	
	#去除中文名，去除为空就加随机数
	if($basename =~ /[\x80-\xFF]{2}/) {
		say "$basename";
		my $oldname = "$basename";
		$basename =~ s/[\x80-\xFF]{2}//g;
		if($basename eq "") {
			say "$basename nulllll";
			$basename = int rand(100000);
			chdir $dirname;
			rename "$oldname", "$basename";
		} else {
			chdir $dirname;
			rename "$oldname", "$basename";
		}
	}
	
}


