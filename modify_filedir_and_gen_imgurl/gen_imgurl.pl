#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use Encode; 
use File::Find;
use File::Basename;

#获取当前目录下的目录名和图片文件名 生成相应的imgurl链接
# example imgurl http://www.feisenkj.com/img/wh/41251/03.jpg

#定义当前目录
my $mydir = "E:/tool/modify_filedir_and_gen_imgurl/imgdir";

open IMGURL_FH,">imgurl.txt";

#读取当面目录中所有文件名和目录名
my (@files, @dirs, @all);
finddepth sub {(push @all, $File::Find::name), $File::Find::name , '/', $_ }, qq($mydir);

#对提取的所有文件名和目录名进行文件和目录分类
for(@all) {
	#say;
	if(-d) {
		push @dirs, $_;
	} elsif(-f) {
		push @files, $_;
	}
}


# 进入到目录生成图片链接
for(@dirs) {
	chdir $_;
	my $basedir = basename($_);
	next if($basedir eq "gen_images_url");
	say IMGURL_FH "";
	say IMGURL_FH "";
	say IMGURL_FH "";
	for my $img (glob '*'){
		my $basename = basename($img);
		#say "http://www.feisenkj.com/img/wh/$basedir/$basename";
		say IMGURL_FH "http://www.feisenkj.com/img/wh/$basedir/$basename";
	}

}

close IMGURL_FH;