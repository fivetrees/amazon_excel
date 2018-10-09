#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use File::Find;
use File::Basename;
use Image::Size;

unlink "imgsize.txt", if(-f "imgsize.txt");
open IMGSIZE_FH,">imgsize.txt";

# 修改递归目录下的图片名字 从01.jpg开始

#定义当前目录
my $mydir = "E:/tool/modify_filedir_and_gen_imgurl/imgdir";

#读取当面目录中所有文件名和目录名
my (@files, @dirs, @all, @imgsizes, @del_dirs);
finddepth sub {(push @all, $File::Find::name), $File::Find::name , '/', $_ }, qq($mydir);

#对提取的所有文件名和目录名进行文件和目录分类
for(@all) {
	if(-d) {
		push @dirs, $_;
	} elsif(-f) {
		push @files, $_;
	}
}

my $imgnum = 0;
#重命名所有图片文件 数字1开始命名
for my $dir (@dirs) {
	chdir $mydir;
	chdir $dir;
	my $file_num = 0;

	
	for (glob '*'){
		unless(-d) {
			#去除txt
			next if($_ =~ /.txt/);
			my ($x, $y) = imgsize("$_");

			$imgnum++;
			my $base_file = basename($dir);
			my $imginfo = "${x}_${y}";
			say IMGSIZE_FH "${x}_${y} :: $_ $base_file";
			#say "${x} * ${y} :: $base_file";
			#say "${x} * ${y} :: $_ :: $dir";
			unless ( grep { $_ eq $imginfo } @imgsizes ){
				#say "fffff";
				push @imgsizes, $imginfo;
			}
		}
		
		$file_num++;
	}
}


say "only exist below size";
say " ";
for(@imgsizes){
	chomp;
	#print;
	my($x, $y) = split '_', $_;
	my $chuchu = $x/$y;
	say "$x $y $chuchu";
	#say "$x $y";
}

say "imgnum $imgnum";

close IMGSIZE_FH;



=pod

use Image::Magick;

my $p = Image::Magick-&gt;new;# 实例化对象

$p-&gt;Read($pic1); # 读取图片
$p-&gt;Resize(geometry=&gt;'165x105+0+0'); # 修改尺寸，中间的“x”就是英文字母“x”
$p-&gt;Write($pic1_name); # 存入新地址
@$p = (); # 清空缓存

$p-&gt;Read($pic2);
$p-&gt;Resize(geometry=&gt;'165x105+0+0');
$p-&gt;Write($pic2_name);
@$p = ();

$p-&gt;Read($pic3);
$p-&gt;Resize(geometry=&gt;'165x105+0+0');
$p-&gt;Write($pic3_name);
@$p = ();
=cut