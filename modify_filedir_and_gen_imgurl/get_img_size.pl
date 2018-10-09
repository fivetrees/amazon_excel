#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use File::Find;
use File::Basename;
use Image::Size;

unlink "imgsize.txt", if(-f "imgsize.txt");
open IMGSIZE_FH,">imgsize.txt";

# �޸ĵݹ�Ŀ¼�µ�ͼƬ���� ��01.jpg��ʼ

#���嵱ǰĿ¼
my $mydir = "E:/tool/modify_filedir_and_gen_imgurl/imgdir";

#��ȡ����Ŀ¼�������ļ�����Ŀ¼��
my (@files, @dirs, @all, @imgsizes, @del_dirs);
finddepth sub {(push @all, $File::Find::name), $File::Find::name , '/', $_ }, qq($mydir);

#����ȡ�������ļ�����Ŀ¼�������ļ���Ŀ¼����
for(@all) {
	if(-d) {
		push @dirs, $_;
	} elsif(-f) {
		push @files, $_;
	}
}

my $imgnum = 0;
#����������ͼƬ�ļ� ����1��ʼ����
for my $dir (@dirs) {
	chdir $mydir;
	chdir $dir;
	my $file_num = 0;

	
	for (glob '*'){
		unless(-d) {
			#ȥ��txt
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

my $p = Image::Magick-&gt;new;# ʵ��������

$p-&gt;Read($pic1); # ��ȡͼƬ
$p-&gt;Resize(geometry=&gt;'165x105+0+0'); # �޸ĳߴ磬�м�ġ�x������Ӣ����ĸ��x��
$p-&gt;Write($pic1_name); # �����µ�ַ
@$p = (); # ��ջ���

$p-&gt;Read($pic2);
$p-&gt;Resize(geometry=&gt;'165x105+0+0');
$p-&gt;Write($pic2_name);
@$p = ();

$p-&gt;Read($pic3);
$p-&gt;Resize(geometry=&gt;'165x105+0+0');
$p-&gt;Write($pic3_name);
@$p = ();
=cut