#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use Encode; 
use File::Find;
use File::Basename;

#��ȡ��ǰĿ¼�µ�Ŀ¼����ͼƬ�ļ��� ������Ӧ��imgurl����
# example imgurl http://www.feisenkj.com/img/wh/41251/03.jpg

#���嵱ǰĿ¼
my $mydir = "E:/tool/modify_filedir_and_gen_imgurl/imgdir";

open IMGURL_FH,">imgurl.txt";

#��ȡ����Ŀ¼�������ļ�����Ŀ¼��
my (@files, @dirs, @all);
finddepth sub {(push @all, $File::Find::name), $File::Find::name , '/', $_ }, qq($mydir);

#����ȡ�������ļ�����Ŀ¼�������ļ���Ŀ¼����
for(@all) {
	#say;
	if(-d) {
		push @dirs, $_;
	} elsif(-f) {
		push @files, $_;
	}
}


# ���뵽Ŀ¼����ͼƬ����
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