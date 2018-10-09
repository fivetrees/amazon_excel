#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use Encode; 
use File::Find;
use File::Basename;

# �޸ĵݹ�Ŀ¼�µ�ͼƬ���� ��01.jpg��ʼ

#���嵱ǰĿ¼
my $mydir = "E:/tool/modify_filedir_and_gen_imgurl/imgdir";

#��ȡ����Ŀ¼�������ļ�����Ŀ¼��
my (@files, @dirs, @all);
finddepth sub {(push @all, $File::Find::name), $File::Find::name , '/', $_ }, qq($mydir);

#����ȡ�������ļ�����Ŀ¼�������ļ���Ŀ¼����
for(@all) {
	if(-d) {
		push @dirs, $_;
	} elsif(-f) {
		push @files, $_;
	}
}

#����Thumbs.db��ɾ��
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


