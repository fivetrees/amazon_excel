#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use Encode; 
use File::Find;
use File::Basename;

# �޸�Ŀ¼��  Ŀ¼��Ϊ���ĵģ�ȥ������

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

for(@dirs) {
	my $dirname = dirname($_);
	my $basename = basename($_);
	say "$dirname :: $basename";
	
}

