#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use Encode; 
use File::Find;
use File::Basename;

# ɾ��ͼƬ����txt�ļ�

#���嵱ǰĿ¼
my $mydir = "E:/tool/modify_filedir_and_gen_imgurl/imgdir";


chdir "$mydir";
my @dirs = <*>;


#����������ͼƬ�ļ� ����1��ʼ����
	
for(@dirs) {
	chdir "$mydir";
	chdir "$_";
	my $num = 1;
	say "$mydir/$_";
	for(glob '*'){
		unless(-d) {
			if($_ =~ /jpg|JPG/) {
				unlink $_;
				say "del $_ ......";
			}
		}
	}
}



