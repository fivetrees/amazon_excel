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


chdir "$mydir";
my @dirs = <*>;


#����Thumbs.db��ɾ��
for(@dirs) {
	chdir "$mydir";
	chdir $_;
	for(glob '*'){
		unless(-d) {
			if("$_" eq "Thumbs.db") {
				unlink $_;
				say "Thumbs.db ...";
			}
		}
	}
}

#����������ͼƬ�ļ� ����1��ʼ����
	
for(@dirs) {
	chdir "$mydir";
	chdir "$_";
	my $num = 1;
	say "$mydir/$_";
	for(glob '*'){
		unless(-d) {
			next if(/txt/i);
			if($num < 10) {
				say "modify $_ ==> fs0$num.jpg";
				rename "$_", "fs0$num.jpg";
			} elsif($num >= 10) {
				say "modify $_ ==> fs$num.jpg";
				rename "$_", "fs$num.jpg";
			}
			$num++;
		}
	}
}

for(@dirs) {
	chdir "$mydir";
	chdir "$_";
	my $num = 1;
	say "$mydir/$_";
	for(glob '*'){
		unless(-d) {
			next if(/txt/i);
			if($num < 10) {
				say "modify $_ ==> 0$num.jpg";
				rename "$_", "0$num.jpg";
			} elsif($num >= 10) {
				say "modify $_ ==> $num.jpg";
				rename "$_", "$num.jpg";
			}
			$num++;
		}
	}
}

