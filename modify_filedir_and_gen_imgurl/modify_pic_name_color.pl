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


#����������ͼƬ�ļ� ����1��ʼ����
for my $dir (@dirs) {
	chdir "$mydir";
	chdir "$dir";
	say "$mydir/$dir";
	my @color_dirs = <*>;
	my $num = 1;
	for(@color_dirs) {
		chdir "$mydir";
		chdir "$dir";
		chdir $_;
		say "$mydir/$dir/$_";
		for(glob '*'){
			unless(-d) {
				if($num < 10) {
					say "modify $_ ==> 0$num.jpg";
					rename "$_", "fskj0$num.jpg";
				} elsif($num >= 10) {
					say "modify $_ ==> $num.jpg";
					rename "$_", "fskj$num.jpg";
				}
				$num++;
			}
		}
	}
	
}


for my $dir (@dirs) {
	chdir "$mydir";
	chdir "$dir";
	say "$mydir/$dir";
	my @color_dirs = <*>;
	my $num = 1;
	for(@color_dirs) {
		chdir "$mydir";
		chdir "$dir";
		chdir $_;
		say "$mydir/$dir/$_";
		for(glob '*'){
			unless(-d) {
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
	
}


