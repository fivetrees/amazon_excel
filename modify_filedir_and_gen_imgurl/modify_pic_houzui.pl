#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use Encode; 
use File::Find;
use File::Basename;

# �޸�ͼƬ�ĺ�׺��

#���嵱ǰĿ¼
my $mydir = "E:/tool/modify_filedir_and_gen_imgurl/imgdir";


chdir "$mydir";
my @dirs = <*>;


#�޸�ͼƬ�ĺ�׺��
for(@dirs) {
	chdir "$mydir";
	chdir "$_";
	say "$mydir/$_";
	for(glob '*'){
		unless(-d) {
			my $old_name = $_;
			if($_ =~ /JPG/) {
				$_ =~ s/JPG/jpg/g;
				rename "$old_name", "$_";
				say "rename $old_name $_";
			}
		}
	}
}
=pod
#����������ͼƬ�ļ� ����1��ʼ����
for(@dirs) {
	chdir "$mydir";
	chdir "$_";
	my $num = 1;
	say "$mydir/$_";
	for(glob '*'){
		unless(-d) {
			my $old_name = $_;
			if($_ =~ /JPG/) {
				$_ =~ s/JPG/jpg/g;
				rename "$old_name", "$_";
				say "rename $old_name $_";
			}
		}
	}
}
=cut

