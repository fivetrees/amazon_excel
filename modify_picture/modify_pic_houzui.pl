#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use Encode; 


# �޸�ͼƬ�ĺ�׺��

#���嵱ǰĿ¼
my $mydir = "E:/tool/modify_picture/imgdir";


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


