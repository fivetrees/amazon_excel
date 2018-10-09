#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use Encode; 
use File::Find;
use File::Basename;



#定义当前目录
my $mydir = "E:/tool/modify_filedir_and_gen_imgurl/imgdir";


say "chdir dir fail" and exit unless(chdir "$mydir");
my @dirs = <*>;



for(@dirs) {
	chdir "$mydir";
	chdir "$_";
	my $num = 1;
	my $basename = basename($_);
	say "$mydir/$_ $basename";
}



