#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use Encode; 
use File::Find;
use File::Basename;


unlink "a.txt", if(-f "a.txt");
open DIRNAME_FH, ">a.txt";

#定义当前目录
my $mydir = "E:/tool/modify_filedir_and_gen_imgurl/imgdir";


say "chdir dir fail" and exit unless(chdir "$mydir");
my @dirs = <*>;
	
for(@dirs) {
	chdir "$mydir";
	chdir "$_";
	my $basename = basename($_);
	say DIRNAME_FH "$basename";
	say "$basename";
}

close DIRNAME_FH;


