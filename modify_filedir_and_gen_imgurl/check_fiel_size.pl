#!/usr/bin/perl

use 5.010;
use strict;
use warnings;

#unlink "imgsize.txt", if(-f "imgsize.txt");

#定义当前目录
my $mydir = "E:/tool/modify_filedir_and_gen_imgurl/imgdir";

chdir $mydir;

for (glob '*'){
	my @args = stat ($_);
	my $size = $args[7];
	
	say "$_ $size";
}

