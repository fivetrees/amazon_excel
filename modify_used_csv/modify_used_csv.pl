#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use Encode; 
use File::Find;
use File::Basename;



#定义当前目录
my $mydir = "E:/tool/modify_used_csv/imgdir";

chdir "$mydir";
my @dirs = <*>;

#加上文件前缀used-	
for(@dirs) {
	say;
	rename "$_", "used-$_";
}



