#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use Encode; 
use File::Find;
use File::Basename;



#���嵱ǰĿ¼
my $mydir = "E:/tool/modify_used_csv/imgdir";

chdir "$mydir";
my @dirs = <*>;

#�����ļ�ǰ׺used-	
for(@dirs) {
	say;
	rename "$_", "used-$_";
}



