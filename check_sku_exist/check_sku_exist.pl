#!/usr/bin/perl

use 5.010;
use strict;
use warnings;


open SKU_FH, "<sku.txt";
open WENZHANG_FH, "<wenzhang.txt";
my @skus = <SKU_FH>;
my @wzs = <WENZHANG_FH>;

say for(@wzs);
for(@skus) {
	chomp;
	#say;
	my @results = grep(/$_/, @wzs);
}

close SKU_FH;
close WENZHANG_FH;

=pod
open FF,"a.txt";
my @text=<FF>;
my @results=grep(/\-a/,@text);
print @results;
=cut