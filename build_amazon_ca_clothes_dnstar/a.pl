#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use Array::Utils ":all";
my @all = qw(A B C D E F G H I J K L M N O P Q R S T U V W X Y Z AA AB AC AD AE AF AG AH AI AJ AK AL AM AN AO AP AQ AR AS AT AU AV AW AX AY AZ BA BB BC BD BE BF BG BH BI BJ BK BL BM BN BO BP BQ BR BS BT BU BV BW BX BY BZ CA CB CC CD CE CF CG CH CI CJ CK CL CM CN CO CP CQ CR CS CT CU CV CW CX CY);
my @ignore = qw(A H I J L M J AM AS AT AU AV AW BF BG BH BO BP CD CE);
my @c = array_diff(@all,@ignore);
for(@c) {
	print "$_ ";
}


