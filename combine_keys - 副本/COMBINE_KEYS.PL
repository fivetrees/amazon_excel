#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use File::Slurp;
use Array::Utils qw(:all);

my @tempkeys = read_file('keys.txt');
my @pinpai = read_file('pinpai.txt');

#去除数组中重复的元素
my %count;
my @keys = grep { ++$count{ $_ } < 2; } @tempkeys;

my @keys2 = @keys;
my @copykeys = @keys;
my @delkeys;
my @newkeys;
my @fakes;
my @fakeword;

chomp(@keys, @keys2, @copykeys);

#删除匹配到的元素
for my $kk (@copykeys){
	#keys2先去除本元素
	#chomp $kk;
	@keys2 = grep { $_ ne "$kk" } @keys2;
	my $res = grep /^\b$kk\b/i, @keys2;
	if($res == 0){
		#say "NO $kk";
		push @newkeys, "$kk\n";
	} else {
		#say "YES $kk";
		#@keys = grep { $_ ne "$kk" } @keys;
		push @delkeys, "$kk\n";
	}
	@keys2 = @keys;
}

@delkeys = sort @delkeys;
@newkeys = sort @newkeys;
#write_file( 'delkeys.txt', @delkeys );


#chomp(@newkeys, @pinpai);
chomp(@pinpai);
my %h;
my @nnewkeys;

for(@newkeys) {
	my $keyword = $_;
	for(@pinpai){
		if($keyword =~ /\b$_\b/i) {
			push @fakes, "$_ :: $keyword";
			push @fakeword, "$keyword";
		}
	}
}

my %count2;
@fakeword = grep { ++$count2{ $_ } < 2; } @fakeword;

@nnewkeys = array_diff(@newkeys,@fakeword);
push @nnewkeys, " \n";
push @nnewkeys, "END_END";

@fakes = sort @fakes;
#write_file( 'fakes.txt', @fakes );

@fakeword = sort @fakeword;
#write_file( 'fakeword.txt', @fakeword );

@nnewkeys = sort @nnewkeys;
write_file( 'newkeys.txt', @nnewkeys );