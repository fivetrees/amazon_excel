#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use Spreadsheet::Read;
use Win32::OLE;
use File::Copy;


say "";
# 产生wish随机数
# wish关键词文件 wish_tags.xlsx
my $wish_book = Spreadsheet::Read->new("wish_tags.xlsx");
my $wish_sheet = $wish_book->sheet(1);
my %h;
my (@h, @tags);
my $tag;
while (@h < 9) {
    my $r = int rand(668);
    push @h, $r if (!exists $h{$r} and $r > 0);
    $h{$r} = 1;
}
push @tags, "fashion";
push @tags, $wish_sheet->cell("A$_") for(@h);
$tag=join(", ",@tags);
say "$tag";