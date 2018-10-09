#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use File::Find;
use File::Basename;

unlink "imgfilenum.txt", if(-f "imgfilenum.txt");
open IMGNUM_FH,">imgfilenum.txt";

# �޸ĵݹ�Ŀ¼�µ�ͼƬ���� ��01.jpg��ʼ

#���嵱ǰĿ¼
my $mydir = "E:/tool/modify_filedir_and_gen_imgurl/imgdir";

#��ȡ����Ŀ¼�������ļ�����Ŀ¼��
my (@files, @dirs, @all, @zero_file, @one_file);
finddepth sub {(push @all, $File::Find::name), $File::Find::name , '/', $_ }, qq($mydir);

#����ȡ�������ļ�����Ŀ¼�������ļ���Ŀ¼����
for(@all) {
	if(-d) {
		push @dirs, $_;
	} elsif(-f) {
		push @files, $_;
	}
}

#���ҳ���ͼƬ�ļ���Ŀ¼
for my $dir (@dirs) {
	chdir $mydir;
	chdir $dir;
	my $file_num = 0;

	for (glob '*'){
		unless(-d) {
			#ȥ��txt
			next if($_ =~ /.txt/);
		}
		$file_num++;
	}
	push @zero_file, $dir if($file_num == 0);
	push @one_file, $dir if($file_num == 1);
	say IMGNUM_FH "file_num is $dir :: $file_num";
}

for my $zerodir (@zero_file) {
	say "$zerodir";
	my $base_file = basename($zerodir);
	#ɾ����ͼƬ�ķ�ͼƬ�ļ�
	chdir $mydir;
	chdir $zerodir;
	for(glob '*') {
		unlink "$_";
		say "delete $_ ......";
	}
	
	#ɾ����ͼƬ��Ŀ¼
	chdir $mydir;
	rmdir $base_file;
	say "delete $base_file ......";
}


#say "one picture file $_" for(@one_file);

close IMGNUM_FH;
