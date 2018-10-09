#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use IO::Stream;
use IO::Stream::Proxy::SOCKSv5;
use WWW::Google::Translate;

IO::Stream->new({
    plugin => [
        proxy   => IO::Stream::Proxy::SOCKSv5->new({
            host    => '127.0.0.1',
            port    => 1080,
        }),
    ],
});


my $wgt = WWW::Google::Translate->new(
    {   key            => '<Your API key here>',
        default_source => 'en',   # optional
        default_target => 'ja',   # optional
    }
);
 
my $r = $wgt->translate( { q => 'My hovercraft is full of eels' } );
 
for my $trans_rh (@{ $r->{data}->{translations} }) {
 
    print $trans_rh->{translatedText}, "\n";
}