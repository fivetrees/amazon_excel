#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use WWW::Google::Translate;
 
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

