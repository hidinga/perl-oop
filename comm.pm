#!/usr/bin/perl

package Comm;

use utf8;
use Time::HiRes qw(gettimeofday);
use File::Basename;
use File::Spec;

no warnings;
require Exporter;

@ISA = qw(Exporter);
@EXPORT = qw(curDir); 
@EXPORT_OK = qw(costTime); 

# 路径

our $cur_dir = File::Spec->rel2abs(dirname(__FILE__));
$Comm::curDir = $cur_dir =~ s/.*\<//r;

# 方法

sub costTime
{
    my $i = gettimeofday();
    return sub {
        my $j = gettimeofday();
        return sprintf("%.03f", $j-$i);
    }
}

1;