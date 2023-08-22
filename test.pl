#!/usr/bin/perl

use comm;
use comm_http;

print $comm::curDir;

# TimeCost

my $cost = Comm::costTime();

# Task

my $http = CommHTTP->new();
$http->{"url"} = "https://www.baidu.com";

$http->simpleGET(sub(){
    printf("%s s\n", $cost->());
})