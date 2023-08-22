#!/usr/bin/perl

package CommHTTP;

use strict;
use utf8;
use LWP::UserAgent;
use HTTP::Cookies;

no warnings;

sub new { 
	my $class = shift;    
	my $self = {
        "ua" => http(),
        "url" => ""
    };
    
    foreach(@_) {
        if(ref $_ eq "HASH"){
            foreach my $k(keys %{$_}){
                $self->{$k} = $_->{$k};
            }
        } 
    }
	bless $self, $class;
	return $self; 
}

sub simpleGET
{
    my $self = shift;
    my $successCallBack = shift;
    my $failCallBack = shift;

    my $ref = $self->{"ua"}->get($self->{"url"});
    my $obj = $ref->content();

    if($ref->is_success) {
        if(ref $successCallBack eq "CODE") {
            $successCallBack->($obj);
        }
    } else {
        if(ref $failCallBack eq "CODE") {
            $failCallBack->();
        }
    }
}

sub simplePOST
{
    my $self = shift;
    my $postData = shift;
    my $successCallBack = shift;
    my $failCallBack = shift;
            
    # 发送请求

    my $ref = $self->{"ua"}->post($self->{"url"}, $postData);
    my $obj = $ref->content();
    
    if($ref->is_success) {
        if(ref $successCallBack eq "CODE") {
            $successCallBack->($obj);
        }
    } else {
        if(ref $failCallBack eq "CODE") {
            $failCallBack->();
        }
    } 
}   

sub http
{
    my $uagent = LWP::UserAgent->new(use_eval => 1);
    $uagent -> agent('Mozilla/5.0 (Windows NT 5.1; rv:15.0) Gecko/20100101 Firefox/15.0');
    $uagent -> timeout(60);

    my $cookie = HTTP::Cookies->new;
    $uagent -> cookie_jar($cookie);
   
    return $uagent;
}

1;