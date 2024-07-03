#!/usr/bin/perl

use strict;
use utf8;
use myExcel;
use warnings;

my $oo = myExcel->new();
$oo -> initExcel();

$oo -> {"head"} = ["你", "好"];
$oo -> {"cols"} = ["A", "B"];

$oo -> {"XLS_data"} = [
    ["9527", "80086"], 
    ["3306", "65535"]
];

# 
$oo -> insertExcel(1, sub{
    
    my $sheet = shift;
    
    # 单元格宽度    
    $sheet -> Columns('A') -> {ColumnWidth} = 20;
    $sheet -> Columns('B') -> {ColumnWidth} = 20;
    $sheet -> Columns('C:D') -> {ColumnWidth} = 15;

    # 对齐(左)
    $sheet -> Columns('B') -> {HorizontalAlignment} = -4131;

    # 首行格式
    $sheet -> Range('A1:D1') -> {RowHeight} = 22;
    $sheet -> Range('A1:D1') -> {Interior} -> {Color} = 0xFFAA88;
    $sheet -> Range('A1:D1') -> {Font} -> {Color} = 0x000000;
    $sheet -> Range('A1:D1') -> {Font} -> {Bold} = 1;
    $sheet -> Range('A1:D1') -> {Borders} -> {LineStyle} = 1;
    $sheet -> Range('A1:D1') -> {Borders} -> {Color} = 0x555555;

    # 格式            
    $sheet -> Columns('A:B') -> {'NumberFormatLocal'} = "@";
}, sub {
    
    my $sheet = shift;
    
    # 单元格宽度    
    $sheet -> Range('A2:B3') -> {RowHeight} = 30;
    
}) 