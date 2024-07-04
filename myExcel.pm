#!/usr/bin/perl

package myExcel;

use utf8;
use Encode;
use Encode::CN;
use File::Basename;
use File::Spec;
use Win32::Console::ANSI;

no warnings;

sub new { 
	my $class = shift;    
	my $self = {
	    "excel" => excel(),
	 "workbook" => "",
	     "path" => File::Spec->rel2abs(dirname(__FILE__)) =~ s/.*\<//r,
	  "xlsPath" => "",
	     "head" => [],				# 需要插入的头 ["你好","谢谢"]
	     "rows" => 2,				# 需要插入的行 [2,2] 代码A列2行开始, C列2行开始
	     "cols" => [],				# 需要插入的列 [A,C]
	 "XLS_data" => [],				# 要插入的数据 [[1,2,3,4,5],[7,8]]
	   "forzen" => 0,				# 冻结行
	     "name" => "Sheet1",
	    "clear" => 0,
	     "file" => ""	
    };
    
    # 合并数据
    
    foreach(@_) {
        next if ref $_ ne "HASH";
        foreach my $k(keys %{$_}){
            $self->{$k} = $_->{$k};
        }
    }
    bless $self, $class;
    return $self; 
}

sub init {
    my $self = shift; 
    $self->{"excel"} = excel(); 
    $self->{"workbook"} = "";
    $self->{"path"} = File::Spec->rel2abs(dirname(__FILE__)) =~ s/.*\<//r;
    $self->{"xlsPath"} = "";
    $self->{"head"} = [];
    $self->{"rows"} = 2;
    $self->{"cols"} = [];
    $self->{"XLS_data"} = [];
    $self->{"forzen"} = 0;
    $self->{"name"} = "Sheet1";			 
    $self->{"clear"} = 0;			 
    $self->{"file"} = "";
    
    foreach(@_) {
        next if ref $_ ne "HASH";
        foreach my $k(keys %{$_}){
            $self->{$k} = $_->{$k};
        }
    }
}

sub excel
{
    use Win32::OLE::Variant;
    use Win32::OLE::Const 'Microsoft Excel';	
    my $excel = CreateObject Win32::OLE 'Excel.Application' or exit ('ERROR : Microsoft Excel Not Found ...');
    $excel -> {'EnableEvents'} = 0;
    $excel -> {'Visible'} = 1;
    $excel -> {'DisplayAlerts'} = 0;
    return $excel;
}

sub initExcel
{
    my $self = shift;
    my $num = shift || 3;				
    
    # 创建表格  
	
    $self->{"workbook"} = $self->{"excel"}->Workbooks->Add();
	my $c = $self->{"workbook"}->Worksheets()->{'Count'};
	if($c < $num && $num > 3) {
	    foreach($c+1..$num){
		$self->{"workbook"}->Worksheets->Add();
	    }
	} 
}

sub getExcel
{
    # getExcel("A")         -- 读取A列
    # getExcel("A2")        -- 读取A2单元格
    # getExcel(["A",2])     -- 读取A列从2行起

    my $self = shift;
    my @range = @_;         # 接收参数 C3, B2, D1
    my $xls_Data = [];	    # 返回结果
    my $xls_Path = encode("gbk", $self->{"xlsPath"});
    
    $self->{"excel"} -> {'Readonly'} = 1;
    $self->{"workbook"} = $self->{"excel"}->Workbooks->Open($xls_Path);
    
    # 读指定表
    
    # $sheet_rs = $self->{"workbook"} -> Worksheets(encode("gbk", $self->{"name"}));
    # $sheet_rs = $self->{"workbook"} -> Worksheets(1);
    my $sheet_rs = $self->{"workbook"}-> ActiveSheet();
    
    # 读取数据
    
    my $x = $sheet_rs->{'UsedRange'}->{'Rows'}->{'Count'};
    my $y = $sheet_rs->{'UsedRange'}->{'Columns'}->{'Count'};
     
    foreach(@range)
    {
        if(ref $_ eq "")
        {
            if($_ =~ /[0-9]$/){
                push(@{$xls_Data}, COM::dgbk($sheet_rs->Range($_)->{'Value'}));
            }
            if($_ =~ /^[A-Z]+$/){
                push(@{$xls_Data}, [map {COM::dgbk($_->[0])} @{$sheet_rs->Range(sprintf("%s1:%s%s",$_, $_, $x))->{'Value'}}]);
            }
        }
        if(ref $_ eq "ARRAY"){
            push(@{$xls_Data}, [map {COM::dgbk($_->[0])} @{$sheet_rs->Range(sprintf("%s%s:%s%s",$_->[0], $_->[1], $_->[0], $x))->{'Value'}}]);
        } 
    }
    
    # 关闭 excel
    
    $self->{"workbook"}->Close({SaveChanges => 0});
    $self->{"excel"}->Close();
    $self->{"excel"}->Quit();
    
    return $xls_Data;
}

sub insertExcel
{
    # 根据 cols 定义的列插入数据 insertExcel(1, sub{}, sub{})
    # 读取 cols = []          -- 写入列名称
    # 读取 head = []          -- 表头行数据
    # 读取 rows = 2           -- 表头行号
    # 读取 XLS_data = []      -- 写入的数据(gbk编码后数据)

    my $self = shift;
    my $num = shift || 1;	    # 插入第几个表格
    
    my $cbStart = shift;            # 开始前回调
    my $cbEnd = shift;              # 结束前回调
	
    # 参数检查
    
    if($self->{"workbook"} eq ""){
        return 1;
    }
    if(scalar @{$self->{"cols"}} < scalar @{$self->{"XLS_data"}}){
        return 1;
    }
	
    # 操作表格

    my $sheet_rs = $self->{"workbook"}-> Worksheets($num);
    $sheet_rs->{"NAME"} = encode("gbk", $self->{"name"});
	
    # 定义样式
	
    $sheet_rs -> Columns -> {Font} -> {Name} = encode("gbk", '微软雅黑');
    $sheet_rs -> Columns -> {Font} -> {Size} = 11;
    $sheet_rs -> Columns -> {HorizontalAlignment} = -4108;	# 居中
    $sheet_rs -> Columns -> {VerticalAlignment} = -4108;	# 居中
    $sheet_rs -> Columns -> {ColumnWidth} = 12; 
    $sheet_rs -> Rows(1) -> {WrapText} = 0;
     
    # 定义格式
     
    if($cbStart) {
        printf("\n%s\n\n", encode("gbk", "设定格式"));
        $cbStart->($sheet_rs);
    } else {
        printf("\n");
    }
    
    printf("%s \e[1;33m%s\e[0m %s [\e[1;33;41m%s\e[0m] ...\n\n", encode("gbk", "写入第"), $num, encode("gbk", "表格"), encode("gbk", $self->{"name"}));
	
    foreach(1..scalar @{$self->{"cols"}})
    {
        my $c = $self->{"cols"}->[$_-1];															# 列字母
        my $h = sprintf("%s%s", $c, ($self->{"rows"}-1));											# 标题定位块
        my $d = [map { [$_] } @{$self->{"XLS_data"}->[$_-1]}];                                      # 数据列
        my $r = sprintf("%s%s:%s%s", $c, $self->{"rows"}, $c, $self->{"rows"}+scalar(@{$d})-1);		# 数据定位块
        
	if(scalar @{$self->{"head"}} > 0){
            printf("%s [\e[1;31m%s\e[0m] %s ...\n", encode("gbk", "正在写入"), $h, encode("gbk", "数据"));
            $sheet_rs -> Range($h) -> {Value} = encode("gbk", $self->{"head"}->[$_-1]) if $self->{"head"}->[$_-1];
	}
        printf("%s [\e[1;31m%s\e[0m] %s ...\n", encode("gbk", "正在写入"), $r, encode("gbk", "数据"));
	$sheet_rs -> Range($r) -> {Value} = $d;
    }

    # 定义样式
	
    if($cbEnd) {
        printf("\n%s\n", encode("gbk", "设定样式"));
        $cbEnd->($sheet_rs);
    }
    
    # 冻结行
    
    if($self->{"forzen"}){
        $sheet_rs->Activate();
        $sheet_rs->Rows($self->{"rows"})->Select();
        $self->{"excel"}->ActiveWindow->{FreezePanes} = "True"; 
    }
}

sub saveExcel
{
	my $self = shift;
	
	# 参数检查
	
	if($self->{"workbook"} eq ""){
		printf("%s\n", encode("gbk", "未初始化"));
		return 1;
	}
    
    # 删除空白表格
    
	if($self->{"clear"})
	{
		my @dels = ();	
		foreach(1..$self->{"workbook"}->Worksheets()->{'Count'}) {
			unshift(@dels, $_) if $self->{"workbook"}->Worksheets($_)->{'Name'} =~ /^Sheet[0-9]{1,2}$/;
		}
		foreach(@dels){
			$self->{"workbook"}->Worksheets($_)->Delete;   
		}
	}
	
	# 保存文件
	
    if($self->{"file"} ne "") {
        my $path_rs = sprintf("%s%s%s.xlsx", encode("gbk", $self->{"path"}), $self->{"path"} =~ /\\$/?"":"\\", encode("gbk", $self->{"file"}));
        $self->{"workbook"} -> SaveAs ($path_rs);
        printf("\n%s : %s\n", encode("gbk", "存放路径"), $path_rs);         
    }
}

1;
