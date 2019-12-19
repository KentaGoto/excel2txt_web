use Mojolicious::Lite;
use File::Copy;
use open IO => qw/:encoding(UTF-8)/;
use HTML::Entities;
use XML::Twig;
use File::Find::Rule;
use File::Basename qw/basename dirname fileparse/;
use FindBin;
use Cwd;
use Archive::Zip qw( :ERROR_CODES :CONSTANTS :MISC_CONSTANTS );
use File::Path 'rmtree';
use File::Spec;
use Spreadsheet::ParseXLSX;
use Encode qw/encode decode/;

#binmode STDOUT, ':encoding(UTF-8)';

require './ParseXLSX.pm';

my $app = app;
my $url = 'http://localhost:3003'; # URL

my $UPLOAD_DIR = app->home->rel_file('/upload');
my $TMP_DIR = app->home->rel_file('/tmp');
my $PROC_TMP_DIR = 'proc_tmp';

my @texts;

get '/excel2txt' => sub {
	my $c = shift;

	# キャッシュを残さない。
	# この設定がないと、ブラウザの戻るボタンでトップページに戻ってきた際に、選択したzipファイルを変更して実行しても以前選択したzipファイルで実行されてしまう。
	# FirefoxやEdgeではこの挙動となった。Chromeではこの挙動とはならなかった。2018/01/18
	$c->res->headers->header("Cache-Control" => "no-store, no-cache, must-revalidate, max-age=0, post-check=0, pre-check=0", 
							 "Pragma"        => "no-cache"
							);

	$c->render('index', top_page => $url);

	@texts = (); # リセット

	# uploadフォルダがなければ作成する
	if ( !-d $UPLOAD_DIR ){
		mkdir $UPLOAD_DIR or die "Can not create directory $UPLOAD_DIR";
	}
	# tmpフォルダがなかったら作る
	if ( -d $TMP_DIR ){
	
	} else {
		mkdir $TMP_DIR, 0700 or die "$!";
	}
	undef($c);
};

post '/excel2txt' => sub {
    my $c = shift;
	my $nfn = $c->param('nfn');

    # 処理対象zipファイル
    my $file = $c->req->upload('files');
    my $zip_filename = $file->filename;
    
    # zip以外受け付けない
    unless ( $zip_filename =~ /\.zip$/ ){
    	return $c->render(
    		template => 'error', 
    		message  => "Error",
    		message2 => "Upload fail. The selected file is not an ZIP file.",
    	);
    }

	# Local time settings
	my $times = time();
	my ($sec, $min, $hour, $mday, $month, $year, $wday, $stime) = localtime($times);
	$month++;
	my $datetime = sprintf '%04d%02d%02d%02d%02d%02d', $year + 1900, $month, $mday, $hour, $min, $sec;

	# uploadディレクトリに日付フォルダ作成
	chdir $UPLOAD_DIR;
	if ( !-d $datetime ){
		mkdir $datetime or die "Can not create directory $datetime";
	}
    
    # アップロードされたファイルを保存
    my $upload_save_file = "$UPLOAD_DIR/" . "$datetime/" . $zip_filename;
    $file->move_to($upload_save_file);
    
	# tmpディレクトリに日付フォルダ作成
	chdir $TMP_DIR;
	if ( !-d $datetime ){
		mkdir $datetime or die "Can not create directory $datetime";
	}

	# tmpフォルダにも処理対象ファイルを移す
	my $tmp_save_file = "$TMP_DIR/" . "$datetime/" . $zip_filename;
	$file->move_to($tmp_save_file);
	
	# Zip解凍
	chdir $datetime;
	my $datetime_fullpath = "$TMP_DIR/$datetime";
	&unzip(\$zip_filename, $datetime_fullpath);

	# zip展開後はzipを削除
	unlink $tmp_save_file;

	# 処理ディレクトリとなるproc_tmpフォルダを作成
	mkdir $PROC_TMP_DIR or die "Can not create directory $PROC_TMP_DIR";
	
	# proc_tmpフォルダのフルパス
	my $PROC_TMP_DIR_abs = "$TMP_DIR/$datetime/$PROC_TMP_DIR";

	#########################################################################
	#  	xlsxからテキストを抽出						                     	 #
	#########################################################################
	my @xlsxs = File::Find::Rule->file->name( '*.xlsx', '*.XLSX' )->in(getcwd);
	chdir $PROC_TMP_DIR_abs;

	foreach (@xlsxs){
		my $xlsx_fullpath = $_;
		my $xlsx_filename = basename($xlsx_fullpath);
		my $xlsx_dirname  = dirname($xlsx_fullpath);

		print "Processing... $xlsx_filename\n";
			
		my $ExcelObj = Spreadsheet::ParseXLSX->new();  # ParseXLSXのオブジェクト定義
		my $Book = $ExcelObj->parse($xlsx_fullpath);   # xlsxファイル読み込み/book扱い
	
		# [ファイル名区切りを出力する]がオンの場合
		if ( defined $nfn ){
			my $xlsx_filename_decode = decode('CP932', $xlsx_filename); # ファイル名をデコードしないとHTML出力で化ける
			push (@texts, "\n\n------------------------------$xlsx_filename_decode------------------------------");
		} else {
			# オフの場合は何もしない
		}
		
		# セルの文字列抽出処理
		for my $Sheet ($Book->worksheets()) {  # worksheetオブジェクト取得
		    &GetValuesFromCell($Sheet);        # worksheetデータ取得へ
		}
	
		# テキストボックスなどの文字列抽出処理
		&GetValuesFromDrawingAndSheetname ($xlsx_filename, $xlsx_dirname, $PROC_TMP_DIR_abs);
	}
	
	# resultsページに移り、抽出したテキストをリダイレクトする
	$c->redirect_to('/excel2txt/results');
	
} => 'upload';

get '/excel2txt/results' => sub {
    my $c = shift;
	@texts = grep $_ !~ /^\s*$/, @texts; # 空白のみまたは空は捨てる
	$c->render('results', 'texts' => \@texts);
};

sub GetValuesFromCell { # worksheetデータ取得
    my ($Sheet) = @_;
    
	# シート名ゲット
	push (@texts, $Sheet->get_name());
    
    my ($Rmin, $Rmax) = $Sheet->row_range(); # 行のデータ範囲(最小,最大)
    my ($Cmin, $Cmax) = $Sheet->col_range(); # 列のデータ範囲(最小,最大)
    
	# セルの値ゲット
    for (my $row=$Rmin; $row<=$Rmax; $row++) {      # rowは行番号
        for (my $col=$Cmin; $col<=$Cmax; $col++) {  # colは列番号
            my $Cell = $Sheet->get_cell($row,$col); # Cellオブジェクト取得
            if (defined($Cell)) {
				push (@texts, $Cell->value(), "\n"); # セルの値
            } else {
				# なにもしない
            }
        }
    }

	# 端末へのログ出力
	my $outputSheetname = $Sheet->get_name();
	print "Get values from cell" . ": " . encode('CP932', $outputSheetname) . "\n";
}

sub GetValuesFromDrawingAndSheetname {
	my ($xlsx_filename, $xlsx_dirname, $PROC_TMP_DIR_abs) = @_;
	my $zip = &xlsxCopy2tmp($xlsx_filename, $xlsx_dirname, $PROC_TMP_DIR_abs); # xlsxをproc_tmpフォルダに移動してzipにする。
	&unzip(\$zip, $PROC_TMP_DIR_abs); # zip解凍
	unlink $zip; # 展開後のzipを削除
	unlink glob '*.xml'; # 要らないxmlを削除する
	my @xmls = File::Find::Rule->file->name( qr/(?:drawing\d+\.xml$|comments\d+\.xml)/ )->in(getcwd);
	&xml_copy(\@xmls, $PROC_TMP_DIR_abs); # 対象のxmlファイルをproc_tmpフォルダにコピーする
	&del_dir($PROC_TMP_DIR_abs); # 要らないフォルダを削除
	my @target_xmls = glob '*.xml'; # proc_tmpフォルダにコピーしたxmlを対象とする
	&xml_parser(\@target_xmls); # xmlをパースしてテキストをゲットする
	unlink glob '*.xml'; # 対象ファイルの*.xmlを削除する
}

sub xlsxCopy2tmp {
	my ($xlsx_filename, $xlsx_dirname, $PROC_TMP_DIR_abs) = @_;
	my $zip;
	$zip = $xlsx_filename;
	$zip =~ s|^(.+)$|$1\.zip|;
	copy($xlsx_dirname . '/' . $xlsx_filename, "$PROC_TMP_DIR_abs/$zip") or die $!;
	return $zip;
}

sub unzip {
	my ($zip, $DIR) = @_;
	my $zip_obj = Archive::Zip->new($$zip);
	my @zip_members = $zip_obj->memberNames();
	foreach (@zip_members) {
		$zip_obj->extractMember($_, "$DIR/$_");
	}
}

sub xml_copy {
	my ($xmls, $PROC_TMP_DIR_abs) = @_;
	foreach (@$xmls){
		print $_ . "\n";
		my $file_src = $_;
		my $file_dst = $PROC_TMP_DIR_abs;
		copy($file_src, $file_dst) or die {$!};
	}
}

sub del_dir {
	my ($PROC_TMP_DIR_abs) = shift;
	rmtree("$PROC_TMP_DIR_abs/xl") or die $!;
	rmtree("$PROC_TMP_DIR_abs/docProps") or die $!;
	rmtree("$PROC_TMP_DIR_abs/_rels") or die $!;
}

sub xml_parser {
	my ($target_xmls) = shift;
	foreach my $xml ( @$target_xmls ){
		my $twig = new XML::Twig( TwigRoots => {
				'//a:p' => \&output_target,
				'//t' => \&output_target,
				'//oddHeader' => \&output_target_sheetXML_header
				});
		$twig->parsefile( $xml );
	}
}

# テキストボックスなどのdrawing XMLから抽出
sub output_target {
	my( $tree, $elem ) = @_;
	my $target = $elem->text;
	push (@texts, $target);
	
	{
		local *STDOUT;
		local *STDERR;
  		open STDOUT, '>', undef;
  		open STDERR, '>', undef;
		$tree->flush_up_to( $elem ); #Memory clear
	}
}

# ヘッダーのsheet XMLから抽出
sub output_target_sheetXML_header {
	my( $tree, $elem ) = @_;
	my $target = $elem->text;
	$target =~ s/^&[A-Z]//; # 要らない文字列削除
	push (@texts, $target);
	
	{
		local *STDOUT;
		local *STDERR;
  		open STDOUT, '>', undef;
  		open STDERR, '>', undef;
		$tree->flush_up_to( $elem ); #Memory clear
	}
}

app->start;

__DATA__

@@ error.html.ep
<h1><%= $message %></h1>
<p><%= $message2 %></p>

@@ layouts/default.html.ep
<html>
<head>
<title><%= title %></title>
<meta http-equiv="Content-type" content="text/html; charset=UTF-8">
<%= stylesheet '/css/style.css' %>
<link type="text/css" rel="stylesheet"
  href="http://code.jquery.com/ui/1.10.3/themes/cupertino/jquery-ui.min.css" />
<script type="text/javascript"
  src="http://code.jquery.com/jquery-1.10.2.min.js"></script>
<script type="text/javascript"
  src="http://code.jquery.com/ui/1.10.3/jquery-ui.min.js"></script>
</head>
<body><%= content %></body>
</html>

@@ index.html.ep
<%
	my $filename = stash('filename');
%>
% layout 'default';
% title 'excel2txt';
%= javascript begin
  // プログレスバー
  $(document).on('click', '#run', function() {
    $('#progress').progressbar({
        max: 100,
        value: false
	});
	// ボタンなどの非表示
	// propやattrのグレーアウトは、Chromeだと処理が実行されないバグがあった。
	$('#run').hide();
	$('#select').hide();
	$('#delimiter_checkbox').hide();
	$('#processing').show();
  });
% end
<div id="out">
<div id="head">
<h1>excel2txt</h1>
<form method="post" action="<%= url_for('upload') %>" enctype ="multipart/form-data">
	<input name="files" type="file" id='select' value="Select File" />
	<input type="submit" id="run" value="Run" />
	</br>
	<p id="delimiter_checkbox">ファイル名区切りを出力する: <%= check_box nfn => 1 %></p>
	</br>
	<p id="processing" style="display: none;">Processing... </p>
	<div id="progress"></div>
</form>
	</div>
	<div id="main">
<h3>Usage</h3>
	<ul>
		<li><strong>*.xlsx</strong> ファイルの入った <strong>zip</strong> ファイルを選択します。</li>
		<li><strong>[Run]</strong> ボタンをクリックします。</li>
		<li>遷移した画面に <strong>*.xlsx</strong> から抽出したテキストが表示されます。</li>
	</ul>
<h3>Option</h3>
	<ul>
		<li><strong>[ファイル名区切りを出力する]</strong> チェックボックスをオンにすると、処理対象となった <strong>*.xlsx</strong> が抽出テキストの区切りとして出力されます。</li>
	</ul>
<h3>Requirements</h3>
	<ul>
		<li>Chrome or Firefox</li>
	</ul>
<h4>Note</h4>
<ul>
	<li>セル、シート名、図形およびテキストボックスに使用されている文字が抽出されます。</li>
</ul>
</div>
<div id="footer">
Copyright &copy; KentaGoto All Rights Reserved.
</div>
</div>

@@ results.html.ep
<html>
<head>
% title 'Results';
<meta http-equiv="Content-type" content="text/html; charset=UTF-8">
<%= stylesheet '/css/style_Results.css' %>
</head>
<body>
% for my $t (@$texts){
	<%= $t %> </br>
% }
</body>
</html>
