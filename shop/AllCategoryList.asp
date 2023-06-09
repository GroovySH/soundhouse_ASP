<%@ LANGUAGE="VBScript" %>
<%
'ネットハウスねっとハウスネットはうす
'サウンドハウス
 Option Explicit
%>
<!--#include file="../common/ADOVBS.inc"-->
<!--#include file="../common/system_common.inc"-->
<!--#include file="../common/shop_common_functions.inc"-->
<!--#include file="../common/bfunctions1.asp"-->

<%
'========================================================================
'
'	全カテゴリー一覧ページ
'
'更新履歴
'2005/09/16 1カテゴリーに複数中カテゴリー対応
'2006/03/27 Web大カテゴリーフラグ対応
'2007/06/05 ハッカーセーフ対応
'2009/04/30 エラー時にerror.aspへ移動
'2009/07/28 デザイン変更（カテゴリ毎に画像を表示）、LargeCategoryCd=""の際は表示順が先頭のカテゴリを表示
'2009/08/05 カテゴリ取得のORDER BYの条件に中カテゴリーコードを追加（複数の中カテゴリーで同じ表示順を指定している場合の対応）
'2011/08/01 an #1087 Error.aspログ出力対応
'2012/09/06 ok リニューアルに伴い新デザインに変更（大カテゴリーは固定とする）
'
'========================================================================

On Error Resume Next

Dim LargeCategoryCd
Dim LargeCategoryName		'2012/09/06 ok Add

Dim wLargeCategoryListHTML
Dim wCategoryListHTML

Dim Connection
Dim RS

Dim w_sql
Dim w_html
Dim wErrDesc   '2011/08/01 an add

'========================================================================

'---- Get input data
LargeCategoryCd = ReplaceInput(Trim(Request("LargeCategoryCd")))

'---- Execute main
call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "AllCategoryList.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

call close_db()

if Err.Description <> "" then
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

'========================================================================
'
'	Function	Connect database
'
'========================================================================
'
Function connect_db()

'---- Connect database
Set Connection = Server.CreateObject("ADODB.Connection")
Connection.Open g_connection

End function

'========================================================================
'
'	Function	Main
'
'========================================================================
'
Function main()

if LargeCategoryCd = "" then
	LargeCategoryCd = "1"
end if

'----- HTML作成
'call CreateLargeCategoryListHTML()		'大カテゴリー一覧	'2012/09/06 ok Del
call CreateCategoryListHTML()					'カテゴリー一覧

End Function

'========================================================================
'
'	Function	大カテゴリー一覧
'		'2012/09/06 ok Del
'========================================================================
'
'Function CreateLargeCategoryListHTML()
'
''---- 大カテゴリー 取り出し
'w_sql = ""
'w_sql = w_sql & "SELECT a.大カテゴリー名"
'w_sql = w_sql & "     , a.大カテゴリーコード"
'w_sql = w_sql & "     , a.大カテゴリー画像ファイル名大"
'w_sql = w_sql & "  FROM 大カテゴリー a WITH (NOLOCK)"
'w_sql = w_sql & " WHERE Web大カテゴリーフラグ = 'Y'"
'w_sql = w_sql & " ORDER BY"
'w_sql = w_sql & "       a.表示順"
'
''@@@@@@@@@@response.write(w_sql)
'
'Set RS = Server.CreateObject("ADODB.Recordset")
'RS.Open w_sql, Connection, adOpenStatic
'
'if RS.EOF = true then 
'	exit function
'end if
'
''----- 大カテゴリー一覧HTML編集
'
'w_html = ""
''2012/09/06 ok Mod Start
''w_html = w_html & "<div class='category_title'><span><h2>全カテゴリー一覧</h2></span></div>" & vbNewLine
''w_html = w_html & "<div id='all_cat_list'>" & vbNewLine
'w_html = w_html & "    <h1 class='title'>全カテゴリー一覧</h1>" & vbNewLine
'w_html = w_html & "    <ul id='allcat'>" & vbNewLine
'
'Do Until RS.EOF = true
'	if LargeCategoryCd = "" then
'		LargeCategoryCd = RS("大カテゴリーコード")
'	end if
''	w_html = w_html & "  <div class='Large_cat' style='background-image:url(images/AllCategoryList/" & RS("大カテゴリー画像ファイル名大") & ")'>" & vbNewLine
''	w_html = w_html & "    <a href='AllCategoryList.asp?LargeCategoryCd=" & RS("大カテゴリーコード")  & "'>" & vbNewLine
''	w_html = w_html & "      <div class='Large_cat_in'><span><h3>" & RS("大カテゴリー名") & "</h3></span></div>" & vbNewLine
''	w_html = w_html & "    </a>" & vbNewLine
''	w_html = w_html & "  </div>" & vbNewLine
'	if LargeCategoryCd = RS("大カテゴリーコード") Then
'		LargeCategoryName = RS("大カテゴリー名")
'		w_html = w_html & "      <li class='l"+ RS("大カテゴリーコード") +" now'>"+ LargeCategoryName +"</li>" & vbNewLine
'	Else
'		w_html = w_html & "      <li class='l"+ RS("大カテゴリーコード") +"'><a href='AllCategoryList.asp?LargeCategoryCd=" + RS("大カテゴリーコード") + "'>" + LargeCategoryName + "</a></li>" & vbNewLine
'	End If
'
'	RS.MoveNext
'Loop
''w_html = w_html & "</div>" & vbNewLine
'w_html = w_html & "    </ul>" & vbNewLine
''2012/09/06 ok Mod End
'
'wLargeCategoryListHTML = w_html
'
'RS.Close
'
'End Function

'========================================================================
'
'	Function	カテゴリー一覧
'
'========================================================================
'
Function CreateCategoryListHTML()

Dim vMidCategoryCd
'Dim vMidCategoryCount	'2012/09/06 ok Del

'---- カテゴリー 取り出し
w_sql = ""
w_sql = w_sql & "SELECT a.中カテゴリーコード"
w_sql = w_sql & "     , a.中カテゴリー名日本語"
w_sql = w_sql & "     , a.中カテゴリー画像ファイル名"
w_sql = w_sql & "     , b.カテゴリーコード"
w_sql = w_sql & "     , b.カテゴリー名"
w_sql = w_sql & "  FROM 中カテゴリー a WITH (NOLOCK)"
w_sql = w_sql & "     , カテゴリー b WITH (NOLOCK)"
w_sql = w_sql & "     , カテゴリー中カテゴリー c WITH (NOLOCK)"
w_sql = w_sql & " WHERE c.中カテゴリーコード = a.中カテゴリーコード"
w_sql = w_sql & "   AND b.カテゴリーコード = c.カテゴリーコード"
w_sql = w_sql & "   AND b.Webカテゴリーフラグ = 'Y'"
w_sql = w_sql & "   AND a.大カテゴリーコード = '" & LargeCategoryCd &"'"
w_sql = w_sql & " ORDER BY"
w_sql = w_sql & "       a.表示順"
w_sql = w_sql & "     , a.中カテゴリーコード"
w_sql = w_sql & "     , c.中カテゴリー区分"
w_sql = w_sql & "     , b.表示順"


'@@@@@@@@@@response.write(w_sql)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open w_sql, Connection, adOpenStatic

if RS.EOF = true then
	exit function
end if

'----- カテゴリー一覧HTML編集

w_html = ""
'2012/09/06 ok Mod Start
'w_html = w_html & "<div class='category_title'><span><h2>カテゴリーから選ぶ</h2></span></div>" & vbNewLine
'w_html = w_html & "<div id='Large_cat_list'>" & vbNewLine
w_html = w_html & "    <h2 class='allcat_title'>" + LargeCategoryName + "</h2>" & vbNewLine
w_html = w_html & "    <ul class='cat_detail'>" & vbNewLine

'vMidCategoryCount = 0

Do Until RS.EOF = true
	vMidCategoryCd = RS("中カテゴリーコード")
'	vMidCategoryCount = vMidCategoryCount + 1
	
'	if vMidCategoryCount Mod 3 = 1 then '中カテゴリ3個ごとに1行としてスタイル設定
'		w_html = w_html & "<div class='line'>" & vbNewLine
'	end if
	
'	w_html = w_html & "  <div class='Mid_cat_list'>" & vbNewLine
'	w_html = w_html & "    <div class='border'>" & vbNewLine
'	w_html = w_html & "      <div class='cat_img'><a href='MidCategoryList.asp?MidCategoryCd=" & RS("中カテゴリーコード") &  "'><img src='cat_img/" & RS("中カテゴリー画像ファイル名") & "' border='0' alt='" & RS("中カテゴリー名日本語") & "'></a></div>" & vbNewLine
'	w_html = w_html & "      <div class='cat_list'>" & vbNewLine
'	w_html = w_html & "        <h4><a href='MidCategoryList.asp?MidCategoryCd=" & RS("中カテゴリーコード") &  "'>" & RS("中カテゴリー名日本語") & "</a></h4>" & vbNewLine
'	w_html = w_html & "        <ul class='list'>" & vbNewLine
	
	w_html = w_html & "      <li>" & vbNewLine
	w_html = w_html & "        <h3 class='allcat_subtitle'><a href='MidCategoryList.asp?MidCategoryCd=" + vMidCategoryCd + "'>" + RS("中カテゴリー名日本語") + "</a></h3>" & vbNewLine
	w_html = w_html & "        <div class='cat_inner m" + vMidCategoryCd + "'>" & vbNewLine
	w_html = w_html & "          <ul class='s_cat_list'>" & vbNewLine

	Do While RS("中カテゴリーコード") =  vMidCategoryCd
'		w_html = w_html & "          <li><a href='SearchList.asp?i_type=c&amp;s_category_cd=" & RS("カテゴリーコード") &  "'>- " & RS("カテゴリー名") & "</a></li>" & vbNewLine
		w_html = w_html & "            <li><a href='SearchList.asp?i_type=c&amp;s_category_cd=" + RS("カテゴリーコード") + "'>" & RS("カテゴリー名") & "</a></li>" & vbNewLine
		RS.MoveNext
		if RS.EOF = true then Exit Do
	Loop
	w_html = w_html & "          </ul>" & vbNewLine
	w_html = w_html & "        </div>" & vbNewLine
	w_html = w_html & "      </li>" & vbNewLine

'	w_html = w_html & "        </ul>" & vbNewLine
'	w_html = w_html & "      </div>" & vbNewLine
'	w_html = w_html & "    </div>" & vbNewLine
'	w_html = w_html & "  </div>" & vbNewLine
	
'	if vMidCategoryCount Mod 3 = 0 then '3の倍数なら<div class='line'>を閉じる
'		w_html = w_html & "</div>" & vbNewLine
'	end if
Loop

'if vMidCategoryCount Mod 3 <> 0 then '3の倍数でない場合も最後のデータであれば<div class='line'>を閉じる
'	w_html = w_html & "</div>" & vbNewLine
'end if

'w_html = w_html & "</div>" & vbNewLine
w_html = w_html & "    </ul>" & vbNewLine
'2012/09/06 ok Mod End

wCategoryListHTML = w_html

RS.Close

End Function

'========================================================================
'
'	Function	Close database
'
'========================================================================
'
Function close_db()

Connection.close
Set Connection= Nothing    '2011/08/01 an add

End function

'========================================================================
%>
<!DOCTYPE html>
<html>
<head>
<meta charset="Shift_JIS">
<title>全カテゴリー一覧｜サウンドハウス</title>
<meta name="Description" content="楽器、PA音響機器、DJ・DTM、照明機器、カラオケ機材をどこよりも【激安特価】でご提供するサウンドハウスの全カテゴリー一覧です。">
<meta name="keywords" content="PAレコーディング,ギター,ベース,ドラム,パーカッション,キーボード,DJ,DTM,レコーダー,スタンド,ラック,ケース,ケーブル,ヘッドホン,イヤホン">
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/categorylist.css?2014812" type="text/css">
</head>
<body>
<!--#include file="../Navi/NaviTop.inc"-->
<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="ここから本文です"></a></span>
  
  <!-- コンテンツstart -->
  <div id="globalContents">
    <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
      <p class="home"><a href="<%=g_HTTP%>"><img src="../images/icon_home.gif" alt="HOME"></a></p>
      <ul id="path">
        <li class="now">全カテゴリー一覧</li>
      </ul>
    </div></div></div>

    <h1 class="title">全カテゴリー一覧</h1>
    <ul id="allcat">
      <li class="l1"><a href="AllCategoryList.asp?LargeCategoryCd=1">PA&amp;レコーディング</a></li>
      <li class="l12"><a href="AllCategoryList.asp?LargeCategoryCd=12">ギター</a></li>
      <li class="l13"><a href="AllCategoryList.asp?LargeCategoryCd=13">ベース</a></li>
      <li class="l14"><a href="AllCategoryList.asp?LargeCategoryCd=14">ドラム &amp;<br>パーカッション</a></li>
      <li class="l15"><a href="AllCategoryList.asp?LargeCategoryCd=15">キーボード</a></li>
      <li class="l16"><a href="AllCategoryList.asp?LargeCategoryCd=16">その他 楽器</a></li>
      <li class="l7"><a href="AllCategoryList.asp?LargeCategoryCd=7">DJ &amp; VJ</a></li>
      <li class="l8"><a href="AllCategoryList.asp?LargeCategoryCd=8">DTM / DAW</a></li>
      <li class="l3"><a href="AllCategoryList.asp?LargeCategoryCd=3">映像機器・<br>レコーダー</a></li>
      <li class="l4"><a href="AllCategoryList.asp?LargeCategoryCd=4">照明・<br>ステージシステム</a></li>
      <li class="l9"><a href="AllCategoryList.asp?LargeCategoryCd=9">スタンド各種</a></li>
      <li class="l5"><a href="AllCategoryList.asp?LargeCategoryCd=5">ラック・ケース</a></li>
      <li class="l10"><a href="AllCategoryList.asp?LargeCategoryCd=10">ケーブル各種</a></li>
      <li class="l6"><a href="AllCategoryList.asp?LargeCategoryCd=6">ヘッドホン・<br>イヤホン</a></li>
      <li class="l51"><a href="AllCategoryList.asp?LargeCategoryCd=51">スタジオ家具</a></li>
    </ul>

<%=wCategoryListHTML%>
  </div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
<script type="text/javascript">
$(function(){
	$(".l<%=LargeCategoryCd%>").addClass("now");
	$(".l<%=LargeCategoryCd%> a").replaceWith("<p>" + $(".l<%=LargeCategoryCd%> a").html() + "</p>");

	$(".allcat_title").text($(".now p").text());

	$(".cat_inner").equalbox();
});</script>
</body>
</html>