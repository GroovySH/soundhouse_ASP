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
'	PremiumGuitar一覧ページ
'
'更新履歴
'2009/09/01 an 新規作成
'2011/08/01 an #1087 Error.aspログ出力対応
'2012/01/20 an SELECT文へLACクエリー案を適用
'2014/03/19 GV 消費税増税に伴う2重表示対応
'
'========================================================================

On Error Resume Next

'ユーザが検索時に選択したデータ
Dim MakerCd
Dim iSort
Dim iPrice
Dim iPage

Dim wItemChar1
Dim wItemChar2
Dim wItemNum1
Dim wItemNum2
Dim wItemDate1
Dim wItemDate2

Dim wMakerName
Dim wMinimumPrice
Dim wMidCategoryCd
Dim wSalesTaxRate
Dim wPriceFrom
Dim wPriceTo
Dim wPageCount
Dim wRecordCount

Dim wMakerHTML
Dim wListHTML
Dim wCountHTML

Dim Connection
Dim RS

Dim wSQL
Dim wErrDesc   '2011/08/01 an add

Const cPageSize = 15 '1ページあたり表示商品件数

'========================================================================

Response.buffer = true

'---- Get input data
MakerCd = ReplaceInput(Trim(Request("MakerCd")))
iSort = ReplaceInput(Trim(Request("iSort")))
iPrice = ReplaceInput(Trim(Request("iPrice")))
iPage = ReplaceInput(Trim(Request("iPage")))

'---- 表示ページ設定
if iPage = "" then
	iPage = 1
else
	iPage = Clng(iPage)
end if

'---- Execute main
call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "PremiumGuitarsList.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
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

Dim vArrayPrice

'---- 対象カテゴリーコード、最低単価取出し
call getCntlMst("商品","PuremiumGuitar","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)
wMinimumPrice = Clng(wItemNum1)
wMidCategoryCd = wItemChar1

'---- 消費税率取出し
call getCntlMst("共通","消費税率","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)
wSalesTaxRate = Clng(wItemNum1)

if iPrice <> "" then
	vArrayPrice = split(iPrice,"-")
	wPriceFrom = vArrayPrice(0)
	wPriceTo   = vArrayPrice(1)
end if

if (ISNumeric(wPriceFrom) = false or wPriceFrom = "") then
	wPriceFrom = wMinimumPrice
end if
if (ISNumeric(wPriceTo) = false or wPriceTo = "") then
	wPriceTo = 9999999
end if

'----- 検索条件HTML作成
call CreateSearchHTML()

'----- PremiumGuitarリストHTML作成
call CreateListHTML()

End Function

'========================================================================
'
'	Function	メーカーリスト HTML作成
'
'========================================================================
'
Function CreateSearchHTML()

'---- 対象となるメーカー取り出し
wSQL = ""
wSQL = wSQL & "SELECT DISTINCT a.メーカー名"
wSQL = wSQL & "              , a.メーカーコード"

'wSQL = wSQL & "  FROM メーカー a WITH (NOLOCK)"     '2012/01/20 an mod s
'wSQL = wSQL & "     , Web商品  b WITH (NOLOCK)"
'wSQL = wSQL & "     , カテゴリー中カテゴリー  c WITH (NOLOCK)"
'wSQL = wSQL & " WHERE  a.メーカーコード = b.メーカーコード"
'wSQL = wSQL & " AND b.販売単価 >" &  wMinimumPrice
'wSQL = wSQL & " AND b.カテゴリーコード = c.カテゴリーコード"
'wSQL = wSQL & " AND c.中カテゴリーコード IN (" & wMidCategoryCd & ")"
'wSQL = wSQL & " AND b.Web商品フラグ = 'Y'"

wSQL = wSQL & " FROM"
wSQL = wSQL & "     メーカー                             a WITH (NOLOCK)"
wSQL = wSQL & "       INNER JOIN Web商品                 b WITH (NOLOCK)"
wSQL = wSQL & "         ON     b.メーカーコード = a.メーカーコード"
wSQL = wSQL & "       INNER JOIN カテゴリー中カテゴリー  c WITH (NOLOCK)"
wSQL = wSQL & "         ON     c.カテゴリーコード = b.カテゴリーコード"
wSQL = wSQL & "       LEFT  JOIN ( SELECT 'Y' AS 'ShohinWebY' ) t1 "
wSQL = wSQL & "         ON     b.Web商品フラグ      = t1.ShohinWebY "
wSQL = wSQL & " WHERE"
wSQL = wSQL & "          t1.ShohinWebY   IS NOT NULL "
wSQL = wSQL & "      AND b.販売単価 >" &  wMinimumPrice
wSQL = wSQL & "      AND c.中カテゴリーコード IN (" & wMidCategoryCd & ")"     '2012/01/20 an mod e
wSQL = wSQL & " ORDER BY a.メーカー名"

'@@@@@@@@@@response.write(wSQL)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

if RS.EOF = true then
	exit function
end if

wMakerHTML = ""

'メーカーリスト
wMakerHTML = wMakerHTML & "	<div class='left'>メーカー: " & vbNewLine
wMakerHTML = wMakerHTML & "	  <select name='MakerCd'>" & vbNewLine
wMakerHTML = wMakerHTML & "	    <option value='' "

if MakerCd = "" then
	wMakerHTML = wMakerHTML & " selected"
end if

wMakerHTML = wMakerHTML & ">ALL</option> " & vbNewLine

Do Until RS.EOF = true
	wMakerHTML = wMakerHTML & "	    <option value='" & RS("メーカーコード") & "'"
	wMakerHTML = wMakerHTML & ">" & RS("メーカー名") & "</option> " & vbNewLine
	If MakerCd = RS("メーカーコード") Then
		wMakerName = RS("メーカー名")
	End If
	RS.MoveNext
Loop

wMakerHTML = wMakerHTML & "	  </select>" & vbNewLine
wMakerHTML = wMakerHTML & "	</div>" & vbNewLine


RS.Close

End function

'========================================================================
'
'	Function	PremiumGuitarリストHTML作成
'
'========================================================================
'
Function CreateListHTML()

Dim vPrice

'---- 該当商品 取り出し
wSQL = ""
wSQL = wSQL & "SELECT DISTINCT"
wSQL = wSQL & "   a.メーカー名"
wSQL = wSQL & " , a.メーカーコード"
wSQL = wSQL & " , b.商品名"
wSQL = wSQL & " , b.商品コード"
wSQL = wSQL & " , b.商品画像ファイル名_小"
wSQL = wSQL & " , b.初回登録日"
wSQL = wSQL & " , CASE"
wSQL = wSQL & "   	WHEN b.個数限定数量 > b.個数限定受注済数量 THEN b.個数限定単価"
wSQL = wSQL & "    	ELSE b.販売単価"
wSQL = wSQL & "   END AS 実販売単価"

'wSQL = wSQL & " FROM メーカー a WITH (NOLOCK)"       '2012/01/20 an mod s
'wSQL = wSQL & "    , Web商品  b WITH (NOLOCK)"
'wSQL = wSQL & "    , カテゴリー中カテゴリー  c WITH (NOLOCK)"
'wSQL = wSQL & " WHERE  a.メーカーコード = b.メーカーコード"
'wSQL = wSQL & "    AND (SELECT CASE"
'wSQL = wSQL & "                   WHEN x.個数限定数量 > x.個数限定受注済数量 THEN (x.個数限定単価 * (100 + " & wSalesTaxRate & " )/100)"
'wSQL = wSQL & "                   ELSE (x.販売単価 * (100 + " & wSalesTaxRate & " )/100)"
'wSQL = wSQL & "                END"
'wSQL = wSQL & "         FROM web商品 x "
'wSQL = wSQL & "         WHERE x.メーカーコード = b.メーカーコード"
'wSQL = wSQL & "            AND x.商品コード = b.商品コード) BETWEEN " & wPriceFrom & " AND " & wPriceTo
'wSQL = wSQL & "    AND b.カテゴリーコード = c.カテゴリーコード"
'wSQL = wSQL & "    AND c.中カテゴリーコード IN (" & wMidCategoryCd & ")"
'wSQL = wSQL & "    AND b.Web商品フラグ = 'Y'"

wSQL = wSQL & " FROM"
wSQL = wSQL & "     メーカー                             a WITH (NOLOCK)"
wSQL = wSQL & "       INNER JOIN Web商品                 b WITH (NOLOCK)"
wSQL = wSQL & "         ON     b.メーカーコード = a.メーカーコード"
wSQL = wSQL & "       INNER JOIN カテゴリー中カテゴリー  c WITH (NOLOCK)"
wSQL = wSQL & "         ON     c.カテゴリーコード = b.カテゴリーコード"
wSQL = wSQL & "       LEFT  JOIN ( SELECT 'Y' AS 'ShohinWebY' ) t1 "
wSQL = wSQL & "         ON     b.Web商品フラグ      = t1.ShohinWebY "
wSQL = wSQL & " WHERE"
wSQL = wSQL & "        t1.ShohinWebY   IS NOT NULL "
wSQL = wSQL & "    AND (SELECT CASE"
wSQL = wSQL & "                   WHEN x.個数限定数量 > x.個数限定受注済数量 THEN (x.個数限定単価 * (100 + " & wSalesTaxRate & " )/100)"
wSQL = wSQL & "                   ELSE (x.販売単価 * (100 + " & wSalesTaxRate & " )/100)"
wSQL = wSQL & "                END"
wSQL = wSQL & "         FROM web商品 x WITH (NOLOCK)"
wSQL = wSQL & "         WHERE x.メーカーコード = b.メーカーコード"
wSQL = wSQL & "            AND x.商品コード = b.商品コード) BETWEEN " & wPriceFrom & " AND " & wPriceTo
wSQL = wSQL & "    AND c.中カテゴリーコード IN (" & wMidCategoryCd & ")"      '2012/01/20 an mod e


if MakerCd <> "" then
	wSQL = wSQL & " AND b.メーカーコード =" & MakerCd
end if

if iSort = "Update_Desc" or iSort = "" then
	wSQL = wSQL & " ORDER BY b.初回登録日 DESC"
elseif iSort = "Price_Desc" then
	wSQL = wSQL & " ORDER BY 実販売単価 DESC"
	wSQL = wSQL & "      , b.初回登録日 DESC"
elseif iSort = "Price_Asc" then
	wSQL = wSQL & " ORDER BY 実販売単価"
		wSQL = wSQL & "    , b.初回登録日 DESC"
elseif iSort = "MakerName" then
	wSQL = wSQL & " ORDER BY a.メーカー名"
	wSQL = wSQL & "        , b.初回登録日 DESC"
else
	wSQL = wSQL & " ORDER BY b.初回登録日 DESC"
end if

'@@@@@@@@@@response.write(wSQL)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

if RS.EOF = true then
	exit function
end if

RS.PageSize = cPageSize
if iPage > ((RS.RecordCount + (cPageSize - 1)) / cPageSize) then		'MAXページを超える場合は最終ページへ
	iPage = Fix(RS.RecordCount / cPageSize)
end if

RS.AbsolutePage = iPage				' start page no

'----- 商品一覧HTML作成

wListHTML = ""
wListHTML = wListHTML & "<div id='pgProductBox'>" & vbNewLine

for i=0 to (RS.PageSize - 1)

	wListHTML = wListHTML & "  <div class='productBox'>" & vbNewLine
	wListHTML = wListHTML & "    <div class='top'></div>" & vbNewLine
    '商品画像表示
	wListHTML = wListHTML & "    <div class='middle'><a href='PremiumGuitarsDetail.asp?Item=" & Server.URLEncode(RS("メーカーコード") & "^" & RS("商品コード"))
	
	wListHTML = wListHTML & "'><img src='prod_img/"
	
	if RS("商品画像ファイル名_小") <> "" then
		wListHTML = wListHTML & RS("商品画像ファイル名_小")
	else
		'商品画像が登録されていない場合はブランク画像表示
		wListHTML = wListHTML & "n/nopict.jpg"
	end if
	wListHTML = wListHTML & "'></a></div>" & vbNewLine
	'商品名表示
	wListHTML = wListHTML & "    <div class='middletextbox'><span class='maker'>" & RS("メーカー名") & "</span><a href='PremiumGuitarsDetail.asp?Item=" & Server.URLEncode(RS("メーカーコード") & "^" & RS("商品コード")) & "'>"
	wListHTML = wListHTML & RS("商品名") & "</a></div>" & vbNewLine
	
	vPrice = calcPrice(RS("実販売単価"), wSalesTaxRate)
	
'2014/03/19 GV mod start ---->
'	wListHTML = wListHTML & "    <div class='Pricetextbox'>価格" & FormatNumber(vPrice,0) & "円</div>" & vbNewLine
	wListHTML = wListHTML & "    <div class='Pricetextbox'>価格" & FormatNumber(RS("実販売単価"),0) & "円(税抜)</div>" & vbNewLine
	wListHTML = wListHTML & "    <div class='Pricetextbox'>(税込&nbsp;" & FormatNumber(vPrice,0) & "円)</div>" & vbNewLine
'2014/03/19 GV mod end   <----
	wListHTML = wListHTML & "    <div class='bottom'></div>" & vbNewLine
	wListHTML = wListHTML & "  </div>" & vbNewLine
	RS.MoveNext

	if RS.EOF Then
			exit for
	end If

Next

wListHTML = wListHTML & "</div>" & vbNewLine

'----- 件数表示HTML作成

Dim i

wCountHTML = ""	
wCountHTML = wCountHTML & "<div class='pgPager'>" & vbNewLine
if iPage <> 1 then
	wCountHTML = wCountHTML & "  <a href='JavaScript:Page_onClick(" & iPage-1 & ");'>[前へ]</a>" & vbNewLine
end if
if iPage <> RS.PageCount then
	wCountHTML = wCountHTML & "  <a href='JavaScript:Page_onClick(" & iPage+1 & ");'>[次へ]</a>　" & vbNewLine
end if
wCountHTML = wCountHTML & RS.RecordCount & "件ありました。" & vbNewLine

for i=1 to RS.PageCount
	wCountHTML = wCountHTML & "  <a href='JavaScript:Page_onClick(" & i & ");'>" & i & "</a>" & vbNewLine
next

wCountHTML = wCountHTML & "&nbsp;(現在" & iPage & "ページ)" & vbNewLine
wCountHTML = wCountHTML & "</div>" & vbNewLine

RS.Close
	
End function

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
<% If MakerCd <> "" Then %>
	<title>プレミアムギター <%= wMakerName %> 一覧｜サウンドハウス</title>
<% Else %>
	<title>プレミアムギター一覧｜サウンドハウス</title>
<% End If %>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/PremiumGuitars.css" type="text/css">
<script type="text/javascript">
//
//	初期設定
//
var isMacIE;
var isNS4;
var isIE4;
var isDynamic;
var s_maker_cd;
isMacIE = ((navigator.userAgent.indexOf("IE 4") > -1) && (navigator.userAgent.indexOf("Mac") > -1));
isNS4 = ((navigator.appName == "Netscape") && (parseInt(navigator.appVersion) >= 4));
isIE4 = ((navigator.appName == "Microsoft Internet Explorer") && (parseInt(navigator.appVersion) >= 4));
isDynamic = (isNS4 || isIE4 && !isMacIE);
//
//	Page onClick
//
function Page_onClick(pPage){
	document.f_search.iPage.value = pPage;
	document.f_search.submit();
}
//=====================================================================
//	ラジオボタン、ドロップダウンリストを以前に選択した状態にする
//=====================================================================
function preset_values(MakerCd,iSort,iPrice){
//	MakerCD
	for (var i=0; i<document.f_search.MakerCd.options.length; i++){
		if (document.f_search.MakerCd.options[i].value == MakerCd){
			document.f_search.MakerCd.options[i].selected = true;
			break;
		}
	}
//	iSort
	for (var i=0; i<document.f_search.iSort.options.length; i++){
		if (document.f_search.iSort.options[i].value == iSort){
			document.f_search.iSort.options[i].selected = true;
			break;
		}
	}
//	iPrice
	for (var i=0; i<document.f_search.iPrice.options.length; i++){
		if (document.f_search.iPrice.options[i].value == iPrice){
			document.f_search.iPrice.options[i].selected = true;
			break;
		}
	}
}
</script>
<style type="text/css">
#globalContents ul.sns {
	overflow: hidden;
	padding: 5px;
}

#globalContents ul.sns li {
	float: right;
	width: 100px;
	height: 20px;
}
</style>
</head>
<body>
<!--#include file="../Navi/Navitop.inc"-->

<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="ここから本文です"></a></span>

<!-- コンテンツstart -->
<div id="globalContents">
    <div id='path_box'><div id='path_box_inner01'><div id='path_box_inner02'>
    <p class='home'><a href="<%=g_HTTP%>"><img src="<%=g_RelLink%>images/icon_home.gif" alt="HOME"></a></p>
    <ul id='path'>
      <li><a href="<%=g_HTTP%>material/">SPECIAL SELECTION一覧</a></li>
      <li><a href="PremiumGuitars.asp">プレミアムギター</a></li>
<% If MakerCd <> "" Then %>
      <li class="now"><%= wMakerName %> 一覧</li>
<% Else %>
      <li class="now">プレミアムギター一覧</li>
<% End If %>
    </ul>
  </div></div></div>
    <ul class="sns">
          <li><a href="https://twitter.com/share" class="twitter-share-button" data-lang="ja">ツイート</a><script>!function(d,s,id){var js,fjs=d.getElementsByTagName(s)[0];if(!d.getElementById(id)){js=d.createElement(s);js.id=id;js.src="//platform.twitter.com/widgets.js";fjs.parentNode.insertBefore(js,fjs);}}(document,"script","twitter-wjs");</script></li>
          <li><iframe src="//www.facebook.com/plugins/like.php?href=http%3A%2F%2Fwww.soundhouse.co.jp%2Fshop%2FPremiumGuitars.asp&amp;send=false&amp;layout=button_count&amp;width=100&amp;show_faces=false&amp;action=like&amp;colorscheme=light&amp;font&amp;height=21&amp;appId=191447484218062" scrolling="no" frameborder="0" style="border:none; overflow:hidden; width:100px; height:21px;" allowTransparency="true"></iframe></li>
        </ul>
<!--
  <h1 class="title">プレミアムギター</h1>
-->
  <div id="pgContainer">
<!-- トップ画像 START -->
<div id="pgHeader">
  <div class="topbox">
    <div class="left"></div>
    <div class="right"></div>
  </div>
</div>
<!-- トップ画像 END -->

<div id="pgSelectBox">
  <form name="f_search">

<%=wMakerHTML%>

    <div class="left">並べ替え:
      <select name="iSort"> 
        <option value="Update_Desc">更新順</option>
        <option value="Price_Asc">価格順△</option>
        <option value="Price_Desc">価格順▽</option>
        <option value="MakerName">メーカー順</option>
      </select>
    </div>
    <div class="left">プライス:
      <select name="iPrice">
        <option value="">ALL</option>
        <option value="<%=wMinimumPrice%>-250000"><%=wMinimumPrice%>円 - 250,000円</option>
        <option value="250001-400000">250,001円 - 400,000円</option>
        <option value="400001-600000">400,001円 - 600,000円</option>
        <option value="600001-1000000">600,001円 - 1,000,000円</option>
        <option value="1000001-9999999">1,000,001円 -</option>
      </select>
    </div>
    <div class="right">
      <input type="hidden" name="iPage" value="<%=iPage%>">
      <input type="submit" id="bottun" value="検索">
    </div>
  </form>
</div>

<!-- 件数表示 -->
<%=wCountHTML%>

<!-- PremiumGuitarリスト -->
<%=wListHTML%>

<!-- 件数表示 -->
<%=wCountHTML%>

<p class="arrow"><a href="#site_title"><img src="images/PremiumGuitars/white_arrow_up.gif" alt=""></a></p>

</div>

  </div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<script type="text/javascript">
	preset_values('<%=MakerCd%>','<%=iSort%>','<%=iPrice%>');
</script>
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>