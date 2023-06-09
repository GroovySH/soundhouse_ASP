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
'	商品比較ページ そのまま比較できるかどうかチェック　(カテゴリー決定/5個以上の時 選択画面表示)
'
'	更新履歴
'2008/05/07 区切り文字変更
'2009/04/30 エラー時にerror.aspへ移動
'2011/08/01 an #1087 Error.aspログ出力対応
'2012/01/20 GV データ取得 SELECT文へ LACクエリー案を適用
'2012/07/18 nt リニューアル用にデータ取得 SELECT文およびasp画面出力を修正
'2014/01/31 GV 本番環境でクッキーを取得できない現象を修正
'
'========================================================================

On Error Resume Next

Dim wNaveWithLink '2012/7/19 nt add
Dim wTitleWithLink

Dim wHikaku
Dim CategoryCd()
Dim MakerCd()
Dim ProductCd()
Dim Iro()
Dim Kikaku()
Dim MakerName()
Dim ProductName()
Dim wRecCnt
Dim wGotoCompareFl
Dim wParm

Dim Connection
Dim RS

Dim i
Dim wHTML
Dim wSQL
Dim wMsg
Dim wErrDesc   '2011/08/01 an add

Dim category_cd '2012/7/19 nt add

'========================================================================

Response.Buffer = true

'---- Execute main

call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "ProductCompareCheck.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc)
end if                                           '2011/08/01 an add e

call close_db()

if Err.Description <> "" then
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

if wGotoCompareFl = true then
	For i= 1 to wRecCnt
		wParm = wParm & "$" & CategoryCd(i) & "^" & MakerCd(i) & "^" & ProductCd(i) & "^" & Trim(Iro(i)) & "^" & Trim(kikaku(i))
	Next
	Response.redirect "ProductCompare.asp?item=" & wParm
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

Dim vMoreThanOneCategoryFl
Dim vMinData
Dim vMinSub
Dim vOldCategoryCd
Dim vProductName
Dim i
Dim j
Dim vTemp
Dim vCookieData		'2014/01/31 GV add

'---- 送信データーの取り出し

'---- 比較Cookie取り出し
'     1件目(添え字0)はダミーデータのため無視
'2014/01/31 GV mod start
vCookieData = Request.Cookies("compare")
'通常のResponse.Cookies()でデータが取得できなかった場合、
'専用のプロシージャ(Shop_common_functions.inc)を使う。
If (Trim(vCookieData) = "") Or (IsNull(Trim(vCookieData)) = true)  Then
	vCookieData = getCookieValue("compare")
End If
'wHikaku = Split(Request.Cookies("compare"), "$")
wHikaku = Split(vCookieData, "$")
'2014/01/31 GV mod end
wRecCnt = Ubound(wHikaku)

ReDim CategoryCd(wRecCnt+1)
ReDim MakerCd(wRecCnt)
ReDim ProductCd(wRecCnt)
ReDim Iro(wRecCnt)
ReDim Kikaku(wRecCnt)
ReDim MakerName(wRecCnt)
ReDim ProductName(wRecCnt)

For i=1 to wRecCnt
	vTemp = Split(wHikaku(i), "^")
	CategoryCd(i) = ReplaceInput(Trim(vTemp(0)))
	MakerCd(i) = ReplaceInput(Trim(vTemp(1)))
	ProductCd(i) = ReplaceInput(Trim(vTemp(2)))
	Iro(i) = ReplaceInput(Trim(vTemp(3)))
	Kikaku(i) = ReplaceInput(Trim(vTemp(4)))
Next

'---- 複数カテゴリーチェック
vMoreThanOneCategoryFl = false
For i=2 to wRecCnt
	if CategoryCd(1) <> CategoryCd(i) then
		vMoreThanOneCategoryFl = true
		Exit For
	end if
Next

if vMoreThanOneCategoryFl = true OR wRecCnt > 5 then

'---- 比較商品データ取り出し
	call getCompareProduct()

'---- カテゴリー，メーカー名，商品名順にソート
	call SortProduct()

'---- カテゴリー別に比較商品一覧作成
	i = 1
	wHTML = ""

'---- ナビゲーションセット
	call SetNavi(i)
	wHTML = wHTML & wNaveWithLink

	Do until i > wRecCnt

		'---- カテゴリータイトルセット (カテゴリーブレーク)
		if vOldCategoryCd <> CategoryCd(i) then
			call SetTitle(i)
			wHTML = wHTML & wTitleWithLink

			'2012/07/18 nt add
			wHTML = wHTML & "<form onSubmit='return Hikaku_onSubmit(this);'>" & vbNewLine
			wHTML = wHTML & "<dl class='productcompare'>" & vbNewLine

			'2012/07/18 nt del
			'wHTML = wHTML & "<table border='1' cellspacing='0' cellpadding='3'>" & vbNewLine
			'wHTML = wHTML & "<form onSubmit='return Hikaku_onSubmit(this);'>" & vbNewLine

			'---- タイトル
			'2012/07/18 nt add
			wHTML = wHTML & "<dt>" & vbNewLine
			wHTML = wHTML & "<ul>" & vbNewLine
			wHTML = wHTML & "<li>比較</li>" & vbNewLine
			wHTML = wHTML & "<li>商品名</li>" & vbNewLine
			wHTML = wHTML & "</ul>" & vbNewLine
			wHTML = wHTML & "</dt>" & vbNewLine
			wHTML = wHTML & "<dd>" & vbNewLine
			wHTML = wHTML & "<ul>" & vbNewLine

			'2012/07/18 nt del
			'wHTML = wHTML & "  <tr bgcolor='#cccccc' class='honbun'>"
			'wHTML = wHTML & "    <td width='50' align='center' nowrap>比較</td>" & vbNewLine
			'wHTML = wHTML & "    <td width='500' align='center' nowrap>商品名</td>" & vbNewLine
			'wHTML = wHTML & "  </tr>"

			vOldCategoryCd = CategoryCd(i)
		end if

		'---- 商品名/色/規格
		'2012/07/18 nt add
		vProductName = ProductName(i)
		if Trim(Iro(i)) <> "" then
			vProductName = vProductName & "/" & Trim(Iro(i))
		end if
		if Trim(Kikaku(i)) <> "" then
			vProductName = vProductName & "/" & Trim(Kikaku(i))
		end if

		'---- 選択チェックボックス
		'2012/07/18 nt add
		wHTML = wHTML & "          <li><span><input type='checkbox' name='iItem' value='$" & CategoryCd(i) & "^" & MakerCd(i) & "^" & ProductCd(i) & "^" & Trim(Iro(i)) & "^" & Trim(kikaku(i)) & "' checked></span>" & MakerName(i) & "<a href='ProductDetail.asp?item=" & MakerCd(i) & "^" & ProductCd(i) & "^" & Trim(Iro(i)) & "^" &  Trim(kikaku(i)) & "'>" & vProductName & "</a></li>" & vbNewLine

		'2012/07/18 nt del
		'wHTML = wHTML & "  <tr>"
		'wHTML = wHTML & "    <td align='center' valign='middle' nowrap>" & vbNewLine
		'wHTML = wHTML & "      <input type='checkbox' name='iItem' value='$" & CategoryCd(i) & "^" & MakerCd(i) & "^" & ProductCd(i) & "^^" & Trim(Iro(i)) & "^" & Trim(kikaku(i)) & "' CHECKED>" & vbNewLine
		'wHTML = wHTML & "    </td>" & vbNewLine

		'2012/07/18 nt del
		'---- メーカー
		'wHTML = wHTML & "    <td align='left' nowrap>" & vbNewLine
		'wHTML = wHTML & "      <span class='honbun'>" & MakerName(i) & "</span><br>" & vbNewLine

		'2012/07/18 nt del
		'---- 商品名/色/規格
		'vProductName = ProductName(i)
		'if Trim(Iro(i)) <> "" then
		'	vProductName = vProductName & "/" & Trim(Iro(i))
		'end if
		'if Trim(Kikaku(i)) <> "" then
		'	vProductName = vProductName & "/" & Trim(Kikaku(i))
		'end if

		'2012/07/18 nt del
		'wHTML = wHTML & "    <a href='ProductDetail.asp?item=" & MakerCd(i) & "^" & ProductCd(i) & "^" & Iro(i) & "^" & Kikaku(i) & "' class='link'>" & vProductName & "</a>" & vbNewLine
		'wHTML = wHTML & "    </td>" & vbNewLine
		'wHTML = wHTML & "  </tr>" & vbNewLine

		'---- 次データをチェック
		i = i + 1

		'---- 比較ボタン
		if i > wRecCnt OR vOldCategoryCd <> CategoryCd(i) then
			'2012/07/18 nt add
			wHTML = wHTML & "</ul>" & vbNewLine
			wHTML = wHTML & "</dd>" & vbNewLine
			wHTML = wHTML & "</dl>" & vbNewLine
			wHTML = wHTML & "<p class='btnBox'><input type='submit' value='比較する' class='opover'></p>"
			wHTML = wHTML & "</form>"

			'2012/07/18 nt del
			'wHTML = wHTML & "  <tr>"
			'wHTML = wHTML & "    <td align='center' valign='middle' colspan=2>" & vbNewLine
			'wHTML = wHTML & "      <input type='image' src='images/Hikaku2.gif' border=0>" & vbNewLine
			'wHTML = wHTML & "    </td>" & vbNewLine
			'wHTML = wHTML & "  </tr>"
			'wHTML = wHTML & "</form>" & vbNewLine
			'wHTML = wHTML & "</table><br>" & vbNewLine
		end if

	Loop
	wGotoCompareFl = false
else
	wGotoCompareFl = true
end if

End Function

'========================================================================
'
'	Function	カテゴリー，メーカー名，商品名順にソート
'
'========================================================================
'
Function SortProduct()

Dim i
Dim RSv

wSQL = ""
For i=1 to wRecCnt
	if i > 1 then
		wSQL = wSQL & " UNION "
	end if
	wSQL = wSQL & "SELECT '" & CategoryCd(i) & "' AS CategoryCd"
	wSQL = wSQL & "     , '" & MakerName(i) & "' AS MakerName"
	wSQL = wSQL & "     , '" & ProductName(i) & "' AS ProductName"
	wSQL = wSQL & "     , '" & MakerCd(i) & "' AS MakerCd"
	wSQL = wSQL & "     , '" & ProductCd(i) & "' AS ProductCd"
	wSQL = wSQL & "     , '" & Iro(i) & "' AS Iro"
	wSQL = wSQL & "     , '" & Kikaku(i) & "' AS Kikaku"
Next
wSQL = wSQL & " ORDER BY 1, 2, 3"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic

i = 1
Do until RSv.EOF = true
	CategoryCd(i) = RSv("CategoryCd")
	MakerCd(i) = RSv("MakerCd")
	ProductCd(i) = RSv("ProductCd")
	Iro(i) = RSv("Iro")
	Kikaku(i) = RSv("Kikaku")
	MakerName(i) = RSv("MakerName")
	ProductName(i) = RSv("ProductName")
	RSv.MoveNext
	i = i + 1
Loop

RSv.Close

End function

'========================================================================
'
'	Function	比較商品データ取り出し
'
'========================================================================
'
Function getCompareProduct()

Dim i
Dim RSv

for i=1 to wRecCnt
	'---- 商品Recordset作成
	wSQL = ""
' 2012/01/20 GV Mod Start
'	wSQL = wSQL & "SELECT b.メーカー名"
'	wSQL = wSQL & "     , a.商品名"
'	wSQL = wSQL & "  FROM Web商品 a WITH (NOLOCK)"
'	wSQL = wSQL & "     , メーカー b WITH (NOLOCK)"
'	wSQL = wSQL & " WHERE b.メーカーコード = a.メーカーコード"
'	wSQL = wSQL & "   AND a.メーカーコード = '" & MakerCd(i) & "'"
'	wSQL = wSQL & "   AND a.商品コード = '" & ProductCd(i) & "'"
	wSQL = wSQL & "SELECT "
	wSQL = wSQL & "      b.メーカー名 "
	wSQL = wSQL & "    , a.商品名 "
	wSQL = wSQL & "FROM "
	wSQL = wSQL & "    Web商品               a WITH (NOLOCK) "
	wSQL = wSQL & "      INNER JOIN メーカー b WITH (NOLOCK) "
	wSQL = wSQL & "        ON     b.メーカーコード = a.メーカーコード "
	wSQL = wSQL & "WHERE "
	wSQL = wSQL & "        a.メーカーコード = '" & MakerCd(i) & "' "
	wSQL = wSQL & "    AND a.商品コード     = '" & Replace(ProductCd(i), "'", "''") & "' "
' 2012/01/20 GV Mod End

	Set RSv = Server.CreateObject("ADODB.Recordset")
	RSv.Open wSQL, Connection, adOpenStatic

	MakerName(i) = RSv("メーカー名")
	ProductName(i) = RSv("商品名")

	RSv.MoveNext
Next

RSv.Close

End function

'========================================================================
'
'	Function	タイトルセット
'
'========================================================================
'
Function SetTitle(i)

Dim RSv

'---- タイトルセット
wSQL = ""
' 2012/01/20 GV Mod Start
'wSQL = wSQL & "SELECT a.大カテゴリーコード"
'wSQL = wSQL & "     , a.大カテゴリー名"
'wSQL = wSQL & "     , b.中カテゴリーコード"
'wSQL = wSQL & "     , b.中カテゴリー名日本語"
'wSQL = wSQL & "     , c.カテゴリーコード"
'wSQL = wSQL & "     , c.カテゴリー名"
'wSQL = wSQL & "     , c.お勧めカテゴリーフラグ"
'wSQL = wSQL & "  FROM 大カテゴリー a"
'wSQL = wSQL & "     , 中カテゴリー b"
'wSQL = wSQL & "     , カテゴリー c"
'wSQL = wSQL & " WHERE b.大カテゴリーコード = a.大カテゴリーコード"
'wSQL = wSQL & "   AND c.中カテゴリーコード = b.中カテゴリーコード"
'wSQL = wSQL & "   AND c.カテゴリーコード = '" & CategoryCd(i) & "'"
wSQL = wSQL & "SELECT "
wSQL = wSQL & "       a.大カテゴリーコード "
wSQL = wSQL & "     , a.大カテゴリー名 "
wSQL = wSQL & "     , b.中カテゴリーコード "
wSQL = wSQL & "     , b.中カテゴリー名日本語 "
wSQL = wSQL & "     , c.カテゴリーコード "
wSQL = wSQL & "     , c.カテゴリー名 "
wSQL = wSQL & "     , c.お勧めカテゴリーフラグ "
wSQL = wSQL & "FROM "
wSQL = wSQL & "    大カテゴリー              a WITH (NOLOCK) "
wSQL = wSQL & "      INNER JOIN 中カテゴリー b WITH (NOLOCK) "
wSQL = wSQL & "        ON     b.大カテゴリーコード = a.大カテゴリーコード "
wSQL = wSQL & "      INNER JOIN カテゴリー   c WITH (NOLOCK) "
wSQL = wSQL & "        ON     c.中カテゴリーコード = b.中カテゴリーコード "
wSQL = wSQL & "WHERE "
wSQL = wSQL & "         c.カテゴリーコード = '" & CategoryCd(i) & "' "
' 2012/01/20 GV Mod End

'@@@@@@@@@@response.write(wSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic

wTitleWithLink = ""
'wTitleWithLink = wTitleWithLink & "<h2 class='title'>" & RSv("大カテゴリー名") & " / " & RSv("中カテゴリー名日本語") & " / " & RSv("カテゴリー名") & "</h2>" & vbNewLine
wTitleWithLink = wTitleWithLink & "<h2 class='title'>" & RSv("カテゴリー名") & "</h2>" & vbNewLine
RSv.close

End Function

'========================================================================
'
'	Function	ナビゲーションセット
'
'========================================================================
'
Function SetNavi(i)

Dim RSv

'---- ナビゲーションセット
wSQL = ""
' 2012/01/20 GV Mod Start
'wSQL = wSQL & "SELECT a.大カテゴリーコード"
'wSQL = wSQL & "     , a.大カテゴリー名"
'wSQL = wSQL & "     , b.中カテゴリーコード"
'wSQL = wSQL & "     , b.中カテゴリー名日本語"
'wSQL = wSQL & "     , c.カテゴリーコード"
'wSQL = wSQL & "     , c.カテゴリー名"
'wSQL = wSQL & "     , c.お勧めカテゴリーフラグ"
'wSQL = wSQL & "  FROM 大カテゴリー a"
'wSQL = wSQL & "     , 中カテゴリー b"
'wSQL = wSQL & "     , カテゴリー c"
'wSQL = wSQL & " WHERE b.大カテゴリーコード = a.大カテゴリーコード"
'wSQL = wSQL & "   AND c.中カテゴリーコード = b.中カテゴリーコード"
'wSQL = wSQL & "   AND c.カテゴリーコード = '" & CategoryCd(i) & "'"
wSQL = wSQL & "SELECT "
wSQL = wSQL & "       a.大カテゴリーコード "
wSQL = wSQL & "     , a.大カテゴリー名 "
wSQL = wSQL & "     , b.中カテゴリーコード "
wSQL = wSQL & "     , b.中カテゴリー名日本語 "
wSQL = wSQL & "     , c.カテゴリーコード "
wSQL = wSQL & "     , c.カテゴリー名 "
wSQL = wSQL & "     , c.お勧めカテゴリーフラグ "
wSQL = wSQL & "FROM "
wSQL = wSQL & "    大カテゴリー              a WITH (NOLOCK) "
wSQL = wSQL & "      INNER JOIN 中カテゴリー b WITH (NOLOCK) "
wSQL = wSQL & "        ON     b.大カテゴリーコード = a.大カテゴリーコード "
wSQL = wSQL & "      INNER JOIN カテゴリー   c WITH (NOLOCK) "
wSQL = wSQL & "        ON     c.中カテゴリーコード = b.中カテゴリーコード "
wSQL = wSQL & "WHERE "
wSQL = wSQL & "         c.カテゴリーコード = '" & CategoryCd(i) & "' "
' 2012/01/20 GV Mod End

'@@@@@@@@@@response.write(wSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic

wNaveWithLink = ""
wNaveWithLink = wNaveWithLink & "<div id='path_box'><div id='path_box_inner01'><div id='path_box_inner02'>" & vbNewLine
wNaveWithLink = wNaveWithLink & " <p class='home'><a href='../'><img src='../images/icon_home.gif' alt='HOME'></a></p>" & vbNewLine
wNaveWithLink = wNaveWithLink & " <ul id='path'>" & vbNewLine
'wNaveWithLink = wNaveWithLink & "  <li><a href='LargeCategoryList.asp?LargeCategoryCd=" & RSv("大カテゴリーコード") & "'>" & RSv("大カテゴリー名") & "</a></li>" & vbNewLine
'wNaveWithLink = wNaveWithLink & "  <li><a href='MidCategoryList.asp?MidCategoryCd=" & RSv("中カテゴリーコード") & "'>" & RSv("中カテゴリー名日本語") & "</a></li>" & vbNewLine
'wNaveWithLink = wNaveWithLink & "  <li><a href='SearchList.asp?i_type=c&s_category_cd=" & RSv("カテゴリーコード") & "'>" &  RSv("カテゴリー名") & "</a></li>" & vbNewLine
wNaveWithLink = wNaveWithLink & "  <li class='now'>商品比較</li>" & vbNewLine
wNaveWithLink = wNaveWithLink & "  </ul>" & vbNewLine
wNaveWithLink = wNaveWithLink & "</div></div></div>" & vbNewLine
wNaveWithLink = wNaveWithLink & "<h1 class='title'>商品比較</h1>" & vbNewLine
wNaveWithLink = wNaveWithLink & "<p class='error'>5個以上または複数のカテゴリーの商品が選択されました。<br>比較する商品を選択してください。</p>" & vbNewLine
RSv.close

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
<title>商品比較｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="Style/productcompare.css" type="text/css">

<script type="text/javascript">
//
//	Hikaku onSubmit
//
function Hikaku_onSubmit(pForm){

	var vParm = "";
	var vCnt = 0;

// Item count check
	for (var i=0; i<pForm.iItem.length; i++){
		if (pForm.iItem[i].checked == true){
			vCnt += 1;
			vParm += pForm.iItem[i].value;
		}
	}
	if (vCnt > 5){
		alert("5個以上の商品が選択されました。5個以内で選択してください。");
		return false;
	}
	if (vCnt < 2){
		alert("2個以上の商品を選択してください。");
		return false;
	}
	window.location = "ProductCompare.asp?item=" + vParm;
	return false;
}

</script>

</head>
<body>
<!--#include file="../Navi/Navitop.inc"-->
<div id="globalMain">
	<span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="ここから本文です"></a></span>
	<!-- コンテンツstart -->
	<div id="globalContents">
		<%=wHTML%>
	</div>
	<div id="globalSide">
		<!--#include file="../Navi/NaviSide.inc"-->
	</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>