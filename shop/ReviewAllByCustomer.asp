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
'	商品レビュー (投稿者別一覧)
'
'更新履歴
'2008/05/23 入力データチェック強化（LEFT, Numeric, EOF他)
'2008/12/24 在庫状況セット関数化
'2009/10/02 会員プロフィール文字色をFFFFFFに変更
'2010/09/09 an レビューが取得できない場合はエラーメッセージ表示
'2011/08/01 an #1087 Error.aspログ出力対応
'2012/07/30 if-web リニューアルレイアウト調整
'2014/03/19 GV 消費税増税に伴う2重表示対応
'
'========================================================================

On Error Resume Next

Dim userID

Dim CNo

Dim wHandleName
Dim wPrefecture
Dim wReviewCnt

Dim wReviewListHTML

Dim wProdTermFl
Dim wPrice
Dim wSalesTaxRate

Dim wItemChar1
Dim wItemChar2
Dim wItemNum1
Dim wItemNum2
Dim wItemDate1
Dim wItemDate2

Dim Connection
Dim RS

Dim wSQL
Dim wHTML
Dim wMSG
Dim wErrDesc   '2011/08/01 an add

'========================================================================

Response.buffer = true

'---- 呼び出し元からのデータ取り出し
CNo = ReplaceInput(Request("CNo"))
if isNumeric(CNo) = false then
	CNO = 0
end if

'---- Execute main
call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "ReviewAllByCustomer.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
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
'	Function	main proc
'
'========================================================================
'
Function main()

Dim vInventoryCd
Dim vInventoryImage
Dim i

'---- 該当顧客のレビュー取り出し
wSQL = ""
wSQL = wSQL & "SELECT a.*"
wSQL = wSQL & "     , b.顧客都道府県"
wSQL = wSQL & "     , c.商品名"
wSQL = wSQL & "     , c.商品画像ファイル名_小"
wSQL = wSQL & "     , c.販売単価"
wSQL = wSQL & "     , c.ASK商品フラグ"
wSQL = wSQL & "     , c.希少数量"
wSQL = wSQL & "     , c.セット商品フラグ"
wSQL = wSQL & "     , c.メーカー直送取寄区分"
wSQL = wSQL & "     , c.取扱中止日"
wSQL = wSQL & "     , c.廃番日"
wSQL = wSQL & "     , c.B品フラグ"
wSQL = wSQL & "     , c.Web納期非表示フラグ"
wSQL = wSQL & "     , c.入荷予定未定フラグ"
wSQL = wSQL & "     , c.個数限定数量"
wSQL = wSQL & "     , c.個数限定受注済数量"
wSQL = wSQL & "     , d.色"
wSQL = wSQL & "     , d.規格"
wSQL = wSQL & "     , d.引当可能入荷予定日"
wSQL = wSQL & "     , d.引当可能数量"
wSQL = wSQL & "     , d.B品引当可能数量"
wSQL = wSQL & "     , e.メーカー名"
wSQL = wSQL & "  FROM 商品レビュー a WITH (NOLOCK)"
wSQL = wSQL & "     , Web顧客住所 b WITH (NOLOCK)"
wSQL = wSQL & "     , Web商品 c WITH (NOLOCK)"
wSQL = wSQL & "     , Web色規格別在庫 d WITH (NOLOCK)"
wSQL = wSQL & "     , メーカー e WITH (NOLOCK)"
wSQL = wSQL & " WHERE b.顧客番号 = a.顧客番号"
wSQL = wSQL & "   AND b.住所連番 = 1"
wSQL = wSQL & "   AND c.メーカーコード = a.メーカーコード"
wSQL = wSQL & "   AND c.商品コード = a.商品コード"
wSQL = wSQL & "   AND d.メーカーコード = a.メーカーコード"
wSQL = wSQL & "   AND d.商品コード = a.商品コード"
wSQL = wSQL & "   AND d.色 = ''"
wSQL = wSQL & "   AND d.規格 = ''"
wSQL = wSQL & "   AND e.メーカーコード = a.メーカーコード"
wSQL = wSQL & "   AND a.顧客番号 = " & CNo 
wSQL = wSQL & " ORDER BY a.ID DESC" 

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic, adLockOptimistic

'@@@@response.write(wSQL)

if RS.EOF = true then
	wMSG = "<p class='error'>該当レビューが登録されていません｡</p>"
else   '2010/0909 an mod

	'---- 消費税率取出し
	call getCntlMst("共通","消費税率","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)
	wSalesTaxRate = Clng(wItemNum1)

	'----
	wHandleName = RS("名前")
	wPrefecture = RS("顧客都道府県")
	wReviewCnt = RS.RecordCount
	wHTML = ""

	Do until RS.EOF = true

	'---- 廃番チェック
		if  (isNull(RS("取扱中止日")) = true AND isNull(RS("廃番日")) = true) _
		 OR (isNull(RS("廃番日")) = false AND RS("引当可能数量") > 0) then
			wProdTermFl = "N"
		else
			wProdTermFl = "Y"
		end if

	'----
		wHTML = wHTML & "<table width='480' cellSpacing='0' cellPadding='0' border='0'>" & vbNewLine
		wHTML = wHTML & "  <tr>" & vbNewLine

	'---- 商品画像
		wHTML = wHTML & "    <td width='110' align='center' valign='top' rowspan='2'>" & vbNewLine
		wHTML = wHTML & "      <a href='ProductDetail.asp?item=" & RS("メーカーコード") & "^" & RS("商品コード") & "'><img src='../shop/prod_img/" & RS("商品画像ファイル名_小") & "' width='100' height='50'></a>" & vbNewLine
		wHTML = wHTML & "    </td>" & vbNewLine

	'---- メーカー名、商品名
		wHTML = wHTML & "    <td width='220'>" & vbNewLine
		wHTML = wHTML & "      " & RS("メーカー名") & "<br>" & vbNewLine
		wHTML = wHTML & "      <a href='ProductDetail.asp?item=" & RS("メーカーコード") & "^" & RS("商品コード") & "'>" & RS("商品名") & "</a>" & vbNewLine
		wHTML = wHTML & "    </td>" & vbNewLine

	'---- 在庫状況
		vInventoryCd = GetInventoryStatus(RS("メーカーコード"),RS("商品コード"),RS("色"),RS("規格"),RS("引当可能数量"),RS("希少数量"),RS("セット商品フラグ"),RS("メーカー直送取寄区分"),RS("引当可能入荷予定日"),wProdTermFl)

		'---- 在庫状況、色を最終セット
		call GetInventoryStatus2(RS("引当可能数量"), RS("Web納期非表示フラグ"), RS("入荷予定未定フラグ"), RS("廃番日"), RS("B品フラグ"), RS("B品引当可能数量"), RS("個数限定数量"), RS("個数限定受注済数量"), wProdTermFl, vInventoryCd, vInventoryImage)

		wHTML = wHTML & "    <td width='150' nowrap>" & vbNewLine
		wHTML = wHTML & "      在庫状況：<img src='images/" & vInventoryImage & "' width='10' height='10' class='inventoryImage'> " & vInventoryCd & "<br>" & vbNewLine

	'----- 衝撃特価
		wHTML = wHTML & "      衝撃特価："

		if RS("ASK商品フラグ") = "Y" then
			wHTML = wHTML & "ASK" & vbNewLine
		else
			wPrice = calcPrice(RS("販売単価"), wSalesTaxRate)
'2014/03/19 GV mod start ---->
'				wHTML = wHTML & FormatNumber(wPrice,0) & "円(税込)" & vbNewLine
				wHTML = wHTML & FormatNumber(RS("販売単価"),0) & "円(税抜)<br>" & vbNewLine
				wHTML = wHTML & "(税込&nbsp;" & FormatNumber(wPrice,0) & "円)" & vbNewLine
'2014/03/19 GV mod end   <----
		end if
		wHTML = wHTML & "    </td>" & vbNewLine
		wHTML = wHTML & "  </tr>" & vbNewLine

	'---- 参考になった人数
		wHTML = wHTML & "  <tr>" & vbNewLine
		wHTML = wHTML & "    <td>参考になった人数：" & RS("参考数") & "人(" & RS("参考数") + RS("不参考数") & "人中)</td>" & vbNewLine

	'---- カート
		if wProdTermFl <> "Y" then
			wHTML = wHTML & "    <td>" & vbNewLine
			wHTML = wHTML & "      <form name='f_data' method='post' action='OrderPreInsert.asp'>" & vbNewLine
			wHTML = wHTML & "        <input type='text' name='qt' size='2' maxsize='4' value='1'>" & vbNewLine
			wHTML = wHTML & "        <input type='image' src='images/btn_cart.png' class='cartBtn opover'>" & vbNewLine
			wHTML = wHTML & "        <input type='hidden' name='maker_cd' value='" & RS("メーカーコード") & "'>" & vbNewLine
			wHTML = wHTML & "        <input type='hidden' name='product_cd' value='" & RS("商品コード") & "'>" & vbNewLine
			wHTML = wHTML & "        <input type='hidden' name='iro' value=''>" & vbNewLine
			wHTML = wHTML & "        <input type='hidden' name='kikaku' value=''>" & vbNewLine
			wHTML = wHTML & "      </form>" & vbNewLine
			wHTML = wHTML & "    </td>" & vbNewLine
		else
			wHTML = wHTML & "    <td><img src='images/icon_sold.gif' alt='完売'></td>" & vbNewLine
		end if
		wHTML = wHTML & "  </tr>" & vbNewLine
		wHTML = wHTML & "</table>" & vbNewLine

	'---- レビュー内容
		wHTML = wHTML & "<table cellSpacing='0' cellPadding='0' width='480' border='0'>" & vbNewLine
		wHTML = wHTML & "  <tr>" & vbNewLine

	'---- おすすめ度
		wHTML = wHTML & "    <td width='130'>" & vbNewLine
		wHTML = wHTML & "      "
		For i=1 to RS("評価")
			wHTML = wHTML & "<img src='images/review_icon10.png'>"
		Next
		For i=RS("評価")+1 to 5
			wHTML = wHTML & "<img src='images/review_icon00.png'>"
		Next
		wHTML = wHTML & " (" & FormatNumber(RS("評価"), 1) & ")" & vbNewLine
		wHTML = wHTML & "    </td>" & vbNewLine

	'---- タイトル, 投稿日
		wHTML = wHTML & "    <td width='270'><b>" & RS("タイトル") & "</b></td>" & vbNewLine
		wHTML = wHTML & "    <td width='80'>" & cf_FormatDate(RS("投稿日"), "YYYY/MM/DD") & "</td>" & vbNewLine
		wHTML = wHTML & "  </tr>" & vbNewLine

	'---- レビュー内容
		wHTML = wHTML & "  <tr>" & vbNewLine
		wHTML = wHTML & "    <td colspan='3' width='480'>" & Replace(RS("レビュー内容"), vbNewline, "<br>") & "</td>" & vbNewLine
		wHTML = wHTML & "  </tr>" & vbNewLine

	'---- 区切り線
		wHTML = wHTML & "  <tr>" & vbNewLine
		wHTML = wHTML & "    <td colSpan='3' height='5'><hr size='1'></td>" & vbNewLine
		wHTML = wHTML & "  </tr>" & vbNewLine

		wHTML = wHTML & "</table>" & vbNewLine

		RS.MoveNext
	Loop
end if     '2010/09/09 an mod

RS.close

wReviewListHTML = wHTML

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
<title>商品レビュー（投稿者別）｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/ReviewAllByCustomer.css" type="text/css">
<link rel="stylesheet" href="style/ask.css?20140401a" type="text/css">
</head>
<body>
<!--#include file="../Navi/Navitop.inc"-->
<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="ここから本文です"></a></span>

<!-- コンテンツstart -->
<div id="globalContents">
<!--
  <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
    <p class="home"><a href="../"><img src="../images/icon_home.gif" alt="HOME"></a></p>
    <ul id="path">
      <li class="now"><%=wHandleName%> さんのレビュー一覧</li>
    </ul>
  </div></div></div>
-->
<% if wMSG <> "" then %>
	<%=wMSG%>
<% else %>

  <h1 class="title"><%=wHandleName%> さんのレビュー一覧</h1>

  <div id="main_container">

    <div id="rewiewlist">

<%=wReviewListHTML%>

    </div>

    <div id="detail_side">

      <div class='detail_side_inner01'><div class='detail_side_inner02'>
        <div class='detail_side_inner_box' id='subtotal'>
          <h4 class='detail_sub'><%=wHandleName%> さんのプロフィール</h4>
          <p>レビュー投稿数：<%=wReviewCnt%>件</p>
          <p>住所：<%=wPrefecture%></p>
        </div>
      </div></div>

    </div>

  </div>
<% end if%>

  </div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/Navibottom.inc"-->
<div class="tooltip"><p>ASK</p></div>
<!--#include file="../Navi/NaviScript.inc"-->
<script type="text/javascript" src="jslib/ask.js?20140401a"></script>
</body>
</html>