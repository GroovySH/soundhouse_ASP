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
'	ウィッシュリスト一覧
'
'更新履歴
'2008/01/18 在庫状況　表示の色を変更
'						個数限定単価もB品と同様の単価表示に変更
'2008/12/24 在庫状況セット関数化
'2009/09/09	カートへ入れるときに、ウィッシュリストから削除するかどうかを問い合わせる
'2011/08/01 an #1087 Error.aspログ出力対応
'2011/10/19 hn 1063 ASK表示方法変更
'2012/07/23 if-web リニューアルレイアウト調整
'2012/08/14 GV #1419 未ログイン時ウィッシュリストからログイン画面を表示する
'2012/09/07 nt ブックマーク等で未ログインかつ直接ページ遷移時もログイン画面を表示
'2014/03/19 GV 消費税増税に伴う2重表示対応
'
'========================================================================

On Error Resume Next

Dim userID

Dim wNotLogin					' ログインしていない	' 2012/08/14 GV #1419 Add

Dim wSalesTaxRate
Dim wPrice
Dim wProdTermFl
Dim wItem

Dim wListHTML
Dim wCartHTML

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

'---- UserID 取り出し
userID = Session("userID")

wMSG = ReplaceInput(Request("msg"))

wNotLogin = False				' 初期状態はログインしている事を前提とする	' 2012/08/14 GV #1419 Add

'---- Execute main
call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "WishList.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

call close_db()

if Err.Description <> "" then	
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

' 2012/08/14 GV #1419 Add Start
If wNotLogin = True Then
	'---- ログインしていない場合はログインページへ
	Session("msg") = wMsg

	'2012/09/07 nt mod Start
	'---- ログイン後、ウィッシュリスト画面表示のため、パラメータ追加
	'Server.Transfer "shop/Login.asp"
	Response.Redirect "../shop/Login.asp?called_from=wishlist"
	'2012/09/07 nt mod End

End If
' 2012/08/14 GV #1419 Add End

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

Dim vPrice
Dim vInventoryCd
Dim vInventoryImage
Dim vProductName
Dim vItem

wListHTML = ""

' 2012/08/14 GV #1419 Mod Start
'if userID = "" then
'	wListHTML = wListHTML & "<p class='error'>ログインをしてください。</p>" & vbNewLine
'	exit function
'end if

Dim vRS

If userID = "" Then
	'---- ログインしていなければエラー　｢ログインしてください。｣
	wNotLogin = True		' ログインされていない
	wMsg = "ログインしてください。"
	Exit Function
End If

' 顧客情報取得
Set vRS = get_customer()

If vRS.EOF = True Then
	'---- Session("userID")で顧客情報が取出せなければエラー　｢ログインしてください。｣
	wNotLogin = True		' ログインされていない
	wMsg = "ログインしてください。"
	Exit Function
End If
' 2012/08/14 GV #1419 Mod End

'---- 消費税率取出し
call getCntlMst("共通","消費税率","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)			'消費税率
wSalesTaxRate = Clng(wItemNum1)

wHTML = ""

'----ウィッシュリスト取り出し
wSQL = ""
wSQL = wSQL & "SELECT DISTINCT"
wSQL = wSQL & "       a.メーカーコード"	
wSQL = wSQL & "     , a.商品コード"
wSQL = wSQL & "     , a.商品名"
wSQL = wSQL & "     , a.商品概略Web"
wSQL = wSQL & "     , a.商品画像ファイル名_小"
wSQL = wSQL & "     , a.販売単価"
wSQL = wSQL & "     , a.個数限定単価"
wSQL = wSQL & "     , a.個数限定数量"
wSQL = wSQL & "     , a.個数限定受注済数量"
wSQL = wSQL & "     , CASE"
wSQL = wSQL & "         WHEN (a.個数限定数量 > a.個数限定受注済数量 AND a.個数限定数量 > 0) THEN 'Y'"
wSQL = wSQL & "         ELSE 'N'"
wSQL = wSQL & "       END AS 個数限定単価フラグ"
wSQL = wSQL & "     , a.メーカー直送取寄区分"
wSQL = wSQL & "     , a.ASK商品フラグ"
wSQL = wSQL & "     , a.取扱中止日"
wSQL = wSQL & "     , a.廃番日"
wSQL = wSQL & "     , a.終了日"
wSQL = wSQL & "     , a.希少数量"
wSQL = wSQL & "     , a.セット商品フラグ"	
wSQL = wSQL & "     , a.Web納期非表示フラグ"	
wSQL = wSQL & "     , a.入荷予定未定フラグ"
wSQL = wSQL & "     , a.B品単価"
wSQL = wSQL & "     , a.完売日"
wSQL = wSQL & "     , a.B品フラグ"
wSQL = wSQL & "     , b.色"
wSQL = wSQL & "     , b.規格"
wSQL = wSQL & "     , b.引当可能数量"
wSQL = wSQL & "     , b.引当可能入荷予定日"
wSQL = wSQL & "     , b.B品引当可能数量"
wSQL = wSQL & "     , c.メーカー名"
wSQL = wSQL & "     , d.登録日"

'---- FROM
wSQL = wSQL & "  FROM Web商品 a WITH (NOLOCK)"
wSQL = wSQL & "     , Web色規格別在庫 b WITH (NOLOCK)"
wSQL = wSQL & "     , メーカー c WITH (NOLOCK)"
wSQL = wSQL & "     , ウィッシュリスト d WITH (NOLOCK)"

'---- WHERE
wSQL = wSQL & " WHERE a.Web商品フラグ = 'Y'"
wSQL = wSQL & "   AND b.メーカーコード = a.メーカーコード"
wSQL = wSQL & "   AND b.商品コード = a.商品コード"
wSQL = wSQL & "   AND c.メーカーコード = a.メーカーコード"
wSQL = wSQL & "   AND b.メーカーコード = d.メーカーコード"
wSQL = wSQL & "   AND b.商品コード = d.商品コード"
wSQL = wSQL & "   AND b.色 = d.色"
wSQL = wSQL & "   AND b.規格 = d.規格"
wSQL = wSQL & "   AND d.顧客番号 = " & userID

'---- ORDER BY
wSQL = wSQL & " ORDER BY c.メーカー名, a.商品名"

'@@@@@@response.write(wSQL)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

'---- 商品一覧作成
if RS.EOF = true then
	wListHTML = wListHTML & "<p class='error'>ウィッシュリストに商品がありません。</p>" & vbNewLine

else
	wListHTML = wListHTML & "<table width='480' border='0' cellspacing='0' cellpadding='0'>" & vbNewLine

'wListHTML = wListHTML & "  <tr>" & vbNewLine
'wListHTML = wListHTML & "    <td height='5' colspan='3'>※商品をカートに入れますとウィッシュリストから削除されます。</td>" & vbNewLine
'wListHTML = wListHTML & "  </tr>" & vbNewLine

'---- 区切り線
wListHTML = wListHTML & "  <tr>" & vbNewLine
wListHTML = wListHTML & "    <td height='5' colspan='3'><hr size='1'></td>" & vbNewLine
wListHTML = wListHTML & "  </tr>" & vbNewLine

	Do until RS.EOF = true
	
		wListHTML = wListHTML & "  <tr align='left' valign='middle'>" & vbNewLine

		wListHTML = wListHTML & "    <form name='f_item' method='post' action='WishListToCartDelete.asp' onSubmit='return order_onClick(this);'>" & vbNewLine

		'---- 終了チェック
		wProdTermFl = "N"
		if isNull(RS("取扱中止日")) = false then		'取扱中止
			wProdTermFl = "Y"
		end if
		if isNull(RS("廃番日")) = false AND RS("引当可能数量") <= 0 then		'廃番で在庫無し
			wProdTermFl = "Y"
		end if
		if isNull(RS("完売日")) = false then		'完売商品
			wProdTermFl = "Y"
		end if

	'----- 商品画像 
		vItem = "Item=" & Server.URLEncode(RS("メーカーコード") & "^" & RS("商品コード") & "^" & Trim(RS("色")) & "^" & Trim(RS("規格")))
		wListHTML = wListHTML & "    <td width='110' align='center' valign='top' rowspan='2'>" & vbNewLine
		wListHTML = wListHTML & "      <a href='" & g_HTTP & "shop/ProductDetail.asp?" & vItem & "'>"
		if Trim(RS("商品画像ファイル名_小")) <> "" then 
			wListHTML = wListHTML & "      <img src='prod_img/" & RS("商品画像ファイル名_小") & "' width='100' height='50'></a>" & vbNewLine
		end if
		wListHTML = wListHTML & "    </td>" & vbNewLine

	'----メーカー名
		wListHTML = wListHTML & "    <td width='220' valign='top' nowrap>" & vbNewLine
		wListHTML = wListHTML & "      <span>"  & RS("メーカー名") & "</span><br>"

	'----- 商品名/色/規格
		wListHTML = wListHTML & "      <a href='" & g_HTTP & "shop/ProductDetail.asp?" & vItem & "'>"
		vProductName = RS("商品名")
		if Trim(RS("色")) <> "" then
			vProductName = vProductName & "/" & Trim(RS("色"))
		end if
		if Trim(RS("規格")) <> "" then
			vProductName = vProductName & "/" & Trim(RS("規格"))
		end if
	 	wListHTML = wListHTML & vProductName & "</a>" & vbNewLine
		wListHTML = wListHTML & "    </td>" & vbNewLine

		wListHTML = wListHTML & "    <td width='150' valign='top' nowrap>" & vbNewLine

	'---- 登録日
		wListHTML = wListHTML & "      登録日：" & fFormatDate(RS("登録日")) & "<br>"

	'----- 販売単価
		vPrice = calcPrice(RS("販売単価"), wSalesTaxRate)

		if RS("B品フラグ") = "Y" OR RS("個数限定単価フラグ") = "Y" then
			wListHTML = wListHTML & "      <del>衝撃特価："
		else
			wListHTML = wListHTML & "      衝撃特価："
		end if

		if RS("ASK商品フラグ") = "Y" then
'2011/10/19 hn mod s
'			wListHTML = wListHTML & "<a href='JavaScript:void(0);' onClick=""askWin=window.open('AskPrice.asp?MakerName=" & Server.URLEncode(RS("メーカー名")) & "&ProductName=" & Server.URLEncode(vProductName) & "&Price=" & vPrice & "' ,'ask', 'width=250 height=80 scrollbars=false memubar=false toolbar=false statusbar=false personalbar=false locationbar=false');"" class='link'>ASK</a>" & vbNewLine
'2014/03/19 GV mod start ---->
'			wListHTML = wListHTML & "<a class='tip'>ASK<span>" & FormatNumber(vPrice,0) & "円(税込)</span></a>" & vbNewLine
			wListHTML = wListHTML & "<a class='tip'>ASK<span class='exc-tax'>" & FormatNumber(RS("販売単価"),0) & "円(税抜)</span><br>"
			wListHTML = wListHTML & "<span class='inc-tax'>(税込&nbsp;" & FormatNumber(vPrice,0) & "円)</span></a>" & vbNewLine
'2014/03/19 GV mod end   <----
'2011/10/19 hn mod e

		else
'2014/03/19 GV mod start ---->
'				wListHTML = wListHTML & FormatNumber(vPrice,0) & "円(税込)" & vbNewLine
				wListHTML = wListHTML & FormatNumber(RS("販売単価"),0) & "円(税抜)<br>"
				wListHTML = wListHTML & "(税込&nbsp;" & FormatNumber(vPrice,0) & "円)"
'2014/03/19 GV mod end   <----
		end if

		if RS("B品フラグ") = "Y" OR RS("個数限定単価フラグ") = "Y" then
			wListHTML = wListHTML & "</del>"
		end if

	'---- B品単価
		if RS("B品フラグ") = "Y" then
			vPrice = calcPrice(RS("B品単価"), wSalesTaxRate)
			wListHTML = wListHTML & "      <br><b>B品特価：</b>"
'2014/03/19 GV mod start ---->
'			wListHTML = wListHTML & FormatNumber(vPrice,0) & "円(税込)" & vbNewLine
			wListHTML = wListHTML & FormatNumber(RS("B品単価"),0) & "円(税抜)<br>"
			wListHTML = wListHTML & "(税抜&nbsp;" & FormatNumber(vPrice,0) & "円)"
'2014/03/19 GV mod end   <----
		end if

	'---- 個数限定単価
		if RS("個数限定単価フラグ") = "Y" then
			vPrice = calcPrice(RS("個数限定単価"), wSalesTaxRate)
			wListHTML = wListHTML & "      <br><b>限定特価：</b>"
'2014/03/19 GV mod start ---->
'			wListHTML = wListHTML & FormatNumber(vPrice,0) & "円(税込)" & vbNewLine
			wListHTML = wListHTML & FormatNumber(RS("個数限定単価"),0) & "円(税抜)<br>"
			wListHTML = wListHTML & "(税込&nbsp;" & FormatNumber(vPrice,0) & "円)"
'2014/03/19 GV mod end   <----
		end if

		wListHTML = wListHTML & "    </td>" & vbNewLine


		wListHTML = wListHTML & "  </tr>" & vbNewLine

	'----- 商品概略Web
		wListHTML = wListHTML & "  <tr align='left' valign='middle'>" & vbNewLine
		if Trim(RS("商品概略Web")) = "" OR isNull(Trim(RS("商品概略Web")))then
			wListHTML = wListHTML & "    <td>　</td>" & vbNewLine
		else
			wListHTML = wListHTML & "    <td>" & RS("商品概略Web") & "</td>" & vbNewLine
		end if

		wListHTML = wListHTML & "    <td valign='top' nowrap>" & vbNewLine

	'----- 在庫状況表示（色規格なし商品のみ）
		vInventoryCd = GetInventoryStatus(RS("メーカーコード"),RS("商品コード"),RS("色"),RS("規格"),RS("引当可能数量"),RS("希少数量"),RS("セット商品フラグ"),RS("メーカー直送取寄区分"),RS("引当可能入荷予定日"),wProdTermFl)

	'---- 在庫状況、色を最終セット
		call GetInventoryStatus2(RS("引当可能数量"), RS("Web納期非表示フラグ"), RS("入荷予定未定フラグ"), RS("廃番日"), RS("B品フラグ"), RS("B品引当可能数量"), RS("個数限定数量"), RS("個数限定受注済数量"), wProdTermFl, vInventoryCd, vInventoryImage)

		wListHTML = wListHTML & "      在庫状況：<img src='images/" & vInventoryImage & "' width='10' height='10' class='inventoryImage'> " & vInventoryCd & "<br>"

		wItem= Trim(RS("メーカーコード")) & "^" & Trim(RS("商品コード")) & "^" & Trim(RS("色")) & "^" & Trim(RS("規格"))

	'---- 登録日、ウィッシュリストから削除
		wListHTML = wListHTML & "      <a href='WishListToCartDelete.asp?DeleteFl=Y&Item=" & wItem & "'>ウィッシュリストから削除</a><br>"

	'----- 数量, カートボタン
		if (IsNull(RS("取扱中止日")) = false) OR (IsNull(RS("完売日")) = false) OR (RS("B品フラグ") = "Y" AND RS("B品引当可能数量") <= 0) OR (IsNull(RS("廃番日")) = false AND RS("引当可能数量") <= 0) then
			wListHTML = wListHTML & "      <input type='hidden' name='qt' value='0'>" & vbNewLine
			wListHTML = wListHTML & "      <img src='images/icon_sold.gif'>" & vbNewLine
		else
			wListHTML = wListHTML & "      <input type='text' name='qt' size='2' value='1'>" & vbNewLine
			wListHTML = wListHTML & "      <input type='image' src='images/btn_cart.png' class='cartBtn'>" & vbNewLine
		end if

		wListHTML = wListHTML & "      <input type='hidden' name='Item' value='" & wItem & "'>" & vbNewLine
		wListHTML = wListHTML & "      <input type='hidden' name='Kubun' value='Cart'>" & vbNewLine

		wListHTML = wListHTML & "      <input type='hidden' name='DeleteFl' value='N'>" & vbNewLine

		wListHTML = wListHTML & "    </td>" & vbNewLine
		wListHTML = wListHTML & "    </form>" & vbNewLine
		wListHTML = wListHTML & "  </tr>" & vbNewLine

	'---- 区切り線
		wListHTML = wListHTML & "  <tr>" & vbNewLine
		wListHTML = wListHTML & "    <td height='5' colspan='3'><hr size='1'></td>" & vbNewLine
		wListHTML = wListHTML & "  </tr>" & vbNewLine

	'----
		RS.MoveNext
	Loop

	wListHTML = wListHTML & "</table>" & vbNewLine

	RS.Close
end if

'---- カートの中身作成
wCartHTML = fCreateCartHtml()

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

' 2012/08/14 GV #1419 Add Start
'========================================================================
'
'	Function	顧客情報の取り出し
'
'========================================================================
Function get_customer()

Dim vRS
Dim vSQL

'---- 顧客情報取り出し
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      a.顧客番号 "
vSQL = vSQL & "FROM "
vSQL = vSQL & "      Web顧客     a WITH (NOLOCK) "
vSQL = vSQL & "        LEFT JOIN 顧客プロフィール c WITH (NOLOCK) "
vSQL = vSQL & "          ON a.顧客番号 = c.顧客番号 "
vSQL = vSQL & "    , Web顧客住所 b WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        a.顧客番号 = b.顧客番号 "
vSQL = vSQL & "    AND b.住所連番 = 1 "
vSQL = vSQL & "    AND a.Web不掲載フラグ <> 'Y'"
vSQL = vSQL & "    AND a.顧客番号 = " & UserID

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, Connection, adOpenStatic, adLockOptimistic

Set get_customer = vRS

End Function
' 2012/08/14 GV #1419 Add End

'========================================================================
%>
<!DOCTYPE html>
<html>
<head>
<meta charset="Shift_JIS">
<title>ウィッシュリスト｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/WishList.css" type="text/css">
<link rel="stylesheet" href="style/ask.css?20140401a" type="text/css">
<script type="text/javascript">
//
//  	Function:	order_onClick
//
function order_onClick(pForm){
	if (pForm.qt.value == ""){
		pForm.qt.value = 0;
	}else{
		if (numericCheck(pForm.qt.value) == false){
			pForm.qt.value = 0;
		}
	}
	if (pForm.qt.value == 0){
		alert("数量を入力してからカートボタンを押してください。");
		return false;
	}

	if (confirm("ウィッシュリストを保持しますか？") == true){
		pForm.DeleteFl.value = "N";
	}else{
		pForm.DeleteFl.value = "Y";
	}
}

</script>
</head>
<body>
<!--#include file="../Navi/Navitop.inc"-->
<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="ここから本文です"></a></span>

<!-- コンテンツstart -->
<div id="globalContents">

  <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
    <p class="home"><a href="../"><img src="../images/icon_home.gif" alt="HOME"></a></p>
    <ul id="path">
      <li class="now">ウィッシュリスト</li>
    </ul>
  </div></div></div>

  <h1 class="title">ウィッシュリスト</h1>

  <div id="wishlist_container">

    <div id="wishlist">

<% if wMSG <> "" then %>
      <p class="error"><%=wMSG%></p>
<% end if %>

<%=wListHTML%>

    </div>

    <div id="detail_side">

<%=wCartHTML%>

    </div>

  </div>

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