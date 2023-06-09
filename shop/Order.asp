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
<!--#include file="../3rdParty/EAgency.inc"-->
<%
'========================================================================
'
'	ショッピングカート
'
'2012/06/14 ok デザイン変更のため旧版を元に新規作成
'2012/07/02 ok ログイン済、カート空かつ保存カートが存在する場合「保存されたカート一覧へ」ボタンを表示するよう修正
'2013/05/20 GV #1505 さぶみっと！レコメンド機能
'2013/08/07 if-web 旧レコメンド（チームラボ）をコメントアウト
'2013/10/21 GV # 大型商品の表示
'
'========================================================================

On Error Resume Next

Dim userID
Dim userName
Dim msg

Dim wSalesTaxRate
Dim wPrice
Dim wNoData
Dim wOrderProductHTML
Dim wSavedCartFl

Dim wRecommendMakerCd
Dim wRecommendProductCd
Dim wRecommendHTML

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

Dim w_error_msg
Dim wErrDesc

Dim wOrderProductId		'2013/05/20 GV #1505 add

'2013/10/21 GV # add start
Dim strLargeItem
Dim wLargeItemFl
Dim wNonLargeItemFl
'2013/10/21 GV # add end

'========================================================================

Response.Expires = -1			' Do not cache

'---- UserID 取り出し
userID = Session("userID")
userName = Session("userName")

'---- Get input data
msg = Session.contents("msg")
Session("msg") = ""

wOrderProductId = ""			'2013/05/20 GV #1505 add

'2013/10/21 GV # add start
strLargeItem = ""
wLargeItemFl = "N"
wNonLargeItemFl = "N"
'2013/10/21 GV # add end

'---- Execute main
call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録
if Err.Description <> "" then
	wErrDesc = "Order.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if

call close_db()

if Err.Description <> "" then
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

'========================================================================
'
'	Function	Connect database
'
'========================================================================
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
Function main()

Dim v_product_nm
Dim vTotalAm
Dim vFreightAm
Dim vFreightForwarder
Dim vSoukoCnt
Dim vKoguchi
'2011/04/14 GV Add Start
Dim vProdTermFl
Dim vInventoryCd
Dim vInventoryImage
'2011/04/14 GV Add End

'---- 保存されたカート情報があるかどうかチェック
wSavedCartFl = "N"
If userID <> "" Then
	Call CheckSavedCart()
End If

'---- 消費税率取出し
Call getCntlMst("共通","消費税率","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)			'消費税率
wSalesTaxRate = Clng(wItemNum1)

vTotalAm = 0
wHTML = ""

'----仮受注データ取り出し
wSQL = ""
wSQL = wSQL & "SELECT a.受注明細番号"
wSQL = wSQL & "     , a.メーカーコード"
wSQL = wSQL & "     , a.商品コード"
wSQL = wSQL & "     , a.色"
wSQL = wSQL & "     , a.規格"
wSQL = wSQL & "     , a.メーカー名"
wSQL = wSQL & "     , a.商品名"
wSQL = wSQL & "     , a.受注数量"
wSQL = wSQL & "     , a.受注単価"
wSQL = wSQL & "     , a.受注金額"
wSQL = wSQL & "     , b.ASK商品フラグ"
wSQL = wSQL & "     , b.取扱中止日"
wSQL = wSQL & "     , b.廃番日"
wSQL = wSQL & "     , b.完売日"
wSQL = wSQL & "     , b.希少数量"
wSQL = wSQL & "     , b.セット商品フラグ"
wSQL = wSQL & "     , b.メーカー直送取寄区分"
wSQL = wSQL & "     , b.空輸禁止フラグ "					'2013/10/21 GV # add
wSQL = wSQL & "     , b.代引不可フラグ "					'2013/10/21 GV # add
wSQL = wSQL & "     , b.送料区分 "							'2013/10/21 GV # add
wSQL = wSQL & "     , b.Web納期非表示フラグ"
wSQL = wSQL & "     , b.入荷予定未定フラグ"
wSQL = wSQL & "     , b.B品フラグ"
wSQL = wSQL & "     , b.個数限定数量"
wSQL = wSQL & "     , b.個数限定受注済数量"
wSQL = wSQL & "     , c.引当可能数量"
wSQL = wSQL & "     , c.引当可能入荷予定日"
wSQL = wSQL & "     , c.B品引当可能数量"
wSQL = wSQL & "     , c.商品ID"							'2013/05/20 GV #1505 add
wSQL = wSQL & "  FROM 仮受注明細 a WITH (NOLOCK)"
wSQL = wSQL & "     ,Web商品 b WITH (NOLOCK)"
wSQL = wSQL & "     ,Web色規格別在庫 c WITH (NOLOCK)"
wSQL = wSQL & " WHERE SessionID = '" & gSessionID & "'"
wSQL = wSQL & "   AND b.メーカーコード = a.メーカーコード"
wSQL = wSQL & "   AND b.商品コード = a.商品コード"
wSQL = wSQL & "   AND c.メーカーコード = a.メーカーコード"
wSQL = wSQL & "   AND c.商品コード = a.商品コード"
wSQL = wSQL & "   AND c.色 = a.色"
wSQL = wSQL & "   AND c.規格 = a.規格"
wSQL = wSQL & " ORDER BY 受注明細番号"

'@@@@response.write(wSQL)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

wNoData = False

'---- 明細HTML作成
If RS.EOF = True Then
	wHTML = wHTML & "      <tr>" & vbNewLine
	wHTML = wHTML & "        <td align='center'>" & vbNewLine
	wHTML = wHTML & "          <b>カートに商品がありません。</b><br>" & vbNewLine
	wHTML = wHTML & "          カートに商品が入らない場合は、ブラウザーのCookieが有効になっていることを確認してください。設定方法については<a href='../guide/site.asp#kankyo'>こちら。</a>" & vbNewLine
	wHTML = wHTML & "        </td>" & vbNewLine
	wHTML = wHTML & "      </tr>" & vbNewLine
	wOrderProductHTML = wHTML
	wNoData = True
	Exit Function
End If

'----- 見出し
wHTML = wHTML & "      <tr>" & vbNewLine
wHTML = wHTML & "        <th class='maker'>メーカー</th>" & vbNewLine
wHTML = wHTML & "        <th class='name'>商品名</th>" & vbNewLine
wHTML = wHTML & "        <th class='stock'>在庫</th>" & vbNewLine
wHTML = wHTML & "        <th class='price'>単価</th>" & vbNewLine
wHTML = wHTML & "        <th class='number'>数量</th>" & vbNewLine
wHTML = wHTML & "        <th class='amount'>金額(税込)</th>" & vbNewLine
wHTML = wHTML & "        <th></th>" & vbNewLine
wHTML = wHTML & "      </tr>" & vbNewLine

Do Until RS.EOF = True

	'---- 2013.10.21 GV # add start
	'---- 大型商品の表示
	strLargeItem = ""
	If (((IsNull(RS("空輸禁止フラグ")) = False) And (RS("空輸禁止フラグ") = "Y")) And _
		((IsNull(RS("代引不可フラグ")) = False) And (RS("代引不可フラグ") = "Y")) And _
		(RS("送料区分") = "重量商品")) Then
		strLargeItem = strLargeItem & "<br><span style='color:red;'>大型商品</span>"
		wLargeItemFl = "Y"
	Else
		wNonLargeItemFl = "Y"
	End If
	'---- 2013.10.21 GV # add end

	'------------- メーカー、商品名
	v_product_nm = RS("商品名")
	If Trim(RS("色")) <> "" Then
		v_product_nm = v_product_nm & "/" & RS("色")
	End If
	If Trim(RS("規格")) <> "" Then
		v_product_nm = v_product_nm & "/" & RS("規格")
	End If
	wHTML = wHTML & "      <tr>" & vbNewLine
	wHTML = wHTML & "        <td>" & RS("メーカー名") & "</td>" & vbNewLine
'	wHTML = wHTML & "        <td><a href='ProductDetail.asp?Item=" & RS("メーカーコード") & "^" & Server.URLEncode(RS("商品コード")) & "^" & RS("色") & "^" & RS("規格") & "' alt=''>" & v_product_nm & "</a></td>" & vbNewLine
	wHTML = wHTML & "        <td><a href='ProductDetail.asp?Item=" & RS("メーカーコード") & "^" & Server.URLEncode(RS("商品コード")) & "^" & RS("色") & "^" & RS("規格") & "' alt=''>" & v_product_nm & "</a>" & strLargeItem & "</td>" & vbNewLine

	'------------- 在庫
	'---- 終了チェック
	vProdTermFl = "N"
	If IsNull(RS("取扱中止日")) = False Then	'取扱中止
		vProdTermFl = "Y"
	End If
	If IsNull(RS("廃番日")) = False And RS("引当可能数量") <= 0 Then	'廃番で在庫無し
		vProdTermFl = "Y"
	End If
	If IsNull(RS("完売日")) = False Then		'完売商品
		vProdTermFl = "Y"
	End If

	'---- 在庫状況
	vInventoryCd = GetInventoryStatus(RS("メーカーコード"), RS("商品コード"), RS("色"), RS("規格"), RS("引当可能数量"), RS("希少数量"), RS("セット商品フラグ"), RS("メーカー直送取寄区分"), RS("引当可能入荷予定日"), vProdTermFl)

	'---- 在庫状況、色を最終セット
	Call GetInventoryStatus2(RS("引当可能数量"), RS("Web納期非表示フラグ"), RS("入荷予定未定フラグ"), RS("廃番日"), RS("B品フラグ"), RS("B品引当可能数量"), RS("個数限定数量"), RS("個数限定受注済数量"), vProdTermFl, vInventoryCd, vInventoryImage)

	'----- 在庫状況表示
	If IsNull(RS("取扱中止日")) = False Or _
	   IsNull(RS("完売日")) = False Or _
	   (RS("B品フラグ") = "Y" And RS("B品引当可能数量") <= 0) Or _
	   (IsNull(RS("廃番日")) = False And RS("引当可能数量") <= 0) Then
		wHTML = wHTML & "        <td><span class='stock'>&nbsp</span></td>" & vbNewLine
	Else
		'---- 完売御礼でない場合のみ、在庫状況を表示
		wHTML = wHTML & "        <td><span class='stock'><img src='images/" & vInventoryImage & "' alt='' > " & vInventoryCd & "</span></td>" & vbNewLine
	End If

	'------------- 単価
	wPrice = calcPrice(RS("受注単価"), wSalesTaxRate)
	vTotalAm = vTotalAm + (wPrice * RS("受注数量"))
	wHTML = wHTML & "        <td>" & FormatNumber(wPrice,0) & "円</td>" & vbNewLine

	'------------- 数量
	wHTML = wHTML & "        <td>" & vbNewLine
	wHTML = wHTML & "          <input type='text' name='qt" & RS("受注明細番号") & "' id='order_form_qt1' value='" & RS("受注数量") & "' size=4 onBlur='qt_onBlur(this);'>" & vbNewLine
	wHTML = wHTML & "          <input type='hidden' name='oldqt" & RS("受注明細番号") & "' value='" & RS("受注数量") & "'>" & vbNewLine
	wHTML = wHTML & "        </td>" & vbNewLine

	'------------- 金額
	wHTML = wHTML & "        <td>" & FormatNumber(wPrice*RS("受注数量"),0) & "円</td>" & vbNewLine

	'------------- 削除ボタン
	wHTML = wHTML & "        <td>" & vbNewLine
	wHTML = wHTML & "          <ul>" & vbNewLine
	wHTML = wHTML & "            <li><a href='JavaScript:delete_onClick(" & RS("受注明細番号") & ");'><img src='images/btn_delete.png' alt='削除' class='opover' ></a></li>" & vbNewLine
	'------------- 後で買うボタン
	If userID <> "" Then
		wHTML = wHTML & "        <li><a href='WishListAdd.asp?OrderDetailNo=" & RS("受注明細番号") & "&Item=" & Server.URLEncode(RS("メーカーコード") & "^" & RS("商品コード") & "^" & RS("色") & "^" & RS("規格")) & "' class='link'><img src='images/btn_later.png' alt='後で買う' class='opover' ></a></li>" & vbNewLine
	End If
	wHTML = wHTML & "          </ul>" & vbNewLine
	wHTML = wHTML & "        </td>" & vbNewLine
	wHTML = wHTML & "      </tr>" & vbNewLine

	'---- 最後にカートに入れた商品のレコメンド表示用
	wRecommendMakerCd = RS("メーカーコード")
	wRecommendProductCd = RS("商品コード")

	'2013/05/20 GV #1505 add start
	'さぶみっと！レコメンド用JSに渡す商品ID
	wOrderProductId = wOrderProductId & "'" & RS("商品ID") & "',"
	'2013/05/20 GV #1505 add end

	RS.MoveNext

Loop

'---- 送料
Call fCalcShipping(gSessionID, "一括", vFreightAm, vFreightForwarder, vSoukoCnt, vKoguchi)		'2011/04/14 hn mod
wPrice = Fix(vFreightAm * (100 + wSalesTaxRate) / 100)

'----商品合計金額，再計算ボタン
wHTML = wHTML & "      <tr>" & vbNewLine
wHTML = wHTML & "        <td colspan='6'>" & vbNewLine
wHTML = wHTML & "          <dl class='total'>" & vbNewLine
wHTML = wHTML & "            <dt>商品合計（税込）</dt><dd>" & FormatNumber(vTotalAm,0) & "円</dd>" & vbNewLine
wHTML = wHTML & "            <dt>送料見積（税込）</dt><dd>" & FormatNumber(wPrice,0) & "円</dd>" & vbNewLine
wHTML = wHTML & "          </dl>" & vbNewLine
wHTML = wHTML & "        </td>" & vbNewLine
wHTML = wHTML & "        <td><a href='JavaScript:calc_onClick();'><img src='images/btn_calculate.png' alt='再計算' class='opover' ></a></td>" & vbNewLine
wHTML = wHTML & "      </tr>" & vbNewLine

wOrderProductHTML = wHTML

'---- レコメンドデータ作成
Call CreateRecommendInfo()

RS.close

End Function

'========================================================================
'
'	Function	保存されたカート情報があるかどうかチェック
'
'========================================================================
'
Function CheckSavedCart()

Dim Rsv

'----保存カートデータ取り出し
wSQL = ""
wSQL = wSQL & "SELECT a.顧客番号"
wSQL = wSQL & "  FROM 保存カート a WITH (NOLOCK)"
wSQL = wSQL & " WHERE 顧客番号 = " & UserID

'@@@@response.write(wSQL)

Set Rsv = Server.CreateObject("ADODB.Recordset")
Rsv.Open wSQL, Connection, adOpenStatic

if RSv.EOF = false then
	wSavedCartFl = "Y"
end if

Rsv.Close

End function


'========================================================================
'
'	Function	レコメンド商品取得  '2010/04/02 an add
'
'========================================================================
'
Function CreateRecommendInfo()

'2013/08/07 if-web del s
'Dim RSv
'
'---- レコメンド商品取得(類似度が大きい5商品)
'wSQL = ""
'
'wSQL = wSQL & "SELECT DISTINCT TOP 5"
'wSQL = wSQL & "       a.メーカーコード"
'wSQL = wSQL & "     , a.商品コード"
'wSQL = wSQL & "     , a.商品名"
'wSQL = wSQL & "     , a.商品画像ファイル名_小"
'wSQL = wSQL & "     , CASE"
'wSQL = wSQL & "         WHEN (a.個数限定数量 > a.個数限定受注済数量 AND a.個数限定数量 > 0) THEN a.個数限定単価"
'wSQL = wSQL & "         ELSE a.販売単価"
'wSQL = wSQL & "       END AS 販売単価"
'wSQL = wSQL & "     , a.ASK商品フラグ"
'wSQL = wSQL & "     , a.カテゴリーコード"
'wSQL = wSQL & "     , b.メーカー名"
'wSQL = wSQL & "     , e.類似度"
'wSQL = wSQL & "  FROM Web商品 a WITH (NOLOCK)"
'wSQL = wSQL & "     , メーカー b WITH (NOLOCK)"
'wSQL = wSQL & "     , Web色規格別在庫 d WITH (NOLOCK)"
'wSQL = wSQL & "     , レコメンド結果購買 e WITH (NOLOCK)"
'wSQL = wSQL & " WHERE a.メーカーコード = e.レコメンドメーカーコード"
'wSQL = wSQL & "   AND a.商品コード = e.レコメンド商品コード"
'wSQL = wSQL & "   AND d.メーカーコード = a.メーカーコード"
'wSQL = wSQL & "   AND d.商品コード = a.商品コード"
'wSQL = wSQL & "   AND b.メーカーコード = a.メーカーコード"
'wSQL = wSQL & "   AND a.Web商品フラグ = 'Y'"
'wSQL = wSQL & "   AND a.取扱中止日 IS NULL"
'wSQL = wSQL & "   AND ((a.廃番日 IS NULL) OR (a.廃番日 IS NOT NULL AND d.引当可能数量 > 0))"
'wSQL = wSQL & "   AND e.メーカーコード = '" & wRecommendMakerCd & "'"
'wSQL = wSQL & "   AND e.商品コード = '" & wRecommendProductCd & "'"
'wSQL = wSQL & " ORDER BY"
'wSQL = wSQL & "       e.類似度 DESC"
'wSQL = wSQL & "     , a.カテゴリーコード"
'
'@@@@response.write(wSQL)
'
'Set RSv = Server.CreateObject("ADODB.Recordset")
'RSv.Open wSQL, Connection, adOpenStatic
'
'wHTML = ""
'
'if RSv.EOF = false then
'
'	wHTML = ""
'	wHTML = wHTML & "  <h2 class='detail_title'>このアイテムを買った人はこんなアイテムも買っています。</h2>" & vbNewLine
'	wHTML = wHTML & "    <ul class='relation'>" & vbNewLine
'
'	Do Until RSv.EOF = True
'
'	wPrice = calcPrice(RSv("販売単価"), wSalesTaxRate)
'
'		wHTML = wHTML & "      <li>" & vbNewLine
'		wHTML = wHTML & "        <p><a href='ProductDetail.asp?Item=" & RSv("メーカーコード") & "^" & RSv("商品コード") & "'><img src='"
'		wHTML = wHTML & "prod_img/" & RSv("商品画像ファイル名_小") & "' alt='" & RSv("メーカー名") & " / " & RSv("商品名") & "' class='opover'><span>"
'		wHTML = wHTML & RSv("メーカー名") & "</span><span>" & RSv("商品名") & "</span></a></p>" & vbNewLine
'		If RSv("ASK商品フラグ") <> "Y" Then
'			wHTML = wHTML & "        <p>" & FormatNumber(wPrice,0) & "円(税込)</p>" & vbNewLine
'		Else
'			wHTML = wHTML & "        <p><a class='tip'>ASK<span>"& FormatNumber(wPrice,0) & "円(税込)</span></p>" & vbNewLine
'		End If
'		
'		wHTML = wHTML & "      </li>" & vbNewLine
'
'		RSv.MoveNext
'
'	Loop
'
'	wHTML = wHTML & "    </ul>"
'
'End if
'
'RSv.Close

'wHTML = wHTML & fEAgency_CreateRecommendCartJS(wOrderProductId)	'2013/05/20 GV #1505 add
'2013/08/07 if-web del e

wHTML = fEAgency_CreateRecommendCartJS(wOrderProductId)	'2013/08/07 if-web add

wRecommendHTML = wHTML

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
<title>ショッピングカート｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/shop.css" type="text/css">
<link rel="stylesheet" href="style/StyleOrder.css?20120703" type="text/css">
<link rel="stylesheet" href="style/ask.css?20140401a" type="text/css">

<script type="text/javascript">

var g_change_fl = false;

//
//	数量変更時
//
function qt_onBlur(p_formItem){

var v_itemName;

	v_itemName = "old" + p_formItem.name;
	if (p_formItem.value != document.f_order_list.elements[v_itemName].value){
		g_change_fl = true;
	}
}

//
//	受注明細行Delete
//
function delete_onClick(p_detail_no){

	document.f_order_list.detail_no.value = p_detail_no;
	document.f_order_list.action = "OrderChange.asp";
	document.f_order_list.submit();
}

//
//	再計算
//
function calc_onClick(){

	g_change_fl = false;
	document.f_order_list.detail_no.value = "all";
	document.f_order_list.action = "OrderChange.asp";
	document.f_order_list.submit();
}

//
//	オーダーSubmit
//
function order_onSubmit(){

//	数量変更があるかどうかチェック
	if (g_change_fl == true){
		alert("数量が変更されています。　再計算ボタンを押してください。");
		return;
	}

	window.location = g_HTTPS + "shop/LoginCheck.asp?called_from=order";
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
      <li class="now">ショッピングカート</li>
    </ul>
  </div></div></div>

  <p class="error"><% = msg %></p>

  <h1 class="title">ショッピングカート</h1>
  <ol id="step">
    <li><img src="images/step01_now.gif" alt="1.ショッピングカート" width="170" height="50"></li>
    <li><img src="images/step02.gif" alt="2.お届け先、お支払方法の選択" width="170" height="50"></li>
    <li><img src="images/step03.gif" alt="3.ご注文内容の確認" width="170" height="50"></li>
    <li><img src="images/step04.gif" alt="4.ご注文完了" width="170" height="50"></li>
  </ol>

  <h2 class="cart_title">カート内容</h2>
  <form method='post' name='f_order_list' action='' onsubmit='calc_onClick();'>
    <table id="cart">
<% = wOrderProductHTML %>
    </table>
    <input type='hidden' name='detail_no' value=''>
  </form>

  <ul id="attention">
    <li>この画面で数量の変更ができます。 数量を変更後「再計算」ボタンを押すと更新されます。</li>
    <li>注文商品を取り消す場合は「削除」ボタンを押してください。</li>
    <li>注文商品を追加する場合は下の「買い物を続ける」ボタンをクリックして商品の画面に戻ってください。</li>
    <li>配送方法、配送先により送料は変わることがあります。</li>
    <li>この画面の金額は商品毎の単価の確認となり最終的な合計金額ではありません。</li>
    <li>最終的なお支払い金額は別途ご案内するご注文確認書をご参照ください。</li>
  </ul>

<% If wNoData = False Then %>
  <div id="btn_box">
    <ul class="btn">
      <li><a href="javascript:history.back();"><img src="images/btn_continue.png" alt="買い物を続ける" class="opover"></a></li>
      <li class="last"><a href="javascript:order_onSubmit();"><img src="images/btn_order.png" alt="ご注文手続きへ" class="opover"></a></li>
    </ul>
  </div>
<% End If %>

<% If (wNoData = False Or wSavedCartFl = "Y") And userID <> "" Then %>
  <div class="btn_box">
<% If wNoData = False And wSavedCartFl = "Y" Then %>
    <ul class="btn">
      <li><a href="SaveCart.asp"><img src="images/btn_cartsave.png" alt="このカート内容を保存" class="opover"></a></li>
      <li class="last"><a href="SaveCartList.asp"><img src="images/btn_cartlist.png" alt="保存されたカート一覧へ" class="opover"></a></li>
    </ul>
<% Elseif  wNoData = False Then %>
    <div class="btn"><a href="SaveCart.asp"><img src="images/btn_cartsave.png" alt="このカート内容を保存" class="opover"></a></div>
<% Else %>
    <div class="btn"><a href="SaveCartList.asp"><img src="images/btn_cartlist.png" alt="保存されたカート一覧へ" class="opover"></a></div>
<% End If %>
  </div>
<% End If %>

  <ul class="info left">
    <li><a href="../guide/change.asp">ご注文商品のキャンセル・返品について</a></li>
    <li><a href="../guide/nouki.asp">商品の納期についてはこちら</a></li>
  </ul>
  <ul class="info right">
    <li class="no"><a href="../shopEng/Order.asp">English</a></li>
  </ul>

  <!-- レコメンド商品 start -->
<% = wRecommendHTML %>
  <!-- レコメンド商品 end -->
  <!--/#contents --></div>
  <div id="globalSide">
    <!--#include file="../Navi/NaviSide.inc"-->
  <!--/#globalSide --></div>
<!--/#main --></div>
<!--#include file="../Navi/Navibottom.inc"-->
<div class="tooltip"><p>ASK</p></div>
<!--#include file="../Navi/NaviScript.inc"-->
<script type="text/javascript" src="jslib/ask.js?20140401a"></script>
</body>
</html>