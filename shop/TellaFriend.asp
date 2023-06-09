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
'	友達にすすめるページ
'
'更新履歴
'2009/04/30 エラー時にerror.aspへ移動
'2011/08/01 an #1087 Error.aspログ出力対応
'2014/03/19 GV 消費税増税に伴う2重表示対応
'
'========================================================================

On Error Resume Next

Dim Item
Dim ItemCnt
Dim ItemList()
Dim MakerCd
Dim ProductCd
Dim MakerNm
Dim ProductNm
Dim FromName

Dim wPrice
Dim wSalesTaxRate
Dim wMailTrailer

Dim wProductHTML
Dim wMessage1
Dim wMessage1HTML

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

'---- Get input data
Item = ReplaceInput(Trim(Request("Item")))

if Item <> "" then
	ItemCnt = cf_unstring(Item, ItemList, "^")
	MakerCd = ItemList(0)
	ProductCd = ItemList(1)
end if

'---- Execute main
call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "TellaFriend.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
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
'	Function	main
'
'========================================================================
Function main()

'---- 消費税率取出し
call getCntlMst("共通","消費税率","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)			'消費税率
wSalesTaxRate = Clng(wItemNum1)

'---- メールトレーラ取り出し
call getCntlMst("Web","Email","一般トレーラ", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)
wMailTrailer = wItemChar1
wMailTrailer = Replace(wMailTrailer, vbNewLine, "<br>")

'---- 商品情報取り出し
wSQL = ""
wSQL = wSQL & "SELECT a.商品コード"
wSQL = wSQL & "     , a.商品名"
wSQL = wSQL & "     , CASE"
wSQL = wSQL & "         WHEN (a.個数限定数量 > a.個数限定受注済数量 AND a.個数限定数量 > 0) THEN a.個数限定単価"
wSQL = wSQL & "         ELSE a.販売単価"
wSQL = wSQL & "       END AS 販売単価"
wSQL = wSQL & "     , a.B品単価"
wSQL = wSQL & "     , a.B品フラグ"
wSQL = wSQL & "     , a.商品概略Web"
wSQL = wSQL & "     , a.商品画像ファイル名_小"
wSQL = wSQL & "     , a.ASK商品フラグ"
wSQL = wSQL & "     , a.メーカーコード"
wSQL = wSQL & "     , b.メーカー名"
wSQL = wSQL & "  FROM Web商品 a WITH (NOLOCK)"
wSQL = wSQL & "     , メーカー b WITH (NOLOCK)"
wSQL = wSQL & " WHERE b.メーカーコード = a.メーカーコード"
wSQL = wSQL & "   AND a.Web商品フラグ = 'Y'"
wSQL = wSQL & "   AND a.メーカーコード = '" & MakerCd & "'"
wSQL = wSQL & "   AND a.商品コード = '" & ProductCd & "'"
		
'@@@@@@response.write(wSQL)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

'---- Move data to work area
if RS.EOF = true then
	exit function
else
	MakerNm = RS("メーカー名")
	ProductNm = RS("商品名")
end if

'---- 商品情報編集
wHTML = ""
wHTML = wHTML & "<dl class='form_product'>" & vbNewLine
wHTML = wHTML & "  <dt>" & vbNewLine
wHTML = wHTML & "      <a href='ProductDetail.asp?Item=" & Server.URLEncode(Item) & "'><img src='../shop/prod_img/" & RS("商品画像ファイル名_小") & "'></a>" & vbNewLine
wHTML = wHTML & "  </dt>" & vbNewLine
wHTML = wHTML & "  <dd>" & RS("メーカー名") & "</dd>" & vbNewLine
wHTML = wHTML & "  <dd>" & "<a href='ProductDetail.asp?Item=" & Server.URLEncode(Item) & "'>" & RS("商品名") & "</a></dd>" & vbNewLine
wHTML = wHTML & "  <dd>" & RS("商品概略Web") & "</dd>" & vbNewLine
wHTML = wHTML & "  <dd>" & vbNewLine

wPrice = calcPrice(RS("販売単価"), wSalesTaxRate)

if RS("ASK商品フラグ") = "Y" then
	wHTML = wHTML & "衝撃特価：<a class='tip'>ASK<span class='exc-tax'>" & FormatNumber(RS("販売単価"),0) & "円(税抜)</span>"
	wHTML = wHTML & "<span class='inc-tax'>(税込&nbsp;" & FormatNumber(wPrice,0) & "円)</span></a>" & vbNewLine
else
	if RS("B品フラグ") = "Y" then
		wHTML = wHTML & "衝撃特価：<del>" & FormatNumber(wPrice,0) & "円(税込)</del><br>" & vbNewLine
		wPrice = calcPrice(RS("B品単価"), wSalesTaxRate)
'2014/03/19 GV mod start ---->
		wHTML = wHTML & "<strong>わけあり品特価：" & FormatNumber(RS("B品単価"),0) & "円(税抜)</strong>" & vbNewLine
		wHTML = wHTML & "(税込&nbsp;" & FormatNumber(wPrice,0) & "円)" & vbNewLine
'2014/03/19 GV mod end   <----
	else
'2014/03/19 GV mod start ---->
'		wHTML = wHTML & "衝撃特価：" & FormatNumber(wPrice,0) & "円(税込)" & vbNewLine
		wHTML = wHTML & "衝撃特価：" & FormatNumber(RS("販売単価"),0) & "円(税抜)" & vbNewLine
		wHTML = wHTML & "(税込&nbsp;" & FormatNumber(wPrice,0) & "円)" & vbNewLine
'2014/03/19 GV mod end   <----
	end if
end if

wHTML = wHTML & "  </dd>" & vbNewLine
wHTML = wHTML & "</dl>" & vbNewLine

wProductHTML = wHTML

RS.close

'---- メッセージヘッダ編集
wHTML = ""
'wHTML = wHTML & FromName & "　様より" & vbNewLine
wHTML = wHTML & MakerNm & "　" & ProductNm & "をおすすめされました。" & vbNewLine
wHTML = wHTML & "ぜひ一度、ご覧ください。" & vbNewLine
wHTML = wHTML & "http://www.soundhouse.co.jp/shop/ProductDetail.asp?Item=" & Server.URLEncode(Item)
wMessage1 = wHTML
wMessage1HTML = Replace(wHTML, vbNewLine, "<br>")

end function

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
<html lang="ja">
<head>
<meta charset="Shift_JIS">
<title>お友達にすすめる｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/ask.css?20140401a" type="text/css">

<script type="text/javascript">

//
// ====== 	Function:	mail on submit
//
function mail_onSubmit(pForm){

	if (pForm.ToAddr.value == ""){
		alert("\n宛先を入力してください。");
		pForm.ToAddr.focus();
		return false;
 	}

		return true;
}

</script>

</head>

<body>

<!--#include file="../Navi/NaviTop.inc"-->

<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="ここから本文です"></a></span>
  
  <!-- コンテンツstart -->
  <div id="globalContents">
    <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
      <p class="home"><a href="<%=g_HTTP%>"><img src="<%=g_RelLink%>images/icon_home.gif" alt="HOME"></a></p>
      <ul id="path">
        <li class="now">お友達にすすめる</li>
      </ul>
    </div></div></div>

    <h1 class="title">お友達にすすめる</h1>

<%=wProductHTML%>

<form name="fMail" type="post" action="TellaFriendSend.asp" onSubmit="return mail_onSubmit(this);">
<table class="form">
  <tr>
    <th>送信者名</th>
    <td><input name="FromName" type="text" size="60"></td>
  </tr>
  <tr>
    <th>宛先</th>
    <td><input name="ToAddr" type="text" size="60"><div>宛先のメールアドレスをご入力ください。</div></td>
  </tr>
  <tr>
    <th>メッセージ</th>
    <td><p><%=wMessage1HTML%></p><textarea name="Message" cols="70" rows="20"></textarea></td>
  </tr>
</table>
<p>ご案内いたしました内容につきまして不明な点がございましたら、弊社営業までご連絡いただきますようお願いいたします。</p>
<p>
	-------------------------------------------------------<br>
	株式会社サウンドハウス<br>
	HP： http://www.soundhouse.co.jp/<br>
	Email ： shop@soundhouse.co.jp<br>
	TEL ： 0476-89-1111<br>
	FAX ： 0476-89-2222<br>
	（月～金：10-19時、土：12-17時、日曜・祝祭日を除く）<br>
	-------------------------------------------------------
</p>
<p class="btnBox"><input type="submit" value="送信" class="opover"></p>
<input type="hidden" name="Item" value="<%=Item%>">
<input type="hidden" name="Message1" value="<%=wMessage1%>">
</form>

</div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<div class="tooltip"><p>ASK</p></div>
<!--#include file="../Navi/NaviScript.inc"-->
<script type="text/javascript" src="jslib/ask.js?20140401a"></script>
</body>
</html>