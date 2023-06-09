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
<!--#include file="../common/HttpsSecurity.inc"-->

<%
'========================================================================
'
'	修理･商品サポート確認ページ
'
'更新履歴
'2010/10/04 an 新規作成
'2011/03/02 hn SetSecureKeyの位置変更
'2011/08/01 an #1087 Error.aspログ出力対応
'2012/06/29 if-web リニューアルレイアウト調整
'
'========================================================================

On Error Resume Next

Dim userID

Dim MakerName
Dim ProductName
Dim Warranty
Dim SerialNo
Dim WhenPurchased
Dim Comment

Dim wCustomerName
Dim wZip
Dim wPrefecture
Dim wAddress
Dim wTelephone
Dim wFax
Dim wEmail

Dim Skey

Dim Connection
Dim RS

Dim wSQL
Dim wHTML
Dim wMSG
Dim wNoData
Dim wErrDesc   '2011/08/01 an add

'========================================================================

Response.buffer = true

'---- セキュリティーキーチェック
if Session("SKey") <> ReplaceInput(Request("SKey")) then
	Response.redirect "SupportInquiry.asp"
end if

'---- get input data
MakerName = ReplaceInput(Left(Request("MakerName"),26))
ProductName = ReplaceInput(Left(Request("ProductName"),51))
Warranty = ReplaceInput(Left(Request("Warranty"),3))
SerialNo = ReplaceInput(Left(Request("SerialNo"),41))
WhenPurchased = ReplaceInput(Left(Request("WhenPurchased"),10))
Comment = ReplaceInput(Left(Request("Comment"),501))

'---- Execute main
call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "SupportInquiryConfirm.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

call close_db()

if Err.Description <> "" OR wNoData = "Y" then
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

'---- 入力データにエラーがある場合は入力画面に戻る
if wMSG <> "" then
	Session("msg") = wMSG
	Server.Transfer("SupportInquiry.asp")
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
'				userIDがCookieにあれば会員情報を表示
'
'========================================================================
Function main()

'セキュリティーキーをセット
Skey = SetSecureKey()

wNoData = ""

'---- 顧客番号取り出し
userID = Session("userID")

if userID = "" then
	wNoData = "Y"
	exit function
end if

'--------- select customer
wSQL = ""
wSQL = wSQL & "SELECT a.顧客番号"
wSQL = wSQL & "     , a.顧客名"
wSQL = wSQL & "     , a.顧客E_mail1"
wSQL = wSQL & "     , b.顧客郵便番号"
wSQL = wSQL & "     , b.顧客都道府県"
wSQL = wSQL & "     , b.顧客住所"
wSQL = wSQL & "     , c.顧客電話番号"
wSQL = wSQL & "  FROM Web顧客 a WITH (NOLOCK)"
wSQL = wSQL & "     , Web顧客住所 b WITH (NOLOCK)"
wSQL = wSQL & "     , Web顧客住所電話番号 c WITH (NOLOCK)"
wSQL = wSQL & " WHERE b.顧客番号 = a.顧客番号" 
wSQL = wSQL & "   AND c.顧客番号 = b.顧客番号" 
wSQL = wSQL & "   AND c.住所連番 = b.住所連番" 
wSQL = wSQL & "   AND b.住所連番 = 1" 
wSQL = wSQL & "   AND c.電話連番 = 1" 
wSQL = wSQL & "   AND a.顧客番号 = " & userID 
		
'@@@@@response.write(wSQL & "<BR>")

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

if RS.EOF = true then
	wNoData = "Y"
	exit function
else
	wCustomerName = RS("顧客名")
	wZip = RS("顧客郵便番号")
	wPrefecture = RS("顧客都道府県")
	wAddress = RS("顧客住所")
	wTelephone = RS("顧客電話番号")
	wEmail = RS("顧客E_mail1")
end if

RS.close

'---- 入力データチェック
call validation()

end function

'========================================================================
'
'    Function    入力内容チェック
'
'========================================================================
'
Function validation()

wMSG = ""

'---- 「メーカー」チェック
if MakerName ="" then
	wMSG = wMSG & "メーカーを入力してください。<br>"
elseif Len(MakerName) > 25 then
	wMSG = wMSG & "メーカーは25文字までです。<br>"
end if

'---- 「商品名」チェック
if ProductName ="" then
	wMSG = wMSG & "商品名を入力してください。<br>"
elseif Len(ProductName) > 50 then
	wMSG = wMSG & "商品名は50文字までです。<br>"
end if

'---- 「保証書あり/なし」チェック
if Warranty = "" then
	wMSG = wMSG & "保証書のあり/なしを選択してください。<br>"
elseif Warranty <> "あり" AND Warranty <> "なし" then
	wMSG = wMSG & "保証書のあり/なしの指定が不正です。<br>"
end if

'---- 「SerialNo」チェック
if Len(SerialNo) > 40 then
	wMSG = wMSG & "シリアル番号は40文字までです。<br>"
end if

if cf_checkHankaku(SerialNo) = false then
	wMSG = wMSG & "シリアル番号は半角で入力してください。<br>"
end if

'---- 「購入後期間」チェック
if WhenPurchased ="" then
	wMSG = wMSG & "ご購入後期間を選択してください。<br>"
elseif WhenPurchased <> "一週間以内" AND WhenPurchased <> "一年未満" AND WhenPurchased <> "一年以上/不明" then
	wMSG = wMSG & "ご購入後期間の指定が不正です。<br>"
end if

'---- 「内容」チェック
if Comment ="" then
	wMSG = wMSG & "内容を入力してください。<br>"
elseif Len(Comment) > 500 then
	wMSG = wMSG & "内容は500文字までです。<br>"
end if

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
<html lang="ja">
<head>
<meta charset="Shift_JIS">
<title>修理･商品サポート内容の確認｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/inquiry.css" type="text/css">
<script type="text/javascript">
//
//    Change onClick
//
function Change_onClick(pReturnURL){
    document.f_data.action = pReturnURL;
    document.f_data.submit();
}
//
//    Store onClick
//
function Store_onClick(pSendURL){
    document.f_data.action = pSendURL;
    document.f_data.submit();
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
      <p class="home"><a href="<%=g_HTTP%>"><img src="../images/icon_home.gif" alt="HOME"></a></p>
      <ul id="path">
        <li class="now">修理･商品サポート</li>
      </ul>
    </div></div></div>

    <h1 class="title">修理･商品サポート内容の確認</h1>
    <p>内容を確認の上、[送信する]ボタンを押してください。</p>

<!-- エラーメッセージ -->
<% if wMSG <> "" then %>
<ul class="error">
  <li><%=wMSG %></li>
</ul>
<% end if %>

<table>
  <tr>
    <th>メーカー<span>*</span></th>
    <td><%=MakerName%></td>
  </tr>
  <tr>
    <th>商品名<span>*</span></th>
    <td><%=ProductName%></td>
  </tr>
  <tr>
    <th>保証書<span>*</span></th>
    <td><%=Warranty%><% if Warranty = "なし" then %><span>（保証書が無い場合、保証を受けられない場合があります）</span><% end if %></td>
  </tr>
  <tr>
    <th>シリアル番号</th>
    <td><%=SerialNo%></td>
  </tr>
  <tr>
    <th>ご購入後期間<span>*</span></th>
    <td><%=WhenPurchased%></td>
  </tr>
  <tr>
    <th>内容<span>*</span></th>
    <td><%=Comment%></td>
  </tr>
  <tr>
    <th>お名前</th>
    <td><%=wCustomerName%></td>
  </tr>
  <tr>
    <th>住 所</th>
    <td>
      〒<%=wZip%><br>
      <%=wPrefecture%><%=wAddress%>
    </td>
  </tr>
  <tr>
    <th>電話番号</th>
    <td><%=wTelephone%></td>
  </tr>
  <tr>
    <th>Fax番号</th>
    <td><%=wFax%></td>
  </tr>
  <tr>
    <th>メールアドレス</th>
    <td><%=wEmail%></td>
  </tr>
</table>

<p>&laquo; <a href="JavaScript:Change_onClick('SupportInquiry.asp');">変更する</a></p>
<form name="f_data" method="post" action="SupportInquiryStore.asp">
  <input type="hidden" name="MakerName" value="<%=MakerName%>">
  <input type="hidden" name="ProductName" value="<%=ProductName%>">
  <input type="hidden" name="Warranty" value="<%=Warranty%>">
  <input type="hidden" name="SerialNo" value="<%=SerialNo%>">
  <input type="hidden" name="WhenPurchased" value="<%=WhenPurchased%>">
  <input type="hidden" name="Comment" value="<%=Comment%>">
  <input type="hidden" name="Skey" value="<%=Skey%>">
  <p class="btnBox"><input type="submit" value="送信する" class="opover"></p>
</form>

</div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>