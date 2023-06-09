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
'	修理･商品サポート送信
'
'更新履歴
'2010/10/04 an SupportInquirySendを元に新規作成
'2011/08/01 an #1087 Error.aspログ出力対応
'2012/06/29 if-web リニューアルレイアウト調整
'
'========================================================================

On Error Resume Next

Dim userID
Dim msg

Dim wCustomerName
Dim wZip
Dim wPrefecture
Dim wAddress
Dim wTelephone
Dim wFax
Dim wEmail

Dim MakerName
Dim ProductName
Dim Warranty
Dim SerialNo
Dim WhenPurchased
Dim Comment
Dim Bikou

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
Dim wNoData
Dim wErrDesc   '2011/08/01 an add

'========================================================================

Response.buffer = true

'---- セキュリティーキーチェック
if Session("SKey") <> ReplaceInput(Request("SKey")) then
	Response.redirect "SupportInquiry.asp"
end if

'---- UserID 取り出し
userID = Session("userID")

'---- 呼び出し元からのデータ取り出し

MakerName = ReplaceInput(Left(Request("MakerName"),25))
ProductName = ReplaceInput(Left(Request("ProductName"),50))
Warranty = ReplaceInput(Left(Request("Warranty"),2))
SerialNo = ReplaceInput(Left(Request("SerialNo"),40))
WhenPurchased = ReplaceInput(Left(Request("WhenPurchased"),10))
Comment = ReplaceInput(Left(Request("Comment"),500))

'---- Execute main
call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "SupportInquiryStore.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

call close_db()

if Err.Description <> "" OR wNoData = "Y" then
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

'---- 顧客情報取得
call GetCustomer()

'---- 顧客情報が取得できない場合はエラー
if wNoData = "Y" then
	exit function
else

	Connection.BeginTrans				'トランザクション開始
	
	'---- サポート依頼へ登録
	wSQL = ""
	wSQL = wSQL & "SELECT *"
	wSQL = wSQL & "  FROM サポート依頼"
	wSQL = wSQL & " WHERE 1 = 2"

	Set RS = Server.CreateObject("ADODB.Recordset")
	RS.Open wSQL, Connection, adOpenStatic, adLockOptimistic

	RS.AddNew

	RS("サポート依頼登録日") = Now()
	RS("顧客番号") = userID
	RS("メーカー名") = MakerName
	RS("商品名") = ProductName
	RS("保証書ありなし") = Warranty
	RS("シリアルNo") = SerialNo
	RS("購入後期間") = WhenPurchased
	RS("内容") = Comment

	RS.Update
	RS.close
	
	if Err.Description = "" then
		Connection.CommitTrans		'Commit
		
		'---- 顧客に受付メール送信
		call SendEmail()
	else
		Connection.RollbackTrans	'Rollback
	end if

end if

End function

'========================================================================
'
'	Function	顧客情報取得
'
'========================================================================

Function GetCustomer()

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

end Function

'========================================================================
'
'	Function	顧客へ受付メール送信
'
'========================================================================

Function SendEmail()

Dim i
Dim v_body
Dim v_body2
Dim v_subject
Dim OBJ_NewMail

'---- 顧客向け本文
v_body = ""
v_body = v_body & "受付日時　　　：" & now() & vbNewLine & vbNewLine
v_body = v_body & "メーカー　　　：" & MakerName & vbNewLine & vbNewLine
v_body = v_body & "商品名　　　　：" & ProductName & vbNewLine & vbNewLine
v_body = v_body & "保証書　　　　：" & Warranty & vbNewLine & vbNewLine
v_body = v_body & "シリアル番号　：" & SerialNo & vbNewLine & vbNewLine
v_body = v_body & "ご購入後期間　：" & WhenPurchased & vbNewLine & vbNewLine
v_body = v_body & "内容　　　　　：" & Comment & vbNewLine & vbNewLine

'---- 社内向け本文
v_body2 = ""
v_body2 = v_body2 & "名前　　　　　：" & wCustomerName & vbNewLine
v_body2 = v_body2 & "住所　　　　　：" & wZip & " " & wPrefecture & wAddress & vbNewLine
v_body2 = v_body2 & "電話番号　　　：" & wTelephone & vbNewLine
v_body2 = v_body2 & "Fax 　　　　　：" & wFax & vbNewLine
v_body2 = v_body2 & "Email　 　　　：" & wEmail & vbNewLine
v_body2 = v_body2 & "顧客番号　　　：" & UserID & vbNewLine

Set OBJ_NewMail = Server.CreateObject("CDO.message") 

'---- 社内向けメール作成
OBJ_NewMail.from = "support@soundhouse.co.jp"
OBJ_NewMail.to = "support@soundhouse.co.jp"

OBJ_NewMail.subject = "修理･商品サポート(" & MakerName & "/" & ProductName & ") " &  wCustomerName & "　様　 [" & UserID & "/RA/Web受付]"
OBJ_NewMail.TextBody = v_body & v_body2
OBJ_NewMail.BodyPart.Charset = "iso-2022-jp"

OBJ_NewMail.Send
'---- 社内向け　ここまで

'---- 顧客向け自動返信メール作成
call getCntlMst("Web","Email","トレーラ", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)

v_body = "お問い合わせありがとうございます｡" & vbNewLine _
       & "以下の内容にて修理･商品サポート依頼を受付いたしました｡" & vbNewLine _
       & "返答まで今しばらくお待ちください｡" & vbNewLine & vbNewLine _
       & v_body & vbNewLine _
       & wItemChar1

OBJ_NewMail.from = "support@soundhouse.co.jp"
OBJ_NewMail.to = wEmail

OBJ_NewMail.subject = "修理･商品サポート依頼を受付いたしました"
OBJ_NewMail.TextBody = v_body
OBJ_NewMail.BodyPart.Charset = "iso-2022-jp"

OBJ_NewMail.Send
'---- 顧客向け　ここまで

Set OBJ_NewMail = Nothing

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
<title>修理･商品サポート依頼を受け付けました｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/inquiry.css" type="text/css">
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

    <h1 class="title">修理･商品サポート依頼を受け付けました</h1>
    <p>
      修理・商品サポート依頼を登録しました。<br>
      後ほど弊社サポート担当からご連絡を差し上げます。<br>
      ありがとうございました。
    </p>
  </div>

<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>