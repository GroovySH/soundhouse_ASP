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
'2007/01/11 コンタクト管理用にSubjectにコンタクトカテゴリー/サブカテゴリーを追加
'2008/04/10 SMTP Server変更
'2008/05/12 改行コードインジェクション対策（i_toパラメータ削除）
'2008/05/13 クロスサイトリクエストフォジェリー対策 Keyパラメータチェック
'2010/01/12 削除する改行コードの指定をvbCr/vbLfに変更
'2011/08/01 an #1087 Error.aspログ出力対応
'
'========================================================================

On Error Resume Next

Dim userID
Dim msg

Dim CustomerName
Dim Furigana
Dim Zip
Dim Prefecture
Dim Address
Dim Telephone
Dim Fax
Dim Email

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

MakerName = ReplaceInput(Request("MakerName"))
ProductName = ReplaceInput(Request("ProductName"))
Warranty = ReplaceInput(Request("Warranty"))
SerialNo = ReplaceInput(Request("SerialNo"))
WhenPurchased = ReplaceInput(Request("WhenPurchased"))
Comment = ReplaceInput(Request("Comment"))
Bikou = ReplaceInput(Request("Bikou"))

CustomerName = ReplaceInput(Request("CustomerName"))
Furigana = ReplaceInput(Request("Furigana"))
Zip = ReplaceInput(Request("Zip"))
Prefecture = ReplaceInput(Request("Prefecture"))
Address = ReplaceInput(Request("Address"))
Email = Replace(Replace(LCase(ReplaceInput(Request("Email"))), vbCr, ""), vbLf, "")
Telephone = ReplaceInput(Request("Telephone"))
Fax = ReplaceInput(Request("Fax"))

'---- Execute main
call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "SupportInquirySend.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
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

Dim i
Dim v_body
Dim v_body2
Dim v_subject
Dim OBJ_NewMail

'---- edit body
v_body = ""
v_body = v_body & "受付日時　　　：" & now() & vbNewLine & vbNewLine
'v_body = v_body & "件名　　　　　：修理･商品サポート(" & MakerName & "/" & ProductName & ")" & vbNewLine & vbNewLine
v_body = v_body & "メーカー　　　：" & MakerName & vbNewLine & vbNewLine
v_body = v_body & "商品名　　　　：" & ProductName & vbNewLine & vbNewLine
v_body = v_body & "保証書　　　　：" & Warranty & vbNewLine & vbNewLine
v_body = v_body & "シリアル番号　：" & SerialNo & vbNewLine & vbNewLine
v_body = v_body & "ご購入後期間　：" & WhenPurchased & vbNewLine & vbNewLine
v_body = v_body & "内容　　　　　：" & Comment & vbNewLine & vbNewLine
v_body = v_body & "備考　　　　　：" & Bikou & vbNewLine & vbNewLine

v_body2 = ""
v_body2 = v_body2 & "名前　　　　　：" & CustomerName & vbNewLine
v_body2 = v_body2 & "ふりがな　　　：" & Furigana & vbNewLine
v_body2 = v_body2 & "住所　　　　　：" & Zip & " " & Prefecture & Address & vbNewLine
v_body2 = v_body2 & "電話番号　　　：" & Telephone & vbNewLine
v_body2 = v_body2 & "Fax 　　　　　：" & Fax & vbNewLine
v_body2 = v_body2 & "Ｅメール　　　：" & Email & vbNewLine
v_body2 = v_body2 & "顧客番号　　　：" & UserID & vbNewLine

'---- send e-mail
Set OBJ_NewMail = Server.CreateObject("CDO.message") 

OBJ_NewMail.from = "support@soundhouse.co.jp"
OBJ_NewMail.to = "support@soundhouse.co.jp"

OBJ_NewMail.subject = "修理･商品サポート(" & MakerName & "/" & ProductName & ") " &  CustomerName & "　様　 [" & UserID & "/RA/Web受付]"
OBJ_NewMail.TextBody = v_body & v_body2
OBJ_NewMail.BodyPart.Charset = "iso-2022-jp"

OBJ_NewMail.Send

'---- 自動返信メール作成（顧客へ)
call getCntlMst("Web","Email","トレーラ", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)

v_body = "お問合せありがとうございます｡" & vbNewLine _
       & "以下の内容にて修理･商品サポート依頼を受付いたしました｡" & vbNewLine _
       & "返答まで今しばらくお待ちください｡" & vbNewLine & vbNewLine _
       & v_body & vbNewLine _
       & wItemChar1

OBJ_NewMail.from = "support@soundhouse.co.jp"
OBJ_NewMail.to = Email

OBJ_NewMail.subject = "修理･商品サポート依頼を受付いたしました"
OBJ_NewMail.TextBody = v_body
OBJ_NewMail.BodyPart.Charset = "iso-2022-jp"

'---- メールサーバー指定
'OBJ_NewMail.Configuration.Fields.Item(g_ItemSMTPSendusing) = g_SMTPSendusing
'OBJ_NewMail.Configuration.Fields.Item(g_ItemSMTPServer) = g_SMTPServer
'OBJ_NewMail.Configuration.Fields.Item(g_ItemSMTPServerPort) = g_SMTPServerPort
'OBJ_NewMail.Configuration.Fields.Update

OBJ_NewMail.Send

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

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<title>サウンドハウス 修理･商品サポート</title>

<!-- 追加SCRIPTはここへ-->

<!--#include file="../Navi/NaviStyle.inc"-->

</head>

<body background="../Navi/Images/back_ground.gif" bgcolor="#ffffff" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<!--#include file="../Navi/NaviTop.inc"-->

<table width="940" height="26" border="0" cellpadding="0" cellspacing="0">
  <tr>

<!--#include file="../Navi/NaviLefta.inc"-->

    <td width="798" align="left" valign="top" bgcolor="#ffffff">

<!------------ ページメイン部分の記述 START ------------>

      <table border="0" cellspacing="0" cellpadding="3">
        <tr>
          <td align="left"><b><font color="#696684">修理･商品サポート</font></b></td>
        </tr>
      </table>

      <table width="798" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td>&nbsp;</td>
          <td class="honbun">
            <br>
            修理･商品サポート依頼を送信しました。<br>
            ありがとうございました。
          </td>
        </tr>
      </table>

<!------------ ページメイン部分の記述 END ------------>

    </td>
  </tr>
</table>

<!--#include file="../Navi/NaviBottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->

</body>
</html>
