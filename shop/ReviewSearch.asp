<%@ LANGUAGE="VBScript" %>
<%
'ネットハウスねっとハウスネットはうす
'サウンドハウス
Option Explicit
%>
<!--#include file="../common/ADOVBS.inc"-->
<!--#include file="../common/Bfunctions1.asp"-->
<!--#include file="../common/shop_common_functions.inc"-->
<!--#include file="../common/system_common.inc"-->
<!--#include file="../common/HttpsSecurity.inc"-->

<%
'========================================================================
'
'    商品レビュー検索
'更新履歴
'2011/09/06 an #816 新規作成
'2012/08/11 nt ログイン情報の引き継ぎを追加
'
'========================================================================

On Error Resume Next
Response.buffer = true
Response.Expires = -1			' Do not cache

Dim Connection

Dim UserID		'2012/08/11 nt add
Dim Password	'2012/08/11 nt add
Dim wErrMSG
Dim wLoginFl

'========================================================================

'2012/08/11 nt add
'---- Get Cookie data
UserID = Request.Cookies("UserID")
Password = Request.Cookies("Password")

'---- Execute main
call connect_db()
call main()
call close_db()

if Err.Description <> "" then
    Response.Redirect g_HTTP & "shop/Error.asp"
end if

'---- 未ログインの場合はログイン画面へ
if wLoginFl <> "Y" then
	Response.Redirect "ReviewMaintLogin.asp"
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
'
Function main()

wErrMSG = ""
wLoginFl = "N"

'---- ログインステータス取得
wLoginFl = fGetSessionData(gSessionID, "ShAdminFl")

if wLoginFl <> "Y" then
	call fSetSessionData(gSessionID, "メッセージ", "ログインしてください。")
	exit function
end if

'---- ReviewMaintのエラーメッセージ、ReviewMain2の削除完了メッセージ取得・クリア
wErrMSG = fGetSessionData(gSessionID, "メッセージ")
call fSetSessionData(gSessionID, "メッセージ", "")

End function

'========================================================================
'
'	Function	Close database
'
'========================================================================
'
Function close_db()

Connection.close
Set Connection= Nothing

End function

'========================================================================
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS" />
<title>商品レビュー検索</title>
<link rel="stylesheet" type="text/css" href="style/review.css" />
</head>
<body>
<div id="content">
<h1>商品レビュー検索</h1>
<td>◆<font color="red"><%=UserID%> </font>さんでログイン中です。</td>
<% if wErrMSG <> "" then %>
<p class="notes"><%=wErrMSG%></p>
<% end if %>
<p>変更・削除を行うレビューIDを入力してください。</p>
<form name="f_data" method="post" action="ReviewMaint.asp">
<ul>
<li>レビューID　<input name="ReviewID" maxlength="30" size="30" autocomplete="off" /></li>
<ul>
<br />
<span style="margin:45px"><input type="submit" value="レビュー検索" /></span>
<!-- 2012/08/11 nt add Start -->
<input type="hidden" name="UserID" value="<% = UserID %>">
<!-- 2012/08/11 nt add End -->
</form>
<br /><br />
<form method="post" action="ReviewMaintLoginCheck.asp?Logout=Y">
<span style="margin:45px"><input type="submit" value="ログアウト" /></span>
</form>
</div>
</body>
</html>