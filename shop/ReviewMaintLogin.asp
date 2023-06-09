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
'    商品レビューメンテナンスログイン
'更新履歴
'2011/09/06 an #816 新規作成
'
'========================================================================

On Error Resume Next
Response.buffer = true
Response.Expires = -1			' Do not cache

Dim wErrMSG
Dim Connection

'========================================================================

'---- Execute main
call connect_db()
call main()
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
'
Function main()

'---- ReviewMaintLoginCheck.aspのエラーメッセージ取得・クリア
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
<title>商品レビューメンテナンスログイン</title>
<link rel="stylesheet" type="text/css" href="style/review.css" />
</head>
<body>
<div id="content">
<h1>商品レビューメンテナンスログイン</h1>
<% if wErrMSG <> "" then %>
<p class="notes"><span style="margin:40px"><%=wErrMSG%></span></p>
<% end if %>
<form name="f_login" method="post" action="ReviewMaintLoginCheck.asp">
<ul>
<li>ユーザーID　<input name="UserID" maxlength="30" size="30" autocomplete="off" /></li><br />
<li>パスワード　 <input type="password" name="Password" maxlength="30" size="30" autocomplete="off" /></li>
<ul>
<br />
<span style="margin:45px"><input type="submit" value="ログイン" /></span>
</form>
</div>
</body>
</html>