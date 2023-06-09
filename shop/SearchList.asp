<%@ LANGUAGE="VBScript" %>
<%
'ネットハウスねっとハウスネットはうす
'サウンドハウス
Option Explicit
'========================================================================
'
'	商品一覧ページ(guide.soundhouse.co.jp 専用)
'
'更新履歴
'2016.02.09 GV PHP版へリダイレクト
'
On Error Resume Next

Dim url
url = "http://www.soundhouse.co.jp/search/index?"
url = url & Request.QueryString

Response.Clear()
Response.Status = "301 Moved Permanently"
Response.AddHeader "Location", url
Response.End()
'Response.Redirect url
%>