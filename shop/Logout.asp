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
'	ログアウト
'
'	更新履歴
'
'2005/10/14 TOPへ戻るケースにinquirysend.asp, catalogrequeststoreを追加
'2006/09/18 LoginFl追加
'2007/12/21 Englishページ追加に伴い戻りさき決定をファイル名のみで行う
'2011/10/01 an #722 Cookie削除
'2011/12/19 hn DC対応
'2012/01/18 hn Cookieにドメイン属性追加
'2012/02/15 GV Cookie の LIFL を ULIFL にキー名変更
'2012/08/14 GV #1419 未ログイン時ウィッシュリストからログイン画面を表示する
'
'========================================================================

'=======================================================================

Dim wHTTP_REFERER

Const LOGIN_FLAG_KEY = "ULIFL"	' 2012/02/15 GV Add

'---- メイン処理
Session("userID") = ""
Session("userName") = ""
Session("userEmail") = ""
Session("LoginFl") = ""

Response.Cookies("CustName").Expires = DateAdd("d", -1, Now())	'2011/10/01 an add 有効期限切れで上書きして削除
Response.Cookies("CustName").Domain = gCookieDomain				'2012/01/18 hn add
' 2012/02/15 GV Mod Start
'Response.Cookies("LIfl").Expires = DateAdd("d", -1, Now())		'2011/12/19 hn add ログインフラグ
'Response.Cookies("LIfl").Domain = gCookieDomain					'2012/01/18 hn add
Response.Cookies(LOGIN_FLAG_KEY).Expires = DateAdd("d", -1, Now())
Response.Cookies(LOGIN_FLAG_KEY).Domain = gCookieDomain
If Len(ReplaceInput(Request.Cookies("LIfl"))) > 0 Then
	' 古い Cookie の LIfl が存在する場合、こちらも削除
	Response.Cookies("LIfl").Expires = DateAdd("d", -1, Now())
	Response.Cookies("LIfl").Domain = gCookieDomain
End If
' 2012/02/15 GV Mod End

Session.Abandon

wHTTP_REFERER = LCase(Request.ServerVariables("HTTP_REFERER"))

'---- 下記該当URLからログアウトされた場合はTOPへ戻る
if InStr(wHTTP_REFERER, "orderinfoenter.asp") > 0 then
	wHTTP_REFERER = g_HTTP
end if
if InStr(wHTTP_REFERER, "orderconfirm.asp") > 0 then
	wHTTP_REFERER = g_HTTP
end if
if InStr(wHTTP_REFERER, "thanks.asp") > 0 then
	wHTTP_REFERER = g_HTTP
end if
if InStr(wHTTP_REFERER, "presentoubo.asp") > 0 then
	wHTTP_REFERER = g_HTTP
end if
if InStr(wHTTP_REFERER, "catalogrequest.asp") > 0 then
	wHTTP_REFERER = g_HTTP
end if
if InStr(wHTTP_REFERER, "inquirysend.asp") > 0 then
	wHTTP_REFERER = g_HTTP
end if
if InStr(wHTTP_REFERER, "catalogrequeststore.asp") > 0 then
	wHTTP_REFERER = g_HTTP
end if
if InStr(wHTTP_REFERER, "/member") > 0 then
	wHTTP_REFERER = g_HTTP
end if
' 2012/08/14 GV #1419 Add Start
if InStr(wHTTP_REFERER, "wishlist.asp") > 0 then
	wHTTP_REFERER = g_HTTP
end if
' 2012/08/14 GV #1419 Add End

Response.Redirect wHTTP_REFERER		'呼び出しもとへもどる

%>
