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
'    商品レビューメンテナンスログインチェック
'更新履歴
'2011/09/06 an 新規作成
'2012/08/11 nt ログイン情報取得先を変更
'             （コントロールマスタ→WEB管理者マスタ[新設]）
'
'========================================================================

On Error Resume Next
Response.Buffer = true
Response.Expires = -1			' Do not cache

Dim UserID
Dim Password
Dim Logout
Dim recCnt		'2012/08/11 nt add
Dim url
Dim Connection

Dim wErrMSG

'========================================================================

'---- Get GET/POST data
UserID = ReplaceInput(Trim(Request("UserID")))
Password = ReplaceInput(Trim(Request("Password")))
Logout = ReplaceInput(Trim(Request("Logout")))

'2012/08/11 nt add
'---- Set Cookie data
Response.Cookies("UserID") = UserID
Response.Cookies("Password") = Password

'---- Execute main
call connect_db()
call main()
call close_db()

if Err.Description <> "" then    
    Response.Redirect g_HTTP & "shop/Error.asp"
end if

'ユーザーID/パスワード不整合時、ログアウト時はログイン画面に戻る
if wErrMSG <> "" then
	Response.Redirect "ReviewMaintLogin.asp"
else
	Response.Redirect "ReviewSearch.asp"
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
'    Function    Main
'
'========================================================================
'
Function main()

Dim vItemChar1
Dim vItemChar2
Dim vItemNum1
Dim vItemNum2
Dim vItemDate1
Dim vItemDate2

'Dim vUserID		2012/08/11 nt del
'Dim vPassword		2012/08/11 nt del

wErrMSG = ""

if Logout = "Y" then
	'---- ログアウト
	call fSetSessionData(gSessionID, "ShAdminFl", "")
	wErrMSG = "ログアウト"
	exit function
end if

'2012/08/11 nt add Start
if UserID = "" And Password = "" then
	wErrMSG = "ユーザーID・パスワードを入力して下さい。"
	call fSetSessionData(gSessionID, "メッセージ", wErrMSG)
	exit function
end if
'2012/08/11 nt add End

'2012/08/11 nt mod Start
'---- コントロールマスタからログイン情報取得
'call getCntlMst("レビュー","ログイン","1", vItemChar1, vItemChar2, vItemNum1, vItemNum2, vItemDate1, vItemDate2)
'vUserID = vItemChar1
'vPassword = vItemChar2

'---- ユーザーID OR パスワードが不一致の場合はセッションデータにエラー登録
'if UserID <> vUserID OR Password <> vPassword then
'	wErrMSG = "ユーザーIDまたはパスワードが不正です。"
'	call fSetSessionData(gSessionID, "メッセージ", wErrMSG)
'---- OKの場合は「ログイン中」に設定
'else
'	call fSetSessionData(gSessionID, "ShAdminFl", "Y")
'end if

'---- WEB管理者マスタから、ログイン情報を取得
Call getWEBMaster()

'---- ログイン可否をセッションデータに登録
if recCnt = 0 then
	'---- 取得したログイン情報が存在しない場合、エラーメッセージを表示し、セッションデータにエラー登録
	wErrMSG = "ユーザーIDまたはパスワードが不正です。"
	call fSetSessionData(gSessionID, "メッセージ", wErrMSG)

else
	'---- 取得したログイン情報が存在すれば、セッションデータにフラグ登録
	call fSetSessionData(gSessionID, "ShAdminFl", "Y")

end if
'2012/08/11 nt mod End

End Function
 

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

'2012/08/11 nt add
'========================================================================
'
'	Function	WEB管理者情報の有無を取得
'
'========================================================================
Function getWEBMaster()

Dim RSv
Dim vSQL

vSQL = ""
vSQL = vSQL & "SELECT * "
vSQL = vSQL & " FROM "
vSQL = vSQL & "    WEB管理者 a WITH (NOLOCK) "
vSQL = vSQL & " WHERE "
vSQL = vSQL & "        a.ユーザID = '" & UserID & "' "
vSQL = vSQL & "    AND a.パスワード = '" & Password & "' "
vSQL = vSQL & "    AND a.削除フラグ = '0'"	'削除フラグ[0]：Active、[1]：Non-Active

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic, adLockOptimistic

'---- WEB管理者情報の有無を取得
recCnt = RSv.RecordCount

RSv.Close

End Function

'========================================================================
%>