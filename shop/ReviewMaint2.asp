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
'	商品レビューメンテナンス2
'     商品レビューの変更/削除を行う
'
'2011/09/06 an #816 新規作成
'2012/08/11 nt ショップコメントの更新項目を追加
'
'========================================================================
On Error Resume Next
Response.buffer = true
Response.Expires = -1			' Do not cache

Dim ReviewID
Dim i_Mode
Dim Title
Dim Hyouka
Dim Review
Dim UserID		'2012/08/11 nt add
Dim Password	'2012/08/11 nt add
Dim recCnt		'2012/08/11 nt add
Dim sCDate		'2012/08/11 nt add
Dim sComment	'2012/08/11 nt add

Dim wReviewDate
Dim wReviewName
Dim wMakerName
Dim wProductName

Dim Connection

Dim wErrMSG
Dim wLoginFl

'========================================================================

'---- Get GET/POST data
ReviewID = ReplaceInput(Request("ReviewID"))
i_Mode = ReplaceInput(Request("i_Mode"))
Title = ReplaceInput(Left(Request("Title"),51))
Hyouka = ReplaceInput(Request("Hyouka"))
Review = ReplaceInput(Left(Request("Review"),1001))
UserID = ReplaceInput(Request.Cookies("UserID"))			'2012/08/11 nt add
Password = ReplaceInput(Request.Cookies("Password"))		'2012/08/11 nt add
sCDate = ReplaceInput(Request("sCDate"))					'2012/08/11 nt add
sComment = ReplaceInput(Left(Request("sComment"),1000))		'2012/08/11 nt add

'---- Execute main
call connect_db()
call main()
call close_db()

if Err.Description <> "" then
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

'---- 未ログインの場合はログイン画面へ
if wLoginFl <> "Y" then
	Response.Redirect g_HTTPS & "shop/ReviewMaintLogin.asp"
end if

'---- エラーの場合はReviewMaintに戻る
if wErrMSG <> "" then
	Server.Transfer "ReviewMaint.asp"
else
	'2012/08/11 nt mod Start
	'if i_Mode = "update" then
	'	Response.Redirect g_HTTPS & "shop/ReviewMaint.asp?ReviewID=" & ReviewID
	'else
	'	Response.Redirect g_HTTPS & "shop/ReviewSearch.asp"
	'end if

	if i_Mode = "update" Or i_Mode = "sCDelete" then
		Response.Redirect g_HTTPS & "shop/ReviewMaint.asp?ReviewID=" & ReviewID
	else
		Response.Redirect g_HTTPS & "shop/ReviewSearch.asp"
	end if
	'2012/08/11 nt mod End
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

wErrMSG = ""
wLoginFl = "N"

'---- セキュリティーキーチェック
If Session("Skey") <> Request("Skey") Then
	Response.Redirect g_HTTP & "index.asp"
End If

'---- ログインステータス取得
wLoginFl = fGetSessionData(gSessionID, "ShAdminFl")

if wLoginFl <> "Y" then
	call fSetSessionData(gSessionID, "メッセージ", "ログインしてください。")
	exit function
end if

'2012/08/11 nt add Start
'---- WEB管理者マスタから、ログイン情報を取得
Call getWEBMaster()

'---- ログイン可否をセッションデータに登録
if recCnt = 0 then
	'---- 取得したログイン情報が存在しない場合、エラーメッセージを表示し、セッションデータにエラー登録
	wErrMSG = "不正なログインです。"
	call fSetSessionData(gSessionID, "メッセージ", wErrMSG)
	exit function
end if
'2012/08/11 nt add End

'---- 入力チェック
call validation()

if wErrMSG = "" then
	call UpdateDeleteReview()
end if

End function

'========================================================================
'
'    Function    入力内容チェック
'
'========================================================================
'
Function validation()

'---- 処理モード
'if i_Mode <> "update" AND i_Mode <> "delete" then 2012/08/11 nt mod
if i_Mode <> "update" AND i_Mode <> "delete" AND i_Mode <> "sCDelete" then
	wErrMSG = wErrMSG & "モードが不正です。<br />"
end if

'---- レビューID
if ReviewID = "" then
	wErrMSG = wErrMSG & "レビューIDを入力してください。<br />"
else
	if cf_checkNumeric(ReviewID) = false then
		wErrMSG = wErrMSG & "レビューIDが不正です。<br />"
	end if
end if

'---- タイトル
if Title = "" then
	wErrMSG = wErrMSG & "タイトルを入力してください。<br />"
else
	if Len(Title) > 50 then
		wErrMSG = wErrMSG & "タイトルは50文字以内で入力してください。<br />"
	end if
end if

'---- 評価
if Hyouka = "" then
	wErrMSG = wErrMSG & "評価を入力してください。<br />"
else
	if Hyouka <> "1" AND Hyouka <> "2" AND Hyouka <> "3" AND Hyouka <> "4" AND Hyouka <> "5" then
		wErrMSG = wErrMSG & "評価は1～5を入力してください。<br />"
	end if
end if

'---- レビュー内容
if Review = "" then
	wErrMSG = wErrMSG & "レビュー内容を入力してください。<br />"
else
	if Len(Review) > 1000 then
		wErrMSG = wErrMSG & "レビュー内容は1000文字以内で入力してください。<br />"
	end if
end if

'2012/08/11 nt add Start
if IsDate(sCDate) = false then
	wErrMSG = wErrMSG & "ショップコメント日が不正です。<br />"
end if

'---- ショップコメント
if Len(sComment) > 1000 then
	wErrMSG = wErrMSG & "ショップコメントは1000文字以内で入力してください。<br />"
end if
'2012/08/11 nt add End

'---- エラーがある場合はセッションデータに記録
if wErrMSG <> "" then
	call fSetSessionData(gSessionID, "メッセージ", wErrMSG)
end if

End function

'========================================================================
'
'    Function    商品レビュー変更、削除
'
'========================================================================
'
Function UpdateDeleteReview()

Dim RSv
Dim vSQL

vSQL = ""
vSQL = vSQL & "SELECT *"
vSQL = vSQL & " FROM 商品レビュー"
vSQL = vSQL & " WHERE ID = " & ReviewID

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic, adLockOptimistic
if RSv.EOF = true then
	wErrMSG = "該当のレビューがありません。 レビューID＝" & ReviewID
	call fSetSessionData(gSessionID, "メッセージ", wErrMSG)
else

	if i_Mode = "update" then
		RSv("タイトル") = Title
		RSv("評価") = Hyouka
		RSv("レビュー内容") = Review

		'2012/08/11 nt add Start
		'---- ショップコメントに入力がなければ、ショップコメント日およびショップコメントは更新しない
		if len(sComment) > 0 then
			if (sCDate <> "") then
				RSv("ショップコメント日") = cf_FormatDate(sCDate, "YYYY/MM/DD")
			end if
			RSv("ショップコメント") = sComment
		end if
		'2012/08/11 nt add End

		RSv.Update

		'2012/08/11 nt mod Start
		'call fSetSessionData(gSessionID, "メッセージ", "更新されました。")
		if len(sComment) > 0 then
			call fSetSessionData(gSessionID, "メッセージ", "更新されました。")
		else
			call fSetSessionData(gSessionID, "メッセージ", "更新されました。<br>※）ショップコメントが入力されなかったため、ショップコメントは更新されません")
		end if
		'2012/08/11 nt mod End

	'2012/08/11 nt add Start
	'---- 「ショップコメントのみ削除」ボタンを追加
	elseif i_Mode = "sCDelete" then

		'2012/08/11 nt add Start
		RSv("ショップコメント日") = NULL
		RSv("ショップコメントタイトル") = NULL
		RSv("ショップコメント") = NULL
		'2012/08/11 nt add End

		RSv.Update
		call fSetSessionData(gSessionID, "メッセージ", "ショップコメントが削除されました。")
	'2012/08/11 nt add End

	else
		RSv.Delete
		call fSetSessionData(gSessionID, "メッセージ", "削除されました。")
	end if
end if

RSv.Close

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