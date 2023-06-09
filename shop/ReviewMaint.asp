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
'	商品レビューメンテナンスページ
'     商品レビューの変更/削除を行う
'
'2011/09/06 an #816 新規作成
'2012/08/11 nt ショップコメントの入力フォーム表示およびその制御を追加
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
Dim Auth		'2012/08/11 nt add
Dim sCDate		'2012/08/11 nt add
Dim sComment	'2012/08/11 nt add
Dim vErrMSG		'2012/08/11 nt add

Dim wReviewDate
Dim wReviewName
Dim wMakerName
Dim wProductName

Dim Skey
Dim Connection

Dim wMSG   'ReviewMaint2からのエラー/完了メッセージ
Dim wNoData
Dim wLoginFl

'========================================================================

'---- Get GET/POST data
ReviewID = ReplaceInput(Request("ReviewID"))
i_Mode = ReplaceInput(Request("i_Mode"))          '以下、エラー時にReviewMaint2から受け取り
Title = ReplaceInput(Left(Request("Title"),50))
Hyouka = ReplaceInput(Request("Hyouka"))
Review = ReplaceInput(Left(Request("Review"),1000))
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
	Response.Redirect "ReviewMaintLogin.asp"
end if

'---- IDの指定が不正、該当レビューがない場合は検索画面へ
if wNoData = "Y" then
	Response.Redirect "ReviewSearch.asp?UserID=" & UserID
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

wMSG = ""
wNoData = "N"
wLoginFl = "N"

'---- セキュリティーキーセット
Skey = SetSecureKey()

'---- ログインステータス取得
wLoginFl = fGetSessionData(gSessionID, "ShAdminFl")

if wLoginFl <> "Y" then
	call fSetSessionData(gSessionID, "メッセージ", "ログインしてください。")
	exit function
end if

'2012/08/13 nt add
'---- ログインユーザの権限を取得
call getWEBMasterAuth

'---- ReviewMaint2.aspのエラーメッセージ取得・クリア
wMSG = fGetSessionData(gSessionID, "メッセージ")
call fSetSessionData(gSessionID, "メッセージ", "")

'---- 入力チェック
call validation()

if wNoData <> "Y" then
	call GetReview()
end if

End function

'========================================================================
'
'    Function    入力内容チェック
'
'========================================================================
'
Function validation()

Dim vErrMSG

vErrMSG = ""

if ReviewID = "" then
	vErrMSG = "レビューIDを入力してください。"
else
	if cf_checkNumeric(ReviewID) = false then
		vErrMSG = "レビューIDが不正です。"
	end if
end if

if vErrMSG <> "" then
	wNoData = "Y"
	call fSetSessionData(gSessionID, "メッセージ", vErrMSG)
end if

End function

'========================================================================
'
'    Function    商品レビュー取得
'
'========================================================================
'
Function GetReview()

Dim RSv
Dim vSQL

vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    a.ID"
vSQL = vSQL & "  , a.投稿日"
vSQL = vSQL & "  , a.タイトル"
vSQL = vSQL & "  , a.評価"
vSQL = vSQL & "  , a.名前"
vSQL = vSQL & "  , a.レビュー内容"
vSQL = vSQL & "  , b.メーカー名"
vSQL = vSQL & "  , c.商品名"
'2012/08/11 nt add Start
vSQL = vSQL & "  , a.ショップコメント日"
vSQL = vSQL & "  , a.ショップコメントタイトル"
vSQL = vSQL & "  , a.ショップコメント"
'2012/08/11 nt add End
vSQL = vSQL & " FROM"
vSQL = vSQL & "    商品レビュー a WITH (NOLOCK)"
vSQL = vSQL & "  , メーカー b WITH (NOLOCK)"
vSQL = vSQL & "  , Web商品 c WITH (NOLOCK)"
vSQL = vSQL & " WHERE a.ID = " & ReviewID
vSQL = vSQL & "   AND b.メーカーコード = a.メーカーコード"
vSQL = vSQL & "   AND c.商品コード = a.商品コード"
vSQL = vSQL & "   AND c.メーカーコード = a.メーカーコード"

'@@@@@response.write(vSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic

if RSv.EOF = true then
	wNoData = "Y"
	vErrMSG = "該当のレビューがありません。 レビューID＝" & ReviewID
	call fSetSessionData(gSessionID, "メッセージ", vErrMSG)
else
	wReviewDate = RSv("投稿日")
	wReviewName = RSv("名前")
	wMakerName = RSv("メーカー名")
	wProductName = RSv("商品名")

	'---- ReviewMaint2からエラーで戻った時は、DBから取得しない
	if i_Mode <> "update" then

		Title = RSv("タイトル")
		Hyouka = RSv("評価")
		Review = RSv("レビュー内容")

		'2012/08/11 nt add Start
		sCDate = RSv("ショップコメント日")
		if (isNull(sCDate) = true) then
			'---- ショップコメント日付がない場合、システム日付をセット
			sCDate = now()
		end if

		sComment = RSv("ショップコメント")
		'2012/08/11 nt add End
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
'	Function	ログインユーザのWEB管理者権限を取得
'
'========================================================================
Function getWEBMasterAuth()

Dim RSv
Dim vSQL

vSQL = ""
vSQL = vSQL & "SELECT 権限 "
vSQL = vSQL & " FROM "
vSQL = vSQL & "    WEB管理者 a WITH (NOLOCK) "
vSQL = vSQL & " WHERE "
vSQL = vSQL & "        a.ユーザID = '" & UserID & "' "
vSQL = vSQL & "    AND a.パスワード = '" & Password & "' "
vSQL = vSQL & "    AND a.削除フラグ = '0'"	'削除フラグ[0]：Active、[1]：Non-Active

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic, adLockOptimistic

if RSv.EOF = true then
	wNoData = "Y"
	vErrMSG = "該当のユーザが存在しません。 ユーザID＝" & UserID
	call fSetSessionData(gSessionID, "メッセージ", vErrMSG)
else
	'---- WEB管理者情報の有無を取得
	Auth = RSv("権限")
end if

RSv.Close

End Function

'========================================================================

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS" />
<title>商品レビューメンテナンス</title>
<link rel="stylesheet" type="text/css" href="style/review.css" />
<script type="text/javascript" src="jslib/review.js?20120817"></script>
</head>
<body>
<div id="content">
<h1>商品レビューメンテナンス</h1>
<td>◆<font color="red"><%=UserID%> </font>さんでログイン中です。</td>
<% if wMSG <> "" then %>
<p class="notes"><%=wMSG%></p>
<% end if %>
<form name="f_data" method="post">
<table>
  <tr>
    <th>レビューID</th>
    <td><%=ReviewID%></td>
  </tr>
  <tr>
    <th>メーカー</th>
    <td><%=wMakerName%></td>
  </tr>
  <tr>
    <th>商品名</th>
    <td><%=wProductName%></td>
  </tr>
  <tr>
    <th>投稿日</th>
    <td><%=wReviewDate%></td>
  </tr>
  <tr>
    <th>タイトル</th>
    <td><input type="text" name="Title" value="<%=Title%>" size="50" maxsize="50" <%if (Auth <> "1") And (Auth <> "2") then%>readonly="readonly" style="background-color:#DCDCDC;"<%end if%> /></td>
  </tr>
  <tr>
    <th>評価</th>
    <td><input type="text" name="Hyouka" value="<%=Hyouka%>" <%if (Auth <> "1") And (Auth <> "2") then%>readonly="readonly" style="background-color:#DCDCDC;"<%end if%> /></td>
  </tr>
  <tr>
    <th>投稿者名</th>
    <td><%=wReviewName%></td>
  </tr>
  <tr>
    <th>レビュー内容</th>
    <td><textarea name="Review" rows="15" cols="60" <%if (Auth <> "1") And (Auth <> "2") then%>readonly="readonly" style="background-color:#DCDCDC;"<%end if%> ><%=Review%></textarea></td>
  </tr>
</table>
<hr>
<h2>ショップコメント</h2>
<table>
  <tr>
    <th>コメント日</th>
    <td><input type="text" name="sCDate" value="<%=sCDate%>" <%if (Auth <> "1") And (Auth <> "3") then%>readonly="readonly" style="background-color:#DCDCDC;"<%end if%> /></td>
  </tr>
  <tr>
    <th>コメント</th>
    <td><textarea name="sComment" rows="15" cols="60" <%if (Auth <> "1") And (Auth <> "3") then%>readonly="readonly" style="background-color:#DCDCDC;" <%end if%> ><%=sComment%></textarea></td>
  </tr>
</table>
<div id="button_div">

<!-- 2012/08/11 nt mod Start -->
<input type="submit" value=" 変更 " onClick="return Update_onClick();" <%if (Auth <> "1") And (Auth <> "2") And (Auth <> "3") then%>disabled<%end if%> />
<input type="submit" value=" 削除 " onClick="return Delete_onClick();" <%if (Auth <> "1") then%>disabled<%end if%> />
<input type="submit" value=" ショップコメントのみ削除 " onClick="return sCDelete_onClick();" <%if (Auth <> "1") And (Auth <> "3") then%>disabled<%end if%> />
<!-- 2012/08/11 nt mod End -->

<input type="submit" value=" 戻る " onClick="Return_onClick();" />
<input type="hidden" name="ReviewID" value="<%=ReviewID%>" />
<!-- 2012/08/11 nt add Start -->
<input type="hidden" name="UserID" value="<%=UserID%>" />
<!-- 2012/08/11 nt add End -->
<input type="hidden" name="i_Mode" value="" />
<input type="hidden" name="Skey" value="<%=Skey%>" />
</div>
</form>
</div>
</body>
</html>