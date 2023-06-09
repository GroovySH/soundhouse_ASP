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
'	商品レビュー 参考数登録
'
'更新履歴
'2007/10/19 ハッカーセーフ対応
'2008/05/23 入力データチェック強化（LEFT, Numeric, EOF他)
'2009/04/30 エラー時にerror.aspへ移動
'2011/08/01 an #1087 Error.aspログ出力対応
'2012/07/30 if-web リニューアルレイアウト調整
'
'========================================================================

On Error Resume Next

Dim userID
Dim msg

Dim ID
Dim Sankou
Dim Item

Dim Connection
Dim RS

Dim w_sql
Dim w_html
Dim w_msg
Dim wErrDesc   '2011/08/01 an add

'========================================================================

'---- UserID 取り出し
userID = Session("userID")

'---- 呼び出し元からのデータ取り出し
ID = ReplaceInput(Request("ID"))
Sankou = ReplaceInput(Request("Sankou"))
Item = ReplaceInput(Request("Item"))

if isNumeric(ID) = false then
	ID = 0
end if

'---- Execute main
call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "ReviewSankou.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
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
'	Function	main proc
'
'========================================================================
'
Function main()

'---- 参考数登録
w_sql = ""
w_sql = w_sql & "SELECT 参考数"
w_sql = w_sql & "     , 不参考数"
w_sql = w_sql & "  FROM 商品レビュー"
w_sql = w_sql & " WHERE ID = " & ID 
	  
Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open w_sql, Connection, adOpenStatic, adLockOptimistic

if RS.EOF = false then

	'---- 参考数/不参考数 更新
	if Sankou = "Y" then
		RS("参考数") = RS("参考数") + 1
	else
		RS("不参考数") = RS("不参考数") + 1
	end if

	RS.Update
end if

RS.close

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
<html>
<head>
<meta charset="Shift_JIS">
<title>商品レビュー｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->
</head>
<body>
<!--#include file="../Navi/Navitop.inc"-->
<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="ここから本文です"></a></span>

<!-- コンテンツstart -->
<div id="globalContents">

  <p>登録されました｡　どうもありがとうございました。</p>
  <p class="btnBox"><a href="ProductDetail.asp?Item=<%=item%>" class="opover">商品ページへ戻る</a></p>

  </div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/Navibottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>