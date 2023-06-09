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
'	カート内容の保存2
'
'更新履歴
'2008/05/23 入力データチェック強化（LEFT, Numeric, EOF他)
'2009/04/30 エラー時にerror.aspへ移動
'2011/04/14 hn SessionID関連変更
'2011/08/01 an #1087 Error.aspログ出力対応
'
'========================================================================

On Error Resume Next

Dim userID

Dim CartName

Dim Connection
Dim RS

Dim wSQL
Dim wHTML
Dim wMSG
Dim wErrDesc   '2011/08/01 an add

'========================================================================

Response.buffer = true

'---- UserID 取り出し
userID = Session("userID")

'---- 呼び出し元からのデータ取り出し
CartName = Left(ReplaceInput(Request("CartName")), 20)

'---- Execute main
call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "SaveCart2.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

call close_db()

if Err.Description <> "" then	
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

if wMSG = "" then
	Response.Redirect "SaveCartList.asp"
else
	Response.Redirect "SaveCart.asp?msg=" & wMSG
end if

'========================================================================
'
'	Function	Connect database
'
'========================================================================
'
Function connect_db()
Dim i

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

Dim RSv

'----仮受注データ取り出し
wSQL = ""
wSQL = wSQL & "SELECT a.受注明細番号"
wSQL = wSQL & "     , a.メーカーコード"
wSQL = wSQL & "     , a.商品コード"
wSQL = wSQL & "     , a.色"
wSQL = wSQL & "     , a.規格"
wSQL = wSQL & "     , a.受注数量"
wSQL = wSQL & "  FROM 仮受注明細 a WITH (NOLOCK)"
wSQL = wSQL & " WHERE SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod

'@@@@@@response.write(wSQL)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

if RS.EOF = true then
	wMSG = "保存するカート情報がありません。"
	exit function
end if

'---- 保存カート情報登録
wSQL = ""
wSQL = wSQL & "SELECT *"
wSQL = wSQL & "  FROM 保存カート"
wSQL = wSQL & " WHERE 顧客番号 = " & userID
wSQL = wSQL & "   AND カート名 = '" & CartName & "'"
	  
Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RSv.EOF = false then
	RSv.Delete
end if

RSv.AddNew

RSv("顧客番号") = userID
RSv("カート名") = CartName
RSv("登録日") = now()

RSv.Update
RSv.close


Do Until RS.EOF = true

	'---- 保存カート明細情報登録
	wSQL = ""
	wSQL = wSQL & "SELECT *"
	wSQL = wSQL & "  FROM 保存カート明細"
	wSQL = wSQL & " WHERE 1 = 2"
		  
	Set RSv = Server.CreateObject("ADODB.Recordset")
	RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

	'---- insert カタログ請求
	RSv.AddNew

	RSv("顧客番号") = userID
	RSv("カート名") = CartName
	RSv("受注明細番号") = RS("受注明細番号")
	RSv("メーカーコード") = RS("メーカーコード")
	RSv("商品コード") = RS("商品コード")
	RSv("色") = RS("色")
	RSv("規格") = RS("規格")
	RSv("受注数量") = RS("受注数量")

	RSv.Update
	
	RS.MoveNext
Loop

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
