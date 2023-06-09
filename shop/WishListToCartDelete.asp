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
'	ウィッシュリストをカートへ移動/ウィッシュリストを削除
'
'更新履歴
'2009/04/30 エラー時にerror.aspへ移動
'2009/09/09	カートへ入れるときに、ウィッシュリストから削除するかどうか
'2011/08/01 an #1087 Error.aspログ出力対応
'
'========================================================================

On Error Resume Next

Dim userID

Dim Kubun
Dim qt
Dim Item
Dim ItemCnt
Dim ItemList()
Dim MakerCd
Dim ProductCd
Dim Iro
Dim Kikaku
Dim DeleteFl

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
Kubun = ReplaceInput(Request("Kubun"))
Item = ReplaceInput(Request("Item"))
qt = ReplaceInput(Request("qt"))
DeleteFl = ReplaceInput(Request("DeleteFl"))

if Item <> "" then
	ItemCnt = cf_unstring(Item, ItemList, "^")
	MakerCd = ItemList(0)
	ProductCd = ItemList(1)
	Iro = ItemList(2)
	Kikaku = ItemList(3)
end if

'---- Execute main
call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "WishListToCartDelete.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

if Err.Description <> "" then	
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

call close_db()

if Kubun = "Cart" then
	Response.Redirect "OrderPreInsert.asp?Item=" & Item & "&qt=" & qt
else
	Response.Redirect "WishList.asp"
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

Dim RSv

if DeleteFl <> "Y" then
	exit function
end if

'---- ウィッシュリストから削除
wSQL = ""
wSQL = wSQL & "SELECT *"
wSQL = wSQL & "  FROM ウィッシュリスト"
wSQL = wSQL & " WHERE 顧客番号 = " & userID
wSQL = wSQL & "   AND メーカーコード = '" & MakerCd & "'"
wSQL = wSQL & "   AND 商品コード = '" & ProductCd & "'"
wSQL = wSQL & "   AND 色 = '" & Iro & "'"
wSQL = wSQL & "   AND 規格 = '" & Kikaku & "'"

'@@@@response.write(wSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RSv.EOF = false then
	RSv.Delete
end if

RSv.Update
RSv.close

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
