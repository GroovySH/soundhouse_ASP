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
'	ウィッシュリストへ追加
'
'更新履歴
'2007/08/23 商品アクセス件数登録（ウィッシュリスト）
'2007/09/10 商品アクセス件数登録（ウィッシュリスト）を月別に変更
'2008/05/23 入力データチェック強化（LEFT, Numeric, EOF他)
'2009/04/30 エラー時にerror.aspへ移動
'2011/04/14 hn SessionID関連変更
'2011/08/01 an #1087 Error.aspログ出力対応
'
'========================================================================

On Error Resume Next

Dim userID

Dim OrderDetailNo
Dim Item
Dim ItemCnt
Dim ItemList()
Dim MakerCd
Dim ProductCd
Dim Iro
Dim Kikaku

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
OrderDetailNo = ReplaceInput(Request("OrderDetailNo"))
Item = ReplaceInput(Request("Item"))

if Item <> "" then
	ItemCnt = cf_unstring(Item, ItemList, "^")
	MakerCd = Left(ItemList(0), 8)
	ProductCd = Left(ItemList(1), 20)
	Iro = Left(ItemList(2), 20)
	Kikaku = Left(ItemList(3), 20)
end if

'---- Execute main
call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "WishListAdd.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

call close_db()

if Err.Description <> "" then	
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

Response.Redirect "WishList.asp?msg=" & wMSG

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
Dim vWishListAdded
Dim vYYYYMM

'---- ウィッシュリスト登録
wSQL = ""
wSQL = wSQL & "SELECT *"
wSQL = wSQL & "  FROM ウィッシュリスト"
wSQL = wSQL & " WHERE 顧客番号 = " & userID
wSQL = wSQL & "   AND メーカーコード = '" & MakerCd & "'"
wSQL = wSQL & "   AND 商品コード = '" & ProductCd & "'"
wSQL = wSQL & "   AND 色 = '" & Iro & "'"
wSQL = wSQL & "   AND 規格 = '" & Kikaku & "'"
	  
Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RSv.EOF = true then
	RSv.AddNew

	RSv("顧客番号") = userID
	RSv("メーカーコード") = MakerCd
	RSv("商品コード") = ProductCd
	RSv("色") = Iro
	RSv("規格") = Kikaku

	vWishListAdded = "Y"
end if

RSv("登録日") = now()

RSv.Update
RSv.close

'---- 仮受注明細削除
if OrderDetailNo <> "" and isNumeric(OrderDetailNo) = true then
	wSQL = ""
	wSQL = wSQL & "SELECT *"
	wSQL = wSQL & "  FROM 仮受注明細"
	wSQL = wSQL & " WHERE SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod
	wSQL = wSQL & "   AND 受注明細番号 = " & OrderDetailNo
		  
	Set RSv = Server.CreateObject("ADODB.Recordset")
	RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

	if RSv.EOF = false then
		RSv.Delete
	end if

	RSv.close
end if

'---- 商品アクセス件数登録（ウィッシュリスト）
if vWishListAdded = "Y" then
	vYYYYMM = Year(Now()) & Right("0" & Month(Now()),2)
	wSQL = ""
	wSQL = wSQL & "SELECT *"
	wSQL = wSQL & "  FROM 商品アクセス件数"
	wSQL = wSQL & " WHERE メーカーコード = '" & MakerCd & "'"
	wSQL = wSQL & "   AND 商品コード = '" & ProductCd & "'"
	wSQL = wSQL & "   AND 年月 = '" & vYYYYMM & "'"
		  
	Set RSv = Server.CreateObject("ADODB.Recordset")
	RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

	if RSv.EOF = true then
		RSv.AddNew

		RSv("メーカーコード") = MakerCd
		RSv("商品コード") = ProductCd
		RSv("年月") = vYYYYMM
		RSv("ウィッシュリスト件数") = 1
	else
		RSv("ウィッシュリスト件数") = RSv("ウィッシュリスト件数") + 1
	end if

	RSv.Update
	RSv.close
end if

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
