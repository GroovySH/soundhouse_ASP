<%@ LANGUAGE="VBScript" %>
<%
'ネットハウスねっとハウスネットはうす
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
'	カードオーダー画面型3Dセキュア/オーソリ 受信 (BlueGate)
'
'		オーソリ番号を受信
'		OrderCard3DAuthSendBG.asp からの戻り
'
'------------------------------------------------------------------------
'	更新履歴
'2008/04/17 作成
'2008/05/14 HTTPSチェック対応
'2009/04/30 エラー時にerror.aspへ移動
'
'========================================================================

On Error Resume Next

Dim w_sessionID
Dim userID
Dim msg

Dim ModeCode	    '電文種別
Dim SID           '加盟店自由域
Dim OrderNo       '注文番号
Dim ApprovalCode  '承認番号
Dim AcqCode       '被仕向会社
Dim TotalAmount   '決済金額合計
Dim ReceiveDateTime '受信日時
Dim PaymentDate   '決済日時
Dim MsgDigest     'MsgDigest
Dim ErrCode       'エラーコード

Dim ResultDigest     'ResultDigest

Dim wSQL
Dim wHTML
Dim wMSG
Dim wNextURL

Dim Connection
Dim RS_order_header

'=======================================================================

w_sessionID = Session.SessionId
userID = Session("UserID")

Session("msg") = ""
wMSG = ""

'---- 受け取り情報取り込み
ModeCode	    = Request("ModeCode")      '電文種別
SID           = ReplaceInput(Request("SID"))           '加盟店自由域
OrderNo       = ReplaceInput(Request("OrderNum"))      '注文番号
ApprovalCode  = ReplaceInput(Request("ApprovalCode"))  '承認番号
AcqCode       = Request("AcqCode")       '被仕向会社
TotalAmount   = Request("TotalAmount")   '決済金額合計
ReceiveDateTime = Request("ReceiveDateTime")  '受信日時
PaymentDate   = Request("PaymentDate")   '決済日時
MsgDigest     = Request("MsgDigest" )    'MsgDigest
ErrCode       = ReplaceInput(Request("ErrCode"))       'エラーコード

'---- execute main process
call ConnectDB()
call main()
call close_db()

if Err.Description <> "" then	
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

if wMSG = "" then
	Response.Redirect "OrderSubmit.asp?OrderNo=" & OrderNo
else
	Session("msg") = wMSG
	Response.Redirect "OrderInfoEnter.asp?CardErrorCd=" & ErrCode
end if

'========================================================================
'========================================================================
'
'	Function	Connect database
'
'========================================================================
'
Function ConnectDB()

'---- Connect database
Set Connection = Server.CreateObject("ADODB.Connection")
Connection.Open g_connection

End function

'========================================================================
'
'	Function	Main 確認用 ダイジェスト作成
'
'========================================================================
'
Function main()

Dim ObjBG
Dim vRetCode

if ErrCode <> "00000000" then
	wMSG = "CardError1"
	exit function
end if

'---- 3DResponseMDCreatorメソッドコール
'Set ObjBG = Server.CreateObject("Aspcompg.aspcom")
Set ObjBG = Server.CreateObject("Memst.MemberStore.1")

ResultDigest = ObjBG.GenerateAuthoriResultMd(g_BlueGate_ID, g_BlueGate_PW, OrderNo, ApprovalCode, ErrCode, AcqCode, TotalAmount, ReceiveDateTime, PaymentDate)

If ResultDigest = "" Then
	wMSG = "CardError1"
end if

call updateOrderHeader()

Set ObjBG = Nothing

end function

'========================================================================
'
'	Function	仮受注情報の更新
'
'========================================================================
'
Function updateOrderHeader()

'---- 仮受注取り出し
wSQL = ""
wSQL = wSQL & "SELECT a.カード与信確認番号"
wSQL = wSQL & "  FROM 仮受注 a"
wSQL = wSQL & " WHERE SessionID = '" & w_sessionID & "'"
	  
Set RS_order_header = Server.CreateObject("ADODB.Recordset")
RS_order_header.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RS_order_header.EOF = true then
	wMSG = "<font color='#ff0000'>注文情報がありません</font>"
	exit function
end if

'---- update 仮受注
RS_order_header("カード与信確認番号")   = ApprovalCode

RS_order_header.update
RS_order_header.close

End function

'========================================================================
'
'	Function	Close database
'
'========================================================================
'
Function close_db()

Connection.close

End function

%>
