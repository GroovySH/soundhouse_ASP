<%@ LANGUAGE="VBScript" %>
<%
 Option Explicit
%>
<!--#include file="../common/ADOVBS.inc"-->
<!--#include file="../common/system_common.inc"-->
<!--#include file="../common/shop_common_functions.inc"-->
<!--#include file="../common/bfunctions1.asp"-->
<%
'========================================================================
'
'	カードオーダー3Dセキュアリクエスト処理 (BlueGate)
'
'		カードの3Dセキュアチェックをリダイレクトでリクエストする。
'		BlueGate 3DSecure からの戻りは、OrderCard3DResponseBG.asp
'
'------------------------------------------------------------------------
'	更新履歴
'2006/09/21 BlueGateアクセスログ追加
'2006/11/03 BlueGateアクセスログ中止
'2006/11/06 BlueGateアクセスログ復活(Open時間を短く)
'2007/02/12 オーソリエラー時の説明ページリンク先を変更
'2007/04/16 BlueGateアクセスログ中止
'2009/04/30 エラー時にerror.aspへ移動
'
'========================================================================

On Error Resume Next

Dim w_sessionID
Dim userID
Dim msg

Dim CardNo
Dim CardExpDt
Dim CardHolderName
Dim OrderTotalAm
Dim OrderTaxShipping
Dim OrderNo

Dim ThreeDDigest
Dim ErrCode

Dim Connection
Dim RS_OrderHeader

Dim wSQL
Dim wHTML
Dim wMSG
Dim wNextURL

Dim FS
Dim FS_Log
Dim LogFileName

'=======================================================================

w_sessionID = Session.SessionId
userID = Session("UserID")

Session("msg") = ""
wMSG = ""

'---- execute main process
call ConnectDB()
call main()
call close_db()

if Err.Description <> "" then	
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

'---- エラーが無いときは3DセキュアBlueGate、エラーがあれば確認ページへ
if wMSG = "" then
	Response.Redirect (g_BlueGate_3DURL _
							  & "?ModeCode=0081" _
							  & "&ShopID="            & Server.URLEncode(g_BlueGate_ID) _
							  & "&OrderNum="          & Server.URLEncode(OrderNo) _
							  & "&Amount="            & Server.URLEncode(OrderTotalAm) _
							  & "&TaxAndDeliCharge="  & Server.URLEncode(OrderTaxShipping) _
							  & "&OrderInfoNum=" _
							  & "&PAN="               & Server.URLEncode(CardNo) _
							  & "&ExpiryDate="        & Server.URLEncode(CardExpDt) _
							  & "&TermURL="           & Server.URLEncode(g_HTTPS & "shop/OrderCard3DResponseBG.asp") _
							  & "&MsgDigest="         & Server.URLEncode(ThreeDDigest) _
							  & "&OptionalAreaName=SID" _
							  & "&OptionalAreaValue=" & Server.URLEncode(w_sessionID) _
						)
else
	Session("msg") = wMSG
	Response.Redirect "OrderInfoEnter.asp"
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
'	Function	Main 3Dセキュア ダイジェスト作成
'
'========================================================================
'
Function main()

'---- カード情報取り出し
call getCard()

if wMSG <> "" then
	exit function
end if

'---- 3Dセキュア ダイジェスト作成
call get3DDigest()

End Function

'========================================================================
'
'	Function	カード情報取り出し
'
'========================================================================
'
Function GetCard()

'---- 仮受注取り出し
wSQL = ""
wSQL = wSQL & "SELECT a.カード番号"
wSQL = wSQL & "     , a.カード有効期限"
wSQL = wSQL & "     , a.カード名義人"
wSQL = wSQL & "     , a.受注合計金額"
wSQL = wSQL & "     , a.カード与信確認番号"
wSQL = wSQL & "  FROM 仮受注 a"
wSQL = wSQL & " WHERE SessionID = '" & w_sessionID & "'"
	  
Set RS_OrderHeader = Server.CreateObject("ADODB.Recordset")
RS_OrderHeader.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RS_OrderHeader.EOF = true then
	wMSG = "<font color='#ff0000'>NoData</font>"
	exit function
end if

CardNo = RS_OrderHeader("カード番号")
CardExpDt = RS_OrderHeader("カード有効期限")
CardExpDt = Left(CardExpDt, 2) & Right(CardExpDt, 2)	'MMYY
CardHolderName = RS_OrderHeader("カード名義人")
OrderTotalAm = RS_OrderHeader("受注合計金額")

RS_OrderHeader.close

End function

'========================================================================
'
'	Function	3Dセキュア ダイジェスト作成
'
'========================================================================
'
Function get3DDigest()

Dim ObjBG
Dim vRetCode

'---- BlueGate Log open
'Set FS = CreateObject("Scripting.FileSystemObject")
'LogFileName = "BlueGateLog/BlueGateLog" & Year(Date()) & Right("0" & Month(Date()), 2) & Right("0" & Day(Date()), 2) & ".txt"
'LogFileName = Server.MapPath(LogFileName)		'Map log file

'---- パラメータのセット
OrderNo          = GetOrderNo()              '注文番号
OrderTaxShipping = 0                         '税送料

'---- 3DRequestMDCreatorメソッドコール
Set ObjBG = Server.CreateObject("Aspcompg.aspcom")

'---- Log before
'Set FS_Log = FS.OpenTextFile(LogFileName, 8, true)			'Log open - Append Mode
'FS_Log.WriteLine(cf_FormatTime(Now(), "HH:MM:SS") & " OrderCard3dSecureBG.asp   ComThreeDRequestMDCreator  BEFORE OrderNo=" & OrderNo)
'FS_Log.Close											'Log close

vRetCode = ObjBG.ComThreeDRequestMDCreator(g_BlueGate_ID, g_BlueGate_PW, OrderNo, OrderTotalAm, OrderTaxShipping, CardNo, CardExpDt)

'----プロパティを設定
ThreeDDigest = ObjBG.ComGetPropValue("MsgDigest") '３Ｄメッセージダイジェスト
ErrCode      = ObjBG.ComGetPropValue("ErrCode")   'エラーコード

'---- Log after
'Set FS_Log = FS.OpenTextFile(LogFileName, 8, true)			'Log open - Append Mode
'FS_Log.WriteLine(cf_FormatTime(Now(), "HH:MM:SS") & " OrderCard3dSecureBG.asp   ComThreeDRequestMDCreator  AFTER  OrderNo=" & OrderNo & " ErrCode=" & ErrCode & " SessionID=" & w_sessionID)
'FS_Log.Close											'Log close

Set ObjBG = Nothing

'---- エラーチェック
call checkError()

end function

'========================================================================
'
'	Function	受注番号取り出し
'
'========================================================================
'
Function GetOrderNo()

Dim vRS_Cntl

'---- コントロールマスタ取り出し
wSQL = ""
wSQL = wSQL & "SELECT item_num1"
wSQL = wSQL & "  FROM コントロールマスタ"
wSQL = wSQL & " WHERE sub_system_cd = '共通'"
wSQL = wSQL & "   AND item_cd = '番号'"
wSQL = wSQL & "   AND item_sub_cd = 'Web受注'"
	  
Set vRS_Cntl = Server.CreateObject("ADODB.Recordset")
vRS_Cntl.Open wSQL, Connection, adOpenStatic, adLockOptimistic

vRS_Cntl("item_num1") = Clng(vRS_Cntl("item_num1")) + 1
GetOrderNo = vRS_Cntl("item_num1")

vRS_Cntl.update
vRS_Cntl.close

End function

'========================================================================
'
'	Function	カードエラーチェック
'
'========================================================================
'
Function checkError()

Dim vNoError
Dim vCardDataError

'---- リターンコード設定
'---- 正常
vNoError = "00000000"

'---- 3D OK
if InStr(vNoError, ErrCode) > 0 then
	wMSG = ""
	exit function
end if

'---- その他カードエラー
wMSG = "<font color='#ff0000'>" _
			& "申し訳ございませんが､御指定のカードでは御注文できません。<br>" _
			& "別のカードまたは､別のお支払方法で御注文願います。<br>" _
			& "Code: " & ErrCode & " (OrderCard3DSecureBG)<br>" _
			& "よくあるご質問は<a href='" & G_HTTP & "guide/qanda8.asp'>こちら</a>" _
			& "</font>"

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
