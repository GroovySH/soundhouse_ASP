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
'	カードオーダー与信確認処理 (BlueGate)
'
'		カードの与信を取りOkならorder_submitへコントロールを渡す。
'		与信OKなら、受注番号の採番も行う。
'
'------------------------------------------------------------------------
'	更新履歴
'2006/07/19 3D用　オーダー番号は3Dプログラムから渡されるため採番不要
'2006/09/21 BlueGateアクセスログ追加
'2006/11/03 BlueGateアクセスログ中止
'2006/11/06 BlueGateアクセスログ復活（Open時間を短く)
'2007/02/12 オーソリエラー時の説明ページリンク先を変更
'2007/04/11 3Dパラメータをオーソリパラメータに追加
'2007/04/16 BlueGateアクセスログ中止
'2007/04/30 BlueGate3DEC用Log採取開始
'2007/05/30 BlueGate3DEC用Log採取中止、ECIを受注情報として取り込み
'2007/08/14 カードエラー時のメッセージ変更
'2009/04/30 エラー時にerror.aspへ移動
'
'========================================================================

On Error Resume Next

Dim w_sessionID
Dim userID
Dim msg

Dim InShopId
Dim InShopPw
Dim InOrderNum
Dim InAmount
Dim IntaxAndDeliCharge
Dim InPan
Dim InExpiryDate
Dim InPaymentMode
Dim InStartPayMonth
Dim InPaymentCount
Dim InInitialAmount
Dim InBonusMonth
Dim InBonusAmount
Dim InBonusCount
Dim InMsgVerNum
Dim InXid
Dim InXStatus
Dim InEci
Dim Incavv
Dim InCavvAlgorithm

Dim CardNo
Dim CardExpDt
Dim CardExpDt1
Dim CardExpDt2
Dim CardHolderName
Dim OrderTotalAm
Dim OrderNo
Dim CardAuthNo

Dim ApprovalCode
Dim ErrCode
Dim AcqCode
Dim TotalAmount
Dim ReceiveDateTime
Dim PaymentDate
Dim DetailCode

Dim Connection
Dim RS_OrderHeader

Dim Auth3DKubun

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

if Session("BlueGate3DReturnCode") = "00000000" then
	Auth3DKubun = "BlueGate3D"
else
	Auth3DKubun = "BlueGate"
end if

OrderNo = Request("OrderNo")
InMsgVerNum = Request("MsgVerNum")
InXid = Request("XID")
InXStatus = Request("XStatus")
InEci = Request("ECI")
Incavv = Request("CAVV")
InCavvAlgorithm = Request("CavvAlgorithm")

'---- execute main process
call ConnectDB()
call main()
call close_db()

if Err.Description <> "" then	
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

'---- エラーが無いときは注文登録処理ページ、エラーがあれば確認ページへ
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
'	Function	Main カード与信確認
'
'========================================================================
'
Function main()

'---- カード情報取り出し
call getCard()

if wMSG <> "" then
	exit function
end if

'---- 与信チェック
call getCardAuth()

'---- 受注情報に与信確認番号をセット
if wMSG = "" then
	call updateOrderHeader()
end if

RS_OrderHeader.close

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
wSQL = wSQL & "     , a.カードネット伝票番号"
wSQL = wSQL & "     , a.BlueGateECI"
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
CardExpDt1 = Left(CardExpDt, 2)
CardExpDt2 = Right(CardExpDt, 2)
CardHolderName = RS_OrderHeader("カード名義人")
OrderTotalAm = RS_OrderHeader("受注合計金額")

End function

'========================================================================
'
'	Function	カード与信確認
'
'========================================================================
'
Function getCardAuth()

Dim ObjBG

Dim vRetCode

'---- BlueGate Log
Set FS = CreateObject("Scripting.FileSystemObject")
LogFileName = "BlueGateLog/BlueGateLog" & Year(Date()) & Right("0" & Month(Date()), 2) & Right("0" & Day(Date()), 2) & ".txt"
LogFileName = Server.MapPath(LogFileName)		'Map log file

'---- パラメータのセット
InShopId           = g_BlueGate_ID             'ショップID
InShopPw           = g_BlueGate_PW             'ショップパスワード

if OrderNo = "" then
	OrderNo            = GetOrderNo()              '注文番号		'3D時は不要
end if

InAmount           = OrderTotalAm              '売上金額
IntaxAndDeliCharge = 0                         '税送料
InPan              = CardNo                    'カード番号
InExpiryDate       = CardExpDt1 & CardExpDt2   '有効期限
InPaymentMode      = "10"                      '支払区分(一括)

'---- オーソリ取得
Set ObjBG = Server.CreateObject("Aspcompg.aspcom")

'---- Log before
'Set FS_Log = FS.OpenTextFile(LogFileName, 8, true)			'Log open - Append Mode
'FS_Log.WriteLine(cf_FormatTime(Now(), "HH:MM:SS") & " OrderCardAuthBG.asp       ComAuthoriRequest          BEFORE OrderNo=" & OrderNo)
'FS_Log.Close											'Log close

vRetCode = ObjBG.ComAuthoriRequest(InShopId, InShopPw, OrderNo, InAmount, IntaxAndDeliCharge, InPan, InExpiryDate, InPaymentMode, InStartPayMonth, InPaymentCount, InInitialAmount, InBonusMonth, InBonusAmount, InBonusCount, InMsgVerNum, InXid, InXStatus, InEci, InCavv, InCavvAlgorithm )

'---- プロパティを設定
ApprovalCode    = ObjBG.ComGetPropValue("ApprovalCode")      '承認番号
ErrCode         = ObjBG.ComGetPropValue("ErrCode")           'エラーコード
AcqCode         = ObjBG.ComGetPropValue("AcqCode")           '被仕向会社
TotalAmount     = ObjBG.ComGetPropValue("TotalAmount")       '決済金額
ReceiveDateTime = ObjBG.ComGetPropValue("ReceiveDateTime")   '受付日時
PaymentDate     = ObjBG.ComGetPropValue("PaymentDate")       '決済処理日付
DetailCode      = ObjBG.ComGetPropValue("DetailCode")        '詳細コード

'---- Log after
'Set FS_Log = FS.OpenTextFile(LogFileName, 8, true)			'Log open - Append Mode
'FS_Log.WriteLine(cf_FormatTime(Now(), "HH:MM:SS") & " OrderCardAuthBG.asp ComAuthoriRequest AFTER  OrderNo=" & OrderNo & " CardNo=" & InPan & " ApprovalCode=" & ApprovalCode & " ECI=" & InECI & " ErrCode=" & ErrCode)
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
'	Function	仮受注情報の更新
'
'========================================================================
'
Function updateOrderHeader()

'---- update 仮受注
RS_OrderHeader("カード与信確認番号")   = ApprovalCode
RS_OrderHeader("カードネット伝票番号") = Auth3DKubun
RS_OrderHeader("BlueGateECI") = InEci

RS_OrderHeader.update

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
'---- カード番号または有効期限エラー
vCardDataError = "S5001060,S5001061,S5001062,S5001069,S5001070,S5001072,S5001079"

'---- オーソリOK
if InStr(vNoError, ErrCode) > 0 then
	wMSG = ""
	exit function
end if

'---- カードエラー
if InStr(vCardDataError, ErrCode) > 0 then
	wMSG = "CardError1"
	exit function
end if

'---- その他カードエラー
wMSG = "CardError2"

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
