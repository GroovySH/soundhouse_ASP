<%@ LANGUAGE="VBScript" %>
<%
'ネットハウスねっとハウスネットはうす
'サウンドハウス
 Option Explicit
%>
<!--#include file="../common/ADOVBS.inc"-->
<!--#include file="../common/system_common_econ.inc"-->
<!--#include file="../common/shop_common_functions.inc"-->
<!--#include file="../common/bfunctions1.asp"-->
<%
'========================================================================
'
'	eontext入金情報登録
'		eContextから呼び出される
'
'更新履歴
'2009/04/30 エラー時にerror.aspへ移動
'2011/02/12 ss 入金済みの場合イーコンへ -2 のエラー戻り値を返しているが、
'              正常値 1 を返しDB更新はしないよう変更
'2011/08/01 an #1087 Error.aspログ出力対応
'2012/01/30 GV 入金情報(受信データ)ログの出力位置 変更
'
'========================================================================

On Error Resume Next

Dim OrderID		'サイト注文番号
Dim ShopID		'サイトショップID
Dim ID				'データID　入金通知：0
Dim PayDate		'入金日 YY/MM/DD HH:MM:SS
Dim PayBy			'入金方法区分 0:現金 1:クレジットカード
Dim CvsCode		'コンビニ企業コード
Dim KssspCode	'コンビニ店舗コード
Dim InputID		'顧客が電話番号欄に入力した内容
Dim OrdAmount	'注文合計

Dim Connection
Dim RS

Dim wSQL
Dim wHTML
Dim wMSG
Dim wErrDesc   '2011/08/01 an add

Dim FS
Dim FS_Data
Dim DataFileName

'=======================================================================

'---- execute main process
call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "eContextPayment.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
'	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc)
end if                                           '2011/08/01 an add e

call close_db()

if Err.Description <> "" then
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

'---- eContextへメッセージ送信
if wMSG = "" then
	wHTML = "1 eContextPayment.asp 正常"
else
	if wMSG = "入金済み" then			'2011/02/12 ss add
		wHTML = "1 eContextPayment.asp 正常"	'2011/02/12 ss add
	else						'2011/02/12 ss add
		wHTML = "-2 eContextPayment.asp " & wMSG
	end if						'2011/02/12 ss add
end if

Response.write(wHTML)

'=======================================================================
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

'---- 受信データーの取り出し
OrderID = ReplaceInput(Trim(Request("OrderID")))
ShopID = ReplaceInput(Trim(Request("ShopID")))
ID = ReplaceInput(Trim(Request("ID")))
PayDate = ReplaceInput(Trim(Request("PayDate")))
PayBy = ReplaceInput(Trim(Request("PayBy")))
CvsCode = ReplaceInput(Trim(Request("CvsCode")))
KssspCode = ReplaceInput(Trim(Request("KssspCode")))
InputID = ReplaceInput(Trim(Request("InputID")))
OrdAmount = ReplaceInput(Trim(Request("OrdAmount")))

'---- 入力データーのチェック
call validate_data()

'---- eContext入金情報登録
if wMSG = "" then
	call InserteContextPayment()
end if

'---- 受信データ格納
Set FS = CreateObject("Scripting.FileSystemObject")
' 2012/01/30 GV Mod Start
'DataFileName = "eContextData/入金通知" & Year(Date()) & Right("0" & Month(Date()), 2) & Right("0" & Day(Date()), 2) & ".txt"
'DataFileName = Server.MapPath(DataFileName)		'Map log file
DataFileName = "入金通知" & Year(Date()) & Right("0" & Month(Date()), 2) & Right("0" & Day(Date()), 2) & ".txt"
DataFileName = g_LogRoot & g_eContextDataLog & DataFileName
' 2012/01/30 GV Mod End
Set FS_Data = FS.OpenTextFile(DataFileName, 8, true)			'File open - Append Mode

FS_Data.WriteLine(cf_FormatTime(Now(), "HH:MM:SS") & " OrderID=" & OrderID & ",ShopID=" & ShopID & ",ID=" & ID & ",PayDate=" & PayDate & ",PayBy=" & PayBy & ",CvsCode=" & CvsCode & ",KssspCode=" & KssspCode & ",InpitID=" & InputID & ",OrdAmount=" & OrdAmount & ",MSG=" & wMSG)

FS_Data.Close

End Function

'========================================================================
'
'	Function	入力データーのチェック
'
'========================================================================
'
Function validate_data()

'---- サイトショップID
if ShopID <> g_eContext_ID then
	wMSG = wMSG & "サイトコード不正 ShopID=" & ShopID & " "
end if

'---- データID　入金通知：0
if ID <> "0" then
	wMSG = wMSG & "データID不正 ID=" & ID & " "
end if

'---- OrderID
if OrderID = "" Or IsNumeric(OrderID) = false then
	wMSG = wMSG & "OrderIDなし "
end if

'---- PayDate
if IsDate(PayDate) = false then
	wMSG = wMSG & "PayDate不正 "
end if

'---- OrdAmount
if IsNumeric(OrdAmount) = false then
	wMSG = wMSG & "OrdAmount不正 "
end if

if wMSG <> "" then
	exit function
end if

'---- 受注番号
wSQL = ""
wSQL = wSQL & "SELECT 受注合計金額"
wSQL = wSQL & "  FROM Web受注 WITH (NOLOCK)"
wSQL = wSQL & " WHERE 受注番号 = " & OrderID

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RS.EOF = true then
	wMSG = wMSG & "該当受注情報なし "
else
	'---- 受注合計
	if isNumeric(OrdAmount) = false then
			wMSG = wMSG & "受注合計不正 OrdAmount=" & OrdAmount & " "
	else
		if Clng(OrdAmount) <> RS("受注合計金額") then
			wMSG = wMSG & "受注合計不一致 OrdAmount=" & OrdAmount & " 受注金額=" & RS("受注合計金額") & " "
		end if
	end if
end if

RS.Close

'2011/02/12 ss add ↓
if wMSG <> "" then
	exit function
end if
'2011/02/12 ss add ↑

'---- 入金済みチェック
wSQL = ""
wSQL = wSQL & "SELECT 入金日"
wSQL = wSQL & "  FROM eContext入金"
wSQL = wSQL & " WHERE 受注番号 = " & OrderID

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RS.EOF = false then
'		wMSG = wMSG & "入金済み 入金日=" & cf_FormatDate(RS("入金日"), "YYYY/MM/DD") & " " & cf_FormatTime(RS("入金日"), "HH:MM:SS") & " "	'2011/02/12 ss del
		wMSG = wMSG & "入金済み"	'2011/02/12 ss add
end if

RS.Close

End Function

'========================================================================
'
'	Function	eContext入金情報登録
'
'========================================================================
'
Function InserteContextPayment()

'----
wSQL = ""
wSQL = wSQL & "SELECT *"
wSQL = wSQL & "  FROM eContext入金"
wSQL = wSQL & " WHERE 1 = 2"

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic, adLockOptimistic

RS.AddNew

RS("受注番号") = OrderID
RS("入金日") = PayDate
RS("eContext入金区分") = Left(PayBy, 1)
RS("企業コード") = Left(CvsCode, 10)
RS("コンビニ店舗コード") = Left(KssspCode, 20)
RS("顧客入力電話番号") = Left(InputID, 20)
RS("入金金額") = OrdAmount
RS("入金受信日") = Now()

RS.Update
RS.Close

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

%>
