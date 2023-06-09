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
'	カードオーダー3Dセキュア結果受け取り処理 (BlueGate)
'
'		カードの3Dセキュアチェックの結果をBlueGateより受け取る。
'		OKなら、OrderCardAuthBG.aspを呼び出し、オーソリを取る。
'
'------------------------------------------------------------------------
'	更新履歴
'2006/09/21 BlueGateアクセスログ追加
'2006/11/03 BlueGateアクセスログ中止
'2006/11/06 BlueGateアクセスログ復活(Open時間を短く)
'2007/02/12 オーソリエラー時の説明ページリンク先を変更
'2007/04/11 OrderCardAuthBG.asp呼び出しパラメータにXstatusを追加
'2007/04/16 BlueGateアクセスログ中止
'2007/08/14 カードエラー時のメッセージ変更
'2009/04/30 エラー時にerror.aspへ移動
'
'========================================================================

On Error Resume Next

Dim w_sessionID
Dim userID
Dim msg

Dim ModeCode	       '電文種別
Dim SID              '加盟店自由域
Dim OrderNo          '注文番号
Dim MsgVerNum        'version
Dim XID              'xid
Dim Xstatus          'status
Dim ECI              'eci
Dim CAVV             'cavv
Dim CavvAlgorithm    'cavvAlgorithm
Dim MsgDigest        'MsgDigest
Dim ErrCode          'エラーコード
Dim ResultDigest     'ResultDigest

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

'---- 受け取り情報取り込み
ModeCode	    = Request("ModeCode")      '電文種別
SID           = Request("SID")           '加盟店自由域
OrderNo       = Request("OrderNum")      '注文番号
MsgVerNum     = Request("MsgVerNum")     '3D version
XID           = Request("XID")           '3D xid
Xstatus       = Request("Xstatus")       '3D status
ECI           = Request("ECI")           '3D eci
CAVV          = Request("CAVV")          '3D cavv
CavvAlgorithm = Request("CavvAlgorithm") '3D cavvAlgorithm
MsgDigest     = Request("MsgDigest" )    '3D MsgDigest
ErrCode       = Request("ErrCode")       'エラーコード

'---- execute main process
call main()

if Err.Description <> "" then	
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

'---- エラーが無いときはオーソリ取得、エラーがあれば確認ページへ
if wMSG = "" then
	Response.Redirect ("OrderCardAuthBG.asp" _
							  & "?OrderNo="        & Server.URLEncode(OrderNo) _
							  & "&MsgVerNum="      & Server.URLEncode(MsgVerNum) _
							  & "&XID="            & Server.URLEncode(XID) _
							  & "&Xstatus="        & Server.URLEncode(Xstatus) _
							  & "&ECI="            & Server.URLEncode(ECI) _
							  & "&CAVV="           & Server.URLEncode(CAVV) _
							  & "&CavvAlgorithm="  & Server.URLEncode(CavvAlgorithm) _
						)
else
	Session("msg") = wMSG
	Response.Redirect "OrderInfoEnter.asp?CardErrorCd=" & ErrCode
end if

'========================================================================
'========================================================================
'
'	Function	Main 3Dセキュア ダイジェスト作成
'
'========================================================================
'
Function main()

Dim ObjBG
Dim vRetCode

'---- BlueGate Log open
'Set FS = CreateObject("Scripting.FileSystemObject")
'LogFileName = "BlueGateLog/BlueGateLog" & Year(Date()) & Right("0" & Month(Date()), 2) & Right("0" & Day(Date()), 2) & ".txt"
'LogFileName = Server.MapPath(LogFileName)		'Map log file

'---- Log after 3d return
'Set FS_Log = FS.OpenTextFile(LogFileName, 8, true)			'Log open - Append Mode
'FS_Log.WriteLine(cf_FormatTime(Now(), "HH:MM:SS") & " OrderCard3dResponseBG.asp Redirect from 3D secure    RETURN OrderNo=" & OrderNo & " ErrCode=" & ErrCode)
'FS_Log.Close											'Log close

'---- エラーチェック
call checkError()
if wMsg <> "" then
	exit function
end if

Session("BlueGate3DReturnCode") = ErrCode

'---- 3DResponseMDCreatorメソッドコール
Set ObjBG = Server.CreateObject("Aspcompg.aspcom")

'---- Log before
'Set FS_Log = FS.OpenTextFile(LogFileName, 8, true)			'Log open - Append Mode
'FS_Log.WriteLine(cf_FormatTime(Now(), "HH:MM:SS") & " OrderCard3dResponseBG.asp ComThreeDResponseMDCreator BEFORE OrderNo=" & OrderNo)
'FS_Log.Close											'Log close

vRetCode = ObjBG.ComThreeDResponseMDCreator(g_BlueGate_ID, g_BlueGate_PW, OrderNo, MsgVerNum, XID, Xstatus, ECI, CAVV, CavvAlgorithm )

'----プロパティを設定
ResultDigest = ObjBG.ComGetPropValue("MsgDigest") '結果ダイジェスト
ErrCode      = ObjBG.ComGetPropValue("ErrCode")   'エラーコード

'---- Log after
'Set FS_Log = FS.OpenTextFile(LogFileName, 8, true)			'Log open - Append Mode
'FS_Log.WriteLine(cf_FormatTime(Now(), "HH:MM:SS") & " OrderCard3dResponseBG.asp ComThreeDResponseMDCreator AFTER  OrderNo=" & OrderNo & " ErrCode=" & ErrCode)
'FS_Log.Close											'Log close

'---- 3Dダイジェストエラー
if (ErrCode <> "00000000") OR (MsgDigest <> ResultDigest) then
	wMSG = "CardError1"
end if

Set ObjBG = Nothing

end function

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
vNoError = "00000000,S102000W"		'S102000W:3DSecureサービス対象外

'---- 3D OK
if InStr(vNoError, ErrCode) > 0 then
	wMSG = ""
	exit function
end if

'---- その他カードエラー
wMSG = "CardError1"

End function

%>
