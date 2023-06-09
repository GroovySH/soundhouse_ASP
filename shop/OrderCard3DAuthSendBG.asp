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
'	カードオーダー画面型3Dセキュア/オーソリリクエスト処理 (BlueGate)
'
'		カード入力と3Dセキュア、オーソリをリクエストする。
'		BlueGate からの戻りは、OrderCard3DAuthReceiveBG.asp
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

Dim OrderTotalAm
Dim OrderTaxShipping
Dim OrderNo

Dim MsgDigest
Dim ErrCode

Dim Connection

Dim wSQL
Dim wHTML
Dim wMSG

'=======================================================================

Response.Buffer = true

w_sessionID = Session.SessionId
userID = Session("UserID")

Session("msg") = ""
wMSG = ""

OrderTotalAm = Session("受注合計金額")

'---- execute main process
call ConnectDB()
call main()
call close_db()

if Err.Description <> "" then	
	Response.Redirect g_HTTP & "shop/Error.asp"
end if


if wMSG <> "" then
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

Dim ObjBG
Dim vMsgDigest

'---- 受注番号生成
OrderNo = GetOrderNo()

'---- メッセージダイジェストを作成します。
'Set ObjBG = Server.CreateObject("Aspcompg.aspcom")
Set ObjBG = Server.CreateObject("Memst.MemberStore.1")

MsgDigest = ObjBG.GenerateOrderReceptionMD(g_BlueGate_ID, g_BlueGate_PW, OrderNo, OrderTotalAm, OrderTaxShipping)

If MsgDigest = "" Then
	ErrCode = ObjBG.GetErrCode()
'---- その他カードエラー
wMSG = "<font color='#ff0000'>" _
			& "申し訳ございませんが､御指定のカードでは御注文できません。<br>" _
			& "別のカードまたは､別のお支払方法で御注文願います。<br>" _
			& "Code: " & ErrCode & " (OrderCard3DAuthSendBG)<br>" _
			& "よくあるご質問は<a href='" & G_HTTP & "guide/qanda8.asp'>こちら</a>" _
			& "</font>"
End If

Set ObjBG = Nothing

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
'	Function	Close database
'
'========================================================================
'
Function close_db()

Connection.close

End function

'========================================================================
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<title>BlueGate 3D、オーソリリクエスト（BlueGate画面型PAN入力決済要求電文）</title>
</head>

<body>
<form action="<%=g_BlueGate_3DURL %>" method="POST" name="f_data" ENCTYPE="application/x-www-form-urlencode">
<input type="hidden" name="ModeCode" value="0071">													<!-- 電文種別 -->
<input type="hidden" name="ShopID" value="<%=g_BlueGate_ID%>">							<!-- ショップID -->
<input type="hidden" name="OrderNum" value="<%=OrderNo%>">									<!-- 注文番号 -->
<input type="hidden" name="Amount" value="<%=OrderTotalAm%>">								<!-- 売上金額 -->
<input type="hidden" name="TaxAndDeliCharge" value="<%=OrderTaxShipping%>">	<!-- 税送料 -->
<input type="hidden" name="TermURL" value="<%=g_HTTPS & "shop/OrderCard3DAuthReceiveBG.asp"%>">	<!-- 戻り先URL -->
<input type="hidden" name="LANG" value="J">																	<!-- 言語 -->
<input type="hidden" name="MsgDigest" value="<%=MsgDigest%>">								<!-- メッセージダイジェスト -->
<input type="hidden" name="OptionalAreaName" value="SID">										<!-- 自由領域名 -->
<input type="hidden" name="OptionalAreaValue" value="<%=Server.URLEncode(w_sessionID)%>">	<!-- 自由領域値 -->
</form>

</body>
</html>

<script language="JavaScript">

	document.f_data.submit();	//Redirect to BuleGate

</script>
