<%@ LANGUAGE="VBScript" %>
<%
'�l�b�g�n�E�X�˂��ƃn�E�X�l�b�g�͂���
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
'	�J�[�h�I�[�_�[��ʌ^3D�Z�L���A/�I�[�\�����N�G�X�g���� (BlueGate)
'
'		�J�[�h���͂�3D�Z�L���A�A�I�[�\�������N�G�X�g����B
'		BlueGate ����̖߂�́AOrderCard3DAuthReceiveBG.asp
'
'------------------------------------------------------------------------
'	�X�V����
'2008/04/17 �쐬
'2008/05/14 HTTPS�`�F�b�N�Ή�
'2009/04/30 �G���[����error.asp�ֈړ�
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

OrderTotalAm = Session("�󒍍��v���z")

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
'	Function	Main 3D�Z�L���A �_�C�W�F�X�g�쐬
'
'========================================================================
'
Function main()

Dim ObjBG
Dim vMsgDigest

'---- �󒍔ԍ�����
OrderNo = GetOrderNo()

'---- ���b�Z�[�W�_�C�W�F�X�g���쐬���܂��B
'Set ObjBG = Server.CreateObject("Aspcompg.aspcom")
Set ObjBG = Server.CreateObject("Memst.MemberStore.1")

MsgDigest = ObjBG.GenerateOrderReceptionMD(g_BlueGate_ID, g_BlueGate_PW, OrderNo, OrderTotalAm, OrderTaxShipping)

If MsgDigest = "" Then
	ErrCode = ObjBG.GetErrCode()
'---- ���̑��J�[�h�G���[
wMSG = "<font color='#ff0000'>" _
			& "�\���󂲂����܂��񂪤��w��̃J�[�h�ł͌䒍���ł��܂���B<br>" _
			& "�ʂ̃J�[�h�܂��ͤ�ʂ̂��x�����@�Ō䒍���肢�܂��B<br>" _
			& "Code: " & ErrCode & " (OrderCard3DAuthSendBG)<br>" _
			& "�悭���邲�����<a href='" & G_HTTP & "guide/qanda8.asp'>������</a>" _
			& "</font>"
End If

Set ObjBG = Nothing

end function

'========================================================================
'
'	Function	�󒍔ԍ����o��
'
'========================================================================
'
Function GetOrderNo()

Dim vRS_Cntl

'---- �R���g���[���}�X�^���o��
wSQL = ""
wSQL = wSQL & "SELECT item_num1"
wSQL = wSQL & "  FROM �R���g���[���}�X�^"
wSQL = wSQL & " WHERE sub_system_cd = '����'"
wSQL = wSQL & "   AND item_cd = '�ԍ�'"
wSQL = wSQL & "   AND item_sub_cd = 'Web��'"
	  
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
<title>BlueGate 3D�A�I�[�\�����N�G�X�g�iBlueGate��ʌ^PAN���͌��ϗv���d���j</title>
</head>

<body>
<form action="<%=g_BlueGate_3DURL %>" method="POST" name="f_data" ENCTYPE="application/x-www-form-urlencode">
<input type="hidden" name="ModeCode" value="0071">													<!-- �d����� -->
<input type="hidden" name="ShopID" value="<%=g_BlueGate_ID%>">							<!-- �V���b�vID -->
<input type="hidden" name="OrderNum" value="<%=OrderNo%>">									<!-- �����ԍ� -->
<input type="hidden" name="Amount" value="<%=OrderTotalAm%>">								<!-- ������z -->
<input type="hidden" name="TaxAndDeliCharge" value="<%=OrderTaxShipping%>">	<!-- �ő��� -->
<input type="hidden" name="TermURL" value="<%=g_HTTPS & "shop/OrderCard3DAuthReceiveBG.asp"%>">	<!-- �߂��URL -->
<input type="hidden" name="LANG" value="J">																	<!-- ���� -->
<input type="hidden" name="MsgDigest" value="<%=MsgDigest%>">								<!-- ���b�Z�[�W�_�C�W�F�X�g -->
<input type="hidden" name="OptionalAreaName" value="SID">										<!-- ���R�̈於 -->
<input type="hidden" name="OptionalAreaValue" value="<%=Server.URLEncode(w_sessionID)%>">	<!-- ���R�̈�l -->
</form>

</body>
</html>

<script language="JavaScript">

	document.f_data.submit();	//Redirect to BuleGate

</script>
