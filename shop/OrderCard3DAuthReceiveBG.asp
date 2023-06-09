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
'	�J�[�h�I�[�_�[��ʌ^3D�Z�L���A/�I�[�\�� ��M (BlueGate)
'
'		�I�[�\���ԍ�����M
'		OrderCard3DAuthSendBG.asp ����̖߂�
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

Dim ModeCode	    '�d�����
Dim SID           '�����X���R��
Dim OrderNo       '�����ԍ�
Dim ApprovalCode  '���F�ԍ�
Dim AcqCode       '��d�����
Dim TotalAmount   '���ϋ��z���v
Dim ReceiveDateTime '��M����
Dim PaymentDate   '���ϓ���
Dim MsgDigest     'MsgDigest
Dim ErrCode       '�G���[�R�[�h

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

'---- �󂯎�����荞��
ModeCode	    = Request("ModeCode")      '�d�����
SID           = ReplaceInput(Request("SID"))           '�����X���R��
OrderNo       = ReplaceInput(Request("OrderNum"))      '�����ԍ�
ApprovalCode  = ReplaceInput(Request("ApprovalCode"))  '���F�ԍ�
AcqCode       = Request("AcqCode")       '��d�����
TotalAmount   = Request("TotalAmount")   '���ϋ��z���v
ReceiveDateTime = Request("ReceiveDateTime")  '��M����
PaymentDate   = Request("PaymentDate")   '���ϓ���
MsgDigest     = Request("MsgDigest" )    'MsgDigest
ErrCode       = ReplaceInput(Request("ErrCode"))       '�G���[�R�[�h

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
'	Function	Main �m�F�p �_�C�W�F�X�g�쐬
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

'---- 3DResponseMDCreator���\�b�h�R�[��
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
'	Function	���󒍏��̍X�V
'
'========================================================================
'
Function updateOrderHeader()

'---- ���󒍎��o��
wSQL = ""
wSQL = wSQL & "SELECT a.�J�[�h�^�M�m�F�ԍ�"
wSQL = wSQL & "  FROM ���� a"
wSQL = wSQL & " WHERE SessionID = '" & w_sessionID & "'"
	  
Set RS_order_header = Server.CreateObject("ADODB.Recordset")
RS_order_header.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RS_order_header.EOF = true then
	wMSG = "<font color='#ff0000'>������񂪂���܂���</font>"
	exit function
end if

'---- update ����
RS_order_header("�J�[�h�^�M�m�F�ԍ�")   = ApprovalCode

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
