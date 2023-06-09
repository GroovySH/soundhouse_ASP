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
'	�J�[�h�I�[�_�[3D�Z�L���A���N�G�X�g���� (BlueGate)
'
'		�J�[�h��3D�Z�L���A�`�F�b�N�����_�C���N�g�Ń��N�G�X�g����B
'		BlueGate 3DSecure ����̖߂�́AOrderCard3DResponseBG.asp
'
'------------------------------------------------------------------------
'	�X�V����
'2006/09/21 BlueGate�A�N�Z�X���O�ǉ�
'2006/11/03 BlueGate�A�N�Z�X���O���~
'2006/11/06 BlueGate�A�N�Z�X���O����(Open���Ԃ�Z��)
'2007/02/12 �I�[�\���G���[���̐����y�[�W�����N���ύX
'2007/04/16 BlueGate�A�N�Z�X���O���~
'2009/04/30 �G���[����error.asp�ֈړ�
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

'---- �G���[�������Ƃ���3D�Z�L���ABlueGate�A�G���[������Ίm�F�y�[�W��
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
'	Function	Main 3D�Z�L���A �_�C�W�F�X�g�쐬
'
'========================================================================
'
Function main()

'---- �J�[�h�����o��
call getCard()

if wMSG <> "" then
	exit function
end if

'---- 3D�Z�L���A �_�C�W�F�X�g�쐬
call get3DDigest()

End Function

'========================================================================
'
'	Function	�J�[�h�����o��
'
'========================================================================
'
Function GetCard()

'---- ���󒍎��o��
wSQL = ""
wSQL = wSQL & "SELECT a.�J�[�h�ԍ�"
wSQL = wSQL & "     , a.�J�[�h�L������"
wSQL = wSQL & "     , a.�J�[�h���`�l"
wSQL = wSQL & "     , a.�󒍍��v���z"
wSQL = wSQL & "     , a.�J�[�h�^�M�m�F�ԍ�"
wSQL = wSQL & "  FROM ���� a"
wSQL = wSQL & " WHERE SessionID = '" & w_sessionID & "'"
	  
Set RS_OrderHeader = Server.CreateObject("ADODB.Recordset")
RS_OrderHeader.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RS_OrderHeader.EOF = true then
	wMSG = "<font color='#ff0000'>NoData</font>"
	exit function
end if

CardNo = RS_OrderHeader("�J�[�h�ԍ�")
CardExpDt = RS_OrderHeader("�J�[�h�L������")
CardExpDt = Left(CardExpDt, 2) & Right(CardExpDt, 2)	'MMYY
CardHolderName = RS_OrderHeader("�J�[�h���`�l")
OrderTotalAm = RS_OrderHeader("�󒍍��v���z")

RS_OrderHeader.close

End function

'========================================================================
'
'	Function	3D�Z�L���A �_�C�W�F�X�g�쐬
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

'---- �p�����[�^�̃Z�b�g
OrderNo          = GetOrderNo()              '�����ԍ�
OrderTaxShipping = 0                         '�ő���

'---- 3DRequestMDCreator���\�b�h�R�[��
Set ObjBG = Server.CreateObject("Aspcompg.aspcom")

'---- Log before
'Set FS_Log = FS.OpenTextFile(LogFileName, 8, true)			'Log open - Append Mode
'FS_Log.WriteLine(cf_FormatTime(Now(), "HH:MM:SS") & " OrderCard3dSecureBG.asp   ComThreeDRequestMDCreator  BEFORE OrderNo=" & OrderNo)
'FS_Log.Close											'Log close

vRetCode = ObjBG.ComThreeDRequestMDCreator(g_BlueGate_ID, g_BlueGate_PW, OrderNo, OrderTotalAm, OrderTaxShipping, CardNo, CardExpDt)

'----�v���p�e�B��ݒ�
ThreeDDigest = ObjBG.ComGetPropValue("MsgDigest") '�R�c���b�Z�[�W�_�C�W�F�X�g
ErrCode      = ObjBG.ComGetPropValue("ErrCode")   '�G���[�R�[�h

'---- Log after
'Set FS_Log = FS.OpenTextFile(LogFileName, 8, true)			'Log open - Append Mode
'FS_Log.WriteLine(cf_FormatTime(Now(), "HH:MM:SS") & " OrderCard3dSecureBG.asp   ComThreeDRequestMDCreator  AFTER  OrderNo=" & OrderNo & " ErrCode=" & ErrCode & " SessionID=" & w_sessionID)
'FS_Log.Close											'Log close

Set ObjBG = Nothing

'---- �G���[�`�F�b�N
call checkError()

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
'	Function	�J�[�h�G���[�`�F�b�N
'
'========================================================================
'
Function checkError()

Dim vNoError
Dim vCardDataError

'---- ���^�[���R�[�h�ݒ�
'---- ����
vNoError = "00000000"

'---- 3D OK
if InStr(vNoError, ErrCode) > 0 then
	wMSG = ""
	exit function
end if

'---- ���̑��J�[�h�G���[
wMSG = "<font color='#ff0000'>" _
			& "�\���󂲂����܂��񂪤��w��̃J�[�h�ł͌䒍���ł��܂���B<br>" _
			& "�ʂ̃J�[�h�܂��ͤ�ʂ̂��x�����@�Ō䒍���肢�܂��B<br>" _
			& "Code: " & ErrCode & " (OrderCard3DSecureBG)<br>" _
			& "�悭���邲�����<a href='" & G_HTTP & "guide/qanda8.asp'>������</a>" _
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
