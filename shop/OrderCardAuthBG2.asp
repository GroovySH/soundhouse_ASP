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
'	�J�[�h�I�[�_�[�^�M�m�F���� (BlueGate)
'
'		�J�[�h�̗^�M�����Ok�Ȃ�order_submit�փR���g���[����n���B
'		�^�MOK�Ȃ�A�󒍔ԍ��̍̔Ԃ��s���B
'
'------------------------------------------------------------------------
'	�X�V����
'2006/07/19 3D�p�@�I�[�_�[�ԍ���3D�v���O��������n����邽�ߍ̔ԕs�v
'2006/09/21 BlueGate�A�N�Z�X���O�ǉ�
'2007/02/12 �I�[�\���G���[���̐����y�[�W�����N���ύX
'2006/09/21 BlueGate�A�N�Z�X���O���~
'2007/04/11 3D�p�����[�^���I�[�\���p�����[�^�ɒǉ�
'2007/04/16 BlueGate�A�N�Z�X���O���~
'2007/05/30 ECI���󒍏��Ƃ��Ď�荞��
'2007/08/14 �J�[�h�G���[���̃��b�Z�[�W�ύX
'2008/10/13 �V�J�[�h���͑Ή��iPCIDSS)
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
Dim CustomerNo

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

'---- �G���[�������Ƃ��͒����o�^�����y�[�W�A�G���[������Ίm�F�y�[�W��
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
'	Function	Main �J�[�h�^�M�m�F
'
'========================================================================
'
Function main()

'---- �J�[�h�����o��
call getCard()
call getCard2()

if wMSG <> "" then
	exit function
end if

'---- �^�M�`�F�b�N
call getCardAuth()

if wMSG = "" then
	'---- �󒍏��ɗ^�M�m�F�ԍ����Z�b�g
	call updateOrderHeader()

	'---- �J�[�h�ԍ���ێ����Ȃ��ꍇ�͍폜
	call chackCardHoji()
end if

RS_OrderHeader.close

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
wSQL = wSQL & "SELECT b.�J�[�h���`�l"
wSQL = wSQL & "     , a.�󒍍��v���z"
wSQL = wSQL & "     , a.�ڋq�ԍ�"
wSQL = wSQL & "     , a.�J�[�h�^�M�m�F�ԍ�"
wSQL = wSQL & "     , a.�J�[�h�l�b�g�`�[�ԍ�"
wSQL = wSQL & "     , a.BlueGateECI"
wSQL = wSQL & "  FROM ���� a"
wSQL = wSQL & "     , Web�ڋq b"
wSQL = wSQL & " WHERE b.�ڋq�ԍ� = a.�ڋq�ԍ�"
wSQL = wSQL & "   AND a.SessionID = '" & w_sessionID & "'"

Set RS_OrderHeader = Server.CreateObject("ADODB.Recordset")
RS_OrderHeader.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RS_OrderHeader.EOF = true then
	wMSG = "<font color='#ff0000'>NoData</font>"
	exit function
end if

CardHolderName = RS_OrderHeader("�J�[�h���`�l")
OrderTotalAm = RS_OrderHeader("�󒍍��v���z")
CustomerNo = RS_OrderHeader("�ڋq�ԍ�")

End function

'========================================================================
'
'	Function	�J�[�h�����o��2
'
'========================================================================
'
Function GetCard2()

Dim Campus

Set Campus = Server.CreateObject("WebCampusAccess.WebCampus")

Campus.Site = g_RegForder
Campus.CustomerNo = CustomerNo

Campus.GetCardNo()

CardNo = Campus.CardNo
CardExpDt = Campus.CardExpDt

CardExpDt = Left(CardExpDt, 2) & Right(CardExpDt, 2)	'MMYY

End function

'========================================================================
'
'	Function	�J�[�h�^�M�m�F
'
'========================================================================
'
Function getCardAuth()

Dim ObjBG

Dim vRetCode

'---- BlueGate Log open
'Set FS = CreateObject("Scripting.FileSystemObject")
'LogFileName = "BlueGateLog/BlueGateLog" & Year(Date()) & Right("0" & Month(Date()), 2) & Right("0" & Day(Date()), 2) & ".txt"
'LogFileName = Server.MapPath(LogFileName)		'Map log file
'Set FS_Log = FS.OpenTextFile(LogFileName, 8, true)			'Log open - Append Mode

'---- �p�����[�^�̃Z�b�g
InShopId           = g_BlueGate_ID             '�V���b�vID
InShopPw           = g_BlueGate_PW             '�V���b�v�p�X���[�h

if OrderNo = "" then
	OrderNo            = GetOrderNo()              '�����ԍ�		'3D���͕s�v
end if

InAmount           = OrderTotalAm              '������z
IntaxAndDeliCharge = 0                         '�ő���
InPan              = CardNo                    '�J�[�h�ԍ�
InExpiryDate       = CardExpDt							   '�L������
InPaymentMode      = "10"                      '�x���敪(�ꊇ)

'---- �I�[�\���擾
Set ObjBG = Server.CreateObject("Aspcompg.aspcom")

'---- Log before
'FS_Log.WriteLine(cf_FormatTime(Now(), "HH:MM:SS") & " OrderCardAuthBG2.asp       ComAuthoriRequest          BEFORE OrderNo=" & OrderNo)

vRetCode = ObjBG.ComAuthoriRequest(InShopId, InShopPw, OrderNo, InAmount, IntaxAndDeliCharge, InPan, InExpiryDate, InPaymentMode, InStartPayMonth, InPaymentCount, InInitialAmount, InBonusMonth, InBonusAmount, InBonusCount, InMsgVerNum, InXid, InXStatus, InEci, InCavv, InCavvAlgorithm )

'---- �v���p�e�B��ݒ�
ApprovalCode    = ObjBG.ComGetPropValue("ApprovalCode")      '���F�ԍ�
ErrCode         = ObjBG.ComGetPropValue("ErrCode")           '�G���[�R�[�h
AcqCode         = ObjBG.ComGetPropValue("AcqCode")           '��d�����
TotalAmount     = ObjBG.ComGetPropValue("TotalAmount")       '���ϋ��z
ReceiveDateTime = ObjBG.ComGetPropValue("ReceiveDateTime")   '��t����
PaymentDate     = ObjBG.ComGetPropValue("PaymentDate")       '���Ϗ������t
DetailCode      = ObjBG.ComGetPropValue("DetailCode")        '�ڍ׃R�[�h

'---- Log after
'FS_Log.WriteLine(cf_FormatTime(Now(), "HH:MM:SS") & " OrderCardAuthBG2.asp       ComAuthoriRequest          AFTER  OrderNo=" & OrderNo & " ErrCode=" & ErrCode)
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
'	Function	���󒍏��̍X�V
'
'========================================================================
'
Function updateOrderHeader()

'---- update ����
RS_OrderHeader("�J�[�h�^�M�m�F�ԍ�")   = ApprovalCode
RS_OrderHeader("�J�[�h�l�b�g�`�[�ԍ�") = Auth3DKubun
RS_OrderHeader("BlueGateECI") = InEci

RS_OrderHeader.update

End function

'========================================================================
'
'	Function	�J�[�h�ԍ���ێ����Ȃ��ꍇ�͍폜
'
'========================================================================
'
Function chackCardHoji()

Dim Campus
Dim RS

'---- �ڋq�J�[�h���
wSQL = ""
wSQL = wSQL & "SELECT b.�J�[�h�ԍ�"
wSQL = wSQL & "  FROM Web�ڋq b"
wSQL = wSQL & " WHERE b.�ڋq�ԍ� = " & CustomerNo

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RS.EOF = true then
	wMSG = "<font color='#ff0000'>NoData</font>"
	exit function
end if

if RS("�J�[�h�ԍ�") = "" then
	Set Campus = Server.CreateObject("WebCampusAccess.WebCampus")

	Campus.Site = g_RegForder
	Campus.CustomerNo = CustomerNo

	Campus.DeleteCardNo()

end if

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
'---- �J�[�h�ԍ��܂��͗L�������G���[
vCardDataError = "S5001060,S5001061,S5001062,S5001069,S5001070,S5001072,S5001079"

'---- �I�[�\��OK
if InStr(vNoError, ErrCode) > 0 then
	wMSG = ""
	exit function
end if

'---- �J�[�h�G���[
if InStr(vCardDataError, ErrCode) > 0 then
	wMSG = "CardError1"
	exit function
end if

'---- ���̑��J�[�h�G���[
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
