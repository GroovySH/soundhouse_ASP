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
'	�J�[�h�I�[�_�[3D�Z�L���A���ʎ󂯎�菈�� (BlueGate)
'
'		�J�[�h��3D�Z�L���A�`�F�b�N�̌��ʂ�BlueGate���󂯎��B
'		OK�Ȃ�AOrderCardAuthBG.asp���Ăяo���A�I�[�\�������B
'
'------------------------------------------------------------------------
'	�X�V����
'2006/09/21 BlueGate�A�N�Z�X���O�ǉ�
'2006/11/03 BlueGate�A�N�Z�X���O���~
'2006/11/06 BlueGate�A�N�Z�X���O����(Open���Ԃ�Z��)
'2007/02/12 �I�[�\���G���[���̐����y�[�W�����N���ύX
'2007/04/11 OrderCardAuthBG.asp�Ăяo���p�����[�^��Xstatus��ǉ�
'2007/04/16 BlueGate�A�N�Z�X���O���~
'2007/08/14 �J�[�h�G���[���̃��b�Z�[�W�ύX
'2009/04/30 �G���[����error.asp�ֈړ�
'
'========================================================================

On Error Resume Next

Dim w_sessionID
Dim userID
Dim msg

Dim ModeCode	       '�d�����
Dim SID              '�����X���R��
Dim OrderNo          '�����ԍ�
Dim MsgVerNum        'version
Dim XID              'xid
Dim Xstatus          'status
Dim ECI              'eci
Dim CAVV             'cavv
Dim CavvAlgorithm    'cavvAlgorithm
Dim MsgDigest        'MsgDigest
Dim ErrCode          '�G���[�R�[�h
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

'---- �󂯎�����荞��
ModeCode	    = Request("ModeCode")      '�d�����
SID           = Request("SID")           '�����X���R��
OrderNo       = Request("OrderNum")      '�����ԍ�
MsgVerNum     = Request("MsgVerNum")     '3D version
XID           = Request("XID")           '3D xid
Xstatus       = Request("Xstatus")       '3D status
ECI           = Request("ECI")           '3D eci
CAVV          = Request("CAVV")          '3D cavv
CavvAlgorithm = Request("CavvAlgorithm") '3D cavvAlgorithm
MsgDigest     = Request("MsgDigest" )    '3D MsgDigest
ErrCode       = Request("ErrCode")       '�G���[�R�[�h

'---- execute main process
call main()

if Err.Description <> "" then	
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

'---- �G���[�������Ƃ��̓I�[�\���擾�A�G���[������Ίm�F�y�[�W��
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
'	Function	Main 3D�Z�L���A �_�C�W�F�X�g�쐬
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

'---- �G���[�`�F�b�N
call checkError()
if wMsg <> "" then
	exit function
end if

Session("BlueGate3DReturnCode") = ErrCode

'---- 3DResponseMDCreator���\�b�h�R�[��
Set ObjBG = Server.CreateObject("Aspcompg.aspcom")

'---- Log before
'Set FS_Log = FS.OpenTextFile(LogFileName, 8, true)			'Log open - Append Mode
'FS_Log.WriteLine(cf_FormatTime(Now(), "HH:MM:SS") & " OrderCard3dResponseBG.asp ComThreeDResponseMDCreator BEFORE OrderNo=" & OrderNo)
'FS_Log.Close											'Log close

vRetCode = ObjBG.ComThreeDResponseMDCreator(g_BlueGate_ID, g_BlueGate_PW, OrderNo, MsgVerNum, XID, Xstatus, ECI, CAVV, CavvAlgorithm )

'----�v���p�e�B��ݒ�
ResultDigest = ObjBG.ComGetPropValue("MsgDigest") '���ʃ_�C�W�F�X�g
ErrCode      = ObjBG.ComGetPropValue("ErrCode")   '�G���[�R�[�h

'---- Log after
'Set FS_Log = FS.OpenTextFile(LogFileName, 8, true)			'Log open - Append Mode
'FS_Log.WriteLine(cf_FormatTime(Now(), "HH:MM:SS") & " OrderCard3dResponseBG.asp ComThreeDResponseMDCreator AFTER  OrderNo=" & OrderNo & " ErrCode=" & ErrCode)
'FS_Log.Close											'Log close

'---- 3D�_�C�W�F�X�g�G���[
if (ErrCode <> "00000000") OR (MsgDigest <> ResultDigest) then
	wMSG = "CardError1"
end if

Set ObjBG = Nothing

end function

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
vNoError = "00000000,S102000W"		'S102000W:3DSecure�T�[�r�X�ΏۊO

'---- 3D OK
if InStr(vNoError, ErrCode) > 0 then
	wMSG = ""
	exit function
end if

'---- ���̑��J�[�h�G���[
wMSG = "CardError1"

End function

%>
