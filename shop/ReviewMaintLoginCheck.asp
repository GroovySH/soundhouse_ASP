<%@ LANGUAGE="VBScript" %>
<%
'�l�b�g�n�E�X�˂��ƃn�E�X�l�b�g�͂���
'�T�E���h�n�E�X
Option Explicit
%>
<!--#include file="../common/ADOVBS.inc"-->
<!--#include file="../common/Bfunctions1.asp"-->
<!--#include file="../common/shop_common_functions.inc"-->
<!--#include file="../common/system_common.inc"-->
<!--#include file="../common/HttpsSecurity.inc"-->

<%
'========================================================================
'
'    ���i���r���[�����e�i���X���O�C���`�F�b�N
'�X�V����
'2011/09/06 an �V�K�쐬
'2012/08/11 nt ���O�C�����擾���ύX
'             �i�R���g���[���}�X�^��WEB�Ǘ��҃}�X�^[�V��]�j
'
'========================================================================

On Error Resume Next
Response.Buffer = true
Response.Expires = -1			' Do not cache

Dim UserID
Dim Password
Dim Logout
Dim recCnt		'2012/08/11 nt add
Dim url
Dim Connection

Dim wErrMSG

'========================================================================

'---- Get GET/POST data
UserID = ReplaceInput(Trim(Request("UserID")))
Password = ReplaceInput(Trim(Request("Password")))
Logout = ReplaceInput(Trim(Request("Logout")))

'2012/08/11 nt add
'---- Set Cookie data
Response.Cookies("UserID") = UserID
Response.Cookies("Password") = Password

'---- Execute main
call connect_db()
call main()
call close_db()

if Err.Description <> "" then    
    Response.Redirect g_HTTP & "shop/Error.asp"
end if

'���[�U�[ID/�p�X���[�h�s�������A���O�A�E�g���̓��O�C����ʂɖ߂�
if wErrMSG <> "" then
	Response.Redirect "ReviewMaintLogin.asp"
else
	Response.Redirect "ReviewSearch.asp"
end if

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
'    Function    Main
'
'========================================================================
'
Function main()

Dim vItemChar1
Dim vItemChar2
Dim vItemNum1
Dim vItemNum2
Dim vItemDate1
Dim vItemDate2

'Dim vUserID		2012/08/11 nt del
'Dim vPassword		2012/08/11 nt del

wErrMSG = ""

if Logout = "Y" then
	'---- ���O�A�E�g
	call fSetSessionData(gSessionID, "ShAdminFl", "")
	wErrMSG = "���O�A�E�g"
	exit function
end if

'2012/08/11 nt add Start
if UserID = "" And Password = "" then
	wErrMSG = "���[�U�[ID�E�p�X���[�h����͂��ĉ������B"
	call fSetSessionData(gSessionID, "���b�Z�[�W", wErrMSG)
	exit function
end if
'2012/08/11 nt add End

'2012/08/11 nt mod Start
'---- �R���g���[���}�X�^���烍�O�C�����擾
'call getCntlMst("���r���[","���O�C��","1", vItemChar1, vItemChar2, vItemNum1, vItemNum2, vItemDate1, vItemDate2)
'vUserID = vItemChar1
'vPassword = vItemChar2

'---- ���[�U�[ID OR �p�X���[�h���s��v�̏ꍇ�̓Z�b�V�����f�[�^�ɃG���[�o�^
'if UserID <> vUserID OR Password <> vPassword then
'	wErrMSG = "���[�U�[ID�܂��̓p�X���[�h���s���ł��B"
'	call fSetSessionData(gSessionID, "���b�Z�[�W", wErrMSG)
'---- OK�̏ꍇ�́u���O�C�����v�ɐݒ�
'else
'	call fSetSessionData(gSessionID, "ShAdminFl", "Y")
'end if

'---- WEB�Ǘ��҃}�X�^����A���O�C�������擾
Call getWEBMaster()

'---- ���O�C���ۂ��Z�b�V�����f�[�^�ɓo�^
if recCnt = 0 then
	'---- �擾�������O�C����񂪑��݂��Ȃ��ꍇ�A�G���[���b�Z�[�W��\�����A�Z�b�V�����f�[�^�ɃG���[�o�^
	wErrMSG = "���[�U�[ID�܂��̓p�X���[�h���s���ł��B"
	call fSetSessionData(gSessionID, "���b�Z�[�W", wErrMSG)

else
	'---- �擾�������O�C����񂪑��݂���΁A�Z�b�V�����f�[�^�Ƀt���O�o�^
	call fSetSessionData(gSessionID, "ShAdminFl", "Y")

end if
'2012/08/11 nt mod End

End Function
 

'========================================================================
'
'	Function	Close database
'
'========================================================================
'
Function close_db()

Connection.close
Set Connection= Nothing

End function

'2012/08/11 nt add
'========================================================================
'
'	Function	WEB�Ǘ��ҏ��̗L�����擾
'
'========================================================================
Function getWEBMaster()

Dim RSv
Dim vSQL

vSQL = ""
vSQL = vSQL & "SELECT * "
vSQL = vSQL & " FROM "
vSQL = vSQL & "    WEB�Ǘ��� a WITH (NOLOCK) "
vSQL = vSQL & " WHERE "
vSQL = vSQL & "        a.���[�UID = '" & UserID & "' "
vSQL = vSQL & "    AND a.�p�X���[�h = '" & Password & "' "
vSQL = vSQL & "    AND a.�폜�t���O = '0'"	'�폜�t���O[0]�FActive�A[1]�FNon-Active

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic, adLockOptimistic

'---- WEB�Ǘ��ҏ��̗L�����擾
recCnt = RSv.RecordCount

RSv.Close

End Function

'========================================================================
%>