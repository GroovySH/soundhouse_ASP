<%@ LANGUAGE="VBScript" %>
<%
'�l�b�g�n�E�X�˂��ƃn�E�X�l�b�g�͂���
'�T�E���h�n�E�X
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
'	���i���r���[�����e�i���X2
'     ���i���r���[�̕ύX/�폜���s��
'
'2011/09/06 an #816 �V�K�쐬
'2012/08/11 nt �V���b�v�R�����g�̍X�V���ڂ�ǉ�
'
'========================================================================
On Error Resume Next
Response.buffer = true
Response.Expires = -1			' Do not cache

Dim ReviewID
Dim i_Mode
Dim Title
Dim Hyouka
Dim Review
Dim UserID		'2012/08/11 nt add
Dim Password	'2012/08/11 nt add
Dim recCnt		'2012/08/11 nt add
Dim sCDate		'2012/08/11 nt add
Dim sComment	'2012/08/11 nt add

Dim wReviewDate
Dim wReviewName
Dim wMakerName
Dim wProductName

Dim Connection

Dim wErrMSG
Dim wLoginFl

'========================================================================

'---- Get GET/POST data
ReviewID = ReplaceInput(Request("ReviewID"))
i_Mode = ReplaceInput(Request("i_Mode"))
Title = ReplaceInput(Left(Request("Title"),51))
Hyouka = ReplaceInput(Request("Hyouka"))
Review = ReplaceInput(Left(Request("Review"),1001))
UserID = ReplaceInput(Request.Cookies("UserID"))			'2012/08/11 nt add
Password = ReplaceInput(Request.Cookies("Password"))		'2012/08/11 nt add
sCDate = ReplaceInput(Request("sCDate"))					'2012/08/11 nt add
sComment = ReplaceInput(Left(Request("sComment"),1000))		'2012/08/11 nt add

'---- Execute main
call connect_db()
call main()
call close_db()

if Err.Description <> "" then
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

'---- �����O�C���̏ꍇ�̓��O�C����ʂ�
if wLoginFl <> "Y" then
	Response.Redirect g_HTTPS & "shop/ReviewMaintLogin.asp"
end if

'---- �G���[�̏ꍇ��ReviewMaint�ɖ߂�
if wErrMSG <> "" then
	Server.Transfer "ReviewMaint.asp"
else
	'2012/08/11 nt mod Start
	'if i_Mode = "update" then
	'	Response.Redirect g_HTTPS & "shop/ReviewMaint.asp?ReviewID=" & ReviewID
	'else
	'	Response.Redirect g_HTTPS & "shop/ReviewSearch.asp"
	'end if

	if i_Mode = "update" Or i_Mode = "sCDelete" then
		Response.Redirect g_HTTPS & "shop/ReviewMaint.asp?ReviewID=" & ReviewID
	else
		Response.Redirect g_HTTPS & "shop/ReviewSearch.asp"
	end if
	'2012/08/11 nt mod End
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
'	Function	main
'
'========================================================================
Function main()

wErrMSG = ""
wLoginFl = "N"

'---- �Z�L�����e�B�[�L�[�`�F�b�N
If Session("Skey") <> Request("Skey") Then
	Response.Redirect g_HTTP & "index.asp"
End If

'---- ���O�C���X�e�[�^�X�擾
wLoginFl = fGetSessionData(gSessionID, "ShAdminFl")

if wLoginFl <> "Y" then
	call fSetSessionData(gSessionID, "���b�Z�[�W", "���O�C�����Ă��������B")
	exit function
end if

'2012/08/11 nt add Start
'---- WEB�Ǘ��҃}�X�^����A���O�C�������擾
Call getWEBMaster()

'---- ���O�C���ۂ��Z�b�V�����f�[�^�ɓo�^
if recCnt = 0 then
	'---- �擾�������O�C����񂪑��݂��Ȃ��ꍇ�A�G���[���b�Z�[�W��\�����A�Z�b�V�����f�[�^�ɃG���[�o�^
	wErrMSG = "�s���ȃ��O�C���ł��B"
	call fSetSessionData(gSessionID, "���b�Z�[�W", wErrMSG)
	exit function
end if
'2012/08/11 nt add End

'---- ���̓`�F�b�N
call validation()

if wErrMSG = "" then
	call UpdateDeleteReview()
end if

End function

'========================================================================
'
'    Function    ���͓��e�`�F�b�N
'
'========================================================================
'
Function validation()

'---- �������[�h
'if i_Mode <> "update" AND i_Mode <> "delete" then 2012/08/11 nt mod
if i_Mode <> "update" AND i_Mode <> "delete" AND i_Mode <> "sCDelete" then
	wErrMSG = wErrMSG & "���[�h���s���ł��B<br />"
end if

'---- ���r���[ID
if ReviewID = "" then
	wErrMSG = wErrMSG & "���r���[ID����͂��Ă��������B<br />"
else
	if cf_checkNumeric(ReviewID) = false then
		wErrMSG = wErrMSG & "���r���[ID���s���ł��B<br />"
	end if
end if

'---- �^�C�g��
if Title = "" then
	wErrMSG = wErrMSG & "�^�C�g������͂��Ă��������B<br />"
else
	if Len(Title) > 50 then
		wErrMSG = wErrMSG & "�^�C�g����50�����ȓ��œ��͂��Ă��������B<br />"
	end if
end if

'---- �]��
if Hyouka = "" then
	wErrMSG = wErrMSG & "�]������͂��Ă��������B<br />"
else
	if Hyouka <> "1" AND Hyouka <> "2" AND Hyouka <> "3" AND Hyouka <> "4" AND Hyouka <> "5" then
		wErrMSG = wErrMSG & "�]����1�`5����͂��Ă��������B<br />"
	end if
end if

'---- ���r���[���e
if Review = "" then
	wErrMSG = wErrMSG & "���r���[���e����͂��Ă��������B<br />"
else
	if Len(Review) > 1000 then
		wErrMSG = wErrMSG & "���r���[���e��1000�����ȓ��œ��͂��Ă��������B<br />"
	end if
end if

'2012/08/11 nt add Start
if IsDate(sCDate) = false then
	wErrMSG = wErrMSG & "�V���b�v�R�����g�����s���ł��B<br />"
end if

'---- �V���b�v�R�����g
if Len(sComment) > 1000 then
	wErrMSG = wErrMSG & "�V���b�v�R�����g��1000�����ȓ��œ��͂��Ă��������B<br />"
end if
'2012/08/11 nt add End

'---- �G���[������ꍇ�̓Z�b�V�����f�[�^�ɋL�^
if wErrMSG <> "" then
	call fSetSessionData(gSessionID, "���b�Z�[�W", wErrMSG)
end if

End function

'========================================================================
'
'    Function    ���i���r���[�ύX�A�폜
'
'========================================================================
'
Function UpdateDeleteReview()

Dim RSv
Dim vSQL

vSQL = ""
vSQL = vSQL & "SELECT *"
vSQL = vSQL & " FROM ���i���r���["
vSQL = vSQL & " WHERE ID = " & ReviewID

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic, adLockOptimistic
if RSv.EOF = true then
	wErrMSG = "�Y���̃��r���[������܂���B ���r���[ID��" & ReviewID
	call fSetSessionData(gSessionID, "���b�Z�[�W", wErrMSG)
else

	if i_Mode = "update" then
		RSv("�^�C�g��") = Title
		RSv("�]��") = Hyouka
		RSv("���r���[���e") = Review

		'2012/08/11 nt add Start
		'---- �V���b�v�R�����g�ɓ��͂��Ȃ���΁A�V���b�v�R�����g������уV���b�v�R�����g�͍X�V���Ȃ�
		if len(sComment) > 0 then
			if (sCDate <> "") then
				RSv("�V���b�v�R�����g��") = cf_FormatDate(sCDate, "YYYY/MM/DD")
			end if
			RSv("�V���b�v�R�����g") = sComment
		end if
		'2012/08/11 nt add End

		RSv.Update

		'2012/08/11 nt mod Start
		'call fSetSessionData(gSessionID, "���b�Z�[�W", "�X�V����܂����B")
		if len(sComment) > 0 then
			call fSetSessionData(gSessionID, "���b�Z�[�W", "�X�V����܂����B")
		else
			call fSetSessionData(gSessionID, "���b�Z�[�W", "�X�V����܂����B<br>���j�V���b�v�R�����g�����͂���Ȃ��������߁A�V���b�v�R�����g�͍X�V����܂���")
		end if
		'2012/08/11 nt mod End

	'2012/08/11 nt add Start
	'---- �u�V���b�v�R�����g�̂ݍ폜�v�{�^����ǉ�
	elseif i_Mode = "sCDelete" then

		'2012/08/11 nt add Start
		RSv("�V���b�v�R�����g��") = NULL
		RSv("�V���b�v�R�����g�^�C�g��") = NULL
		RSv("�V���b�v�R�����g") = NULL
		'2012/08/11 nt add End

		RSv.Update
		call fSetSessionData(gSessionID, "���b�Z�[�W", "�V���b�v�R�����g���폜����܂����B")
	'2012/08/11 nt add End

	else
		RSv.Delete
		call fSetSessionData(gSessionID, "���b�Z�[�W", "�폜����܂����B")
	end if
end if

RSv.Close

End function

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