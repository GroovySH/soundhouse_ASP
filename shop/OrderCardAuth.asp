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
'	�J�[�h�I�[�_�[�^�M�m�F����
'
'		�J�[�h�̗^�M�����Ok�Ȃ�order_submit�փR���g���[����n���B
'
'------------------------------------------------------------------------
'	
'		���̃v���O������Cardnet����̃T���v���v���O���������ɍ���Ă��܂��B
'		���͂��ꂽ�J�[�h�̗^�M�`�F�b�N���s���
'		OK�Ȃ�󒍓o�^������
'
'------------------------------------------------------------------------
'	�X�V����
'2005/04/05 �J�[�h�����󒍃f�[�^������o���悤�ɕύX
'2006/06/30 �󒍏��Ȃ��̂Ƃ��̓G���[
'2009/04/30 �G���[����error.asp�ֈړ�
'
'========================================================================

On Error Resume Next

Dim w_sessionID
Dim userID
Dim msg

Dim card_no
Dim card_exp_dt
Dim card_exp_dt1
Dim card_exp_dt2
Dim card_holder_nm
Dim order_total_am
Dim card_order_no
Dim card_net_no
Dim card_auth_no

Dim Connection
Dim RS_order_header

Dim w_sql
Dim w_html
Dim w_msg
Dim w_next_URL

'=======================================================================

w_sessionID = Session.SessionId
userID = Request.cookies("UserID")

Session("msg") = ""
w_msg = ""

'---- execute main process
call connect_db()
call main()
call close_db()

if Err.Description <> "" then	
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

'---- �G���[�������Ƃ��͒����o�^�����y�[�W�A�G���[������Ίm�F�y�[�W��

if w_msg = "" then
	Response.Redirect "OrderSubmit.asp"
else
	Session("msg") = w_msg
	Response.Redirect "OrderInfoEnter.asp"
end if

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
'	Function	Main �J�[�h�^�M�m�F
'
'========================================================================
'
Function main()

'---- �J�[�h�����o��
call get_card()

if w_msg <> "" then
	exit function
end if

'************************************** �ύX����
card_order_no = 00001

'---- �^�M�`�F�b�N
call card_auth()

'---- �󒍏��ɗ^�M�m�F�ԍ����Z�b�g
if w_msg = "" then
	call update_order_header()
end if

'---- �O�̂��߃`�F�b�N
if RS_order_header("�J�[�h�^�M�m�F�ԍ�") = "" then
		w_msg = "<font color='#ff0000'>�J�[�h�^�M�̎擾���o���܂���ł����B<br>�ʂ̃J�[�h�ōēx�䒍�����������B</font>"
end if

RS_order_header.close

End Function

'========================================================================
'
'	Function	�J�[�h�����o��
'
'========================================================================
'
Function get_card()

'---- ���󒍎��o��
w_sql = ""
w_sql = w_sql & "SELECT a.�J�[�h�ԍ�"
w_sql = w_sql & "     , a.�J�[�h�L������"
w_sql = w_sql & "     , a.�J�[�h���`�l"
w_sql = w_sql & "     , a.�󒍍��v���z"
w_sql = w_sql & "     , a.�J�[�h�^�M�m�F�ԍ�"
w_sql = w_sql & "     , a.�J�[�h�l�b�g�`�[�ԍ�"
w_sql = w_sql & "  FROM ���� a"
w_sql = w_sql & " WHERE SessionID = " & w_sessionID
	  
Set RS_order_header = Server.CreateObject("ADODB.Recordset")
RS_order_header.Open w_sql, Connection, adOpenStatic, adLockOptimistic

if RS_order_header.EOF = true then
	w_msg = "<font color='#ff0000'>NoData</font>"
	exit function
end if

card_no = RS_order_header("�J�[�h�ԍ�")
card_exp_dt = RS_order_header("�J�[�h�L������")
card_exp_dt1 = Left(card_exp_dt, 2)
card_exp_dt2 = Right(card_exp_dt, 2)
card_holder_nm = RS_order_header("�J�[�h���`�l")
order_total_am = RS_order_header("�󒍍��v���z")

End function

'========================================================================
'
'	Function	���󒍏��̍X�V
'
'========================================================================
'
Function update_order_header()

'---- update ����
RS_order_header("�J�[�h�^�M�m�F�ԍ�") = card_auth_no
RS_order_header("�J�[�h�l�b�g�`�[�ԍ�") = card_net_no

RS_order_header.update

End function

'========================================================================
'
'	Function	�J�[�h�^�M�m�F
'
'========================================================================
'
Function card_auth()

REM	/*==================================================================*/

REM -- ���σp�b�P�[�W��W�J�����f�B���N�g����ݒ肵�܂��B
REM -- (�Q�l)	sgsv003z.exe, sgsv004z.exe, sgsv012z.exe, sgsv00001a.prm
REM --			��W�J�����f�B���N�g���ł��B

Dim HomeDir
HomeDir = "d:\soundhouse\wwwroot\cardnet"			'�{��@@@@@@@@@@@@@@@@@@@@@@@@
'''''HomeDir = "\\Emax2\Web\SH_New\cardnet"		'�e�X�g@@@@@@@@@@@@@@@@@@@@@@@@


REM	/*==================================================================*/
REM	/* ���̐ݒ�̓p�b�P�[�W�V�X�e���̐ݒ�l�ł��B						*/
REM	/*==================================================================*/

REM	--	�`�r�o�A�g�c�k�k���W�X�g���o�^��

Dim DllRegist
DllRegist = "Sgsv011z.SSLAuth"

REM	--	�V�X�e���G���[�����������ꍇ�ɕ\������g�s�l�k

Dim ErrorURL
ErrorURL = "OrderInfoEnter.asp"

REM	--	�^�M�����������ꍇ�ɕ\������g�s�l�k

Dim SuccessURL
SuccessURL = "OrderSubmit.asp"

REM	--	�^�M�����ۂ��ꂽ�ꍇ�ɕ\������g�s�l�k

Dim FailureURL
FailureURL = "OrderInfoEnter.asp"


REM	/*==================================================================*/
REM	/*                                                                  */
REM	/*		�t�@�C����		�F	sgsv00013a.asp		(Original name)           */
REM	/*                                                                  */
REM	/*		�T�v			�F	�J�[�g���b�W�A�g���s�`�r�o                      */
REM	/*                                                                  */
REM	/*		�쐬��			�F 2000/01/04                                     */
REM	/*                                                                  */
REM	/*		�X�V����                                                      */
REM	/*		  ���t	  �ύX��				���R                                  */
REM	/*                                                                  */
REM	/*==================================================================*/

REM --
REM -- �I�u�W�F�N�g�����̉����܂��B(�ŏ��ɕK���K�v�ł��B)
REM --

'@@@@@@On Error Resume Next

Dim SSLAuth
Set SSLAuth = CreateObject(DllRegist)

REM --
REM -- ���σp�b�P�[�W��W�J�����f�B���N�g����ݒ肵�܂��B
REM -- (�Q�l)
REM -- sgsv003z.exe, sgsv004z.exe, sgsv012z.exe, sgsv00001a.prm��
REM -- ���݂���f�B���N�g���ł��B
SSLAuth.HomeDir = HomeDir

REM --
REM -- �d�b���σZ���^�[�֑��M���邽�߂̏������̂悤�ɐݒ肵�܂��B
REM --

REM -- �i�K�{�j�T�[�o�h�c��ݒ�i������Г��{�J�[�h�l�b�g����ʒm�����j
REM -- <<�Œ�l>>
SSLAuth.ServerID = "1680"

REM -- �i�K�{�j�V���b�v�h�c��ݒ�i������Г��{�J�[�h�l�b�g����ʒm�����j
REM -- <<�Œ�l>>
SSLAuth.ShopID = "0001"

REM -- �i�K�{�j�I�[�\�����z��ݒ�
SSLAuth.Amount = order_total_am

REM -- �i�K�{�j�x�����@��ݒ�
SSLAuth.PayMode = "10"			' �ꊇ

REM -- �i�x���敪������(61)�̎��̂ݕK�{�j
REM -- �����񐔂�ݒ肵�܂��B�����ȊO�̏ꍇ�ɐݒ肵�Ă��\���܂���B
If Request("card_payment_method") = "61" Then
	SSLAuth.InstallCount = 1
End If

REM -- �i�K�{�j�J�[�h�ԍ���ݒ�
SSLAuth.PAN = card_no

REM -- �i�K�{�j�J�[�h�L��������ݒ� �����ӁF(Month2��+Year�Q��)
SSLAuth.CardExp = card_exp_dt1 & card_exp_dt2

REM -- �i�K�{�j�`�[�ԍ���ݒ�
REM -- <<�����X�l���J�X�^�}�C�Y���āA�`�[�ԍ���ݒ肵�Ă��������B>>
SSLAuth.SalesSlipNo = cf_NumToChar(card_order_no, 5)

REM -- �i�I�v�V�����j���i�ԍ���ݒ�
SSLAuth.GoodsCode = "0990"

REM -- �i�I�v�V�����j�����X�_��J�[�h��ЃR�[�h��ݒ�
If Request("RECV_CO_COD") <> "" Then
	SSLAuth.CardCoCode = ""
End If

REM -- �i�I�v�V�����j�����X�[���ԍ��ݒ�
If Request("MER_TERM_NUM") <> "" Then
	SSLAuth.MerchantID = ""
End If

REM -- �i�I�v�V�����j�[�����ʔԍ�
If Request("TERM_NUM") <> "" Then
	SSLAuth.TerminalID = ""
End If

REM --
REM -- �d�b���σZ���^�[�փf�[�^�𑗐M���܂��B
REM --
SSLAuth.Send()

REM -- �V�X�e���I�ȃG���[�������������𒲂ׂ܂��B
Dim SystemErrorMessage
SystemErrorMessage = ""
If Err.Description <> "" Then
	w_next_url = ErrorURL
	w_msg = "<font color='#ff0000'>" _
				& "system error�\���󂲂����܂��񂪤�Z���^�[�V�X�e�����Ŏ�t���~���Ă���܂��B<br>" _
				& "���΂炭���Ă���䒍�����������B<br>" _
				& "Code: " & p_ErrorCode _
				& "</font>"
	Exit Function
End If

REM --
REM -- �戵���ʕ\��
REM --

If Err.Number <> 0 Then
	REM -- �V�X�e���G���[�����������ꍇ
	w_next_url = ErrorURL
	call card_error(SSLAuth.ErrorCode)
Else
	REM -- �V�X�e���I�ɐ���ȏꍇ
	'If SSLAuth.ErrorCode = "   " Then
	If Trim(SSLAuth.ErrorCode) = "" Then
		if Trim(SSLAuth.AuthCode) = "" then
			REM �I�[�\���擾�Ɏ��s
			w_next_url = FailureURL
			w_msg = "<font color='#ff0000'" _
						& "�J�[�h�^�M�̎擾���o���܂���ł����B<br>" _
						& "�ʂ̃J�[�h�ōēx�䒍�����������B<br>" _
						& "</font>"
		else
			REM �I�[�\���擾�ɐ���
			card_auth_no = SSLAuth.AuthCode
			card_net_no = SSLAuth.CardNetNo
			if trim(card_auth_no) = "" then			'�O�̂��߃`�F�b�N	'020924
				w_next_url = FailureURL
				w_msg = "<font color='#ff0000'" _
							& "�J�[�h�^�M�̎擾�͏o���܂������A�^�M�ԍ��擾���ɃG���[���������܂����B<br>" _
							& "���Љc�Ƃ܂ł��A����������<br>" _
							& "</font>"
			else
				w_next_url = SuccessURL
				w_msg = ""
			end if
		end if
	else
		REM �I�[�\���擾�Ɏ��s
		w_next_url = FailureURL
		call card_error(SSLAuth.ErrorCode)
	End If
End If

REM --
REM -- �Ō�ɕK���I�u�W�F�N�g��������܂��B
REM --
Set SSLAuth = Nothing

end function

'========================================================================
'
' �J�[�h���� Error
'
'========================================================================

Dim ErrorCode

Dim error_input
Dim error_card
Dim error_system
Dim error_system_l
Dim error_system_h
Dim error_package

error_input = "G65,G83"

error_card = "G12,G55,G56,G60,G61,S06"

error_system = "V12,J01,J02,J10,J11,J20,J21,J22,J30,J31,J32,S01,S02,S03," _
			 & "S04,S05,S10,S12,S13,S15,S90,S99,P01,P12,P30,P31,P50,P51," _
			 & "P52,P53,P54,P55,P65,P68,P69,P70,P71,P72,P73,P74,P75,P76," _
			 & "P78,P80,P81,P83,P84,P90,E90,K01,K02,K40,K50"
			 
error_system_l = "C00"
error_system_h = "C99"

error_package = "V01,V02,V03,V10,V11,V14,V15,V99"

'=======================================================================

'========================================================================
'
'	Function	�J�[�h�G���[���b�Z�[�W�쐬
'
'========================================================================
'
Function card_error(p_ErrorCode)

'---- set error message
'---- ���̓G���[
if InStr(error_input, p_ErrorCode) > 0 then	'input error
	w_msg = "<font color='#ff0000'>" _
				& "�J�[�h�ԍ��܂��͗L�������̓��͂Ɍ�肪����܂����B<br>" _
				& "���͓��e���m�F���Ă���ēx�䒍�����������B<br>" _
				& "Code: " & p_ErrorCode & "<br>" _
				& "�悭���邲�����<a href='http://www.soundhouse.co.jp/information/t_qanda.htm#card'>������</a>" _
				& "</font>"
	exit function
end if

'---- �J�[�h�G���[
if InStr(error_card, p_ErrorCode) > 0 then	'card error
	w_msg = "<font color='#ff0000'>" _
				& "�\���󂲂����܂��񂪤��w��̃J�[�h�ł͌䒍���ł��܂���B<br>" _
				& "�ʂ̃J�[�h�܂��ͤ�ʂ̂��x�����@�Ō䒍���肢�܂��B<br>" _
				& "Code: " & p_ErrorCode & "<br>" _
				& "�悭���邲�����<a href='http://www.soundhouse.co.jp/information/t_qanda.htm#card'>������</a>" _
				& "</font>"
	exit function
end if

'---- �p�b�P�[�W�G���[
if InStr(error_package, p_ErrorCode) > 0 then	'package error
	w_msg = "<font color='#ff0000'>" _
				& "�\���󂲂����܂��񂪤�������ɃG���[���������܂����B<br>" _
				& "�ēx�䒍�����������B�i��d�\���ɂ͂Ȃ�܂���j<br>" _
				& "Code: " & p_ErrorCode & "<br>" _
				& "�悭���邲�����<a href='http://www.soundhouse.co.jp/information/t_qanda.htm#card'>������</a>" _
				& "</font>"
	exit function
end if

'---- �V�X�e���G���[
if InStr(error_system, p_ErrorCode) > 0 then	'system error
	w_msg = "<font color='#ff0000'>" _
				& "�\���󂲂����܂��񂪤�Z���^�[�V�X�e�����Ŏ�t���~���Ă���܂��B<br>" _
				& "���΂炭���Ă���䒍�����������B<br>" _
				& "Code: " & p_ErrorCode & "<br>" _
				& "�悭���邲�����<a href='http://www.soundhouse.co.jp/information/t_qanda.htm#card'>������</a>" _
				& "</font>"
	exit function
end if

'---- �V�X�e���G���[
if (p_ErrorCode >= error_system_l) AND (p_ErrorCode <= error_system_h) then	'system error
	w_msg = "<font color='#ff0000'>" _
				& "�\���󂲂����܂��񂪤�Z���^�[�V�X�e�����Ŏ�t���~���Ă���܂��B<br>" _
				& "���΂炭���Ă���䒍�����������B<br>" _
				& "Code: " & p_ErrorCode & "<br>" _
				& "�悭���邲�����<a href='http://www.soundhouse.co.jp/information/t_qanda.htm#card'>������</a>" _
				& "</font>"
	exit function
end if

'---- ���̑��G���[
w_msg = "<font color='#ff0000'>" _
			& "�\���󂲂����܂��񂪤����̌䒍���͂��󂯂ł��܂���ł����B<br>" _
			& "�ʂ̂��x�����@�Ō䒍���肢�܂��B<br>" _
			& "Code: " & p_ErrorCode & "<br>" _
			& "�悭���邲�����<a href='http://www.soundhouse.co.jp/information/t_qanda.htm#card'>������</a>" _
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
