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
<%
'========================================================================
'
'	�I�[�_�[���א��ʕύX
'
'�X�V����
'2006/06/26 �p�Տ��i�̏ꍇ�A�����\���ȏ�Ɏ󒍂��Ȃ��悤�ɕύX
'2006/10/19 ���ʂɋ󕶎�������ꂽ��G���[�Ώ�
'2011/04/14 hn SessionID�֘A�ύX
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2012/08/07 GV #1400 �J�[�g�y�[�W�Čv�Z
'
'========================================================================
'
'		�I�[�_�[���׍s�̐��ʕύX���s���
'		���ʂ�0�̏ꍇ�͖��׍s���폜�
'		�s�ԍ��w��̂Ƃ��͊Y���s�݂̂̍X�V�B(�폜�{�^��)
'		�s�ԍ���all�̂Ƃ��͑S�s�X�V�B�i�Čv�Z�{�^��)
'
'------------------------------------------------------------------------

'	�X�V����
'2008/05/23 ���̓f�[�^�`�F�b�N�����iLEFT, Numeric, EOF��)
'2009/04/30 �G���[����error.asp�ֈړ�
'
'========================================================================

On Error Resume Next

Dim userID

Dim detail_no
Dim qt(100)

Dim Connection
Dim RS

Dim w_sql
Dim w_msg
Dim w_html

Dim w_detail_cnt
Dim wErrDesc   '2011/08/01 an add

'=======================================================================

Session("msg") = ""
w_msg = ""

'---- execute main process
call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "OrderChange.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

call close_db()

if Err.Description <> "" then	
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

if w_msg <> "" then
	w_msg = "<font color='#ff0000'>" & w_msg & "</font>"
	Session("msg") = w_msg
end if

Response.Redirect "Order.asp"

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
'	Function	Main
'
'========================================================================
'
Function main()
Dim i
Dim v_item

'---- ���M�f�[�^�[�̎��o��
detail_no = ReplaceInput(Trim(Request("detail_no")))

for i=1 to 100
	v_item = "qt" & i
	qt(i) = ReplaceInput(Trim(Request(v_item)))

	if qt(i) <> "" then
		if isNumeric(qt(i)) = false then
			w_msg = "<center><font color='#ff0000'>���ʂɐ�������͂��Ă��������B</font></center>"
			exit function
		end if
		if qt(i) > 100000 then
			w_msg = "<center><font color='#ff0000'>���͂��ꂽ���ʂ��傫�����܂��B</font></center>"
			exit function
		end if
	end if
Next

'---- ���󒍖��׏��X�V

if detail_no = "all" then
	call update_all()
else
	call delete_one()
end if

End Function

'========================================================================
'
'	Function	Update ���󒍖��� �S��
'
'========================================================================
'
Function update_all()

'---- �󒍖��׏���o��
w_sql = ""
w_sql = w_sql & "SELECT a.�󒍖��הԍ�"
w_sql = w_sql & "     , a.���i�R�[�h"
w_sql = w_sql & "     , a.�󒍐���"
w_sql = w_sql & "     , a.�󒍒P��"
w_sql = w_sql & "     , a.�󒍋��z"
w_sql = w_sql & "     , b.�p�ԓ�"
w_sql = w_sql & "     , c.�����\����"
w_sql = w_sql & "  FROM ���󒍖��� a"
w_sql = w_sql & "     , Web���i b"
w_sql = w_sql & "     , Web�F�K�i�ʍ݌� c"
w_sql = w_sql & " WHERE b.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
w_sql = w_sql & "   AND b.���i�R�[�h = a.���i�R�[�h"
w_sql = w_sql & "   AND c.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
w_sql = w_sql & "   AND c.���i�R�[�h = a.���i�R�[�h"
w_sql = w_sql & "   AND c.�F = a.�F"
w_sql = w_sql & "   AND c.�K�i = a.�K�i"
w_sql = w_sql & "   AND a.SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod
	  
Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open w_sql, Connection, adOpenStatic, adLockOptimistic

'---- �󒍐��ʁC���z �X�V
Do while RS.EOF = false
	if qt(RS("�󒍖��הԍ�")) <> "" then
		if CLng(qt(RS("�󒍖��הԍ�"))) > 0 then
			if isNull(RS("�p�ԓ�")) = false AND RS("�����\����") < CLng(qt(RS("�󒍖��הԍ�"))) then
				w_msg = w_msg & RS("���i�R�[�h") & "�́A�݌ɂ�" & RS("�����\����") & "��������܂���B�@���ʂ�ύX���Ă��������������B<br>"
			else
				RS("�󒍐���") = CLng(qt(RS("�󒍖��הԍ�")))
				RS("�󒍋��z") = Fix(RS("�󒍒P��")) * CLng(qt(RS("�󒍖��הԍ�")))

				RS.Update
			end if
		else
			'2012/08/07 GV add start #1400
 			'RS.Delete
			RS("�󒍐���") = 0
			RS("�󒍋��z") = 0

			RS.Update
			'2012/08/07 GV add end   #1400 
		end if
	end if

	RS.MoveNext
Loop

Rs.Close

call delete_zero()	'2012/08/07 GV add #1400

End function

'2012/08/07 GV add start #1400
'========================================================================
'
'	Function	Delete ���󒍖��ׁi�󒍐��� = 0�j
'
'========================================================================
'
Function delete_zero()

Dim CMD
Set CMD = Server.CreateObject("ADODB.Command")
CMD.ActiveConnection = Connection

'---- ���󒍖��׍폜
w_sql = ""
w_sql = w_sql & "DELETE FROM ���󒍖���"
w_sql = w_sql & " WHERE SessionID = '" & gSessionID & "'"
w_sql = w_sql & " AND   �󒍐��� = 0"

CMD.CommandText = w_sql
CMD.Execute

Set CMD = Nothing

End function
'2012/08/07 GV add end   #1400 

'========================================================================
'
'	Function	Delete ���󒍖���
'
'========================================================================
'
Function delete_one()

if isNumeric(detail_no) = false then
	exit function
end if

'---- �󒍖��׏���o��
w_sql = ""
w_sql = w_sql & "SELECT �󒍐���"
w_sql = w_sql & "  FROM ���󒍖���"
w_sql = w_sql & " WHERE SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod
w_sql = w_sql & "   AND �󒍖��הԍ� = " & detail_no
	  
Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open w_sql, Connection, adOpenStatic, adLockOptimistic

'---- �󒍐��ʍ폜
RS.Delete
Rs.Close

End function

'========================================================================
'
'	Function	Close database
'
'========================================================================
'
Function close_db()

Connection.close
Set Connection= Nothing    '2011/08/01 an add

End function

%>
