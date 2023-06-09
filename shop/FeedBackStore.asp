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
'	���q�l�A���P�[�g �o�^
'
'�X�V����
'2008/05/23 ���̓f�[�^�`�F�b�N�����iLEFT��)
'2009/04/30 �G���[����error.asp�ֈړ�
'2010/07/16 st �A���P�[�g����������500�����ȓ��ɑΉ�
'2010/12/07 an ����3,4�̏��ԓ���ւ��B����4�͔C�Ӎ��ڂɕύX���t�H�[�����͕ύX���Ȃ�
'2011/02/23 hn �p�����[�^��掞�����500�����Œ����J�b�g���폜
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'
'========================================================================

On Error Resume Next

Dim userID
Dim msg

Dim OrderNo
Dim q1
Dim q1Name
Dim q1Department
Dim q1Comment
Dim q2
Dim q2Comment
Dim q3           '����4�ɊY��
Dim q3Comment    '����4�ɊY��
Dim q4           '����3�ɊY��
Dim q4Other      '����3�ɊY��
Dim q5
Dim q5Comment
Dim q6
Dim q6Other
Dim q7
Dim q7Comment

'2010 07/29 st del s
'Dim q8
'Dim q9
'2010 07/29 e

Dim Connection
Dim RS

Dim w_sql
Dim w_html
Dim w_msg
Dim wErrDesc   '2011/08/01 an add

Dim wCanWriteFl

'========================================================================

Response.buffer = true

'---- UserID ���o��
userID = Session("userID")

'---- �Ăяo��������̃f�[�^���o��
OrderNo = ReplaceInput(Request("OrderNo"))
q1 = ReplaceInput(Left(Request("q1"), 10))
q1Name = ReplaceInput(Left(Request("q1Name"), 10))
q1Department = ReplaceInput(Left(Request("q1Department"), 10))
q1Comment = ReplaceInput(Request("q1Comment"))	'2011/02/23 hn mod
q2 = ReplaceInput(Left(Request("q2"), 10))
q2Comment = ReplaceInput(Request("q2Comment"))		'2011/02/23 hn mod
q3 = ReplaceInput(Left(Request("q3"), 10))                   '����4�ɊY��
q3Comment = ReplaceInput(Request("q3Comment"))						   '����4�ɊY��			'2011/02/23 hn mod
q4 = ReplaceInput(Left(Request("q4"), 150))                  '����3�ɊY��
q4Other = ReplaceInput(Left(Request("q4Other"), 50))         '����3�ɊY��
q5 = ReplaceInput(Left(Request("q5"), 10))
q5Comment = ReplaceInput(Request("q5Comment"))		'2011/02/23 hn mod
q6 = ReplaceInput(Left(Request("q6"), 50))
q6Other = ReplaceInput(Request("q6Other"))			'2011/02/23 hn mod
q7 = ReplaceInput(Left(Request("q7"), 10))
q7Comment = ReplaceInput(Request("q7Comment"))		'2011/02/23 hn mod

'2010 07/16 st del s
'q8 = ReplaceInput(Left(Request("q8"), 10))
'q9 = ReplaceInput(Request("q9"))
'2010 07/16 e

if isNumeric(OrderNo) = true then
	'---- Execute main
	call connect_db()
	call main()
	
	'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
	if Err.Description <> "" then
		wErrDesc = "FeedBackStore.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
		call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
	end if                                           '2011/08/01 an add e

	call close_db()
	
	if Err.Description <> "" then	
		Response.Redirect g_HTTP & "shop/Error.asp"
	end if
	
	'2010 07/16 st ad s
	if w_msg <> "" then
		Session("msg") = w_msg
		Session("CanWriteFl") = wCanWriteFl
		Server.Transfer "FeedBack.asp"
	end if
	'2010 07/16 e
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
'	Function	main proc
'
'========================================================================
'
Function main()

call ValidateData()

if w_msg <> "" then
	wCanWriteFl = "Y" '------�����������K�v�ȏꍇ 2010/07/17
else
	'---- �A���P�[�g���ʓo�^
	w_sql = ""
	w_sql = w_sql & "SELECT *"
	w_sql = w_sql & "  FROM �o�׌�A���P�[�g"
	w_sql = w_sql & " WHERE �󒍔ԍ� = " & OrderNo 
		  
	Set RS = Server.CreateObject("ADODB.Recordset")
	RS.Open w_sql, Connection, adOpenStatic, adLockOptimistic

	if RS.EOF = false then
		w_msg = "�󒍔ԍ�[" &  OrderNo & "] �Ɋւ���A���P�[�g�́A���ɓo�^����Ă��܂��B"
		wCanWriteFl = "N"
	else
		'---- insert �A���P�[�g
		RS.AddNew

		RS("�󒍔ԍ�") = OrderNo
		RS("����1") = q1
		RS("����1���O") = q1Name
		RS("����1����") = q1Department
		RS("����1�ӌ�") = q1Comment
		RS("����2") = q2
		RS("����2�ӌ�") = q2Comment
		RS("����3") =q3               '����4�ɊY��
		RS("����3�ӌ�") = q3Comment   '����4�ɊY��
		RS("����4") = q4              '����3�ɊY��
		RS("����4���̑�") = q4Other   '����3�ɊY��
		RS("����5") = q5
		RS("����5�ӌ�") = q5Comment
		RS("����6") = q6
		RS("����6���̑�") = q6Other
		RS("����7") = q7
		RS("����7�ӌ�") = q7Comment
		
		'2010 07/16 st del s
'		RS("����8") = q8
'		RS("����9") = q9
		'2010 07/16 e

		RS("�o�^��") = now()

		RS.Update
	end if
	RS.close
end if

End function

'========================================================================
'
'	Function	���̓f�[�^�`�F�b�N '2010/07/16 st add
'
'========================================================================
'
Function ValidateData()

if q1 = "" then
	w_msg = w_msg & "����1�����͂���Ă��܂���B<br>"
end if

if (Len(q1Comment)) > 500 then
	w_msg = w_msg & "����1�ɓ��͂��ꂽ��������500�����𒴂��Ă��܂���@500�����ȓ��ł��肢���܂��B<br>"
end if

if q2 = "" then
	w_msg = w_msg & "����2�����͂���Ă��܂���B<br>"
end if

if (Len(q2Comment)) > 500 then
	w_msg = w_msg & "����2�ɓ��͂��ꂽ��������500�����𒴂��Ă��܂���@500�����ȓ��ł��肢���܂��B<br>"
end if

'if q3 = "" then          '2010/12/07 an del q3�͎���4�ɊY���B����4�͕K�{�łȂ���
'	w_msg = w_msg & "����4�����͂���Ă��܂���B<br>"
'end if

if (Len(q3Comment)) > 500 then
	w_msg = w_msg & "����4�ɓ��͂��ꂽ��������500�����𒴂��Ă��܂���@500�����ȓ��ł��肢���܂��B<br>"  '2010/12/07 an mod q3�͎���4�ɊY��
end if

if q5 = "" then
	w_msg = w_msg & "����5�����͂���Ă��܂���B<br>"
end if

if (Len(q5Comment)) > 500 then
	w_msg = w_msg & "����5�ɓ��͂��ꂽ��������500�����𒴂��Ă��܂���@500�����ȓ��ł��肢���܂��B<br>"
end if

if q6 = "" then
	w_msg = w_msg & "����6�����͂���Ă��܂���B<br>"
end if

if (Len(q6Other)) > 500 then
	w_msg = w_msg & "����6�ɓ��͂��ꂽ��������500�����𒴂��Ă��܂���@500�����ȓ��ł��肢���܂��B<br>"
end if

if q7= "" then
	w_msg = w_msg & "����7�����͂���Ă��܂���B<br>"
end if

if (Len(q7Comment)) > 500 then
	w_msg = w_msg & "����7�ɓ��͂��ꂽ��������500�����𒴂��Ă��܂���@500�����ȓ��ł��肢���܂��B<br>"
end if

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

'========================================================================
%>

<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="Shift_JIS">
<title>�A���P�[�g���肪�Ƃ��������܂����b�T�E���h�n�E�X</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/feedback.css" type="text/css">

</head>

<body>

<!--#include file="../Navi/NaviTop.inc"-->

<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="��������{���ł�"></a></span>
  
  <!-- �R���e���cstart -->
  <div id="globalContents" class="feedback">
    <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
      <p class="home"><a href="<%=g_HTTP%>"><img src="<%=g_RelLink%>images/icon_home.gif" alt="HOME"></a></p>
      <ul id="path">
        <li class="now">���w���җl�����A���P�[�g</li>
      </ul>
    </div></div></div>
    
    <p><strong>�A���P�[�g�ɂ����͂��肪�Ƃ��������܂����B</strong></p>
    <p>����Ƃ���낵�����������Ă��������܂��悤���肢�������܂��B</p>

</div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>
