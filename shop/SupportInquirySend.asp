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
'	�C������i�T�|�[�g���M
'
'�X�V����
'2007/01/11 �R���^�N�g�Ǘ��p��Subject�ɃR���^�N�g�J�e�S���[/�T�u�J�e�S���[��ǉ�
'2008/04/10 SMTP Server�ύX
'2008/05/12 ���s�R�[�h�C���W�F�N�V�����΍�ii_to�p�����[�^�폜�j
'2008/05/13 �N���X�T�C�g���N�G�X�g�t�H�W�F���[�΍� Key�p�����[�^�`�F�b�N
'2010/01/12 �폜������s�R�[�h�̎w���vbCr/vbLf�ɕύX
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'
'========================================================================

On Error Resume Next

Dim userID
Dim msg

Dim CustomerName
Dim Furigana
Dim Zip
Dim Prefecture
Dim Address
Dim Telephone
Dim Fax
Dim Email

Dim MakerName
Dim ProductName
Dim Warranty
Dim SerialNo
Dim WhenPurchased
Dim Comment
Dim Bikou

Dim wItemChar1
Dim wItemChar2
Dim wItemNum1
Dim wItemNum2
Dim wItemDate1
Dim wItemDate2

Dim Connection
Dim RS

Dim wSQL
Dim wHTML
Dim wMSG
Dim wErrDesc   '2011/08/01 an add

'========================================================================

Response.buffer = true

'---- �Z�L�����e�B�[�L�[�`�F�b�N
if Session("SKey") <> ReplaceInput(Request("SKey")) then
	Response.redirect "SupportInquiry.asp"
end if

'---- UserID ���o��
userID = Session("userID")

'---- �Ăяo��������̃f�[�^���o��

MakerName = ReplaceInput(Request("MakerName"))
ProductName = ReplaceInput(Request("ProductName"))
Warranty = ReplaceInput(Request("Warranty"))
SerialNo = ReplaceInput(Request("SerialNo"))
WhenPurchased = ReplaceInput(Request("WhenPurchased"))
Comment = ReplaceInput(Request("Comment"))
Bikou = ReplaceInput(Request("Bikou"))

CustomerName = ReplaceInput(Request("CustomerName"))
Furigana = ReplaceInput(Request("Furigana"))
Zip = ReplaceInput(Request("Zip"))
Prefecture = ReplaceInput(Request("Prefecture"))
Address = ReplaceInput(Request("Address"))
Email = Replace(Replace(LCase(ReplaceInput(Request("Email"))), vbCr, ""), vbLf, "")
Telephone = ReplaceInput(Request("Telephone"))
Fax = ReplaceInput(Request("Fax"))

'---- Execute main
call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "SupportInquirySend.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

call close_db()

if Err.Description <> "" then
	Response.Redirect g_HTTP & "shop/Error.asp"
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
'	Function	Main
'
'========================================================================
'
Function main()

Dim i
Dim v_body
Dim v_body2
Dim v_subject
Dim OBJ_NewMail

'---- edit body
v_body = ""
v_body = v_body & "��t�����@�@�@�F" & now() & vbNewLine & vbNewLine
'v_body = v_body & "�����@�@�@�@�@�F�C������i�T�|�[�g(" & MakerName & "/" & ProductName & ")" & vbNewLine & vbNewLine
v_body = v_body & "���[�J�[�@�@�@�F" & MakerName & vbNewLine & vbNewLine
v_body = v_body & "���i���@�@�@�@�F" & ProductName & vbNewLine & vbNewLine
v_body = v_body & "�ۏ؏��@�@�@�@�F" & Warranty & vbNewLine & vbNewLine
v_body = v_body & "�V���A���ԍ��@�F" & SerialNo & vbNewLine & vbNewLine
v_body = v_body & "���w������ԁ@�F" & WhenPurchased & vbNewLine & vbNewLine
v_body = v_body & "���e�@�@�@�@�@�F" & Comment & vbNewLine & vbNewLine
v_body = v_body & "���l�@�@�@�@�@�F" & Bikou & vbNewLine & vbNewLine

v_body2 = ""
v_body2 = v_body2 & "���O�@�@�@�@�@�F" & CustomerName & vbNewLine
v_body2 = v_body2 & "�ӂ肪�ȁ@�@�@�F" & Furigana & vbNewLine
v_body2 = v_body2 & "�Z���@�@�@�@�@�F" & Zip & " " & Prefecture & Address & vbNewLine
v_body2 = v_body2 & "�d�b�ԍ��@�@�@�F" & Telephone & vbNewLine
v_body2 = v_body2 & "Fax �@�@�@�@�@�F" & Fax & vbNewLine
v_body2 = v_body2 & "�d���[���@�@�@�F" & Email & vbNewLine
v_body2 = v_body2 & "�ڋq�ԍ��@�@�@�F" & UserID & vbNewLine

'---- send e-mail
Set OBJ_NewMail = Server.CreateObject("CDO.message") 

OBJ_NewMail.from = "support@soundhouse.co.jp"
OBJ_NewMail.to = "support@soundhouse.co.jp"

OBJ_NewMail.subject = "�C������i�T�|�[�g(" & MakerName & "/" & ProductName & ") " &  CustomerName & "�@�l�@ [" & UserID & "/RA/Web��t]"
OBJ_NewMail.TextBody = v_body & v_body2
OBJ_NewMail.BodyPart.Charset = "iso-2022-jp"

OBJ_NewMail.Send

'---- �����ԐM���[���쐬�i�ڋq��)
call getCntlMst("Web","Email","�g���[��", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)

v_body = "���⍇�����肪�Ƃ��������܂��" & vbNewLine _
       & "�ȉ��̓��e�ɂďC������i�T�|�[�g�˗�����t�������܂����" & vbNewLine _
       & "�ԓ��܂ō����΂炭���҂����������" & vbNewLine & vbNewLine _
       & v_body & vbNewLine _
       & wItemChar1

OBJ_NewMail.from = "support@soundhouse.co.jp"
OBJ_NewMail.to = Email

OBJ_NewMail.subject = "�C������i�T�|�[�g�˗�����t�������܂���"
OBJ_NewMail.TextBody = v_body
OBJ_NewMail.BodyPart.Charset = "iso-2022-jp"

'---- ���[���T�[�o�[�w��
'OBJ_NewMail.Configuration.Fields.Item(g_ItemSMTPSendusing) = g_SMTPSendusing
'OBJ_NewMail.Configuration.Fields.Item(g_ItemSMTPServer) = g_SMTPServer
'OBJ_NewMail.Configuration.Fields.Item(g_ItemSMTPServerPort) = g_SMTPServerPort
'OBJ_NewMail.Configuration.Fields.Update

OBJ_NewMail.Send

Set OBJ_NewMail = Nothing

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

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<title>�T�E���h�n�E�X �C������i�T�|�[�g</title>

<!-- �ǉ�SCRIPT�͂�����-->

<!--#include file="../Navi/NaviStyle.inc"-->

</head>

<body background="../Navi/Images/back_ground.gif" bgcolor="#ffffff" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<!--#include file="../Navi/NaviTop.inc"-->

<table width="940" height="26" border="0" cellpadding="0" cellspacing="0">
  <tr>

<!--#include file="../Navi/NaviLefta.inc"-->

    <td width="798" align="left" valign="top" bgcolor="#ffffff">

<!------------ �y�[�W���C�������̋L�q START ------------>

      <table border="0" cellspacing="0" cellpadding="3">
        <tr>
          <td align="left"><b><font color="#696684">�C������i�T�|�[�g</font></b></td>
        </tr>
      </table>

      <table width="798" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td>&nbsp;</td>
          <td class="honbun">
            <br>
            �C������i�T�|�[�g�˗��𑗐M���܂����B<br>
            ���肪�Ƃ��������܂����B
          </td>
        </tr>
      </table>

<!------------ �y�[�W���C�������̋L�q END ------------>

    </td>
  </tr>
</table>

<!--#include file="../Navi/NaviBottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->

</body>
</html>
