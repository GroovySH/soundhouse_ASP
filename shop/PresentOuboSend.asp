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
'	�v���[���g�t�H�[���̑��M
'
'�X�V����
'2005/05/13 OBJ_NewMail.BodyPart.Charset = "iso-2022-jp"���Z�b�g
'2005/09/29 ���[���T�u�W�F�N�g�ɃR���^�N�g�Ǘ��p����ǉ�
'2008/05/12 ���s�R�[�h�C���W�F�N�V�����΍�ii_to�p�����[�^�폜�j
'2008/05/13 �N���X�T�C�g���N�G�X�g�t�H�W�F���[�΍� Key�p�����[�^�`�F�b�N
'2009/04/30 �G���[����error.asp�ֈړ�
'
'========================================================================

On Error Resume Next

Dim userID
Dim msg

Dim customer_nm
Dim furigana
Dim zip
Dim prefecture
Dim address
Dim telephone
Dim e_mail
Dim purchase
Dim comment

Dim w_sql
Dim w_html
Dim w_msg

'========================================================================

Response.buffer = true

'---- �Z�L�����e�B�[�L�[�`�F�b�N
if Session("Skey") <> ReplaceInput(Request("SKey")) then
	Response.redirect "PresentOubo.asp"
end if

'---- UserID ���o��
userID = Session("userID")

'---- �Ăяo��������̃f�[�^���o��
customer_nm = ReplaceInput(Request("customer_nm"))
furigana = ReplaceInput(Request("furigana"))
zip = ReplaceInput(Request("zip"))
prefecture = ReplaceInput(Request("prefecture"))
address = ReplaceInput(Request("address"))
telephone = ReplaceInput(Request("telephone"))
e_mail = ReplaceInput(Request("e_mail"))
purchase = ReplaceInput(Request("purchase"))
comment = ReplaceInput(Request("comment"))

'---- Execute main
call main()

if Err.Description <> "" then	
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

'========================================================================
'
'	Function	Main
'
'========================================================================
'
Function main()

Dim i
Dim v_body
Dim v_item
Dim OBJ_NewMail

'---- edit body
v_body = ""
v_body = v_body & "��t�����F" & now() & vbNewLine & vbNewLine
v_body = v_body & "�ڋq�ԍ��F" & userID & vbNewLine & vbNewLine
v_body = v_body & "���O�@�@�F" & customer_nm & vbNewLine
v_body = v_body & "�ӂ肪�ȁF" & furigana & vbNewLine
v_body = v_body & "�Z���@�@�F" & zip & " " & prefecture & address & vbNewLine
v_body = v_body & "�d�b�ԍ��F" & telephone & vbNewLine
v_body = v_body & "�d���[���F" & e_mail & vbNewLine
v_body = v_body & "�w�����@�F" & purchase & vbNewLine
v_body = v_body & "�R�����g�F" & comment & vbNewLine

'@@@@@response.write(v_body)

'---- send e-mail
Set OBJ_NewMail = Server.CreateObject("CDO.Message") 

OBJ_NewMail.from = "present@soundhouse.co.jp"
OBJ_NewMail.to = "present@soundhouse.co.jp"
OBJ_NewMail.subject = "�v���[���g���� " & customer_nm & " [" & userID & "/Web-Emax/�v���[���g����]"
OBJ_NewMail.TextBody = v_body
OBJ_NewMail.BodyPart.Charset = "iso-2022-jp"

OBJ_NewMail.Send

Set OBJ_NewMail = Nothing

End function

'========================================================================
%>

<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="Shift_JIS">
<title>�v���[���g����b�T�E���h�n�E�X</title>
<!--#include file="../Navi/NaviStyle.inc"-->
</head>
<body>

<!--#include file="../Navi/NaviTop.inc"-->

<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="��������{���ł�"></a></span>
  
  <!-- �R���e���cstart -->
  <div id="globalContents">
    <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
      <p class="home"><a href="<%=g_HTTP%>"><img src="<%=g_RelLink%>images/icon_home.gif" alt="HOME"></a></p>
      <ul id="path">
        <li class="now">�v���[���g����</li>
      </ul>
    </div></div></div>

    <h1 class="title">�v���[���g����</h1>
    <p>�v���[���g�̂����������܂����B<br>���肪�Ƃ��������܂����B</p>
    
</div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>