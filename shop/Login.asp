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
'	���O�C���y�[�W
'
'	�X�V����
'2004/12/20 �Ăяo������URL�ݒ�ǉ�
'2006/09/18 LoginFl�ǉ��@���O�C�����ێ�
'2007/08/13 �G���[���b�Z�[�W�\���ύX
'2008/05/12 �p�X���[�h���Z�b�g��HTTPS�ɕύX
'2008/05/14 HTTPS�`�F�b�N�Ή�
'2010/07/30 st RtnURL������ꍇ�͂��̂܂܌Ăяo�����փ��_�C���N�g'
'2011/04/20 an #843 ���O�C�����AEmail�̑���Ƀ��[�U�[ID���g�p
'2011/05/09 an #843�֘A �u���[�U�[ID��Y�ꂽ���́v�ǉ�
'2011/05/11 an �u�p�X���[�h��Y�ꂽ���́v�̓��[�U�[ID/�d�b�ԍ��̓��͂ɕύX

'========================================================================

Dim member_email  '2011/04/20 an del, 2011/05/09 an re-add
Dim telephone     '2011/05/09 an add
Dim MemberID      '2011/05/11 an add
Dim telephone_password  '2011/05/11 an add
Dim msg

Dim called_from
Dim logoff_fl
Dim userID
Dim RtnURL

Dim w_html
Dim w_msg

'========================================================================

'gHTTPSPage = true		'HTTPS�y�[�W

Response.buffer = true

'---- �Ăяo�����v���O��������̃��b�Z�[�W���o��

msg = Session.contents("msg")
Session("msg") = ""
'userID = Session("userID")

called_from = ReplaceInput(Request("called_from"))
logoff_fl = ReplaceInput(Request("logoff_fl"))
RtnURL = replace(ReplaceInput(Request("RtnURL")), "��", "&")		'�Ăяo����URL '2010/07/30 st mod
member_email = ReplaceInput_NoCRLF(Left(Request("member_email"),60))  '2011/05/09 an add �G���[���Ɏ󂯎��
telephone = ReplaceInput_NoCRLF(Left(Request("telephone"),20))        '2011/05/09 an add
MemberID = ReplaceInput_NoCRLF(Left(Request("MemberID"),60))      '2011/05/11 an add �G���[���Ɏ󂯎��
telephone_password = ReplaceInput_NoCRLF(Left(Request("telephone_password"),20)) '2011/05/11 an add

if logoff_fl = "Y" then
	'---- �ڋq�ԍ�, �ڋq��Cookie���N���A
	Session("userID") = ""
	Session("userName") = ""
	'Session("userEmail") = ""  '2011/04/20 an del
	Session("LoginFl") = ""

	Response.Redirect g_HTTP		' Log out ��@TOP�֖߂�
'else
	'---- �ȑO�ɓ��͂���Email�����o��
	'member_email = Session.contents("member_email")
	'if userID <> "" AND member_email = "" then
	'	member_email = Session("userEmail")
	'end if

	'Session("member_email") = ""
end if

'========================================================================

%>

<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="Shift_JIS">
<title>���O�C���b�T�E���h�n�E�X</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link href="<%=g_HTTPS%>/shop/style/login.css?20120718" rel="stylesheet" type="text/css">

<script type="text/javascript">
//
//	Login onSubmit
//
function Login_onSubmit(p_form){

	if (p_form.MemberID.value == ""){
		alert("���[�U�[ID����͂��Ă��������B");
		p_form.MemberID.focus();
		return false;
	}
	if (p_form.member_password.value == ""){
		alert("�p�X���[�h����͂��Ă��������B");
		p_form.member_password.focus();
		return false;
	}
		return true;
}

//
//	Password onSubmit
//
function Password_onSubmit(p_form){

	if (p_form.MemberID.value == ""){
		alert("���[�U�[ID����͂��Ă��������B");
		p_form.MemberID.focus();
		return false;	
	}
	if (p_form.telephone_password.value == ""){
		alert("�d�b�ԍ�����͂��Ă��������B");
		p_form.telephone_password.focus();
		return false;
	}
		return true;
}

//
//	UserID onSubmit
//
function UserID_onSubmit(p_form){

	if (p_form.member_email.value == ""){
		alert("���[���A�h���X����͂��Ă��������B");
		p_form.member_email.focus();
		return false;	
	}
	if (p_form.telephone.value == ""){
		alert("�d�b�ԍ�����͂��Ă��������B");
		p_form.telephone.focus();
		return false;
	}
		return true;
}

</script>

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
        <li class="now">���O�C��</li
      ></ul>
    </div></div></div>


    <h1 class="title">���O�C��</h1>

	<% if msg <> "" then %>
  	<p class="error"><%=msg%></p>
	<% end if %>
    
    <ul id="login">
      <li>
      	<form name="fLogin" method="post" action="<%=g_HTTPS%>shop/LoginCheck.asp" onSubmit="return Login_onSubmit(this);">
    	<h2>����o�^������Ă����</h2>
    	<p>���[�U�[ID�E�p�X���[�h����͂�[���O�C��]�{�^���������Ă��������B<br><a href="#login00">�����O�C���ł��Ȃ��ꍇ�͂�����</a></p>
        <table class="form">
            <tr>
              <th>���[�U�[ID</th>
              <td><input name="MemberID" type="text">���p�p����</td>
            </tr>
            <tr>
              <th>�p�X���[�h</th>
              <td><input name="member_password" type="password">���p�p����</td>
            </tr>
          </table>
          <p class="btnBox"><input type="submit" value="���O�C��" class="opover"></p>
          <input name="called_from" type="hidden" value="<%=called_from%>">
          <input name="RtnURL" type="hidden" value="<%=RtnURL%>">
        </form>
      </li>
      <li>
    	<h2>����o�^������Ă��Ȃ���</h2>
    	<p>����o�^�͖����ŃJ���^���ł��I���o�^����������΁A������Z���̓��͂��K�v����܂���B<br>�܂��AWEB��������̃��[���j���[�X�������ȏ�񂪂����ς��ł��B</p>
        <p class="btnBox"><a href="<%=g_HTTPS%>Member/MemberAgreement.asp?called_from=navi" class="opover">���o�^�͂�����</a></p>
      </li>
      <li class="forget">
    	<h3>�p�X���[�h��Y�ꂽ����</h3>
        <form name="fForgotPassword" method="post" action="<%=g_HTTPS%>Member/MemberPasswordSend.asp?called_from=<%=called_from%>" onSubmit="return Password_onSubmit(this);">
    	<p>�p�X���[�h��Y�ꂽ���́A���o�^�̃��[�U�[ID�E�d�b�ԍ�����͂�[�p�X���[�h���Z�b�g]�{�^���������Ă��������B<br>���o�^�̃��[���A�h���X���Ƀ��[�����t����܂��̂ŁA���ē����e�����m�F���������B</p>
        <table class="form">
            <tr>
              <th>���[�U�[ID</th>
              <td><input name="MemberID" type="text" value="<%=MemberID%>">���p�p����</td>
            </tr>
            <tr>
              <th>�d�b�ԍ�</th>
              <td><input name="telephone_password" type="text" value="<%=telephone_password%>">���p����</td>
            </tr>
          </table>
          <p class="btnBox"><input type="submit" value="�p�X���[�h���Z�b�g" class="opover"></p>
          <input name="called_from" type="hidden" value="<%=called_from%>">
          <input name="i_function" type="hidden" value="send">
        </form>
      </li>
      <li class="forget">
    	<h3>���[�U�[ID��Y�ꂽ����</h3>
        <form name="fForgotUserID" method="post" action="<%=g_HTTPS%>Member/MemberUserIDSend.asp?called_from=<%=called_from%>" onSubmit="return UserID_onSubmit(this);">
    	<p>���[�U�[ID��Y�ꂽ���́A���o�^�̃��[���A�h���X�E�d�b�ԍ�����͂�[���[�U�[ID�m�F]�{�^���������Ă��������B<br>���o�^�̃��[���A�h���X���Ƀ��[�U�[ID�����m�点���܂��̂ŁA���m�F���������B</p>
        <table class="form">
            <tr>
              <th>���[���A�h���X</th>
              <td><input name="member_email" type="text" value="<%=member_email%>">���p�p����</td>
            </tr>
            <tr>
              <th>�d�b�ԍ�</th>
              <td><input name="telephone" type="text" value="<%=telephone%>">���p����</td>
            </tr>
          </table>
          <p class="btnBox"><input type="submit" value="���[�U�[ID�m�F" class="opover"></p>
        </form>
      </li>
    </ul>
    
    <div id="login00">
    	<h4>���O�C���ł��Ȃ��ꍇ��</h4>
        <p>���O�C���ł��Ȃ��ꍇ�A�ȉ��̍��ڂ����m�F���������B</p>
        <h5>�y�G���[���b�Z�[�W�z</h5>
        <p>��ʏ㕔�ɐԎ��̃G���[���b�Z�[�W���\�������ꍇ�́A���b�Z�[�W���e�ɏ]�������͓��e�̏C�������肢���܂��B</p>
        <h5>�y�����ݒ�z</h5>
        <p>���g���̃p�\�R���̓����ݒ肪�����������ɂȂ��Ă��邩���m�F���������B</p>
        <h5>�y�u���E�U�̍ċN���z</h5>
        <p>���g���̃u���E�U�̊J���Ă���S�ẴE�C���h�E����U���A�ēx�J���Ă��m�F���������B</p>
        <h5>�y�N�b�L�[�̍폜�z</h5>
        <p>��L���m�F���Ă����O�C�����ł��Ȃ��ꍇ�́A�N�b�L�[�̃N���A�����������������B</p>
        <ul>
        	<li><a href="<%=g_HTTP%>guide/qanda14.asp#ie">Internet Explorer�������p�̕�</a></li>
            <li><a href="<%=g_HTTP%>guide/qanda14.asp#ff">Firefox�������p�̕�</a></li>
            <li><a href="<%=g_HTTP%>guide/qanda14.asp#sf">Safari�������p�̕�</a></li>
            <li><a href="<%=g_HTTP%>guide/qanda14.asp#cr">Chrome�������p�̕�</a></li>
        </ul>
        <p>���L���b�V��/�N�b�L�[�̃N���A��̓u���E�U���ċN�����Ă��������B</p>
        <h5>�y���₢���킹��z</h5>
        <p>���O�C���ɂ��Ă̂��₢���킹�́A<a href="<%=g_HTTPS%>shop/Inquiry.asp">���₢���킹�t�H�|��</a>�܂��͉��L�ւ��肢�������܂��B</p>
        <ul>
        	<li>TEL�F0476-89-1111</li>
            <li>FAX�F0476-89-2222</li>
            <li>MAIL�F<a href="mailto:shop@soundhouse.co.jp">shop@soundhouse.co.jp</a></li>
        </ul>
    </div>


</div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>