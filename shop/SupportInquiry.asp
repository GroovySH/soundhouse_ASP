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
'	�C������i�T�|�[�g�y�[�W
'
'�X�V����
'2008/05/12 ���s�R�[�h�C���W�F�N�V�����΍�ii_to�p�����[�^�폜�j
'2008/05/13 �N���X�T�C�g���N�G�X�g�t�H�W�F���[�΍� Key�p�����[�^�Z�b�g
'2009/04/30 �G���[����error.asp�ֈړ�
'2010/10/04 an ���j���[�A���Ή��B�˗��f�[�^��DB�ɓo�^����悤�ɕύX
'2011/02/21 hn RtnURL�g�p����g_HTTP/g_HTTPS���g�p����悤�ɕύX�iPCIDSS)
'2011/03/02 hn SetSecureKey�̈ʒu�ύX
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2012/06/29 if-web ���j���[�A�����C�A�E�g����
'
'========================================================================

On Error Resume Next

Dim userID

Dim MakerName       '2010/10/4 an add
Dim ProductName     '2010/10/4 an add
Dim Warranty        '2010/10/4 an add
Dim SerialNo        '2010/10/4 an add
Dim WhenPurchased   '2010/10/4 an add
Dim Comment         '2010/10/4 an add

Dim wCustomerName
Dim wZip
Dim wPrefecture
Dim wAddress
Dim wTelephone
Dim wFax
Dim wEmail

Dim Skey

Dim Connection
Dim RS

Dim wSQL
Dim wHTML
Dim wMSG
Dim wNoData
Dim wErrDesc   '2011/08/01 an add

'========================================================================

'---- �Ăяo�����v���O��������̃G���[���b�Z�[�W���o��  '2010/10/4 an add
wMSG = Session("msg")
Session("msg") = ""

'---- �ڋq�ԍ����o��
userID = Session("userID")

'---- �G���[���͓��̓f�[�^���󂯎���čĕ\��   '2010/10/4 an add
MakerName = ReplaceInput(Left(Request("MakerName"),25))
ProductName = ReplaceInput(Left(Request("ProductName"),50))
Warranty = ReplaceInput(Left(Request("Warranty"),2))
SerialNo = ReplaceInput(Left(Request("SerialNo"),40))
WhenPurchased = ReplaceInput(Left(Request("WhenPurchased"),10))
Comment = ReplaceInput(Left(Request("Comment"),500))

'---- ���O�C�����Ă��Ȃ���΃��O�C����ʂ�
if userID = "" then
	Response.Redirect g_HTTPS & "shop/LoginCheck.asp?RtnURL=" & g_HTTPS & "shop/SupportInquiry.asp"	'2011/02/21 hn mod
end if

'---- Execute main
call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "SupportInquiry.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

call close_db()

if Err.Description <> "" OR wNoData = "Y" then  'LoginCheck���Ă���͂��Ȃ̂Ōڋq��񂪎擾�ł��Ȃ���΃G���[
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
'	Function	main
'
'				userID��Cookie�ɂ���Ή������\��
'
'========================================================================
Function main()

'---- �Z�L�����e�B�[�L�[�Z�b�g 
Skey = SetSecureKey()

wNoData = ""

'--------- select customer
wSQL = ""
wSQL = wSQL & "SELECT a.�ڋq�ԍ�"
wSQL = wSQL & "     , a.�ڋq��"
wSQL = wSQL & "     , a.�ڋqE_mail1"
wSQL = wSQL & "     , b.�ڋq�X�֔ԍ�"
wSQL = wSQL & "     , b.�ڋq�s���{��"
wSQL = wSQL & "     , b.�ڋq�Z��"
wSQL = wSQL & "     , c.�ڋq�d�b�ԍ�"
wSQL = wSQL & "  FROM Web�ڋq a WITH (NOLOCK)"
wSQL = wSQL & "     , Web�ڋq�Z�� b WITH (NOLOCK)"
wSQL = wSQL & "     , Web�ڋq�Z���d�b�ԍ� c WITH (NOLOCK)"
wSQL = wSQL & " WHERE b.�ڋq�ԍ� = a.�ڋq�ԍ�" 
wSQL = wSQL & "   AND c.�ڋq�ԍ� = b.�ڋq�ԍ�" 
wSQL = wSQL & "   AND c.�Z���A�� = b.�Z���A��" 
wSQL = wSQL & "   AND b.�Z���A�� = 1" 
wSQL = wSQL & "   AND c.�d�b�A�� = 1" 
wSQL = wSQL & "   AND a.�ڋq�ԍ� = " & userID 
		
'@@@@@response.write(wSQL & "<BR>")

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

if RS.EOF = true then
	wNoData = "Y"
	exit function
else
	wCustomerName = RS("�ڋq��")
	wZip = RS("�ڋq�X�֔ԍ�")
	wPrefecture = RS("�ڋq�s���{��")
	wAddress = RS("�ڋq�Z��")
	wTelephone = RS("�ڋq�d�b�ԍ�")
	wEmail = RS("�ڋqE_mail1")
end if

RS.close

end function

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
<title>�C������i�T�|�[�g�b�T�E���h�n�E�X</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/inquiry.css" type="text/css">
<script type="text/javascript">
//
// ====== 	Function:	check if some data was entered other than spaces
//		Parm:		p_val		Check value
//		Return value:	If entered --> True,  Not entered --> False
//
function check_required(p_val){
	if (p_val == ""){return(false);}
	for(i=0; i<p_val.length; i++){
		if (p_val.substring(i, i+1)!=" " && p_val.substring(i, i+1)!="�@"){
			return(true);
		}
	}
	return(false);
}
//
// ====== 	Function:	post on submit
//
function post_onSubmit(){
	var vChecked = false;
	if (check_required(document.f_data.MakerName.value) == false){
		alert("\n���[�J�[����͂��Ă��������B");
		document.f_data.MakerName.focus();
		return false;
 	}
	if (check_required(document.f_data.ProductName.value) == false){
		alert("\n���i������͂��Ă��������B");
		document.f_data.ProductName.focus();
		return false;
 	}
	if ((document.f_data.Warranty[0].checked == false) && (document.f_data.Warranty[1].checked == false)){
		alert("\n�ۏ؏��̂���/�Ȃ����`�F�b�N���Ă��������B");
		return false;
 	}
	if (document.f_data.WhenPurchased[0].selected == true){
		alert("\n���w������Ԃ�I�����Ă��������B");
		return false;
 	}
	if (check_required(document.f_data.Comment.value) == false){
		alert("\n���e����͂��Ă��������B");
		document.f_data.Comment.focus();
		return false;
 	}
	return true;
}
//========================================================================
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
        <li class="now">�C������i�T�|�[�g</li>
      </ul>
    </div></div></div>

    <h1 class="title">�C������i�T�|�[�g</h1>

<!-- �G���[���b�Z�[�W -->
<% if wMSG <> "" then %>
<ul class="error">
  <li><%=wMSG %></li>
</ul>
<% end if %>

<form name="f_data" id="inquiry" action="SupportInquiryConfirm.asp" method="post" onSubmit="return post_onSubmit();">
  <table>
    <tr>
      <th>���[�J�[<span>*</span></th>
      <td><input type="text" name="MakerName" value="<%=MakerName%>" size="50" maxlength="25"></td>
    </tr>
    <tr>
      <th>���i��<span>*</span></th>
      <td><input type="text" name="ProductName" value="<%=ProductName%>" size="65" maxlength="50"></td>
    </tr>
    <tr>
      <th>�ۏ؏�<span>*</span></th>
      <td>
        <label><input type="radio" name="Warranty" value="����"<% if Warranty = "����" then %> checked="checked"<% end if %>>����</label>�@
        <label><input type="radio" name="Warranty" value="�Ȃ�"<% if Warranty = "�Ȃ�" then %> checked="checked"<% end if %>>�Ȃ�</label>�@<span>(�ۏ؏��������ꍇ�A�ۏ؂��󂯂��Ȃ��ꍇ������܂�)</span>
      </td>
    </tr>
    <tr>
      <th>�V���A���ԍ�</th>
      <td><input name="SerialNo" type="text" value="<%=SerialNo%>" size="60" maxlength="40"><br><span>(�ۏ؏��A�{�̂ɋL�ڂ��Ȃ��ꍇ�͕K�v����܂���)</span></td>
    </tr>
    <tr>
      <th>���w�������<span>*</span></th>
      <td>
        <select name="WhenPurchased">
          <option value=""<% if WhenPurchased = "" then%> selected="selected"<% end if %>>�I�����Ă������� 
          <option value="��T�Ԉȓ�"<% if WhenPurchased = "��T�Ԉȓ�" then%> selected="selected"<% end if %>>��T�Ԉȓ� 
          <option value="��N����"<% if WhenPurchased = "��N����" then%> selected="selected"<% end if %>>��N����
          <option value="��N�ȏ�/�s��"<% if WhenPurchased = "��N�ȏ�/�s��" then%> selected="selected"<% end if %>>��N�ȏ�/�s��
        </select>
      </td>
    </tr>
    <tr>
      <th>���e<span>*</span><br>�i500�����܂Łj</th>
      <td><textarea name="Comment" rows="5" cols="55"><%=Comment%></textarea></td>
    </tr>
    <tr>
      <th>�����O</th>
      <td><%=wCustomerName%></td>
    </tr>
    <tr>
      <th>�Z��</th>
      <td>
        ��<%=wZip%><br>
        <%=wPrefecture%><%=wAddress%>
      </td>
    </tr>
    <tr>
      <th>�d�b�ԍ�</th>
      <td><%=wTelephone%></td>
    </tr>
    <tr>
      <th>Fax�ԍ�</th>
      <td><%=wFax%></td>
    </tr>
    <tr>
      <th>���[���A�h���X</th>
      <td><%=wEmail%></td>
    </tr>
  </table>
  <p>�u*�v�̂��Ă��鍀�ڂ͕K�{���͍��ڂł��B</p>
  <input type="hidden" name="Skey" value="<%=Skey%>">
  <p class="btnBox"><input type="submit" value="���e���m�F����" class="opover"></p>
</form>

</div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>