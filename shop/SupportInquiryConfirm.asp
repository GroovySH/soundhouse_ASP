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
'	�C������i�T�|�[�g�m�F�y�[�W
'
'�X�V����
'2010/10/04 an �V�K�쐬
'2011/03/02 hn SetSecureKey�̈ʒu�ύX
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2012/06/29 if-web ���j���[�A�����C�A�E�g����
'
'========================================================================

On Error Resume Next

Dim userID

Dim MakerName
Dim ProductName
Dim Warranty
Dim SerialNo
Dim WhenPurchased
Dim Comment

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

Response.buffer = true

'---- �Z�L�����e�B�[�L�[�`�F�b�N
if Session("SKey") <> ReplaceInput(Request("SKey")) then
	Response.redirect "SupportInquiry.asp"
end if

'---- get input data
MakerName = ReplaceInput(Left(Request("MakerName"),26))
ProductName = ReplaceInput(Left(Request("ProductName"),51))
Warranty = ReplaceInput(Left(Request("Warranty"),3))
SerialNo = ReplaceInput(Left(Request("SerialNo"),41))
WhenPurchased = ReplaceInput(Left(Request("WhenPurchased"),10))
Comment = ReplaceInput(Left(Request("Comment"),501))

'---- Execute main
call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "SupportInquiryConfirm.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

call close_db()

if Err.Description <> "" OR wNoData = "Y" then
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

'---- ���̓f�[�^�ɃG���[������ꍇ�͓��͉�ʂɖ߂�
if wMSG <> "" then
	Session("msg") = wMSG
	Server.Transfer("SupportInquiry.asp")
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

'�Z�L�����e�B�[�L�[���Z�b�g
Skey = SetSecureKey()

wNoData = ""

'---- �ڋq�ԍ����o��
userID = Session("userID")

if userID = "" then
	wNoData = "Y"
	exit function
end if

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

'---- ���̓f�[�^�`�F�b�N
call validation()

end function

'========================================================================
'
'    Function    ���͓��e�`�F�b�N
'
'========================================================================
'
Function validation()

wMSG = ""

'---- �u���[�J�[�v�`�F�b�N
if MakerName ="" then
	wMSG = wMSG & "���[�J�[����͂��Ă��������B<br>"
elseif Len(MakerName) > 25 then
	wMSG = wMSG & "���[�J�[��25�����܂łł��B<br>"
end if

'---- �u���i���v�`�F�b�N
if ProductName ="" then
	wMSG = wMSG & "���i������͂��Ă��������B<br>"
elseif Len(ProductName) > 50 then
	wMSG = wMSG & "���i����50�����܂łł��B<br>"
end if

'---- �u�ۏ؏�����/�Ȃ��v�`�F�b�N
if Warranty = "" then
	wMSG = wMSG & "�ۏ؏��̂���/�Ȃ���I�����Ă��������B<br>"
elseif Warranty <> "����" AND Warranty <> "�Ȃ�" then
	wMSG = wMSG & "�ۏ؏��̂���/�Ȃ��̎w�肪�s���ł��B<br>"
end if

'---- �uSerialNo�v�`�F�b�N
if Len(SerialNo) > 40 then
	wMSG = wMSG & "�V���A���ԍ���40�����܂łł��B<br>"
end if

if cf_checkHankaku(SerialNo) = false then
	wMSG = wMSG & "�V���A���ԍ��͔��p�œ��͂��Ă��������B<br>"
end if

'---- �u�w������ԁv�`�F�b�N
if WhenPurchased ="" then
	wMSG = wMSG & "���w������Ԃ�I�����Ă��������B<br>"
elseif WhenPurchased <> "��T�Ԉȓ�" AND WhenPurchased <> "��N����" AND WhenPurchased <> "��N�ȏ�/�s��" then
	wMSG = wMSG & "���w������Ԃ̎w�肪�s���ł��B<br>"
end if

'---- �u���e�v�`�F�b�N
if Comment ="" then
	wMSG = wMSG & "���e����͂��Ă��������B<br>"
elseif Len(Comment) > 500 then
	wMSG = wMSG & "���e��500�����܂łł��B<br>"
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
<title>�C������i�T�|�[�g���e�̊m�F�b�T�E���h�n�E�X</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/inquiry.css" type="text/css">
<script type="text/javascript">
//
//    Change onClick
//
function Change_onClick(pReturnURL){
    document.f_data.action = pReturnURL;
    document.f_data.submit();
}
//
//    Store onClick
//
function Store_onClick(pSendURL){
    document.f_data.action = pSendURL;
    document.f_data.submit();
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
      <p class="home"><a href="<%=g_HTTP%>"><img src="../images/icon_home.gif" alt="HOME"></a></p>
      <ul id="path">
        <li class="now">�C������i�T�|�[�g</li>
      </ul>
    </div></div></div>

    <h1 class="title">�C������i�T�|�[�g���e�̊m�F</h1>
    <p>���e���m�F�̏�A[���M����]�{�^���������Ă��������B</p>

<!-- �G���[���b�Z�[�W -->
<% if wMSG <> "" then %>
<ul class="error">
  <li><%=wMSG %></li>
</ul>
<% end if %>

<table>
  <tr>
    <th>���[�J�[<span>*</span></th>
    <td><%=MakerName%></td>
  </tr>
  <tr>
    <th>���i��<span>*</span></th>
    <td><%=ProductName%></td>
  </tr>
  <tr>
    <th>�ۏ؏�<span>*</span></th>
    <td><%=Warranty%><% if Warranty = "�Ȃ�" then %><span>�i�ۏ؏��������ꍇ�A�ۏ؂��󂯂��Ȃ��ꍇ������܂��j</span><% end if %></td>
  </tr>
  <tr>
    <th>�V���A���ԍ�</th>
    <td><%=SerialNo%></td>
  </tr>
  <tr>
    <th>���w�������<span>*</span></th>
    <td><%=WhenPurchased%></td>
  </tr>
  <tr>
    <th>���e<span>*</span></th>
    <td><%=Comment%></td>
  </tr>
  <tr>
    <th>�����O</th>
    <td><%=wCustomerName%></td>
  </tr>
  <tr>
    <th>�Z ��</th>
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

<p>&laquo; <a href="JavaScript:Change_onClick('SupportInquiry.asp');">�ύX����</a></p>
<form name="f_data" method="post" action="SupportInquiryStore.asp">
  <input type="hidden" name="MakerName" value="<%=MakerName%>">
  <input type="hidden" name="ProductName" value="<%=ProductName%>">
  <input type="hidden" name="Warranty" value="<%=Warranty%>">
  <input type="hidden" name="SerialNo" value="<%=SerialNo%>">
  <input type="hidden" name="WhenPurchased" value="<%=WhenPurchased%>">
  <input type="hidden" name="Comment" value="<%=Comment%>">
  <input type="hidden" name="Skey" value="<%=Skey%>">
  <p class="btnBox"><input type="submit" value="���M����" class="opover"></p>
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