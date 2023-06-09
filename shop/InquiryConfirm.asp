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
'	�⍇�����e�m�F�y�[�W
'     �u���͓��e�ɃG���[������/�ύX���N���b�N�v��Inquiry.asp�ɖ߂�
'     �u���M�N���b�N�v��InquirySend.asp��
'
'2011/04/13 an �V�K�쐬 #725
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2012/06/29 if-web ���j���[�A�����C�A�E�g����
'
'========================================================================
On Error Resume Next

Dim message
Dim subject
Dim ContactCategory
Dim ContactSubCategory
Dim ContactSubCategoryFl
Dim customer_nm
Dim furigana
Dim zip
Dim prefecture
Dim address
Dim telephone
Dim fax
Dim e_mail

Dim Skey
Dim Connection

Dim wMessage
Dim wMSG
Dim wErrDesc   '2011/08/01 an add

'========================================================================
Response.buffer = true

'---- �Ăяo��������̃f�[�^���o��
ContactCategory = ReplaceInput_NoCRLF(Left(Request("ContactCategory"),20))
ContactSubCategory = ReplaceInput_NoCRLF(Left(Request("ContactSubCategory"),20))
ContactSubCategoryFl = ReplaceInput_NoCRLF(Left(Request("ContactSubCategoryFl"),1))
subject = ReplaceInput_NoCRLF(Left(Request("subject"),151))
message = ReplaceInput(Left(Request("message"),2001))
customer_nm = ReplaceInput_NoCRLF(Left(Request("customer_nm"),31))
furigana = ReplaceInput_NoCRLF(Left(Request("furigana"),31))
zip = ReplaceInput_NoCRLF(Left(Request("zip"),9))
prefecture = ReplaceInput_NoCRLF(Left(Request("prefecture"),9))
address = ReplaceInput_NoCRLF(Left(Request("address"),41))
telephone = ReplaceInput_NoCRLF(Left(Request("telephone"),21))
fax = ReplaceInput_NoCRLF(Left(Request("fax"),21))
e_mail = ReplaceInput_NoCRLF(Left(Request("e_mail"),61))

'---- Execute main
call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "InquiryConfirm.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

call close_db()

if wMsg <> "" then
    Server.Transfer "Inquiry.asp"
end if

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
'	Function	main
'
'========================================================================
Function main()

'---- �Z�L�����e�B�[�L�[�Z�b�g 
Skey = SetSecureKey()

'---- ���̓`�F�b�N
call validation()

'---- ���̓`�F�b�NOK�Ȃ�m�F��ʍ쐬
if wMsg = "" then
	'���s�R�[�h��<br>�ɕϊ�
    wMessage = Replace(message, vbNewLine, "<br>")
else
	Session("Msg") = wMsg
	exit Function
end if

End function

'========================================================================
'
'    Function    �⍇�����͓��e�`�F�b�N
'
'========================================================================
'
Function validation()

Dim vAddress

wMsg = ""

'---- �u��ʁv
if ContactCategory = "" OR (ContactSubCategoryFl <> "N" AND ContactSubCategory = "" ) Then
    wMsg = wMsg & "��ʂ�I�����Ă��������B<br>"
end if

'---- �u�����v
if subject = "" Then
    wMsg = wMsg & "��������͂��Ă��������B<br>"
else
	if Len(subject) > 150 then
    	wMsg = wMsg & "������150�����ȓ��œ��͂��Ă��������B<br>"
	end if
end if

'---- �u���b�Z�[�W�v
if message = "" Then
    wMsg = wMsg & "���b�Z�[�W����͂��Ă��������B<br>"
else
	if Len(message) > 2000 then
    	wMsg = wMsg & "���b�Z�[�W��2000�����ȓ��œ��͂��Ă��������B<br>"
	end if
end if

'---- �u�����O�v
if customer_nm = "" Then
    wMsg = wMsg & "�����O����͂��Ă��������B<br>"
else
	if Len(customer_nm) > 30 then
    	wMsg = wMsg & "�����O��30�����ȓ��œ��͂��Ă��������B<br>"
	end if
end if

'---- �u�t���K�i�v
if cf_checkKataKana(furigana) = false Then
    wMsg = wMsg & "�t���K�i�͑S�p�J�i�œ��͂��Ă��������B<br>"
else
	if Len(furigana) > 30 then
    	wMsg = wMsg & "�t���K�i��30�����ȓ��œ��͂��Ă��������B<br>"
	end if
end if

'---- �u�X�֔ԍ��v
if zip <> "" then
	if IsNumeric(Replace(zip, "-", "")) = False Or cf_checkHankaku2(zip) = False Then
		wMsg = wMsg & "�X�֔ԍ��𔼊p�����ƃn�C�t��(�|)�œ��͂��Ă��������B<br>"
	else
		if Len(zip) > 10 then
	    	wMsg = wMsg & "�X�֔ԍ���10�����ȓ��œ��͂��Ă��������B<br>"
	    else
	    	if check_zip(zip, vAddress) = False Then
				wMsg = wMsg & "�X�֔ԍ����X�֔ԍ������ɂ���܂���B<br>"
			else
				'�s���{�����I������Ă���ꍇ�͕s�������Ȃ����`�F�b�N
				if prefecture <> "" then
					if InStr(vAddress, Trim(prefecture)) <= 0  Then
						wMsg = wMsg & "���͂��ꂽ�X�֔ԍ��Ɠs���{������v���܂���B<br>"
					end if
				end if
			end if
		end if
	end if
end if

'---- �u�Z���v
if Len(address) > 40 then
    wMsg = wMsg & "�Z����40�����ȓ��œ��͂��Ă��������B<br>"
end if

'---- �u�d�b�ԍ��v
if telephone = "" Then
    wMsg = wMsg & "�d�b�ԍ�����͂��Ă��������B<br>"
else

	if IsNumeric(Replace(telephone, "-", "")) = False Or cf_checkHankaku2(telephone) = False Then
		wMsg = wMsg & "�d�b�ԍ��𔼊p�����ƃn�C�t��(�|)�œ��͂��Ă��������B<br>"
	else
        if Len(telephone) > 20 then
            wMsg = wMsg & "�d�b�ԍ���20�����ȓ��œ��͂��Ă��������B<br>"
        end if
    end if
end if

'---- �uFAX�ԍ��v
if fax <> "" then
	if IsNumeric(Replace(fax, "-", "")) = False Or cf_checkHankaku2(fax) = False Then
		wMsg = wMsg & "FAX�ԍ��𔼊p�����ƃn�C�t��(�|)�œ��͂��Ă��������B<br>"
	else
		if Len(fax) > 20 then
	    	wMsg = wMsg & "FAX�ԍ���20�����ȓ��œ��͂��Ă��������B<br>"
	    end if
	end if
end if

 '---- �uE mail�v
if e_mail = "" Then
    wMsg = wMsg & "���[���A�h���X����͂��Ă��������B<br>"
else
	if Len(e_mail) > 60 then
    	wMsg = wMsg & "���[���A�h���X��60�����ȓ��œ��͂��Ă��������B<br>"
    else
		if fCheckEmail(e_mail) = false then
    		wMsg = wMsg & "���[���A�h���X���K�؂ł͂���܂���B<br>"
    	end if
    end if
end if

End function

'========================================================================
'
'	Function	�X�֔ԍ���������
'
'========================================================================
Function check_zip(vZip, vAddress)

Dim RSv
Dim vSQL

'---- �X�֔ԍ���������
vSQL = ""
vSQL = vSQL & "SELECT �s���{��������"
vSQL = vSQL & "       , �s�撬��������"
vSQL = vSQL & "       , ���於����"
vSQL = vSQL & "  FROM �X�֔ԍ����� WITH (NOLOCK)"
vSQL = vSQL & " WHERE �X�֔ԍ� = '" & Replace(vZip, "-", "") & "'"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic, adLockOptimistic

If RSv.EOF = False Then
	check_zip = True
	vAddress = Trim(RSv("�s���{��������")) & Trim(RSv("�s�撬��������"))
Else
	check_zip = False
	vAddress = ""
End If

RSv.Close

End Function

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
<title>���₢���킹���e�̊m�F�b�T�E���h�n�E�X</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/inquiry.css" type="text/css">
<script type="text/javascript">
// ======	Function:	���M�{�^��on click
function send_onClick(){
	document.f_data.submit();
}
// ======	Function:	�ύX�{�^�� on click
function return_onClick(){
	document.f_data.action = 'Inquiry.asp';
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
        <li class="now">���₢���킹���e�̊m�F</li>
      </ul>
    </div></div></div>

    <h1 class="title">���₢���킹���e�̊m�F</h1>
    <p>���₢���킹���e���m�F�̏�A[���M����]�{�^���������Ă��������B</p>
<table>
  <tr>
    <th>�� ��</th>
    <td><%=ContactCategory%><br><%=ContactSubCategory%></td>
  </tr>
  <tr>
    <th>�� ��</th>
    <td><%=subject%></td>
  </tr>
  <tr>
    <th>���b�Z�[�W</th>
    <td><%=wMessage%></td>
  </tr>
  <tr>
    <th>�����O</th>
    <td><%=customer_nm%></td>
  </tr>
  <tr>
    <th>�t���K�i</th>
    <td><%=furigana%></td>
  </tr>
  <tr>
    <th>�Z ��</th>
    <td><%=zip%><br><%=prefecture%><%=address%></td>
  </tr>
  <tr>
    <th>�d�b�ԍ�</th>
    <td><%=telephone%></td>
  </tr>
  <tr>
    <th>FAX�ԍ�</th>
    <td><%=fax%></td>
  </tr>
  <tr>
    <th>E mail</th>
    <td><%=e_mail%></td>
  </tr>
</table>

<p>&laquo; <a href="JavaScript:return_onClick();">�ύX����</a></p>
      <form name="f_data" method="post" action="InquirySend.asp">
        <input type="hidden" name="message"            value="<% = message %>">
        <input type="hidden" name="subject"            value="<% = subject %>">
        <input type="hidden" name="ContactCategory"    value="<% = ContactCategory %>">
        <input type="hidden" name="ContactSubCategory" value="<% = ContactSubCategory %>">
        <input type="hidden" name="ContactSubCategoryFl" value="<% = ContactSubCategoryFl %>">
        <input type="hidden" name="customer_nm"        value="<% = customer_nm %>">
        <input type="hidden" name="furigana"           value="<% = furigana %>">
        <input type="hidden" name="zip"                value="<% = zip %>">
        <input type="hidden" name="prefecture"         value="<% = prefecture %>">
        <input type="hidden" name="address"            value="<% = address %>">
        <input type="hidden" name="telephone"          value="<% = telephone %>">
        <input type="hidden" name="fax"                value="<% = fax %>">
        <input type="hidden" name="e_mail"             value="<% = e_mail %>">
        <input type="hidden" name="Skey"               value="<% = Skey %>">
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