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
'	�⍇���y�[�W
'
'�X�V����
'2005/01/31 �p�����[�^��n���ꂽ���́A�����֎����\��
'2005/05/24 Submit��<a>����<input type="image" �ɕύX
'2005/08/18 �R���^�N�g�J�e�S���[�A�R���^�N�g�T�u�J�e�S���[��ǉ�
'2005/09/29 �R���^�N�g�J�e�S���[(Web-Emax)�͑ΏۊO
'2006/01/09 Emal�`�F�b�N����
'2007/10/19 �n�b�J�[�Z�[�t�Ή�
'2008/05/12 ���s�R�[�h�C���W�F�N�V�����΍�ii_to�p�����[�^�폜�j
'2008/05/13 �N���X�T�C�g���N�G�X�g�t�H�W�F���[�΍� Key�p�����[�^�Z�b�g
'2009/04/30 �G���[����error.asp�ֈړ�
'2011/03/02 hn SetSecureKey�̈ʒu�ύX
'2011/04/13 an #725 �m�F���(InquiryConfirm)�ǉ��A�G���[���b�Z�[�W�\���Ή�
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2012/06/29 if-web ���j���[�A�����C�A�E�g����
'
'========================================================================

On Error Resume Next

Dim userID

Dim CategoryNm
Dim MakerNm
Dim ProductCd
'Dim wSubject     '2011/04/13 an del
Dim wCategoryListHTML

Dim ContactCategory     '2011/04/13 an add
Dim ContactSubCategory  '2011/04/13 an add
Dim subject             '2011/04/13 an add
Dim message             '2011/04/13 an add
Dim customer_nm
Dim furigana
Dim zip
Dim prefecture
Dim address
Dim telephone
Dim fax
Dim e_mail

'Dim Skey   '2011/04/13 an del

Dim Connection
Dim RS

Dim w_sql
Dim w_error_msg
Dim wHTML
Dim wMsg
Dim wErrDesc   '2011/08/01 an add

'========================================================================

wMsg = ""

'---- Get Session data  2011/04/13 an add
wMsg = Session("msg")
Session("msg") = ""

'---- Get input data
CategoryNm = ReplaceInput(Trim(Request("CategoryNm")))
MakerNm = ReplaceInput(Trim(Request("MakerNm")))
ProductCd = ReplaceInput(Trim(Request("ProductCd")))

'---- ���̓`�F�b�N�G���[����InquiryConfirm.asp����f�[�^�󂯎��  2011/04/13 an add s
ContactCategory = ReplaceInput(Left(Request("ContactCategory"),20))
ContactSubCategory = ReplaceInput(Left(Request("ContactSubCategory"),20))
subject = ReplaceInput(Left(Request("subject"),150))
message = ReplaceInput(Left(Request("message"),2000))
customer_nm = ReplaceInput(Left(Request("customer_nm"),30))
furigana = ReplaceInput(Left(Request("furigana"),30))
zip = ReplaceInput(Left(Request("zip"),8))
prefecture = ReplaceInput(Left(Request("prefecture"),8))
address = ReplaceInput(Left(Request("address"),40))
telephone = ReplaceInput(Left(Request("telephone"),20))
fax = ReplaceInput(Left(Request("fax"),20))
e_mail = ReplaceInput_NoCRLF(Left(Request("e_mail"),60))  '2011/04/13 an add e

'wSubject = ""  '2011/04/13 an del
if (CategoryNm <> "") OR (MakerNm <> "" ) OR (ProductCd <> "") then
	'wSubject = "(" & MakerNm & "/" & ProductCd & ")"   '2011/04/13 an del
	subject = "�i" & MakerNm & "/" & ProductCd & "�j"
end if

'---- Execute main
call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "Inquiry.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
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
'	Function	main
'
'				userID��Cookie�ɂ���Ή������\��
'
'========================================================================
Function main()

'---- �Z�L�����e�B�[�L�[�Z�b�g 
'Skey = SetSecureKey() '2011/04/13 an del

'---- �R���^�N�g�J�e�S���[�A�R���^�N�g�T�u�J�e�S���ꗗ�쐬
call CreateCategoryListHTML()

if wMsg = "" then  '���O�C�����Ă���Όڋq�����o��(InquiryConfirm.asp����߂����ꍇ�͓��̓f�[�^��\��) 2011/04/13 an add

    userID = Session("userID")

    if userID = "" then
    	exit function
    end if

    '--------- select customer
    w_sql = ""
    w_sql = w_sql & "SELECT a.�ڋq�ԍ�"
    w_sql = w_sql & "     , a.�ڋq��"
    w_sql = w_sql & "     , a.�ڋq�t���K�i"
    w_sql = w_sql & "     , a.�ڋqE_mail1"
    w_sql = w_sql & "     , b.�ڋq�X�֔ԍ�"
    w_sql = w_sql & "     , b.�ڋq�s���{��"
    w_sql = w_sql & "     , b.�ڋq�Z��"
    w_sql = w_sql & "     , c.�ڋq�d�b�ԍ�"
    w_sql = w_sql & "  FROM Web�ڋq a WITH (NOLOCK)"
    w_sql = w_sql & "     , Web�ڋq�Z�� b WITH (NOLOCK)"
    w_sql = w_sql & "     , Web�ڋq�Z���d�b�ԍ� c WITH (NOLOCK)"
    w_sql = w_sql & " WHERE b.�ڋq�ԍ� = a.�ڋq�ԍ�" 
    w_sql = w_sql & "   AND c.�ڋq�ԍ� = b.�ڋq�ԍ�" 
    w_sql = w_sql & "   AND c.�Z���A�� = b.�Z���A��" 
    w_sql = w_sql & "   AND b.�Z���A�� = 1" 
    w_sql = w_sql & "   AND c.�d�b�A�� = 1" 
    w_sql = w_sql & "   AND a.�ڋq�ԍ� = " & userID 
    		
    '@@@@@response.write(w_sql & "<BR>")

    Set RS = Server.CreateObject("ADODB.Recordset")
    RS.Open w_sql, Connection, adOpenStatic

    '-------- Move data to work area
    if RS.EOF = true then
    	exit function
    else
    	customer_nm = RS("�ڋq��")
    	furigana = RS("�ڋq�t���K�i")
    	zip = RS("�ڋq�X�֔ԍ�")
    	prefecture = RS("�ڋq�s���{��")
    	address = RS("�ڋq�Z��")
    	telephone = RS("�ڋq�d�b�ԍ�")
    	e_mail = RS("�ڋqE_mail1")
    end if

    RS.close
end if    '2011/04/13 an add

end function

'========================================================================
'
'	Function	CreateCategoryListHTML
'
'========================================================================
Function CreateCategoryListHTML()

Dim RSv
Dim vBreakKey1
Dim vBreakNextKey1
Dim vRecCount

'--------- �R���^�N�g�J�e�S���[�A�R���^�N�g�T�u�J�e�S������o��
w_sql = ""
w_sql = w_sql & "SELECT a.�R���^�N�g�J�e�S���[��"
w_sql = w_sql & "     , b.�R���^�N�g�T�u�J�e�S���[��"
w_sql = w_sql & "  FROM �R���^�N�g�J�e�S���[ a WITH (NOLOCK)"
w_sql = w_sql & "       LEFT JOIN �R���^�N�g�T�u�J�e�S���[ b WITH (NOLOCK)"
w_sql = w_sql & "              ON b.�R���^�N�g�J�e�S���[�R�[�h = a.�R���^�N�g�J�e�S���[�R�[�h" 
w_sql = w_sql & " WHERE a.Web��\���t���O != 'Y'"
w_sql = w_sql & " ORDER BY"
w_sql = w_sql & "       a.�R���^�N�g�J�e�S���[�R�[�h" 
w_sql = w_sql & "     , b.�R���^�N�g�T�u�J�e�S���[�R�[�h" 
		
'@@@@@response.write(w_sql & "<BR>")

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open w_sql, Connection, adOpenStatic

'-------- �ꗗ�쐬
if RSv.EOF = true then
	exit function
end if

vBreakNextKey1 = RSv("�R���^�N�g�J�e�S���[��")
vBreakKey1 = vBreakNextKey1

wHTML = ""
vRecCount = 0

'---- Main loop
Do Until vBreakNextKey1 = "@EOF"
	vRecCount = vRecCount + 1
	'---- �J�e�S���[���W�I�{�^��
	wHTML = wHTML & "          <li>" & vbNewLine
	wHTML = wHTML & "            <input type=""radio"" name=""ContactCategorySel"" value=""" & RSv("�R���^�N�g�J�e�S���[��") & """"

	If (CategoryNm = "") And (vRecCount = 1) Then
		wHTML = wHTML & " checked=""checked"""
	End If

	wHTML = wHTML & " id=""type_" & vRecCount & """>"
	wHTML = wHTML & "<label for=""type_" & vRecCount & """>" & RSv("�R���^�N�g�J�e�S���[��") & "</label>" & vbNewLine

	If IsNull(RSv("�R���^�N�g�T�u�J�e�S���[��")) = False Then
		wHTML = wHTML & "            <select name=""ContactSubCategorySel"" id=""subcategory" & vRecCount - 1 & """ onChange=""SubCategory_onChange('" & RSv("�R���^�N�g�J�e�S���[��") & "')"">" & vbNewLine
		wHTML = wHTML & "              <option value="""">�I�����Ă�������" & vbNewLine

		vBreakKey1 = vBreakNextKey1

		Do Until vBreakKey1 <> vBreakNextKey1      '�J�e�S���[�u���[�N�܂�
			'---- �T�u�J�e�S���[SELECT OPTIONS
			wHTML = wHTML & "              <option value=""" & RSv("�R���^�N�g�T�u�J�e�S���[��") & """"
			If CategoryNm = RSv("�R���^�N�g�T�u�J�e�S���[��") Then
				wHTML = wHTML & " SELECTED"
			End If
			wHTML = wHTML & ">" & RSv("�R���^�N�g�T�u�J�e�S���[��") & vbNewLine

			RSv.MoveNext
			If RSv.EOF = True Then
				vBreakNextKey1 = "@EOF"
			Else
				vBreakNextKey1 = RSv("�R���^�N�g�J�e�S���[��")
			End If
		Loop

		wHTML = wHTML & "            </select>" & vbNewLine
		wHTML = wHTML & "          </li>" & vbNewLine
	Else
		wHTML = wHTML & "          </li>" & vbNewLine
		RSv.MoveNext
		If RSv.EOF = True Then
			vBreakNextKey1 = "@EOF"
		Else
			vBreakNextKey1 = RSv("�R���^�N�g�J�e�S���[��")
		End If
	End If

Loop

RSv.Close

wCategoryListHTML = wHTML

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
<title>���₢���킹�b�T�E���h�n�E�X</title>
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

//=====================================================================
//	�Z������ onClick
//=====================================================================
function address_search_onClick(){

	if (document.f_data.zip.value == ""){
		alert("�X�֔ԍ�����͂��Ă��������B");
		return;
	}
 
	AddrWin = window.open("../comasp/Address_search.asp?zip=" + document.f_data.zip.value +"&name_prefecture=i_selected_prefecture&name_address=address","AddrSearch","width=200,height=100");
}

//
// ====== 	Function:	next on submit
//
function next_onSubmit(){
	for (var i=0; i<document.f_data.ContactCategorySel.length; i++){
		if (document.f_data.ContactCategorySel[i].checked == true){
			document.f_data.ContactCategory.value = document.f_data.ContactCategorySel[i].value;
			
			var subcategory = document.getElementById('subcategory' + i);
			if (subcategory != null){  //�T�u�J�e�S���[�L�̏ꍇ
				for (var j=0; j<subcategory.options.length; j++){
					if (subcategory.options[j].selected == true){
					    document.f_data.ContactSubCategory.value = subcategory.options[j].value;
						break;
					}
				}
			}else{
				document.f_data.ContactSubCategory.value = "";
				document.f_data.ContactSubCategoryFl.value = "N";
			}
			break;
		}
	}
	return;
}

//
// ======	Function:	���W�I�{�^���A�h���b�v�_�E�����X�g���ȑO�ɑI��������Ԃɂ���
//
function preset_values(){

    // ��ʃJ�e�S���[
	for (var i=0; i<document.f_data.ContactCategorySel.length; i++){
		if (document.f_data.ContactCategorySel[i].value == document.f_data.ContactCategory.value){
			document.f_data.ContactCategorySel[i].checked = true;
			//��ʃT�u�J�e�S���[
			if (document.f_data.ContactSubCategory.value != "" ){
			
				var subcategory = document.getElementById('subcategory' + i);
				for (var j=0; j<subcategory.options.length; j++){
			        if (subcategory.options[j].value == document.f_data.ContactSubCategory.value){
				        subcategory.options[j].selected = true;
				        break;
			        }
		        }
		    }
		break;
		}
	}

    // �s���{��
	for (var i=0; i<document.f_data.prefecture.options.length; i++){
		if (document.f_data.prefecture.options[i].value == document.f_data.i_selected_prefecture.value){
			document.f_data.prefecture.options[i].selected = true;
			break;
		}
	}
	return;
}

//
// Function: �T�u�J�e�S���[�ύX���ɐe�J�e�S���[��I��
//
function SubCategory_onChange(pSubCategoryValue){

	for (var i=0; i<document.f_data.ContactCategorySel.length; i++){
		if (document.f_data.ContactCategorySel[i].value == pSubCategoryValue){
			document.f_data.ContactCategorySel[i].checked = true;
			break;
		}
	}
	return;
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
        <li class="now">���₢���킹</li>
      </ul>
    </div></div></div>

    <h1 class="title">���₢���킹</h1>
    <div id="notice">
      <p>���₢���킹���������O�ɂ悭���邲����������m�F��������</p>
      <ul>
        <li><a href="http://guide.soundhouse.co.jp/guide/qanda7.asp">E-Mail���͂��Ȃ�</a></li>
        <li><a href="http://guide.soundhouse.co.jp/guide/qanda5.asp">�ԕi�E����</a></li>
        <li><a href="http://guide.soundhouse.co.jp/guide/qanda1.asp">������</a></li>
        <li><a href="http://guide.soundhouse.co.jp/guide/qanda4.asp">���͂�</a></li>
        <li><a href="http://guide.soundhouse.co.jp/guide/qanda11.asp">�̎���</a></li>
      </ul>
    </div>

<form name="f_data" id="inquiry" action="<%=g_HTTPS%>shop/InquiryConfirm.asp" method="post" onSubmit="return next_onSubmit();">

  <!-- �G���[���b�Z�[�W -->
  <% If wMsg <> "" Then %>
  <ul class="error">
    <li><%=wMsg %></li>
  </ul>
  <% End If %>

<table>
  <tr>
    <th>���<span>*</span></th>
    <td>
      <ul>
<% = wCategoryListHTML %>
      </ul>
    </td>
  </tr>
  <tr>
    <th>����<span>*</span></th>
    <td><input type="text" name="subject" size="60" value="<%=subject%>"></td>
  </tr>
  <tr>
    <th>���b�Z�[�W<span>*</span></th>
    <td><textarea name="message" rows="5" cols="60"><%=message%></textarea></td>
  </tr>
  <tr>
    <th>�����O<span>*</span></th>
    <td><input type="text" name="customer_nm" size="40" maxlength="30" value="<%=customer_nm%>"></td>
  </tr>
  <tr>
    <th>�t���K�i</th>
    <td><input type="text" name="furigana" size="40" maxlength="30" value="<%=furigana%>"><span>�i�S�p�J�i�j</span></td>
  </tr>
  <tr>
    <th>�Z ��</th>
    <td>
      ��<input type="text" name="zip" size="10" maxlength="8" value="<%=zip%>">�i���p�����j<a href="JavaScript:address_search_onClick();" class="tipBtn">�Z������</a>�X�֔ԍ�����͂��ă{�^���������Ă��������<br>
      <input type="hidden" name="i_selected_prefecture" value="<%=prefecture%>">
      <select name="prefecture" size="1">
        <option value="">�s���{��</option>
        <option value="�k�C��">�k�C��</option>
        <option value="�X��">�X��</option>
        <option value="�H�c��">�H�c��</option>
        <option value="��茧">��茧</option>
        <option value="�{�錧">�{�錧</option>
        <option value="�R�`��">�R�`��</option>
        <option value="������">������</option>
        <option value="�Ȗ،�">�Ȗ،�</option>
        <option value="�V����">�V����</option>
        <option value="�Q�n��">�Q�n��</option>
        <option value="��ʌ�">��ʌ�</option>
        <option value="��錧">��錧</option>
        <option value="��t��">��t��</option>
        <option value="�����s">�����s</option>
        <option value="�_�ސ쌧">�_�ސ쌧</option>
        <option value="�R����">�R����</option>
        <option value="���쌧">���쌧</option>
        <option value="�򕌌�">�򕌌�</option>
        <option value="�x�R��">�x�R��</option>
        <option value="�ΐ쌧">�ΐ쌧</option>
        <option value="�É���">�É���</option>
        <option value="���m��">���m��</option>
        <option value="�O�d��">�O�d��</option>
        <option value="�ޗǌ�">�ޗǌ�</option>
        <option value="�a�̎R��">�a�̎R��</option>
        <option value="���䌧">���䌧</option>
        <option value="���ꌧ">���ꌧ</option>
        <option value="���s�{">���s�{</option>
        <option value="���{">���{</option>
        <option value="���Ɍ�">���Ɍ�</option>
        <option value="���R��">���R��</option>
        <option value="���挧">���挧</option>
        <option value="������">������</option>
        <option value="�L����">�L����</option>
        <option value="�R����">�R����</option>
        <option value="���쌧">���쌧</option>
        <option value="������">������</option>
        <option value="���Q��">���Q��</option>
        <option value="���m��">���m��</option>
        <option value="������">������</option>
        <option value="���ꌧ">���ꌧ</option>
        <option value="�啪��">�啪��</option>
        <option value="�F�{��">�F�{��</option>
        <option value="�{�茧">�{�茧</option>
        <option value="���茧">���茧</option>
        <option value="��������">��������</option>
        <option value="���ꌧ">���ꌧ</option>
      </select>
      <input type="text" name="address" size="60" maxlength="40"  value="<%=address%>">
    </td>
  </tr>
  <tr>
    <th>�d�b�ԍ�<span>*</span></th>
    <td><input type="text" name="telephone" size="30" maxlength="20" value="<%=telephone%>" class="validate required">�i���p�����j</td>
  </tr>
  <tr>
    <th>FAX�ԍ�</th>
    <td><input type="text" name="fax" size="30" maxlength="20" value="<%=fax%>">�i���p�����j</td>
  </tr>
  <tr>
    <th>E-mail<span>*</span></th>
    <td><input type="text" name="e_mail" size="30" maxlength="60" value="<%=e_mail%>" class="validate mail required">�i���p�p�����j</td>
  </tr>
</table>
<p>�u*�v�̂��Ă��鍀�ڂ͕K�{���͍��ڂł��B</p>
<input type="hidden" name="ContactCategory" value="<%=ContactCategory%>">
<input type="hidden" name="ContactSubCategory" value="<%=ContactSubCategory%>">
<input type="hidden" name="ContactSubCategoryFl" value="<%=ContactSubCategoryFl%>">
<p class="btnBox"><input type="submit" value="���e���m�F����" class="opover"></p>
</form>

</div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
<script type="text/javascript">
	preset_values();
</script>
</body>
</html>