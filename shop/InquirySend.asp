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
'	�⍇�����M
'
'�X�V����
'2005/05/13 OBJ_NewMail.BodyPart.Charset = "iso-2022-jp"���Z�b�g
'2005/08/18 �R���^�N�g�J�e�S���[�A�R���^�N�g�T�u�J�e�S���[��ǉ�
'2005/09/06 �R���^�N�g�Ǘ��e�X�g�Ή�
'2005/09/07 �����ԐM���[�����M
'2005/09/08 �⍇�����e��Web�R���^�N�g�e�[�u���֊i�[����悤�ɕύX
'2006/08/11 ���̓f�[�^�`�F�b�N����
'2008/05/12 ���s�R�[�h�C���W�F�N�V�����΍�ii_to�p�����[�^�폜�j
'2008/05/13 �N���X�T�C�g���N�G�X�g�t�H�W�F���[�΍� Key�p�����[�^�`�F�b�N
'2008/05/23 ���̓f�[�^�`�F�b�N�����iLEFT��)
'2009/04/30 �G���[����error.asp�ֈړ�
'2009/09/03 �����ԐM���e��ǋL
'2010/01/08 an �폜������s�R�[�h�̎w���vbCr/vbLf�ɕύX
'2011/04/13 an #725�ɍ��킹�G���[�`�F�b�N����
'2011/04/14 hn Session�֘A�ύX
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�, wErr��Err.Description�ɖ߂����i�@�\���Ă��Ȃ����߁j
'2012/06/29 if-web ���j���[�A�����C�A�E�g����
'
'========================================================================

On Error Resume Next

Dim userID
'Dim msg  '2011/04/13 an del

Dim message
Dim subject
Dim ContactCategory
Dim ContactSubCategory
Dim ContactSubCategoryFl  '2011/04/20 an add
Dim customer_nm
Dim furigana
Dim zip
Dim prefecture
Dim address
Dim telephone
Dim fax
Dim e_mail

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
'Dim wErr		'2011/04/13, 2011/08/01 an del
Dim wErrDesc    '2011/08/01 an add

'========================================================================

Response.buffer = true

'---- �Z�L�����e�B�[�L�[�`�F�b�N
if Session("SKey") <> ReplaceInput(Request("SKey")) then
	Response.redirect "Inquiry.asp"
end if

'---- UserID ���o��
userID = Session("userID")

'---- �Ăяo��������̃f�[�^���o��
message = ReplaceInput(Left(Request("message"),2000))
subject = ReplaceInput(Left(Request("subject"),150))
ContactCategory = ReplaceInput(Left(Request("ContactCategory"),20))
ContactSubCategory = ReplaceInput(Left(Request("ContactSubCategory"),20))
ContactSubCategoryFl = ReplaceInput_NoCRLF(Left(Request("ContactSubCategoryFl"),1))  '2011/04/13 an add
customer_nm = ReplaceInput(Left(Request("customer_nm"),30))
furigana = ReplaceInput(Left(Request("furigana"),30))
zip = ReplaceInput(Left(Request("zip"),8))
prefecture = ReplaceInput(Left(Request("prefecture"),8))
address = ReplaceInput(Left(Request("address"),40))
e_mail = ReplaceInput_NoCRLF(Left(Request("e_mail"),60)) '2010/01/08 an  2011/04/13 an mod
telephone = ReplaceInput(Left(Request("telephone"),20))
fax = ReplaceInput(Left(Request("fax"),20))

'---- Execute main
call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	Connection.RollbackTrans
	wErrDesc = "InquirySend.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

call close_db()

if wMsg <> "" then  '2011/04/13 an add s �ʏ��InquiryConfirm���o�R����̂ł����ŃG���[�͋N���Ȃ�����Transfer�͂��Ȃ�
    Response.Redirect g_HTTPS & "shop/Inquiry.asp"
end if              '2011/04/13 an add e

'if wErr <> "" then	           '2011/08/01 an del
if Err.Description <> "" then  '2011/08/01 an add
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
Dim v_subject
Dim OBJ_NewMail

'---- ���̓`�F�b�N      '2011/04/13 an add s
call validation()

if wMsg <> "" then
	exit function
end if

Connection.BeginTrans   '2011/04/13 an add e

'---- Web�R���^�N�g�֊i�[
wSQL = ""
wSQL = wSQL & "SELECT *"
wSQL = wSQL & "  FROM Web�R���^�N�g"
wSQL = wSQL & " WHERE 1 = 2"
	  
Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic, adLockOptimistic

RS.AddNew

RS("�R���^�N�g��") = Now()
if userID <> "" then
	RS("�ڋq�ԍ�") = userID
else
	RS("�ڋq�ԍ�") = 0
end if

RS("���͌ڋq��") = customer_nm
RS("���͌ڋq�t���K�i") = furigana
RS("���͌ڋqEmail") = e_mail
RS("���͌ڋq�d�b�ԍ�") = telephone
RS("���͌ڋqFax") = fax
RS("���͌ڋq�X�֔ԍ�") = zip
RS("���͌ڋq�s���{��") = prefecture
RS("���͌ڋq�Z��") = address
RS("�R���^�N�g�J�e�S���[") = ContactCategory
RS("�R���^�N�g�T�u�J�e�S���[") = ContactSubCategory
RS("����") = "shop@soundhouse.co.jp"

v_subject = subject & " " & customer_nm & "�l"
RS("����") = v_subject
RS("�{��") = message

RS.Update

RS.close

'wErr = Err.Description    '2011/04/13, 2011/08/01 an del

if Err.Description = "" then    '2011/04/13 an add 2011/08/01 an mod
                                
	Connection.CommitTrans		'Commit   '2011/04/13 an add
                                          
	'---- �����ԐM���[���쐬�i�ڋq��)     
	call getCntlMst("Web","Email","�g���[��", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)
                                          
	v_body = "���₢���킹���肪�Ƃ��������܂��" & vbNewLine _
	       & "�ȉ��̓��e�ɂĂ��₢���킹����t�������܂����" & vbNewLine _
	       & "�ԓ��܂ō����΂炭���҂����������" & vbNewLine & vbNewLine _
	       & "���Ђ���̕ԐM��24���Ԉȓ��i�x�Ɠ��������j�ɓ������Ȃ��ꍇ�ɂ́A���萔�ł���" & vbNewLine _
	       & "���̎|���₢���킹���������܂��悤���肢�\���グ�܂��B" & vbNewLine & vbNewLine

	v_body = v_body & "��t�����@�@�@�F" & now() & vbNewLine & vbNewLine
	v_body = v_body & "�����@�@�@�@�@�F" & subject & vbNewLine & vbNewLine
	v_body = v_body & "�J�e�S���[�@�@�F" & ContactCategory & vbNewLine
	v_body = v_body & "�T�u�J�e�S���[�F" & ContactSubCategory & vbNewLine
	v_body = v_body & "���b�Z�[�W�@�@�F" & message & vbNewLine & vbNewLine

	v_body = v_body & wItemChar1

	Set OBJ_NewMail = Server.CreateObject("CDO.Message") 

	OBJ_NewMail.from = "shop@soundhouse.co.jp"
	OBJ_NewMail.to = e_mail

	OBJ_NewMail.subject = "���₢���킹����t�������܂���"
	OBJ_NewMail.TextBody = v_body
	OBJ_NewMail.BodyPart.Charset = "iso-2022-jp"

	OBJ_NewMail.Send

	Set OBJ_NewMail = Nothing

else                                               '2011/04/13 an add s
	Connection.RollbackTrans	'Rollback
end if

End function                                       '2011/04/13 an add e

'========================================================================
'
'    Function    �⍇�����͓��e�`�F�b�N   '2011/04/13 an add
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
'	Function	�X�֔ԍ���������   '2011/04/13 an add
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
<title>���₢���킹�𑗐M���܂����b�T�E���h�n�E�X</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/inquiry.css" type="text/css">
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
        <li class="now">���₢���킹�𑗐M���܂���</li>
      </ul>
    </div></div></div>

    <h1 class="title">���₢���킹�𑗐M���܂���</h1>
    <p>���⍇���𑗐M���܂����B<br>���肪�Ƃ��������܂����B</p>
  </div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>