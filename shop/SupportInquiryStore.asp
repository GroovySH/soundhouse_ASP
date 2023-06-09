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
'2010/10/04 an SupportInquirySend�����ɐV�K�쐬
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2012/06/29 if-web ���j���[�A�����C�A�E�g����
'
'========================================================================

On Error Resume Next

Dim userID
Dim msg

Dim wCustomerName
Dim wZip
Dim wPrefecture
Dim wAddress
Dim wTelephone
Dim wFax
Dim wEmail

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
Dim wNoData
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

MakerName = ReplaceInput(Left(Request("MakerName"),25))
ProductName = ReplaceInput(Left(Request("ProductName"),50))
Warranty = ReplaceInput(Left(Request("Warranty"),2))
SerialNo = ReplaceInput(Left(Request("SerialNo"),40))
WhenPurchased = ReplaceInput(Left(Request("WhenPurchased"),10))
Comment = ReplaceInput(Left(Request("Comment"),500))

'---- Execute main
call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "SupportInquiryStore.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

call close_db()

if Err.Description <> "" OR wNoData = "Y" then
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

'---- �ڋq���擾
call GetCustomer()

'---- �ڋq��񂪎擾�ł��Ȃ��ꍇ�̓G���[
if wNoData = "Y" then
	exit function
else

	Connection.BeginTrans				'�g�����U�N�V�����J�n
	
	'---- �T�|�[�g�˗��֓o�^
	wSQL = ""
	wSQL = wSQL & "SELECT *"
	wSQL = wSQL & "  FROM �T�|�[�g�˗�"
	wSQL = wSQL & " WHERE 1 = 2"

	Set RS = Server.CreateObject("ADODB.Recordset")
	RS.Open wSQL, Connection, adOpenStatic, adLockOptimistic

	RS.AddNew

	RS("�T�|�[�g�˗��o�^��") = Now()
	RS("�ڋq�ԍ�") = userID
	RS("���[�J�[��") = MakerName
	RS("���i��") = ProductName
	RS("�ۏ؏�����Ȃ�") = Warranty
	RS("�V���A��No") = SerialNo
	RS("�w�������") = WhenPurchased
	RS("���e") = Comment

	RS.Update
	RS.close
	
	if Err.Description = "" then
		Connection.CommitTrans		'Commit
		
		'---- �ڋq�Ɏ�t���[�����M
		call SendEmail()
	else
		Connection.RollbackTrans	'Rollback
	end if

end if

End function

'========================================================================
'
'	Function	�ڋq���擾
'
'========================================================================

Function GetCustomer()

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

end Function

'========================================================================
'
'	Function	�ڋq�֎�t���[�����M
'
'========================================================================

Function SendEmail()

Dim i
Dim v_body
Dim v_body2
Dim v_subject
Dim OBJ_NewMail

'---- �ڋq�����{��
v_body = ""
v_body = v_body & "��t�����@�@�@�F" & now() & vbNewLine & vbNewLine
v_body = v_body & "���[�J�[�@�@�@�F" & MakerName & vbNewLine & vbNewLine
v_body = v_body & "���i���@�@�@�@�F" & ProductName & vbNewLine & vbNewLine
v_body = v_body & "�ۏ؏��@�@�@�@�F" & Warranty & vbNewLine & vbNewLine
v_body = v_body & "�V���A���ԍ��@�F" & SerialNo & vbNewLine & vbNewLine
v_body = v_body & "���w������ԁ@�F" & WhenPurchased & vbNewLine & vbNewLine
v_body = v_body & "���e�@�@�@�@�@�F" & Comment & vbNewLine & vbNewLine

'---- �Г������{��
v_body2 = ""
v_body2 = v_body2 & "���O�@�@�@�@�@�F" & wCustomerName & vbNewLine
v_body2 = v_body2 & "�Z���@�@�@�@�@�F" & wZip & " " & wPrefecture & wAddress & vbNewLine
v_body2 = v_body2 & "�d�b�ԍ��@�@�@�F" & wTelephone & vbNewLine
v_body2 = v_body2 & "Fax �@�@�@�@�@�F" & wFax & vbNewLine
v_body2 = v_body2 & "Email�@ �@�@�@�F" & wEmail & vbNewLine
v_body2 = v_body2 & "�ڋq�ԍ��@�@�@�F" & UserID & vbNewLine

Set OBJ_NewMail = Server.CreateObject("CDO.message") 

'---- �Г��������[���쐬
OBJ_NewMail.from = "support@soundhouse.co.jp"
OBJ_NewMail.to = "support@soundhouse.co.jp"

OBJ_NewMail.subject = "�C������i�T�|�[�g(" & MakerName & "/" & ProductName & ") " &  wCustomerName & "�@�l�@ [" & UserID & "/RA/Web��t]"
OBJ_NewMail.TextBody = v_body & v_body2
OBJ_NewMail.BodyPart.Charset = "iso-2022-jp"

OBJ_NewMail.Send
'---- �Г������@�����܂�

'---- �ڋq���������ԐM���[���쐬
call getCntlMst("Web","Email","�g���[��", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)

v_body = "���₢���킹���肪�Ƃ��������܂��" & vbNewLine _
       & "�ȉ��̓��e�ɂďC������i�T�|�[�g�˗�����t�������܂����" & vbNewLine _
       & "�ԓ��܂ō����΂炭���҂����������" & vbNewLine & vbNewLine _
       & v_body & vbNewLine _
       & wItemChar1

OBJ_NewMail.from = "support@soundhouse.co.jp"
OBJ_NewMail.to = wEmail

OBJ_NewMail.subject = "�C������i�T�|�[�g�˗�����t�������܂���"
OBJ_NewMail.TextBody = v_body
OBJ_NewMail.BodyPart.Charset = "iso-2022-jp"

OBJ_NewMail.Send
'---- �ڋq�����@�����܂�

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

<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="Shift_JIS">
<title>�C������i�T�|�[�g�˗����󂯕t���܂����b�T�E���h�n�E�X</title>
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
        <li class="now">�C������i�T�|�[�g</li>
      </ul>
    </div></div></div>

    <h1 class="title">�C������i�T�|�[�g�˗����󂯕t���܂���</h1>
    <p>
      �C���E���i�T�|�[�g�˗���o�^���܂����B<br>
      ��قǕ��ЃT�|�[�g�S�����炲�A���������グ�܂��B<br>
      ���肪�Ƃ��������܂����B
    </p>
  </div>

<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>