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
'	�ۑ��J�[�g�̈ꗗ
'
'�X�V����
'2009/04/30 �G���[����error.asp�ֈړ�
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2012/07/17 if-web ���j���[�A�����C�A�E�g����
'
'========================================================================

On Error Resume Next

Dim userID

Dim wSalesTaxRate
Dim wPrice
Dim wCartHTML

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

'---- UserID ���o��
userID = Session("userID")

wMSG = ReplaceInput(Request("msg"))

'---- Execute main
call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "SaveCartList.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
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

Dim vDateStored
Dim vCartName
Dim vTotalAm
Dim vBreakKey
Dim vBreakNextKey

wHTML = "" & vbNewLine

if userID = "" then
	wHTML = wHTML & "<p class='error'>���O�C�������Ă��������B</p>" & vbNewLine
	wCartHTML = wHTML
	exit function
end if

'---- ����ŗ���o��
call getCntlMst("����","����ŗ�","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)			'����ŗ�
wSalesTaxRate = Clng(wItemNum1)

wHTML = ""

'----�ۑ��J�[�g�f�[�^���o��
wSQL = ""
wSQL = wSQL & "SELECT a.�J�[�g��"
wSQL = wSQL & "     , a.�o�^��"
wSQL = wSQL & "     , b.���[�J�[�R�[�h"
wSQL = wSQL & "     , b.���i�R�[�h"
wSQL = wSQL & "     , b.�F"
wSQL = wSQL & "     , b.�K�i"
wSQL = wSQL & "     , b.�󒍐���"
wSQL = wSQL & "     , CASE"
wSQL = wSQL & "         WHEN (c.�����萔�� > c.������󒍍ϐ��� AND c.�����萔�� > 0) THEN c.������P��"
wSQL = wSQL & "         ELSE c.�̔��P��"
wSQL = wSQL & "       END AS �̔��P��"
wSQL = wSQL & "     , c.�I����"
wSQL = wSQL & "     , c.�戵���~��"
wSQL = wSQL & "     , c.�p�ԓ�"
wSQL = wSQL & "     , c.������"
wSQL = wSQL & "     , c.B�i�P��"
wSQL = wSQL & "     , c.B�i�t���O"
wSQL = wSQL & "  FROM �ۑ��J�[�g a WITH (NOLOCK)"
wSQL = wSQL & "     , �ۑ��J�[�g���� b WITH (NOLOCK)"
wSQL = wSQL & "     , Web���i c WITH (NOLOCK)"
wSQL = wSQL & " WHERE b.�ڋq�ԍ� = a.�ڋq�ԍ�"
wSQL = wSQL & "   AND b.�J�[�g�� = a.�J�[�g��"
wSQL = wSQL & "   AND c.���[�J�[�R�[�h = b.���[�J�[�R�[�h"
wSQL = wSQL & "   AND c.���i�R�[�h = b.���i�R�[�h"
wSQL = wSQL & "   AND a.�ڋq�ԍ� = " & userID
wSQL = wSQL & " ORDER BY a.�o�^�� DESC"

'@@@@@@response.write(wSQL)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

'----- ���o��
wHTML = wHTML & "<table id='saveCart'>" & vbNewLine
wHTML = wHTML & "  <tr>" & vbNewLine
wHTML = wHTML & "    <th class='date'>�o�^��</th>" & vbNewLine
wHTML = wHTML & "    <th class='name'>�J�[�g��</th>" & vbNewLine
wHTML = wHTML & "    <th class='total'>���i���v(�ō�)</th>" & vbNewLine
wHTML = wHTML & "    <th class='cart'>&nbsp;</th>" & vbNewLine
wHTML = wHTML & "    <th class='delete'>&nbsp;</th>" & vbNewLine
wHTML = wHTML & "  </tr>" & vbNewLine

if RS.EOF = true then
	wHTML = wHTML & "  <tr><td colspan='5'><p class='error'>�ۑ����ꂽ�J�[�g������܂���B</p></td></tr>" 
	wHTML = wHTML & "</table>" & vbNewLine
	wCartHTML = wHTML
	exit function
end if

vBreakNextKey = RS("�J�[�g��")
vBreakKey = vBreakNextKey
vTotalAm = 0

Do Until RS.EOF = true
	if RS("B�i�t���O") = "Y" then
		wPrice = calcPrice(RS("B�i�P��"), wSalesTaxRate)
	else
		wPrice = calcPrice(RS("�̔��P��"), wSalesTaxRate)
	end if

	vTotalAm = vTotalAm + (wPrice * RS("�󒍐���"))
	vDateStored = fFormatDate(RS("�o�^��"))
	vCartName = RS("�J�[�g��")

	RS.MoveNext

	if RS.EOF = false then
		vBreakNextKey = RS("�J�[�g��")
	else
		vBreakNextKey = "@EOF"
	end if

	if vBreakKey <> vBreakNextKey then
		'------------- �o�^��
		wHTML = wHTML & "  <tr>" & vbNewLine
		wHTML = wHTML & "    <td class='date'>" & vDateStored & "</td>" & vbNewLine

		'------------- �J�[�g��
		wHTML = wHTML & "    <td class='name'>" & vCartName & "</td>" & vbNewLine

			'------------- ���i���v
		wHTML = wHTML & "    <td class='total'>" & FormatNumber(vTotalAm,0) & "�~</td>" & vbNewLine

			'------------- �J�[�g�փ{�^��
		wHTML = wHTML & "    <td class='cart'><a href='SaveCartMoveToOrder.asp?CartName=" & Server.URLencode(vCartName) & "'><img src='images/btn_cart.png' alt='�J�[�g��' class='opover'></a></td>" & vbNewLine

			'------------- �폜�{�^��
		wHTML = wHTML & "    <td class='delete'><a href='SaveCartDelete.asp?CartName=" & Server.URLencode(vCartName) & "' class='tipBtn'>�폜</a></td>" & vbNewLine

		wHTML = wHTML & "  </tr>" & vbNewLine

		vBreakKey = vBreakNextKey
		vTotalAm = 0
	end if

Loop

wHTML = wHTML & "</table>" & vbNewLine

RS.close
wCartHTML = wHTML

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
<html>
<head>
<meta charset="Shift_JIS">
<title>�ۑ��J�[�g�ꗗ�b�T�E���h�n�E�X</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/shop.css" type="text/css">
<link rel="stylesheet" href="style/StyleOrder.css?20120717" type="text/css">
</head>
<body>
<!--#include file="../Navi/Navitop.inc"-->
<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="��������{���ł�"></a></span>

<!-- �R���e���cstart -->
<div id="globalContents">

  <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
    <p class="home"><a href="../"><img src="../images/icon_home.gif" alt="HOME"></a></p>
    <ul id="path">
      <li class="now">�ۑ��J�[�g�ꗗ</li>
    </ul>
  </div></div></div>

  <h1 class="title">�ۑ��J�[�g�ꗗ</h1>

<% if wMSG <> "" then %>
	<p class="error"><%=wMSG%></p>
<% end if %>

<%=wCartHTML%>

  </div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/Navibottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>