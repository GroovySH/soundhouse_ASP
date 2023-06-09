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
'	�J�[�g���e�̕ۑ�
'
'�X�V����
'2009/04/10 fCalcShipping�̃p�����[�^�ǉ��i�����j
'2009/04/30 �G���[����error.asp�ֈړ�
'2011/04/14 hn SessionID�֘A�ύX
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2012/07/17 if-web ���j���[�A�����C�A�E�g����
'
'========================================================================

On Error Resume Next

Dim userID

Dim wSalesTaxRate
Dim wPrice
Dim wNoData
Dim wOrderProductHTML

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
	wErrDesc = "SaveCart.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
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

Dim vProductNm
Dim vTotalAm
Dim vFreightAm
Dim vFreightForwarder
Dim vSoukoCnt
Dim vKoguchi

'---- ����ŗ���o��
call getCntlMst("����","����ŗ�","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)			'����ŗ�
wSalesTaxRate = Clng(wItemNum1)

vTotalAm = 0
wHTML = ""

'----���󒍃f�[�^���o��
wSQL = ""
wSQL = wSQL & "SELECT a.�󒍖��הԍ�"
wSQL = wSQL & "     , a.���[�J�[�R�[�h"
wSQL = wSQL & "     , a.���i�R�[�h"
wSQL = wSQL & "     , a.�F"
wSQL = wSQL & "     , a.�K�i"
wSQL = wSQL & "     , a.���[�J�[��"
wSQL = wSQL & "     , a.���i��"
wSQL = wSQL & "     , a.�󒍐���"
wSQL = wSQL & "     , a.�󒍒P��" 
wSQL = wSQL & "     , a.�󒍋��z" 
wSQL = wSQL & "     , b.ASK���i�t���O" 
wSQL = wSQL & "  FROM ���󒍖��� a WITH (NOLOCK)"
wSQL = wSQL & "     ,Web���i b WITH (NOLOCK)"
wSQL = wSQL & " WHERE SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod
wSQL = wSQL & "   AND b.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
wSQL = wSQL & "   AND b.���i�R�[�h = a.���i�R�[�h"
wSQL = wSQL & " ORDER BY �󒍖��הԍ�"

'@@@@@@response.write(wSQL)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

wNoData = false

'---- ����HTML�쐬
if RS.EOF = true then
	wNoData = true
	exit function
end if

'----- ���o��
wHTML = wHTML & "<table id='cart'>" & vbNewLine
wHTML = wHTML & "  <tr>" & vbNewLine
wHTML = wHTML & "    <th class='maker'>���[�J�[</th>" & vbNewLine
wHTML = wHTML & "    <th class='name'>���i��</th>" & vbNewLine
wHTML = wHTML & "    <th class='price'>�P��(�ō�)</th>" & vbNewLine
wHTML = wHTML & "    <th class='number'>����</th>" & vbNewLine
wHTML = wHTML & "    <th class='amount'>���z(�ō�)</th>" & vbNewLine
wHTML = wHTML & "  </tr>" & vbNewLine

Do Until RS.EOF = true
	'------------- ���[�J�[�A���i��
	vProductNm = RS("���i��")
	if Trim(RS("�F")) <> "" then
		vProductNm = vProductNm & "/" & RS("�F")
	end if
	if Trim(RS("�K�i")) <> "" then
		vProductNm = vProductNm & "/" & RS("�K�i")
	end if
	wHTML = wHTML & "  <tr>" & vbNewLine
	wHTML = wHTML & "    <td>" & RS("���[�J�[��") & "</td>" & vbNewLine
	wHTML = wHTML & "    <td><a href='ProductDetail.asp?Item=" & RS("���[�J�[�R�[�h") & "^" & Server.URLEncode(RS("���i�R�[�h")) & "^" & RS("�F") & "^" & RS("�K�i") & "'>" & vProductNm & "</a></td>" & vbNewLine

		'------------- �P��
	wPrice = calcPrice(RS("�󒍒P��"), wSalesTaxRate)
	vTotalAm = vTotalAm + (wPrice * RS("�󒍐���"))
	wHTML = wHTML & "    <td class='num'>" & FormatNumber(wPrice,0) & "�~</td>" & vbNewLine

		'------------- ����
	wHTML = wHTML & "    <td class='num'>" & RS("�󒍐���") & "</td>" & vbNewLine

		'------------- ���z
	wHTML = wHTML & "    <td class='num'>" & FormatNumber(wPrice*RS("�󒍐���"),0) & "�~</td>" & vbNewLine

	RS.MoveNext
Loop

wHTML = wHTML & "  <tr>" & vbNewLine
wHTML = wHTML & "    <td colspan='5'>" & vbNewLine
wHTML = wHTML & "      <dl class='total'>" & vbNewLine
'----���i���v���z
wHTML = wHTML & "        <dt>���i���v(�ō�)</dt><dd>" & FormatNumber(vTotalAm,0) & "�~</dd>" & vbNewLine
'---- ����
Call fCalcShipping(gSessionID, "�ꊇ", vFreightAm, vFreightForwarder, vSoukoCnt, vKoguchi)		'2011/04/14 hn mod
wPrice = Fix(vFreightAm * (100 + wSalesTaxRate) / 100)
wHTML = wHTML & "        <dt>��������(�ō�)</dt><dd>" & FormatNumber(wPrice,0) & "�~</dd>" & vbNewLine
wHTML = wHTML & "      </dl>" & vbNewLine
wHTML = wHTML & "    </td>" & vbNewLine
wHTML = wHTML & "  </tr>" & vbNewLine

wHTML = wHTML & "</table>" & vbNewLine

RS.close
wOrderProductHTML = wHTML

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
<title>�J�[�g�ۑ��b�T�E���h�n�E�X</title>
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
      <li class="now">�J�[�g�ۑ�</li>
    </ul>
  </div></div></div>

  <h1 class="title">�J�[�g�ۑ�</h1>

<% if wMSG <> "" then %>
  <p class="error"><%=wMSG%></p>
<% end if %>

  <h2 class="cart_title">�J�[�g���e</h2>
<%=wOrderProductHTML%>

  <form name="fData" method="post" action="SaveCart2.asp">

    <p>�J�[�g������͂��A[�ۑ�����]�{�^���������Ă��������B<br>�������O�̃J�[�g������ꍇ�͏㏑�����܂��B</p>
    <table class="form">
      <tr>
        <th>�J�[�g��</th>
        <td><input name="CartName" type="text" size="20" maxlength="10"><span>(10�����ȓ��j</span></td>
      </tr>
    </table>

    <p>&laquo; <a href="Order.asp">�߂�</a></p>
    <p class="btnBox"><input type="submit" value="�ۑ�����" class="opover"></p>

  </form>

  </div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/Navibottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>