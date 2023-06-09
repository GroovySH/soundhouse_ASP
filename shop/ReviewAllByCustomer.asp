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
'	���i���r���[ (���e�ҕʈꗗ)
'
'�X�V����
'2008/05/23 ���̓f�[�^�`�F�b�N�����iLEFT, Numeric, EOF��)
'2008/12/24 �݌ɏ󋵃Z�b�g�֐���
'2009/10/02 ����v���t�B�[�������F��FFFFFF�ɕύX
'2010/09/09 an ���r���[���擾�ł��Ȃ��ꍇ�̓G���[���b�Z�[�W�\��
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2012/07/30 if-web ���j���[�A�����C�A�E�g����
'2014/03/19 GV ����ő��łɔ���2�d�\���Ή�
'
'========================================================================

On Error Resume Next

Dim userID

Dim CNo

Dim wHandleName
Dim wPrefecture
Dim wReviewCnt

Dim wReviewListHTML

Dim wProdTermFl
Dim wPrice
Dim wSalesTaxRate

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

Response.buffer = true

'---- �Ăяo��������̃f�[�^���o��
CNo = ReplaceInput(Request("CNo"))
if isNumeric(CNo) = false then
	CNO = 0
end if

'---- Execute main
call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "ReviewAllByCustomer.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
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
'	Function	main proc
'
'========================================================================
'
Function main()

Dim vInventoryCd
Dim vInventoryImage
Dim i

'---- �Y���ڋq�̃��r���[���o��
wSQL = ""
wSQL = wSQL & "SELECT a.*"
wSQL = wSQL & "     , b.�ڋq�s���{��"
wSQL = wSQL & "     , c.���i��"
wSQL = wSQL & "     , c.���i�摜�t�@�C����_��"
wSQL = wSQL & "     , c.�̔��P��"
wSQL = wSQL & "     , c.ASK���i�t���O"
wSQL = wSQL & "     , c.�󏭐���"
wSQL = wSQL & "     , c.�Z�b�g���i�t���O"
wSQL = wSQL & "     , c.���[�J�[�������敪"
wSQL = wSQL & "     , c.�戵���~��"
wSQL = wSQL & "     , c.�p�ԓ�"
wSQL = wSQL & "     , c.B�i�t���O"
wSQL = wSQL & "     , c.Web�[����\���t���O"
wSQL = wSQL & "     , c.���ח\�薢��t���O"
wSQL = wSQL & "     , c.�����萔��"
wSQL = wSQL & "     , c.������󒍍ϐ���"
wSQL = wSQL & "     , d.�F"
wSQL = wSQL & "     , d.�K�i"
wSQL = wSQL & "     , d.�����\���ח\���"
wSQL = wSQL & "     , d.�����\����"
wSQL = wSQL & "     , d.B�i�����\����"
wSQL = wSQL & "     , e.���[�J�[��"
wSQL = wSQL & "  FROM ���i���r���[ a WITH (NOLOCK)"
wSQL = wSQL & "     , Web�ڋq�Z�� b WITH (NOLOCK)"
wSQL = wSQL & "     , Web���i c WITH (NOLOCK)"
wSQL = wSQL & "     , Web�F�K�i�ʍ݌� d WITH (NOLOCK)"
wSQL = wSQL & "     , ���[�J�[ e WITH (NOLOCK)"
wSQL = wSQL & " WHERE b.�ڋq�ԍ� = a.�ڋq�ԍ�"
wSQL = wSQL & "   AND b.�Z���A�� = 1"
wSQL = wSQL & "   AND c.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
wSQL = wSQL & "   AND c.���i�R�[�h = a.���i�R�[�h"
wSQL = wSQL & "   AND d.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
wSQL = wSQL & "   AND d.���i�R�[�h = a.���i�R�[�h"
wSQL = wSQL & "   AND d.�F = ''"
wSQL = wSQL & "   AND d.�K�i = ''"
wSQL = wSQL & "   AND e.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
wSQL = wSQL & "   AND a.�ڋq�ԍ� = " & CNo 
wSQL = wSQL & " ORDER BY a.ID DESC" 

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic, adLockOptimistic

'@@@@response.write(wSQL)

if RS.EOF = true then
	wMSG = "<p class='error'>�Y�����r���[���o�^����Ă��܂���</p>"
else   '2010/0909 an mod

	'---- ����ŗ���o��
	call getCntlMst("����","����ŗ�","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)
	wSalesTaxRate = Clng(wItemNum1)

	'----
	wHandleName = RS("���O")
	wPrefecture = RS("�ڋq�s���{��")
	wReviewCnt = RS.RecordCount
	wHTML = ""

	Do until RS.EOF = true

	'---- �p�ԃ`�F�b�N
		if  (isNull(RS("�戵���~��")) = true AND isNull(RS("�p�ԓ�")) = true) _
		 OR (isNull(RS("�p�ԓ�")) = false AND RS("�����\����") > 0) then
			wProdTermFl = "N"
		else
			wProdTermFl = "Y"
		end if

	'----
		wHTML = wHTML & "<table width='480' cellSpacing='0' cellPadding='0' border='0'>" & vbNewLine
		wHTML = wHTML & "  <tr>" & vbNewLine

	'---- ���i�摜
		wHTML = wHTML & "    <td width='110' align='center' valign='top' rowspan='2'>" & vbNewLine
		wHTML = wHTML & "      <a href='ProductDetail.asp?item=" & RS("���[�J�[�R�[�h") & "^" & RS("���i�R�[�h") & "'><img src='../shop/prod_img/" & RS("���i�摜�t�@�C����_��") & "' width='100' height='50'></a>" & vbNewLine
		wHTML = wHTML & "    </td>" & vbNewLine

	'---- ���[�J�[���A���i��
		wHTML = wHTML & "    <td width='220'>" & vbNewLine
		wHTML = wHTML & "      " & RS("���[�J�[��") & "<br>" & vbNewLine
		wHTML = wHTML & "      <a href='ProductDetail.asp?item=" & RS("���[�J�[�R�[�h") & "^" & RS("���i�R�[�h") & "'>" & RS("���i��") & "</a>" & vbNewLine
		wHTML = wHTML & "    </td>" & vbNewLine

	'---- �݌ɏ�
		vInventoryCd = GetInventoryStatus(RS("���[�J�[�R�[�h"),RS("���i�R�[�h"),RS("�F"),RS("�K�i"),RS("�����\����"),RS("�󏭐���"),RS("�Z�b�g���i�t���O"),RS("���[�J�[�������敪"),RS("�����\���ח\���"),wProdTermFl)

		'---- �݌ɏ󋵁A�F���ŏI�Z�b�g
		call GetInventoryStatus2(RS("�����\����"), RS("Web�[����\���t���O"), RS("���ח\�薢��t���O"), RS("�p�ԓ�"), RS("B�i�t���O"), RS("B�i�����\����"), RS("�����萔��"), RS("������󒍍ϐ���"), wProdTermFl, vInventoryCd, vInventoryImage)

		wHTML = wHTML & "    <td width='150' nowrap>" & vbNewLine
		wHTML = wHTML & "      �݌ɏ󋵁F<img src='images/" & vInventoryImage & "' width='10' height='10' class='inventoryImage'> " & vInventoryCd & "<br>" & vbNewLine

	'----- �Ռ�����
		wHTML = wHTML & "      �Ռ������F"

		if RS("ASK���i�t���O") = "Y" then
			wHTML = wHTML & "ASK" & vbNewLine
		else
			wPrice = calcPrice(RS("�̔��P��"), wSalesTaxRate)
'2014/03/19 GV mod start ---->
'				wHTML = wHTML & FormatNumber(wPrice,0) & "�~(�ō�)" & vbNewLine
				wHTML = wHTML & FormatNumber(RS("�̔��P��"),0) & "�~(�Ŕ�)<br>" & vbNewLine
				wHTML = wHTML & "(�ō�&nbsp;" & FormatNumber(wPrice,0) & "�~)" & vbNewLine
'2014/03/19 GV mod end   <----
		end if
		wHTML = wHTML & "    </td>" & vbNewLine
		wHTML = wHTML & "  </tr>" & vbNewLine

	'---- �Q�l�ɂȂ����l��
		wHTML = wHTML & "  <tr>" & vbNewLine
		wHTML = wHTML & "    <td>�Q�l�ɂȂ����l���F" & RS("�Q�l��") & "�l(" & RS("�Q�l��") + RS("�s�Q�l��") & "�l��)</td>" & vbNewLine

	'---- �J�[�g
		if wProdTermFl <> "Y" then
			wHTML = wHTML & "    <td>" & vbNewLine
			wHTML = wHTML & "      <form name='f_data' method='post' action='OrderPreInsert.asp'>" & vbNewLine
			wHTML = wHTML & "        <input type='text' name='qt' size='2' maxsize='4' value='1'>" & vbNewLine
			wHTML = wHTML & "        <input type='image' src='images/btn_cart.png' class='cartBtn opover'>" & vbNewLine
			wHTML = wHTML & "        <input type='hidden' name='maker_cd' value='" & RS("���[�J�[�R�[�h") & "'>" & vbNewLine
			wHTML = wHTML & "        <input type='hidden' name='product_cd' value='" & RS("���i�R�[�h") & "'>" & vbNewLine
			wHTML = wHTML & "        <input type='hidden' name='iro' value=''>" & vbNewLine
			wHTML = wHTML & "        <input type='hidden' name='kikaku' value=''>" & vbNewLine
			wHTML = wHTML & "      </form>" & vbNewLine
			wHTML = wHTML & "    </td>" & vbNewLine
		else
			wHTML = wHTML & "    <td><img src='images/icon_sold.gif' alt='����'></td>" & vbNewLine
		end if
		wHTML = wHTML & "  </tr>" & vbNewLine
		wHTML = wHTML & "</table>" & vbNewLine

	'---- ���r���[���e
		wHTML = wHTML & "<table cellSpacing='0' cellPadding='0' width='480' border='0'>" & vbNewLine
		wHTML = wHTML & "  <tr>" & vbNewLine

	'---- �������ߓx
		wHTML = wHTML & "    <td width='130'>" & vbNewLine
		wHTML = wHTML & "      "
		For i=1 to RS("�]��")
			wHTML = wHTML & "<img src='images/review_icon10.png'>"
		Next
		For i=RS("�]��")+1 to 5
			wHTML = wHTML & "<img src='images/review_icon00.png'>"
		Next
		wHTML = wHTML & " (" & FormatNumber(RS("�]��"), 1) & ")" & vbNewLine
		wHTML = wHTML & "    </td>" & vbNewLine

	'---- �^�C�g��, ���e��
		wHTML = wHTML & "    <td width='270'><b>" & RS("�^�C�g��") & "</b></td>" & vbNewLine
		wHTML = wHTML & "    <td width='80'>" & cf_FormatDate(RS("���e��"), "YYYY/MM/DD") & "</td>" & vbNewLine
		wHTML = wHTML & "  </tr>" & vbNewLine

	'---- ���r���[���e
		wHTML = wHTML & "  <tr>" & vbNewLine
		wHTML = wHTML & "    <td colspan='3' width='480'>" & Replace(RS("���r���[���e"), vbNewline, "<br>") & "</td>" & vbNewLine
		wHTML = wHTML & "  </tr>" & vbNewLine

	'---- ��؂��
		wHTML = wHTML & "  <tr>" & vbNewLine
		wHTML = wHTML & "    <td colSpan='3' height='5'><hr size='1'></td>" & vbNewLine
		wHTML = wHTML & "  </tr>" & vbNewLine

		wHTML = wHTML & "</table>" & vbNewLine

		RS.MoveNext
	Loop
end if     '2010/09/09 an mod

RS.close

wReviewListHTML = wHTML

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
<html>
<head>
<meta charset="Shift_JIS">
<title>���i���r���[�i���e�ҕʁj�b�T�E���h�n�E�X</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/ReviewAllByCustomer.css" type="text/css">
<link rel="stylesheet" href="style/ask.css?20140401a" type="text/css">
</head>
<body>
<!--#include file="../Navi/Navitop.inc"-->
<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="��������{���ł�"></a></span>

<!-- �R���e���cstart -->
<div id="globalContents">
<!--
  <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
    <p class="home"><a href="../"><img src="../images/icon_home.gif" alt="HOME"></a></p>
    <ul id="path">
      <li class="now"><%=wHandleName%> ����̃��r���[�ꗗ</li>
    </ul>
  </div></div></div>
-->
<% if wMSG <> "" then %>
	<%=wMSG%>
<% else %>

  <h1 class="title"><%=wHandleName%> ����̃��r���[�ꗗ</h1>

  <div id="main_container">

    <div id="rewiewlist">

<%=wReviewListHTML%>

    </div>

    <div id="detail_side">

      <div class='detail_side_inner01'><div class='detail_side_inner02'>
        <div class='detail_side_inner_box' id='subtotal'>
          <h4 class='detail_sub'><%=wHandleName%> ����̃v���t�B�[��</h4>
          <p>���r���[���e���F<%=wReviewCnt%>��</p>
          <p>�Z���F<%=wPrefecture%></p>
        </div>
      </div></div>

    </div>

  </div>
<% end if%>

  </div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/Navibottom.inc"-->
<div class="tooltip"><p>ASK</p></div>
<!--#include file="../Navi/NaviScript.inc"-->
<script type="text/javascript" src="jslib/ask.js?20140401a"></script>
</body>
</html>