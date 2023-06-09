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
'	�E�B�b�V�����X�g�ꗗ
'
'�X�V����
'2008/01/18 �݌ɏ󋵁@�\���̐F��ύX
'						������P����B�i�Ɠ��l�̒P���\���ɕύX
'2008/12/24 �݌ɏ󋵃Z�b�g�֐���
'2009/09/09	�J�[�g�֓����Ƃ��ɁA�E�B�b�V�����X�g����폜���邩�ǂ�����₢���킹��
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2011/10/19 hn 1063 ASK�\�����@�ύX
'2012/07/23 if-web ���j���[�A�����C�A�E�g����
'2012/08/14 GV #1419 �����O�C�����E�B�b�V�����X�g���烍�O�C����ʂ�\������
'2012/09/07 nt �u�b�N�}�[�N���Ŗ����O�C�������ڃy�[�W�J�ڎ������O�C����ʂ�\��
'2014/03/19 GV ����ő��łɔ���2�d�\���Ή�
'
'========================================================================

On Error Resume Next

Dim userID

Dim wNotLogin					' ���O�C�����Ă��Ȃ�	' 2012/08/14 GV #1419 Add

Dim wSalesTaxRate
Dim wPrice
Dim wProdTermFl
Dim wItem

Dim wListHTML
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

Response.buffer = true

'---- UserID ���o��
userID = Session("userID")

wMSG = ReplaceInput(Request("msg"))

wNotLogin = False				' ������Ԃ̓��O�C�����Ă��鎖��O��Ƃ���	' 2012/08/14 GV #1419 Add

'---- Execute main
call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "WishList.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

call close_db()

if Err.Description <> "" then	
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

' 2012/08/14 GV #1419 Add Start
If wNotLogin = True Then
	'---- ���O�C�����Ă��Ȃ��ꍇ�̓��O�C���y�[�W��
	Session("msg") = wMsg

	'2012/09/07 nt mod Start
	'---- ���O�C����A�E�B�b�V�����X�g��ʕ\���̂��߁A�p�����[�^�ǉ�
	'Server.Transfer "shop/Login.asp"
	Response.Redirect "../shop/Login.asp?called_from=wishlist"
	'2012/09/07 nt mod End

End If
' 2012/08/14 GV #1419 Add End

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

Dim vPrice
Dim vInventoryCd
Dim vInventoryImage
Dim vProductName
Dim vItem

wListHTML = ""

' 2012/08/14 GV #1419 Mod Start
'if userID = "" then
'	wListHTML = wListHTML & "<p class='error'>���O�C�������Ă��������B</p>" & vbNewLine
'	exit function
'end if

Dim vRS

If userID = "" Then
	'---- ���O�C�����Ă��Ȃ���΃G���[�@����O�C�����Ă��������B�
	wNotLogin = True		' ���O�C������Ă��Ȃ�
	wMsg = "���O�C�����Ă��������B"
	Exit Function
End If

' �ڋq���擾
Set vRS = get_customer()

If vRS.EOF = True Then
	'---- Session("userID")�Ōڋq��񂪎�o���Ȃ���΃G���[�@����O�C�����Ă��������B�
	wNotLogin = True		' ���O�C������Ă��Ȃ�
	wMsg = "���O�C�����Ă��������B"
	Exit Function
End If
' 2012/08/14 GV #1419 Mod End

'---- ����ŗ���o��
call getCntlMst("����","����ŗ�","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)			'����ŗ�
wSalesTaxRate = Clng(wItemNum1)

wHTML = ""

'----�E�B�b�V�����X�g���o��
wSQL = ""
wSQL = wSQL & "SELECT DISTINCT"
wSQL = wSQL & "       a.���[�J�[�R�[�h"	
wSQL = wSQL & "     , a.���i�R�[�h"
wSQL = wSQL & "     , a.���i��"
wSQL = wSQL & "     , a.���i�T��Web"
wSQL = wSQL & "     , a.���i�摜�t�@�C����_��"
wSQL = wSQL & "     , a.�̔��P��"
wSQL = wSQL & "     , a.������P��"
wSQL = wSQL & "     , a.�����萔��"
wSQL = wSQL & "     , a.������󒍍ϐ���"
wSQL = wSQL & "     , CASE"
wSQL = wSQL & "         WHEN (a.�����萔�� > a.������󒍍ϐ��� AND a.�����萔�� > 0) THEN 'Y'"
wSQL = wSQL & "         ELSE 'N'"
wSQL = wSQL & "       END AS ������P���t���O"
wSQL = wSQL & "     , a.���[�J�[�������敪"
wSQL = wSQL & "     , a.ASK���i�t���O"
wSQL = wSQL & "     , a.�戵���~��"
wSQL = wSQL & "     , a.�p�ԓ�"
wSQL = wSQL & "     , a.�I����"
wSQL = wSQL & "     , a.�󏭐���"
wSQL = wSQL & "     , a.�Z�b�g���i�t���O"	
wSQL = wSQL & "     , a.Web�[����\���t���O"	
wSQL = wSQL & "     , a.���ח\�薢��t���O"
wSQL = wSQL & "     , a.B�i�P��"
wSQL = wSQL & "     , a.������"
wSQL = wSQL & "     , a.B�i�t���O"
wSQL = wSQL & "     , b.�F"
wSQL = wSQL & "     , b.�K�i"
wSQL = wSQL & "     , b.�����\����"
wSQL = wSQL & "     , b.�����\���ח\���"
wSQL = wSQL & "     , b.B�i�����\����"
wSQL = wSQL & "     , c.���[�J�[��"
wSQL = wSQL & "     , d.�o�^��"

'---- FROM
wSQL = wSQL & "  FROM Web���i a WITH (NOLOCK)"
wSQL = wSQL & "     , Web�F�K�i�ʍ݌� b WITH (NOLOCK)"
wSQL = wSQL & "     , ���[�J�[ c WITH (NOLOCK)"
wSQL = wSQL & "     , �E�B�b�V�����X�g d WITH (NOLOCK)"

'---- WHERE
wSQL = wSQL & " WHERE a.Web���i�t���O = 'Y'"
wSQL = wSQL & "   AND b.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
wSQL = wSQL & "   AND b.���i�R�[�h = a.���i�R�[�h"
wSQL = wSQL & "   AND c.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
wSQL = wSQL & "   AND b.���[�J�[�R�[�h = d.���[�J�[�R�[�h"
wSQL = wSQL & "   AND b.���i�R�[�h = d.���i�R�[�h"
wSQL = wSQL & "   AND b.�F = d.�F"
wSQL = wSQL & "   AND b.�K�i = d.�K�i"
wSQL = wSQL & "   AND d.�ڋq�ԍ� = " & userID

'---- ORDER BY
wSQL = wSQL & " ORDER BY c.���[�J�[��, a.���i��"

'@@@@@@response.write(wSQL)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

'---- ���i�ꗗ�쐬
if RS.EOF = true then
	wListHTML = wListHTML & "<p class='error'>�E�B�b�V�����X�g�ɏ��i������܂���B</p>" & vbNewLine

else
	wListHTML = wListHTML & "<table width='480' border='0' cellspacing='0' cellpadding='0'>" & vbNewLine

'wListHTML = wListHTML & "  <tr>" & vbNewLine
'wListHTML = wListHTML & "    <td height='5' colspan='3'>�����i���J�[�g�ɓ���܂��ƃE�B�b�V�����X�g����폜����܂��B</td>" & vbNewLine
'wListHTML = wListHTML & "  </tr>" & vbNewLine

'---- ��؂��
wListHTML = wListHTML & "  <tr>" & vbNewLine
wListHTML = wListHTML & "    <td height='5' colspan='3'><hr size='1'></td>" & vbNewLine
wListHTML = wListHTML & "  </tr>" & vbNewLine

	Do until RS.EOF = true
	
		wListHTML = wListHTML & "  <tr align='left' valign='middle'>" & vbNewLine

		wListHTML = wListHTML & "    <form name='f_item' method='post' action='WishListToCartDelete.asp' onSubmit='return order_onClick(this);'>" & vbNewLine

		'---- �I���`�F�b�N
		wProdTermFl = "N"
		if isNull(RS("�戵���~��")) = false then		'�戵���~
			wProdTermFl = "Y"
		end if
		if isNull(RS("�p�ԓ�")) = false AND RS("�����\����") <= 0 then		'�p�Ԃō݌ɖ���
			wProdTermFl = "Y"
		end if
		if isNull(RS("������")) = false then		'�������i
			wProdTermFl = "Y"
		end if

	'----- ���i�摜 
		vItem = "Item=" & Server.URLEncode(RS("���[�J�[�R�[�h") & "^" & RS("���i�R�[�h") & "^" & Trim(RS("�F")) & "^" & Trim(RS("�K�i")))
		wListHTML = wListHTML & "    <td width='110' align='center' valign='top' rowspan='2'>" & vbNewLine
		wListHTML = wListHTML & "      <a href='" & g_HTTP & "shop/ProductDetail.asp?" & vItem & "'>"
		if Trim(RS("���i�摜�t�@�C����_��")) <> "" then 
			wListHTML = wListHTML & "      <img src='prod_img/" & RS("���i�摜�t�@�C����_��") & "' width='100' height='50'></a>" & vbNewLine
		end if
		wListHTML = wListHTML & "    </td>" & vbNewLine

	'----���[�J�[��
		wListHTML = wListHTML & "    <td width='220' valign='top' nowrap>" & vbNewLine
		wListHTML = wListHTML & "      <span>"  & RS("���[�J�[��") & "</span><br>"

	'----- ���i��/�F/�K�i
		wListHTML = wListHTML & "      <a href='" & g_HTTP & "shop/ProductDetail.asp?" & vItem & "'>"
		vProductName = RS("���i��")
		if Trim(RS("�F")) <> "" then
			vProductName = vProductName & "/" & Trim(RS("�F"))
		end if
		if Trim(RS("�K�i")) <> "" then
			vProductName = vProductName & "/" & Trim(RS("�K�i"))
		end if
	 	wListHTML = wListHTML & vProductName & "</a>" & vbNewLine
		wListHTML = wListHTML & "    </td>" & vbNewLine

		wListHTML = wListHTML & "    <td width='150' valign='top' nowrap>" & vbNewLine

	'---- �o�^��
		wListHTML = wListHTML & "      �o�^���F" & fFormatDate(RS("�o�^��")) & "<br>"

	'----- �̔��P��
		vPrice = calcPrice(RS("�̔��P��"), wSalesTaxRate)

		if RS("B�i�t���O") = "Y" OR RS("������P���t���O") = "Y" then
			wListHTML = wListHTML & "      <del>�Ռ������F"
		else
			wListHTML = wListHTML & "      �Ռ������F"
		end if

		if RS("ASK���i�t���O") = "Y" then
'2011/10/19 hn mod s
'			wListHTML = wListHTML & "<a href='JavaScript:void(0);' onClick=""askWin=window.open('AskPrice.asp?MakerName=" & Server.URLEncode(RS("���[�J�[��")) & "&ProductName=" & Server.URLEncode(vProductName) & "&Price=" & vPrice & "' ,'ask', 'width=250 height=80 scrollbars=false memubar=false toolbar=false statusbar=false personalbar=false locationbar=false');"" class='link'>ASK</a>" & vbNewLine
'2014/03/19 GV mod start ---->
'			wListHTML = wListHTML & "<a class='tip'>ASK<span>" & FormatNumber(vPrice,0) & "�~(�ō�)</span></a>" & vbNewLine
			wListHTML = wListHTML & "<a class='tip'>ASK<span class='exc-tax'>" & FormatNumber(RS("�̔��P��"),0) & "�~(�Ŕ�)</span><br>"
			wListHTML = wListHTML & "<span class='inc-tax'>(�ō�&nbsp;" & FormatNumber(vPrice,0) & "�~)</span></a>" & vbNewLine
'2014/03/19 GV mod end   <----
'2011/10/19 hn mod e

		else
'2014/03/19 GV mod start ---->
'				wListHTML = wListHTML & FormatNumber(vPrice,0) & "�~(�ō�)" & vbNewLine
				wListHTML = wListHTML & FormatNumber(RS("�̔��P��"),0) & "�~(�Ŕ�)<br>"
				wListHTML = wListHTML & "(�ō�&nbsp;" & FormatNumber(vPrice,0) & "�~)"
'2014/03/19 GV mod end   <----
		end if

		if RS("B�i�t���O") = "Y" OR RS("������P���t���O") = "Y" then
			wListHTML = wListHTML & "</del>"
		end if

	'---- B�i�P��
		if RS("B�i�t���O") = "Y" then
			vPrice = calcPrice(RS("B�i�P��"), wSalesTaxRate)
			wListHTML = wListHTML & "      <br><b>B�i�����F</b>"
'2014/03/19 GV mod start ---->
'			wListHTML = wListHTML & FormatNumber(vPrice,0) & "�~(�ō�)" & vbNewLine
			wListHTML = wListHTML & FormatNumber(RS("B�i�P��"),0) & "�~(�Ŕ�)<br>"
			wListHTML = wListHTML & "(�Ŕ�&nbsp;" & FormatNumber(vPrice,0) & "�~)"
'2014/03/19 GV mod end   <----
		end if

	'---- ������P��
		if RS("������P���t���O") = "Y" then
			vPrice = calcPrice(RS("������P��"), wSalesTaxRate)
			wListHTML = wListHTML & "      <br><b>��������F</b>"
'2014/03/19 GV mod start ---->
'			wListHTML = wListHTML & FormatNumber(vPrice,0) & "�~(�ō�)" & vbNewLine
			wListHTML = wListHTML & FormatNumber(RS("������P��"),0) & "�~(�Ŕ�)<br>"
			wListHTML = wListHTML & "(�ō�&nbsp;" & FormatNumber(vPrice,0) & "�~)"
'2014/03/19 GV mod end   <----
		end if

		wListHTML = wListHTML & "    </td>" & vbNewLine


		wListHTML = wListHTML & "  </tr>" & vbNewLine

	'----- ���i�T��Web
		wListHTML = wListHTML & "  <tr align='left' valign='middle'>" & vbNewLine
		if Trim(RS("���i�T��Web")) = "" OR isNull(Trim(RS("���i�T��Web")))then
			wListHTML = wListHTML & "    <td>�@</td>" & vbNewLine
		else
			wListHTML = wListHTML & "    <td>" & RS("���i�T��Web") & "</td>" & vbNewLine
		end if

		wListHTML = wListHTML & "    <td valign='top' nowrap>" & vbNewLine

	'----- �݌ɏ󋵕\���i�F�K�i�Ȃ����i�̂݁j
		vInventoryCd = GetInventoryStatus(RS("���[�J�[�R�[�h"),RS("���i�R�[�h"),RS("�F"),RS("�K�i"),RS("�����\����"),RS("�󏭐���"),RS("�Z�b�g���i�t���O"),RS("���[�J�[�������敪"),RS("�����\���ח\���"),wProdTermFl)

	'---- �݌ɏ󋵁A�F���ŏI�Z�b�g
		call GetInventoryStatus2(RS("�����\����"), RS("Web�[����\���t���O"), RS("���ח\�薢��t���O"), RS("�p�ԓ�"), RS("B�i�t���O"), RS("B�i�����\����"), RS("�����萔��"), RS("������󒍍ϐ���"), wProdTermFl, vInventoryCd, vInventoryImage)

		wListHTML = wListHTML & "      �݌ɏ󋵁F<img src='images/" & vInventoryImage & "' width='10' height='10' class='inventoryImage'> " & vInventoryCd & "<br>"

		wItem= Trim(RS("���[�J�[�R�[�h")) & "^" & Trim(RS("���i�R�[�h")) & "^" & Trim(RS("�F")) & "^" & Trim(RS("�K�i"))

	'---- �o�^���A�E�B�b�V�����X�g����폜
		wListHTML = wListHTML & "      <a href='WishListToCartDelete.asp?DeleteFl=Y&Item=" & wItem & "'>�E�B�b�V�����X�g����폜</a><br>"

	'----- ����, �J�[�g�{�^��
		if (IsNull(RS("�戵���~��")) = false) OR (IsNull(RS("������")) = false) OR (RS("B�i�t���O") = "Y" AND RS("B�i�����\����") <= 0) OR (IsNull(RS("�p�ԓ�")) = false AND RS("�����\����") <= 0) then
			wListHTML = wListHTML & "      <input type='hidden' name='qt' value='0'>" & vbNewLine
			wListHTML = wListHTML & "      <img src='images/icon_sold.gif'>" & vbNewLine
		else
			wListHTML = wListHTML & "      <input type='text' name='qt' size='2' value='1'>" & vbNewLine
			wListHTML = wListHTML & "      <input type='image' src='images/btn_cart.png' class='cartBtn'>" & vbNewLine
		end if

		wListHTML = wListHTML & "      <input type='hidden' name='Item' value='" & wItem & "'>" & vbNewLine
		wListHTML = wListHTML & "      <input type='hidden' name='Kubun' value='Cart'>" & vbNewLine

		wListHTML = wListHTML & "      <input type='hidden' name='DeleteFl' value='N'>" & vbNewLine

		wListHTML = wListHTML & "    </td>" & vbNewLine
		wListHTML = wListHTML & "    </form>" & vbNewLine
		wListHTML = wListHTML & "  </tr>" & vbNewLine

	'---- ��؂��
		wListHTML = wListHTML & "  <tr>" & vbNewLine
		wListHTML = wListHTML & "    <td height='5' colspan='3'><hr size='1'></td>" & vbNewLine
		wListHTML = wListHTML & "  </tr>" & vbNewLine

	'----
		RS.MoveNext
	Loop

	wListHTML = wListHTML & "</table>" & vbNewLine

	RS.Close
end if

'---- �J�[�g�̒��g�쐬
wCartHTML = fCreateCartHtml()

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

' 2012/08/14 GV #1419 Add Start
'========================================================================
'
'	Function	�ڋq���̎��o��
'
'========================================================================
Function get_customer()

Dim vRS
Dim vSQL

'---- �ڋq�����o��
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      a.�ڋq�ԍ� "
vSQL = vSQL & "FROM "
vSQL = vSQL & "      Web�ڋq     a WITH (NOLOCK) "
vSQL = vSQL & "        LEFT JOIN �ڋq�v���t�B�[�� c WITH (NOLOCK) "
vSQL = vSQL & "          ON a.�ڋq�ԍ� = c.�ڋq�ԍ� "
vSQL = vSQL & "    , Web�ڋq�Z�� b WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        a.�ڋq�ԍ� = b.�ڋq�ԍ� "
vSQL = vSQL & "    AND b.�Z���A�� = 1 "
vSQL = vSQL & "    AND a.Web�s�f�ڃt���O <> 'Y'"
vSQL = vSQL & "    AND a.�ڋq�ԍ� = " & UserID

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, Connection, adOpenStatic, adLockOptimistic

Set get_customer = vRS

End Function
' 2012/08/14 GV #1419 Add End

'========================================================================
%>
<!DOCTYPE html>
<html>
<head>
<meta charset="Shift_JIS">
<title>�E�B�b�V�����X�g�b�T�E���h�n�E�X</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/WishList.css" type="text/css">
<link rel="stylesheet" href="style/ask.css?20140401a" type="text/css">
<script type="text/javascript">
//
//  	Function:	order_onClick
//
function order_onClick(pForm){
	if (pForm.qt.value == ""){
		pForm.qt.value = 0;
	}else{
		if (numericCheck(pForm.qt.value) == false){
			pForm.qt.value = 0;
		}
	}
	if (pForm.qt.value == 0){
		alert("���ʂ���͂��Ă���J�[�g�{�^���������Ă��������B");
		return false;
	}

	if (confirm("�E�B�b�V�����X�g��ێ����܂����H") == true){
		pForm.DeleteFl.value = "N";
	}else{
		pForm.DeleteFl.value = "Y";
	}
}

</script>
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
      <li class="now">�E�B�b�V�����X�g</li>
    </ul>
  </div></div></div>

  <h1 class="title">�E�B�b�V�����X�g</h1>

  <div id="wishlist_container">

    <div id="wishlist">

<% if wMSG <> "" then %>
      <p class="error"><%=wMSG%></p>
<% end if %>

<%=wListHTML%>

    </div>

    <div id="detail_side">

<%=wCartHTML%>

    </div>

  </div>

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