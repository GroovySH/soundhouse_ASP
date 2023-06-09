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
<!--#include file="../3rdParty/EAgency.inc"-->
<%
'========================================================================
'
'	�V���b�s���O�J�[�g
'
'2012/06/14 ok �f�U�C���ύX�̂��ߋ��ł����ɐV�K�쐬
'2012/07/02 ok ���O�C���ρA�J�[�g�󂩂ۑ��J�[�g�����݂���ꍇ�u�ۑ����ꂽ�J�[�g�ꗗ�ցv�{�^����\������悤�C��
'2013/05/20 GV #1505 ���Ԃ݂��ƁI���R�����h�@�\
'2013/08/07 if-web �����R�����h�i�`�[�����{�j���R�����g�A�E�g
'2013/10/21 GV # ��^���i�̕\��
'
'========================================================================

On Error Resume Next

Dim userID
Dim userName
Dim msg

Dim wSalesTaxRate
Dim wPrice
Dim wNoData
Dim wOrderProductHTML
Dim wSavedCartFl

Dim wRecommendMakerCd
Dim wRecommendProductCd
Dim wRecommendHTML

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

Dim w_error_msg
Dim wErrDesc

Dim wOrderProductId		'2013/05/20 GV #1505 add

'2013/10/21 GV # add start
Dim strLargeItem
Dim wLargeItemFl
Dim wNonLargeItemFl
'2013/10/21 GV # add end

'========================================================================

Response.Expires = -1			' Do not cache

'---- UserID ���o��
userID = Session("userID")
userName = Session("userName")

'---- Get input data
msg = Session.contents("msg")
Session("msg") = ""

wOrderProductId = ""			'2013/05/20 GV #1505 add

'2013/10/21 GV # add start
strLargeItem = ""
wLargeItemFl = "N"
wNonLargeItemFl = "N"
'2013/10/21 GV # add end

'---- Execute main
call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^
if Err.Description <> "" then
	wErrDesc = "Order.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if

call close_db()

if Err.Description <> "" then
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

'========================================================================
'
'	Function	Connect database
'
'========================================================================
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
Function main()

Dim v_product_nm
Dim vTotalAm
Dim vFreightAm
Dim vFreightForwarder
Dim vSoukoCnt
Dim vKoguchi
'2011/04/14 GV Add Start
Dim vProdTermFl
Dim vInventoryCd
Dim vInventoryImage
'2011/04/14 GV Add End

'---- �ۑ����ꂽ�J�[�g��񂪂��邩�ǂ����`�F�b�N
wSavedCartFl = "N"
If userID <> "" Then
	Call CheckSavedCart()
End If

'---- ����ŗ���o��
Call getCntlMst("����","����ŗ�","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)			'����ŗ�
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
wSQL = wSQL & "     , b.�戵���~��"
wSQL = wSQL & "     , b.�p�ԓ�"
wSQL = wSQL & "     , b.������"
wSQL = wSQL & "     , b.�󏭐���"
wSQL = wSQL & "     , b.�Z�b�g���i�t���O"
wSQL = wSQL & "     , b.���[�J�[�������敪"
wSQL = wSQL & "     , b.��A�֎~�t���O "					'2013/10/21 GV # add
wSQL = wSQL & "     , b.����s�t���O "					'2013/10/21 GV # add
wSQL = wSQL & "     , b.�����敪 "							'2013/10/21 GV # add
wSQL = wSQL & "     , b.Web�[����\���t���O"
wSQL = wSQL & "     , b.���ח\�薢��t���O"
wSQL = wSQL & "     , b.B�i�t���O"
wSQL = wSQL & "     , b.�����萔��"
wSQL = wSQL & "     , b.������󒍍ϐ���"
wSQL = wSQL & "     , c.�����\����"
wSQL = wSQL & "     , c.�����\���ח\���"
wSQL = wSQL & "     , c.B�i�����\����"
wSQL = wSQL & "     , c.���iID"							'2013/05/20 GV #1505 add
wSQL = wSQL & "  FROM ���󒍖��� a WITH (NOLOCK)"
wSQL = wSQL & "     ,Web���i b WITH (NOLOCK)"
wSQL = wSQL & "     ,Web�F�K�i�ʍ݌� c WITH (NOLOCK)"
wSQL = wSQL & " WHERE SessionID = '" & gSessionID & "'"
wSQL = wSQL & "   AND b.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
wSQL = wSQL & "   AND b.���i�R�[�h = a.���i�R�[�h"
wSQL = wSQL & "   AND c.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
wSQL = wSQL & "   AND c.���i�R�[�h = a.���i�R�[�h"
wSQL = wSQL & "   AND c.�F = a.�F"
wSQL = wSQL & "   AND c.�K�i = a.�K�i"
wSQL = wSQL & " ORDER BY �󒍖��הԍ�"

'@@@@response.write(wSQL)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

wNoData = False

'---- ����HTML�쐬
If RS.EOF = True Then
	wHTML = wHTML & "      <tr>" & vbNewLine
	wHTML = wHTML & "        <td align='center'>" & vbNewLine
	wHTML = wHTML & "          <b>�J�[�g�ɏ��i������܂���B</b><br>" & vbNewLine
	wHTML = wHTML & "          �J�[�g�ɏ��i������Ȃ��ꍇ�́A�u���E�U�[��Cookie���L���ɂȂ��Ă��邱�Ƃ��m�F���Ă��������B�ݒ���@�ɂ��Ă�<a href='../guide/site.asp#kankyo'>������B</a>" & vbNewLine
	wHTML = wHTML & "        </td>" & vbNewLine
	wHTML = wHTML & "      </tr>" & vbNewLine
	wOrderProductHTML = wHTML
	wNoData = True
	Exit Function
End If

'----- ���o��
wHTML = wHTML & "      <tr>" & vbNewLine
wHTML = wHTML & "        <th class='maker'>���[�J�[</th>" & vbNewLine
wHTML = wHTML & "        <th class='name'>���i��</th>" & vbNewLine
wHTML = wHTML & "        <th class='stock'>�݌�</th>" & vbNewLine
wHTML = wHTML & "        <th class='price'>�P��</th>" & vbNewLine
wHTML = wHTML & "        <th class='number'>����</th>" & vbNewLine
wHTML = wHTML & "        <th class='amount'>���z(�ō�)</th>" & vbNewLine
wHTML = wHTML & "        <th></th>" & vbNewLine
wHTML = wHTML & "      </tr>" & vbNewLine

Do Until RS.EOF = True

	'---- 2013.10.21 GV # add start
	'---- ��^���i�̕\��
	strLargeItem = ""
	If (((IsNull(RS("��A�֎~�t���O")) = False) And (RS("��A�֎~�t���O") = "Y")) And _
		((IsNull(RS("����s�t���O")) = False) And (RS("����s�t���O") = "Y")) And _
		(RS("�����敪") = "�d�ʏ��i")) Then
		strLargeItem = strLargeItem & "<br><span style='color:red;'>��^���i</span>"
		wLargeItemFl = "Y"
	Else
		wNonLargeItemFl = "Y"
	End If
	'---- 2013.10.21 GV # add end

	'------------- ���[�J�[�A���i��
	v_product_nm = RS("���i��")
	If Trim(RS("�F")) <> "" Then
		v_product_nm = v_product_nm & "/" & RS("�F")
	End If
	If Trim(RS("�K�i")) <> "" Then
		v_product_nm = v_product_nm & "/" & RS("�K�i")
	End If
	wHTML = wHTML & "      <tr>" & vbNewLine
	wHTML = wHTML & "        <td>" & RS("���[�J�[��") & "</td>" & vbNewLine
'	wHTML = wHTML & "        <td><a href='ProductDetail.asp?Item=" & RS("���[�J�[�R�[�h") & "^" & Server.URLEncode(RS("���i�R�[�h")) & "^" & RS("�F") & "^" & RS("�K�i") & "' alt=''>" & v_product_nm & "</a></td>" & vbNewLine
	wHTML = wHTML & "        <td><a href='ProductDetail.asp?Item=" & RS("���[�J�[�R�[�h") & "^" & Server.URLEncode(RS("���i�R�[�h")) & "^" & RS("�F") & "^" & RS("�K�i") & "' alt=''>" & v_product_nm & "</a>" & strLargeItem & "</td>" & vbNewLine

	'------------- �݌�
	'---- �I���`�F�b�N
	vProdTermFl = "N"
	If IsNull(RS("�戵���~��")) = False Then	'�戵���~
		vProdTermFl = "Y"
	End If
	If IsNull(RS("�p�ԓ�")) = False And RS("�����\����") <= 0 Then	'�p�Ԃō݌ɖ���
		vProdTermFl = "Y"
	End If
	If IsNull(RS("������")) = False Then		'�������i
		vProdTermFl = "Y"
	End If

	'---- �݌ɏ�
	vInventoryCd = GetInventoryStatus(RS("���[�J�[�R�[�h"), RS("���i�R�[�h"), RS("�F"), RS("�K�i"), RS("�����\����"), RS("�󏭐���"), RS("�Z�b�g���i�t���O"), RS("���[�J�[�������敪"), RS("�����\���ח\���"), vProdTermFl)

	'---- �݌ɏ󋵁A�F���ŏI�Z�b�g
	Call GetInventoryStatus2(RS("�����\����"), RS("Web�[����\���t���O"), RS("���ח\�薢��t���O"), RS("�p�ԓ�"), RS("B�i�t���O"), RS("B�i�����\����"), RS("�����萔��"), RS("������󒍍ϐ���"), vProdTermFl, vInventoryCd, vInventoryImage)

	'----- �݌ɏ󋵕\��
	If IsNull(RS("�戵���~��")) = False Or _
	   IsNull(RS("������")) = False Or _
	   (RS("B�i�t���O") = "Y" And RS("B�i�����\����") <= 0) Or _
	   (IsNull(RS("�p�ԓ�")) = False And RS("�����\����") <= 0) Then
		wHTML = wHTML & "        <td><span class='stock'>&nbsp</span></td>" & vbNewLine
	Else
		'---- �������łȂ��ꍇ�̂݁A�݌ɏ󋵂�\��
		wHTML = wHTML & "        <td><span class='stock'><img src='images/" & vInventoryImage & "' alt='' > " & vInventoryCd & "</span></td>" & vbNewLine
	End If

	'------------- �P��
	wPrice = calcPrice(RS("�󒍒P��"), wSalesTaxRate)
	vTotalAm = vTotalAm + (wPrice * RS("�󒍐���"))
	wHTML = wHTML & "        <td>" & FormatNumber(wPrice,0) & "�~</td>" & vbNewLine

	'------------- ����
	wHTML = wHTML & "        <td>" & vbNewLine
	wHTML = wHTML & "          <input type='text' name='qt" & RS("�󒍖��הԍ�") & "' id='order_form_qt1' value='" & RS("�󒍐���") & "' size=4 onBlur='qt_onBlur(this);'>" & vbNewLine
	wHTML = wHTML & "          <input type='hidden' name='oldqt" & RS("�󒍖��הԍ�") & "' value='" & RS("�󒍐���") & "'>" & vbNewLine
	wHTML = wHTML & "        </td>" & vbNewLine

	'------------- ���z
	wHTML = wHTML & "        <td>" & FormatNumber(wPrice*RS("�󒍐���"),0) & "�~</td>" & vbNewLine

	'------------- �폜�{�^��
	wHTML = wHTML & "        <td>" & vbNewLine
	wHTML = wHTML & "          <ul>" & vbNewLine
	wHTML = wHTML & "            <li><a href='JavaScript:delete_onClick(" & RS("�󒍖��הԍ�") & ");'><img src='images/btn_delete.png' alt='�폜' class='opover' ></a></li>" & vbNewLine
	'------------- ��Ŕ����{�^��
	If userID <> "" Then
		wHTML = wHTML & "        <li><a href='WishListAdd.asp?OrderDetailNo=" & RS("�󒍖��הԍ�") & "&Item=" & Server.URLEncode(RS("���[�J�[�R�[�h") & "^" & RS("���i�R�[�h") & "^" & RS("�F") & "^" & RS("�K�i")) & "' class='link'><img src='images/btn_later.png' alt='��Ŕ���' class='opover' ></a></li>" & vbNewLine
	End If
	wHTML = wHTML & "          </ul>" & vbNewLine
	wHTML = wHTML & "        </td>" & vbNewLine
	wHTML = wHTML & "      </tr>" & vbNewLine

	'---- �Ō�ɃJ�[�g�ɓ��ꂽ���i�̃��R�����h�\���p
	wRecommendMakerCd = RS("���[�J�[�R�[�h")
	wRecommendProductCd = RS("���i�R�[�h")

	'2013/05/20 GV #1505 add start
	'���Ԃ݂��ƁI���R�����h�pJS�ɓn�����iID
	wOrderProductId = wOrderProductId & "'" & RS("���iID") & "',"
	'2013/05/20 GV #1505 add end

	RS.MoveNext

Loop

'---- ����
Call fCalcShipping(gSessionID, "�ꊇ", vFreightAm, vFreightForwarder, vSoukoCnt, vKoguchi)		'2011/04/14 hn mod
wPrice = Fix(vFreightAm * (100 + wSalesTaxRate) / 100)

'----���i���v���z�C�Čv�Z�{�^��
wHTML = wHTML & "      <tr>" & vbNewLine
wHTML = wHTML & "        <td colspan='6'>" & vbNewLine
wHTML = wHTML & "          <dl class='total'>" & vbNewLine
wHTML = wHTML & "            <dt>���i���v�i�ō��j</dt><dd>" & FormatNumber(vTotalAm,0) & "�~</dd>" & vbNewLine
wHTML = wHTML & "            <dt>�������ρi�ō��j</dt><dd>" & FormatNumber(wPrice,0) & "�~</dd>" & vbNewLine
wHTML = wHTML & "          </dl>" & vbNewLine
wHTML = wHTML & "        </td>" & vbNewLine
wHTML = wHTML & "        <td><a href='JavaScript:calc_onClick();'><img src='images/btn_calculate.png' alt='�Čv�Z' class='opover' ></a></td>" & vbNewLine
wHTML = wHTML & "      </tr>" & vbNewLine

wOrderProductHTML = wHTML

'---- ���R�����h�f�[�^�쐬
Call CreateRecommendInfo()

RS.close

End Function

'========================================================================
'
'	Function	�ۑ����ꂽ�J�[�g��񂪂��邩�ǂ����`�F�b�N
'
'========================================================================
'
Function CheckSavedCart()

Dim Rsv

'----�ۑ��J�[�g�f�[�^���o��
wSQL = ""
wSQL = wSQL & "SELECT a.�ڋq�ԍ�"
wSQL = wSQL & "  FROM �ۑ��J�[�g a WITH (NOLOCK)"
wSQL = wSQL & " WHERE �ڋq�ԍ� = " & UserID

'@@@@response.write(wSQL)

Set Rsv = Server.CreateObject("ADODB.Recordset")
Rsv.Open wSQL, Connection, adOpenStatic

if RSv.EOF = false then
	wSavedCartFl = "Y"
end if

Rsv.Close

End function


'========================================================================
'
'	Function	���R�����h���i�擾  '2010/04/02 an add
'
'========================================================================
'
Function CreateRecommendInfo()

'2013/08/07 if-web del s
'Dim RSv
'
'---- ���R�����h���i�擾(�ގ��x���傫��5���i)
'wSQL = ""
'
'wSQL = wSQL & "SELECT DISTINCT TOP 5"
'wSQL = wSQL & "       a.���[�J�[�R�[�h"
'wSQL = wSQL & "     , a.���i�R�[�h"
'wSQL = wSQL & "     , a.���i��"
'wSQL = wSQL & "     , a.���i�摜�t�@�C����_��"
'wSQL = wSQL & "     , CASE"
'wSQL = wSQL & "         WHEN (a.�����萔�� > a.������󒍍ϐ��� AND a.�����萔�� > 0) THEN a.������P��"
'wSQL = wSQL & "         ELSE a.�̔��P��"
'wSQL = wSQL & "       END AS �̔��P��"
'wSQL = wSQL & "     , a.ASK���i�t���O"
'wSQL = wSQL & "     , a.�J�e�S���[�R�[�h"
'wSQL = wSQL & "     , b.���[�J�[��"
'wSQL = wSQL & "     , e.�ގ��x"
'wSQL = wSQL & "  FROM Web���i a WITH (NOLOCK)"
'wSQL = wSQL & "     , ���[�J�[ b WITH (NOLOCK)"
'wSQL = wSQL & "     , Web�F�K�i�ʍ݌� d WITH (NOLOCK)"
'wSQL = wSQL & "     , ���R�����h���ʍw�� e WITH (NOLOCK)"
'wSQL = wSQL & " WHERE a.���[�J�[�R�[�h = e.���R�����h���[�J�[�R�[�h"
'wSQL = wSQL & "   AND a.���i�R�[�h = e.���R�����h���i�R�[�h"
'wSQL = wSQL & "   AND d.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'wSQL = wSQL & "   AND d.���i�R�[�h = a.���i�R�[�h"
'wSQL = wSQL & "   AND b.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'wSQL = wSQL & "   AND a.Web���i�t���O = 'Y'"
'wSQL = wSQL & "   AND a.�戵���~�� IS NULL"
'wSQL = wSQL & "   AND ((a.�p�ԓ� IS NULL) OR (a.�p�ԓ� IS NOT NULL AND d.�����\���� > 0))"
'wSQL = wSQL & "   AND e.���[�J�[�R�[�h = '" & wRecommendMakerCd & "'"
'wSQL = wSQL & "   AND e.���i�R�[�h = '" & wRecommendProductCd & "'"
'wSQL = wSQL & " ORDER BY"
'wSQL = wSQL & "       e.�ގ��x DESC"
'wSQL = wSQL & "     , a.�J�e�S���[�R�[�h"
'
'@@@@response.write(wSQL)
'
'Set RSv = Server.CreateObject("ADODB.Recordset")
'RSv.Open wSQL, Connection, adOpenStatic
'
'wHTML = ""
'
'if RSv.EOF = false then
'
'	wHTML = ""
'	wHTML = wHTML & "  <h2 class='detail_title'>���̃A�C�e���𔃂����l�͂���ȃA�C�e���������Ă��܂��B</h2>" & vbNewLine
'	wHTML = wHTML & "    <ul class='relation'>" & vbNewLine
'
'	Do Until RSv.EOF = True
'
'	wPrice = calcPrice(RSv("�̔��P��"), wSalesTaxRate)
'
'		wHTML = wHTML & "      <li>" & vbNewLine
'		wHTML = wHTML & "        <p><a href='ProductDetail.asp?Item=" & RSv("���[�J�[�R�[�h") & "^" & RSv("���i�R�[�h") & "'><img src='"
'		wHTML = wHTML & "prod_img/" & RSv("���i�摜�t�@�C����_��") & "' alt='" & RSv("���[�J�[��") & " / " & RSv("���i��") & "' class='opover'><span>"
'		wHTML = wHTML & RSv("���[�J�[��") & "</span><span>" & RSv("���i��") & "</span></a></p>" & vbNewLine
'		If RSv("ASK���i�t���O") <> "Y" Then
'			wHTML = wHTML & "        <p>" & FormatNumber(wPrice,0) & "�~(�ō�)</p>" & vbNewLine
'		Else
'			wHTML = wHTML & "        <p><a class='tip'>ASK<span>"& FormatNumber(wPrice,0) & "�~(�ō�)</span></p>" & vbNewLine
'		End If
'		
'		wHTML = wHTML & "      </li>" & vbNewLine
'
'		RSv.MoveNext
'
'	Loop
'
'	wHTML = wHTML & "    </ul>"
'
'End if
'
'RSv.Close

'wHTML = wHTML & fEAgency_CreateRecommendCartJS(wOrderProductId)	'2013/05/20 GV #1505 add
'2013/08/07 if-web del e

wHTML = fEAgency_CreateRecommendCartJS(wOrderProductId)	'2013/08/07 if-web add

wRecommendHTML = wHTML

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
<title>�V���b�s���O�J�[�g�b�T�E���h�n�E�X</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/shop.css" type="text/css">
<link rel="stylesheet" href="style/StyleOrder.css?20120703" type="text/css">
<link rel="stylesheet" href="style/ask.css?20140401a" type="text/css">

<script type="text/javascript">

var g_change_fl = false;

//
//	���ʕύX��
//
function qt_onBlur(p_formItem){

var v_itemName;

	v_itemName = "old" + p_formItem.name;
	if (p_formItem.value != document.f_order_list.elements[v_itemName].value){
		g_change_fl = true;
	}
}

//
//	�󒍖��׍sDelete
//
function delete_onClick(p_detail_no){

	document.f_order_list.detail_no.value = p_detail_no;
	document.f_order_list.action = "OrderChange.asp";
	document.f_order_list.submit();
}

//
//	�Čv�Z
//
function calc_onClick(){

	g_change_fl = false;
	document.f_order_list.detail_no.value = "all";
	document.f_order_list.action = "OrderChange.asp";
	document.f_order_list.submit();
}

//
//	�I�[�_�[Submit
//
function order_onSubmit(){

//	���ʕύX�����邩�ǂ����`�F�b�N
	if (g_change_fl == true){
		alert("���ʂ��ύX����Ă��܂��B�@�Čv�Z�{�^���������Ă��������B");
		return;
	}

	window.location = g_HTTPS + "shop/LoginCheck.asp?called_from=order";
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
      <li class="now">�V���b�s���O�J�[�g</li>
    </ul>
  </div></div></div>

  <p class="error"><% = msg %></p>

  <h1 class="title">�V���b�s���O�J�[�g</h1>
  <ol id="step">
    <li><img src="images/step01_now.gif" alt="1.�V���b�s���O�J�[�g" width="170" height="50"></li>
    <li><img src="images/step02.gif" alt="2.���͂���A���x�����@�̑I��" width="170" height="50"></li>
    <li><img src="images/step03.gif" alt="3.���������e�̊m�F" width="170" height="50"></li>
    <li><img src="images/step04.gif" alt="4.����������" width="170" height="50"></li>
  </ol>

  <h2 class="cart_title">�J�[�g���e</h2>
  <form method='post' name='f_order_list' action='' onsubmit='calc_onClick();'>
    <table id="cart">
<% = wOrderProductHTML %>
    </table>
    <input type='hidden' name='detail_no' value=''>
  </form>

  <ul id="attention">
    <li>���̉�ʂŐ��ʂ̕ύX���ł��܂��B ���ʂ�ύX��u�Čv�Z�v�{�^���������ƍX�V����܂��B</li>
    <li>�������i���������ꍇ�́u�폜�v�{�^���������Ă��������B</li>
    <li>�������i��ǉ�����ꍇ�͉��́u�������𑱂���v�{�^�����N���b�N���ď��i�̉�ʂɖ߂��Ă��������B</li>
    <li>�z�����@�A�z����ɂ�著���͕ς�邱�Ƃ�����܂��B</li>
    <li>���̉�ʂ̋��z�͏��i���̒P���̊m�F�ƂȂ�ŏI�I�ȍ��v���z�ł͂���܂���B</li>
    <li>�ŏI�I�Ȃ��x�������z�͕ʓr���ē����邲�����m�F�������Q�Ƃ��������B</li>
  </ul>

<% If wNoData = False Then %>
  <div id="btn_box">
    <ul class="btn">
      <li><a href="javascript:history.back();"><img src="images/btn_continue.png" alt="�������𑱂���" class="opover"></a></li>
      <li class="last"><a href="javascript:order_onSubmit();"><img src="images/btn_order.png" alt="�������葱����" class="opover"></a></li>
    </ul>
  </div>
<% End If %>

<% If (wNoData = False Or wSavedCartFl = "Y") And userID <> "" Then %>
  <div class="btn_box">
<% If wNoData = False And wSavedCartFl = "Y" Then %>
    <ul class="btn">
      <li><a href="SaveCart.asp"><img src="images/btn_cartsave.png" alt="���̃J�[�g���e��ۑ�" class="opover"></a></li>
      <li class="last"><a href="SaveCartList.asp"><img src="images/btn_cartlist.png" alt="�ۑ����ꂽ�J�[�g�ꗗ��" class="opover"></a></li>
    </ul>
<% Elseif  wNoData = False Then %>
    <div class="btn"><a href="SaveCart.asp"><img src="images/btn_cartsave.png" alt="���̃J�[�g���e��ۑ�" class="opover"></a></div>
<% Else %>
    <div class="btn"><a href="SaveCartList.asp"><img src="images/btn_cartlist.png" alt="�ۑ����ꂽ�J�[�g�ꗗ��" class="opover"></a></div>
<% End If %>
  </div>
<% End If %>

  <ul class="info left">
    <li><a href="../guide/change.asp">���������i�̃L�����Z���E�ԕi�ɂ���</a></li>
    <li><a href="../guide/nouki.asp">���i�̔[���ɂ��Ă͂�����</a></li>
  </ul>
  <ul class="info right">
    <li class="no"><a href="../shopEng/Order.asp">English</a></li>
  </ul>

  <!-- ���R�����h���i start -->
<% = wRecommendHTML %>
  <!-- ���R�����h���i end -->
  <!--/#contents --></div>
  <div id="globalSide">
    <!--#include file="../Navi/NaviSide.inc"-->
  <!--/#globalSide --></div>
<!--/#main --></div>
<!--#include file="../Navi/Navibottom.inc"-->
<div class="tooltip"><p>ASK</p></div>
<!--#include file="../Navi/NaviScript.inc"-->
<script type="text/javascript" src="jslib/ask.js?20140401a"></script>
</body>
</html>