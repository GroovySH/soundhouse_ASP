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
'	�A�N�Z�X�����L���O�@(���i�r���[�A�F�B�Ɋ��߂�A�E�B�b�V�����X�g)
'
'	�X�V����
'2007/10/16 �O�������L���O�\���ɕύX
'2009/04/30 �G���[����error.asp�ֈړ�
'2010/02/18 an ASK���i�p�����[�^��Server.URLEncode���s�Ȃ�
'2010/05/10 an ���j���[�A���Ή��i�J�[�g�{�^���A�݌ɏ��A���iID�\���ǉ��j
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2012/01/19 GV �f�[�^�擾 SELECT���� LAC�N�G���[�Ă�K�p
'2012/01/19 GV NOLOCK�I�v�V�����t�^�R��Ή�
'2012/01/23 GV �u���i���r���[�v�e�[�u������u���i���r���[�W�v�v�e�[�u���g�p�ɕύX (CreateReviewImg()�v���V�[�W��)
'2012/08/07 if-web ���j���[�A�����C�A�E�g����
'2014/03/19 GV ����ő��łɔ���2�d�\���Ή�
'
'========================================================================

On Error Resume Next

Dim RankType

Dim wSalesTaxRate
Dim wYYYYMM

Dim wRank
Dim wItem
Dim wPrice
Dim wPriceNoTax		'2014/03/19 GV add


Dim wProdTermFl '�̔��I�����i�t���O
Dim wInventoryCd
Dim wInventoryImage

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
Dim wTop3HTML
Dim wUnder4HTML
Dim wMSG
Dim wErrDesc   '2011/08/01 an add

'========================================================================

'---- ���M�f�[�^�[�̎��o��
RankType = ReplaceInput(Request("RankType"))

if RankType = "" then
	RankType = "���i�r���["
end if

'---- Execute main
call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "RankingAccess.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
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

Dim i
Dim vPrevMonth

'---- ����ŗ���o��
call getCntlMst("����","����ŗ�","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)			'����ŗ�
wSalesTaxRate = Clng(wItemNum1)

'---- �O��
vPrevMonth = DateAdd("m", -1, Date())
wYYYYMM = Year(vPrevMonth) &  Right("0" & Month(vPrevMonth), 2)

'---- �����L���O���o��
wSQL = ""
' 2012/01/19 GV Mod Start
'wSQL = wSQL & "SELECT TOP 20"
'wSQL = wSQL & "       a.���[�J�[�R�[�h"
'wSQL = wSQL & "     , a.���i�R�[�h"
'wSQL = wSQL & "     , b.���i��"
'wSQL = wSQL & "     , b.���i�摜�t�@�C����_��"
'wSQL = wSQL & "     , b.�����ߏ��i�R�����g"
'wSQL = wSQL & "     , b.���i�T��Web"
'wSQL = wSQL & "     , b.�戵���~��"
'wSQL = wSQL & "     , b.������"
'wSQL = wSQL & "     , b.�p�ԓ�"
'wSQL = wSQL & "     , b.B�i�t���O"
'wSQL = wSQL & "     , b.ASK���i�t���O"
'wSQL = wSQL & "     , b.�󏭐���"
'wSQL = wSQL & "     , b.�����萔��"
'wSQL = wSQL & "     , b.������󒍍ϐ���"
'wSQL = wSQL & "     , b.�Z�b�g���i�t���O"
'wSQL = wSQL & "     , b.���[�J�[�������敪"
'wSQL = wSQL & "     , b.���ח\�薢��t���O"
'wSQL = wSQL & "     , b.Web�[����\���t���O"
'wSQL = wSQL & "     , CASE"
'wSQL = wSQL & "         WHEN (b.�����萔�� > b.������󒍍ϐ��� AND b.�����萔�� > 0) THEN b.������P��"
'wSQL = wSQL & "         ELSE b.�̔��P��"
'wSQL = wSQL & "       END AS �̔��P��"
'wSQL = wSQL & "     , c.���[�J�[��"
'wSQL = wSQL & "     , d.�J�e�S���[�R�[�h"
'wSQL = wSQL & "     , d.�J�e�S���[��"
'wSQL = wSQL & "     , e.�F"
'wSQL = wSQL & "     , e.�K�i"
'wSQL = wSQL & "     , e.B�i�����\����"
'wSQL = wSQL & "     , e.�����\����"
'wSQL = wSQL & "     , e.�����\���ח\���"
'wSQL = wSQL & "     , e.���iID"
'
''�F�K�i�����邩�ǂ��� 2007/05/30
'wSQL = wSQL & "     , (SELECT COUNT(*)"
'wSQL = wSQL & "          FROM Web�F�K�i�ʍ݌� f WITH (NOLOCK)"
'wSQL = wSQL & "         WHERE f.���[�J�[�R�[�h = b.���[�J�[�R�[�h"
'wSQL = wSQL & "           AND f.���i�R�[�h = b.���i�R�[�h"
'wSQL = wSQL & "           AND (f.�F != '' OR f.�K�i != '')"
'wSQL = wSQL & "           AND f.�I���� IS NULL"
'wSQL = wSQL & "       ) AS �F�K�iCNT"
'
'wSQL = wSQL & "  FROM ���i�A�N�Z�X���� a WITH (NOLOCK)"
'wSQL = wSQL & "     , Web���i b WITH (NOLOCK)"
'wSQL = wSQL & "     , ���[�J�[ c WITH (NOLOCK)"
'wSQL = wSQL & "     , �J�e�S���[ d WITH (NOLOCK)"
'wSQL = wSQL & "     , Web�F�K�i�ʍ݌� e WITH (NOLOCK)"
'
'wSQL = wSQL & " WHERE b.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'wSQL = wSQL & "   AND b.���i�R�[�h = a.���i�R�[�h"
'wSQL = wSQL & "   AND c.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'wSQL = wSQL & "   AND d.�J�e�S���[�R�[�h = b.�J�e�S���[�R�[�h"
'wSQL = wSQL & "   AND e.���[�J�[�R�[�h = b.���[�J�[�R�[�h"
'wSQL = wSQL & "   AND e.���i�R�[�h = b.���i�R�[�h"
'wSQL = wSQL & "   AND b.Web���i�t���O = 'Y'"
'wSQL = wSQL & "   AND e.�F = ''"
'wSQL = wSQL & "   AND e.�K�i = ''"
'wSQL = wSQL & "   AND a.�N�� = '" & wYYYYMM & "'"
'
'wSQL = wSQL & " ORDER BY"
'
'if RankType = "���i�r���[" then
'	wSQL = wSQL & "       a.�y�[�W�r���[���� DESC"
'end if
'if RankType = "�F�B�ɂ�����" then
'	wSQL = wSQL & "       a.�F�B�ɂ����ߌ��� DESC"
'end if
'if RankType = "�~�������̃��X�g" then
'	wSQL = wSQL & "       a.�E�B�b�V�����X�g���� DESC"
'end if
'
'wSQL = wSQL & "     , c.���[�J�[��"
'wSQL = wSQL & "     , b.���i��"
wSQL = wSQL & "SELECT TOP 20 "
wSQL = wSQL & "      a.���[�J�[�R�[�h "
wSQL = wSQL & "    , a.���i�R�[�h "
wSQL = wSQL & "    , b.���i�� "
wSQL = wSQL & "    , b.���i�摜�t�@�C����_�� "
wSQL = wSQL & "    , b.�����ߏ��i�R�����g "
wSQL = wSQL & "    , b.���i�T��Web "
wSQL = wSQL & "    , b.�戵���~�� "
wSQL = wSQL & "    , b.������ "
wSQL = wSQL & "    , b.�p�ԓ� "
wSQL = wSQL & "    , b.B�i�t���O "
wSQL = wSQL & "    , b.ASK���i�t���O "
wSQL = wSQL & "    , b.�󏭐��� "
wSQL = wSQL & "    , b.�����萔�� "
wSQL = wSQL & "    , b.������󒍍ϐ��� "
wSQL = wSQL & "    , b.�Z�b�g���i�t���O "
wSQL = wSQL & "    , b.���[�J�[�������敪 "
wSQL = wSQL & "    , b.���ח\�薢��t���O "
wSQL = wSQL & "    , b.Web�[����\���t���O "
wSQL = wSQL & "    , CASE "
wSQL = wSQL & "        WHEN (b.�����萔�� > b.������󒍍ϐ��� AND b.�����萔�� > 0) THEN b.������P�� "
wSQL = wSQL & "        ELSE b.�̔��P�� "
wSQL = wSQL & "      END AS �̔��P�� "
wSQL = wSQL & "    , c.���[�J�[�� "
wSQL = wSQL & "    , d.�J�e�S���[�R�[�h "
wSQL = wSQL & "    , d.�J�e�S���[�� "
wSQL = wSQL & "    , e.�F "
wSQL = wSQL & "    , e.�K�i "
wSQL = wSQL & "    , e.B�i�����\���� "
wSQL = wSQL & "    , e.�����\���� "
wSQL = wSQL & "    , e.�����\���ח\��� "
wSQL = wSQL & "    , e.���iID "
wSQL = wSQL & "    , (SELECT COUNT(f.���i�R�[�h) "
wSQL = wSQL & "         FROM Web�F�K�i�ʍ݌� f WITH (NOLOCK) "
wSQL = wSQL & "        WHERE     f.���[�J�[�R�[�h = b.���[�J�[�R�[�h "
wSQL = wSQL & "              AND f.���i�R�[�h = b.���i�R�[�h "
wSQL = wSQL & "              AND (   f.�F   != '' "
wSQL = wSQL & "                   OR f.�K�i != '') "
wSQL = wSQL & "              AND f.�I���� IS NULL "
wSQL = wSQL & "      ) AS �F�K�iCNT "
wSQL = wSQL & "FROM "
wSQL = wSQL & "    ���i�A�N�Z�X����             a WITH (NOLOCK) "
wSQL = wSQL & "      INNER JOIN Web���i         b WITH (NOLOCK) "
wSQL = wSQL & "        ON     b.���[�J�[�R�[�h   = a.���[�J�[�R�[�h "
wSQL = wSQL & "           AND b.���i�R�[�h       = a.���i�R�[�h "
wSQL = wSQL & "      INNER JOIN ���[�J�[        c WITH (NOLOCK) "
wSQL = wSQL & "        ON     c.���[�J�[�R�[�h   = a.���[�J�[�R�[�h "
wSQL = wSQL & "      INNER JOIN �J�e�S���[      d WITH (NOLOCK) "
wSQL = wSQL & "        ON     d.�J�e�S���[�R�[�h = b.�J�e�S���[�R�[�h "
wSQL = wSQL & "      INNER JOIN Web�F�K�i�ʍ݌� e WITH (NOLOCK) "
wSQL = wSQL & "        ON     e.���[�J�[�R�[�h   = b.���[�J�[�R�[�h "
wSQL = wSQL & "           AND e.���i�R�[�h       = b.���i�R�[�h "
wSQL = wSQL & "      LEFT JOIN ( SELECT 'Y' AS 'ShohinWebY' ) t1 "
wSQL = wSQL & "        ON     b.Web���i�t���O    = t1.ShohinWebY "
wSQL = wSQL & "      LEFT JOIN ( SELECT ''  AS 'Iro' )        t2 "
wSQL = wSQL & "        ON     e.�F               = t2.Iro "
wSQL = wSQL & "      LEFT JOIN ( SELECT ''  AS 'Kikaku' )     t3 "
wSQL = wSQL & "        ON     e.�K�i             = t3.Kikaku "
wSQL = wSQL & "WHERE "
wSQL = wSQL & "        t1.ShohinWebY   IS NOT NULL "
wSQL = wSQL & "    AND t2.Iro          IS NOT NULL "
wSQL = wSQL & "    AND t3.Kikaku       IS NOT NULL "
wSQL = wSQL & "    AND a.�N�� = '" & wYYYYMM & "' "
wSQL = wSQL & "ORDER BY "
If RankType     = "���i�r���[" Then
	wSQL = wSQL & "      a.�y�[�W�r���[���� DESC "
ElseIf RankType = "�F�B�ɂ�����" Then
	wSQL = wSQL & "      a.�F�B�ɂ����ߌ��� DESC "
ElseIf RankType = "�~�������̃��X�g" Then
	wSQL = wSQL & "      a.�E�B�b�V�����X�g���� DESC "
End If
wSQL = wSQL & "    , c.���[�J�[�� "
wSQL = wSQL & "    , b.���i�� "
' 2012/01/19 GV Mod End

'@@@@response.write(wSQL)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

if RS.EOF = true then
	exit function
end if

wRank = 0
wHTML = ""
wTop3HTML = ""
wUnder4HTML = ""

Do Until RS.EOF = true
	wRank = wRank + 1
	wItem = Server.URLEncode(RS("���[�J�[�R�[�h") & "^" & RS("���i�R�[�h") & "^" & "^")
	wPrice = calcPrice(RS("�̔��P��"), wSalesTaxRate)
	wPriceNoTax = RS("�̔��P��")							'2014/03/19 GV add

	'---- �݌ɏ󋵕\���̂��߁A�I���`�F�b�N
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
	if RS("B�i�t���O") = "Y" AND RS("B�i�����\����") <= 0 then    'B�i�ō݌ɂȂ�
		wProdTermFl = "Y"
	end if

	'---- �݌ɏ󋵍쐬
	if RS("�F�K�iCNT") = 0 then
		wInventoryCd = GetInventoryStatus(RS("���[�J�[�R�[�h"),RS("���i�R�[�h"),RS("�F"),RS("�K�i"),RS("�����\����"),RS("�󏭐���"),RS("�Z�b�g���i�t���O"),RS("���[�J�[�������敪"),RS("�����\���ח\���"),wProdTermFl)

		'---- �݌ɏ󋵁A�F���ŏI�Z�b�g
		call GetInventoryStatus2(RS("�����\����"), RS("Web�[����\���t���O"), RS("���ח\�薢��t���O"), RS("�p�ԓ�"), RS("B�i�t���O"), RS("B�i�����\����"), RS("�����萔��"), RS("������󒍍ϐ���"), wProdTermFl, wInventoryCd, wInventoryImage)

	end if

	'---- 1�`3��
	if wRank <= 3 then
		call CreateTop3ItemHTML()
	'---- 4�`20��
	else
		call CreateUnder4ItemHTML()
	end if

	RS.MoveNext
Loop

RS.Close

End function


'========================================================================
'
'	Function	�����L���O���3���i�̕\��
'
'========================================================================

Function CreateTop3ItemHTML()

Dim vBGColor

'---- �����s�Ɗ�s�ŕ\���F��ς���
if wRank Mod 2 <> 0 then
	vBGColor = "bg_color1"   '��s
else
	vBGColor = "bg_color2"   '�����s
end if

wHTML = wHTML & "<!-- bigbox1 START -->" & vbNewLine
wHTML = wHTML & "    <div class='rankingaccess_bigbox'>" & vbNewLine
wHTML = wHTML & "      <div class='left " & vBGColor & "'>" & vbNewLine
wHTML = wHTML & "        <div class='crown_box'><img src='images/ranking/ico_no" & wRank & "crown.gif' alt='' width='41' height='30'></div>" & vbNewLine
wHTML = wHTML & "      </div>" & vbNewLine
wHTML = wHTML & "      <div class='centerleft'>" & vbNewLine
wHTML = wHTML & "        <a href='ProductDetail.asp?Item=" & wItem & "'>" & vbNewLine
if RS("���i�摜�t�@�C����_��") <> "" then
	wHTML = wHTML & "          <img src='prod_img/" & RS("���i�摜�t�@�C����_��") & "' alt='" & RS("���[�J�[��") & " " & RS("���i��") & "'>" & vbNewLine
else
	wHTML = wHTML & "          <img src='prod_img/n/nopict.jpg' alt='" & RS("���[�J�[��") & " " & RS("���i��") & "'>" & vbNewLine
end if
wHTML = wHTML & "        </a>" & vbNewLine
wHTML = wHTML & "      </div>" & vbNewLine
wHTML = wHTML & "      <div class='center'>" & vbNewLine
wHTML = wHTML & "        <h2>" & vbNewLine
wHTML = wHTML & "          <a href='ProductDetail.asp?Item=" & wItem & "'>" & vbNewLine
wHTML = wHTML & "            <span class='txt_maker'>" & RS("���[�J�[��") & "</span>&nbsp;" & vbNewLine
wHTML = wHTML & "            <span class='txt_product'>" & RS("���i��") & "</span>" & vbNewLine
wHTML = wHTML & "          </a>" & vbNewLine
wHTML = wHTML & "        </h2>" & vbNewLine
wHTML = wHTML & "        <h3>" & vbNewLine
wHTML = wHTML & "          <a href='SearchList.asp?i_type=c&s_category_cd=" & RS("�J�e�S���[�R�[�h") & "'>" & RS("�J�e�S���[��") & "</a>" & vbNewLine
wHTML = wHTML & "        </h3>" & vbNewLine
wHTML = wHTML & "        <div class='bg'>" & vbNewLine
wHTML = wHTML & "          <div class='price_box'>" & vbNewLine
wHTML = wHTML & "            �Ռ������F "

if RS("ASK���i�t���O") = "Y" then
	wHTML = wHTML & "ASK" & vbNewLine
else
'2014/03/19 GV mod start ---->
'	wHTML = wHTML & "<strong>" & FormatNumber(wPrice,0) & "�~(�ō�)</strong>�@" & vbNewLine
	wHTML = wHTML & "<strong>" & FormatNumber(wPriceNoTax,0) & "�~(�Ŕ�)</strong>�@" & vbNewLine
	wHTML = wHTML & "(�ō�&nbsp;" & FormatNumber(wPrice,0) & "�~)�@" & vbNewLine
'2014/03/19 GV mod end   <----
end if

wHTML = wHTML & "          </div>" & vbNewLine
wHTML = wHTML & "          <div class='notes_box'>" & vbNewLine

'----- ���i���r���[
wHTML = wHTML & "            " & CreateReviewImg(RS("���[�J�[�R�[�h"), RS("���i�R�[�h"))

wHTML = wHTML & "          </div>" & vbNewLine
wHTML = wHTML & "        </div>" & vbNewLine

'---- ���i����
wHTML = wHTML & "        <p>"
if RS("�����ߏ��i�R�����g") <> "" then
	wHTML = wHTML & Replace(RS("�����ߏ��i�R�����g"), vbNewLine, "<br>")
else
	wHTML = wHTML & Replace(RS("���i�T��Web"), vbNewLine, "<br>")
end if
wHTML = wHTML & "        </p>" & vbNewLine

wHTML = wHTML & "      </div>" & vbNewLine
wHTML = wHTML & "      <div class='rankingaccess_shopbox'>" & vbNewLine
wHTML = wHTML & "        <form name='f_item' method='post' action='OrderPreInsert.asp' onSubmit='return order_onClick(this);'>" & vbNewLine
'wHTML = wHTML & "        <div class='right'>" & vbNewLine

'---- �F�K�i�Ȃ�
if RS("�F�K�iCNT") = 0 then
	if wProdTermFl = "Y" then
		wHTML = wHTML & "            <img src='images/icon_sold.gif'><br>" & vbNewLine
	else
		wHTML = wHTML & "            <input type='hidden' name='qt' value='1'>" & vbNewLine
		wHTML = wHTML & "            <input type='hidden' name='maker_cd' value='" & RS("���[�J�[�R�[�h") & "'>" & vbNewLine
		wHTML = wHTML & "            <input type='hidden' name='product_cd' value='" & RS("���i�R�[�h") & "'>" & vbNewLine
		wHTML = wHTML & "            <input type='hidden' name='category_cd' value='" & RS("�J�e�S���[�R�[�h") & "'>" & vbNewLine
		wHTML = wHTML & "            <input type='image' src='images/btn_cart.png' style='vertical-align:middle' alt='�J�[�g��' class='opover'><br>" & vbNewLine
	end if

'----�F�K�i����
else
	if wProdTermFl = "Y" then
		wHTML = wHTML & "            <img src='images/icon_sold.gif'><br>" & vbNewLine
	else
		wHTML = wHTML & "            <input type='hidden' name='qt' value='0'>" & vbNewLine
		wHTML = wHTML & "            <a href='ProductDetail.asp?Item=" & wItem & "'><img src='images/btn_detail.png'></a><br>" & vbNewLine
	end if
end if

'wHTML = wHTML & "        </div>" & vbNewLine
wHTML = wHTML & "        <div class='shopid'>���iID:" & RS("���iID") & "</div>" & vbNewLine
wHTML = wHTML & "        </form>" & vbNewLine
wHTML = wHTML & "        <div class='itemstock'><img src='images/" & wInventoryImage & "' alt=''> " & wInventoryCd & "</div>" & vbNewLine
wHTML = wHTML & "      </div>" & vbNewLine
wHTML = wHTML & "    </div>" & vbNewLine
wHTML = wHTML & "<!-- bigbox END -->" & vbNewLine

wTop3HTML = wHTML

End function

'========================================================================
'
'	Function	�����L���O4�ʈȉ��̏��i�̕\��
'
'========================================================================

Function CreateUnder4ItemHTML()

Dim vBGColor
Dim vMakerProduct
Dim vProductName

wHTML = ""

'---- 4�ʂ̎�O�ɍ��ږ��̍s��\��
if wRank = 4 then
	wHTML = wHTML & "    <!-- s_box TH START -->" & vbNewLine
	wHTML = wHTML & "    <div id='s_box_th_ra'>" & vbNewLine
	wHTML = wHTML & "      <div id='th_no'>����</div>" & vbNewLine
	wHTML = wHTML & "      <div id='th_prod'><div class='cell'>���[�J�[�@���i</div></div>" & vbNewLine
	wHTML = wHTML & "      <div id='th_cat'>�J�e�S���[</div>" & vbNewLine
	wHTML = wHTML & "      <div id='th_point'>���r���[�|�C���g</div>" & vbNewLine
	wHTML = wHTML & "      <div id='th_stock'>�݌ɏ�</div>" & vbNewLine
	wHTML = wHTML & "      <div id='th_cart'>�J�[�g</div>" & vbNewLine
	wHTML = wHTML & "    </div>" & vbNewLine
	wHTML = wHTML & "    <!-- s_box TH END -->" & vbNewLine
end if

'---- �����s�Ɗ�s�ŕ\���F��ς���
if wRank Mod 2 <> 0 then
	vBGColor = "s_box1"
else
	vBGColor = "s_box2"
end if

wHTML = wHTML & "    <!-- s_box START -->" & vbNewLine
wHTML = wHTML & "    <div class='" & vBGColor & "'>" & vbNewLine
wHTML = wHTML & "      <div class='s_box_height'>" & vbNewLine
wHTML = wHTML & "        <div class='num_box'>" & wRank & "</div>" & vbNewLine
wHTML = wHTML & "        <div class='text_box'>" & vbNewLine
wHTML = wHTML & "          <a href='ProductDetail.asp?Item=" & wItem & "'>" & vbNewLine

'--- ���[�J�[���{���i����������2�s�ɂȂ�ꍇ��"..."�ŏȗ�
vMakerProduct = RS("���[�J�[��") & " " & RS("���i��")
if Len(vMakerProduct) > 33 then

	vProductName = Left(RS("���i��"), 30-Len(RS("���[�J�[��"))) &  "..."
else
	vProductName = RS("���i��")
end if

wHTML = wHTML & "            <span class='txt_maker'>" & RS("���[�J�[��") & "</span>&nbsp;" & vbNewLine
wHTML = wHTML & "            <span class='txt_product'>" & vProductName & "</span><br>" & vbNewLine
wHTML = wHTML & "            <span class='txt_price_h'>�Ռ������F</span>" & vbNewLine

if RS("ASK���i�t���O") = "Y" then
	wHTML = wHTML & "ASK"
else
'2014/03/19 GV mod start ---->
'	wHTML = wHTML & "            <span class='txt_price_d'>" & FormatNumber(wPrice,0) & "�~(�ō�)</span>" & vbNewLine
	wHTML = wHTML & "            <span class='txt_price_d'>" & FormatNumber(wPriceNoTax,0) & "�~(�Ŕ�)</span>�@"
	wHTML = wHTML & "<span class='txt_price_t'>(�ō�&nbsp;"&FormatNumber(wPrice,0) & "�~)</span>" & vbNewLine
'2014/03/19 GV mod end   <----
end if
wHTML = wHTML & "          </a>" & vbNewLine
wHTML = wHTML & "        </div>" & vbNewLine
wHTML = wHTML & "        <div class='cat_box'>" & vbNewLine
wHTML = wHTML & "          <a href='SearchList.asp?i_type=c&s_category_cd=" & RS("�J�e�S���[�R�[�h")  & "'>" & RS("�J�e�S���[��") & "</a>" & vbNewLine
wHTML = wHTML & "        </div>" & vbNewLine

'----- ���i���r���[
wHTML = wHTML & "        <div class='note_box'>" & vbNewLine
wHTML = wHTML & "          <div class='pt8'>" & vbNewLine
wHTML = wHTML & "            " & CreateReviewImg(RS("���[�J�[�R�[�h"), RS("���i�R�[�h")) & vbNewLine
wHTML = wHTML & "          </div>" & vbNewLine
wHTML = wHTML & "        </div>" & vbNewLine

'----- �݌ɕ\��
wHTML = wHTML & "        <div class='stock_box'>" & vbNewLine
wHTML = wHTML & "          <div class='pt12'>" & vbNewLine
wHTML = wHTML & "            <img height='10' src='images/" & wInventoryImage & "' width='10' alt=''> " & wInventoryCd & vbNewLine
wHTML = wHTML & "          </div>" & vbNewLine
wHTML = wHTML & "        </div>" & vbNewLine

'---- �J�[�g
wHTML = wHTML & "        <div class='cart_box'>" & vbNewLine
wHTML = wHTML & "        <form name='f_item' method='post' action='OrderPreInsert.asp' onSubmit='return order_onClick(this);'>" & vbNewLine

'---- �F�K�i�Ȃ�
if RS("�F�K�iCNT") = 0 then
	if wProdTermFl = "Y" then
		wHTML = wHTML & "            <img src='images/icon_sold.gif' alt='����'><br>" & vbNewLine
	else
		wHTML = wHTML & "            <input type='hidden' name='qt' value='1'>" & vbNewLine
		wHTML = wHTML & "            <input type='hidden' name='maker_cd' value='" & RS("���[�J�[�R�[�h") & "'>" & vbNewLine
		wHTML = wHTML & "            <input type='hidden' name='product_cd' value='" & RS("���i�R�[�h") & "'>" & vbNewLine
		wHTML = wHTML & "            <input type='hidden' name='category_cd' value='" & RS("�J�e�S���[�R�[�h") & "'>" & vbNewLine
		wHTML = wHTML & "            <input type='image' src='images/btn_cart.png' style='vertical-align:middle' alt='�J�[�g��' class='opover'><br>" & vbNewLine
	end if

'----�F�K�i����
else
	if wProdTermFl = "Y" then
		wHTML = wHTML & "            <img src='images/icon_sold.gif' alt='����'><br>" & vbNewLine
	else
		wHTML = wHTML & "            <input type='hidden' name='qt' value='0'>" & vbNewLine
		wHTML = wHTML & "            <a href='ProductDetail.asp?Item=" & wItem & "'><img src='images/btn_detail.png' alt='�ڍׂ�����'></a><br>" & vbNewLine
	end if
end if

wHTML = wHTML & "          ���iID:" & RS("���iID") & vbNewLine
wHTML = wHTML & "        </form>" & vbNewLine
wHTML = wHTML & "        </div>" & vbNewLine
wHTML = wHTML & "      </div>" & vbNewLine
wHTML = wHTML & "    </div>" & vbNewLine
wHTML = wHTML & "<!-- s_box END -->" & vbNewLine

wUnder4HTML = wUnder4HTML & wHTML

End function

'========================================================================
'
'	Function	���i���r���[�摜�쐬
'
'========================================================================
'
Function CreateReviewImg(pMakerCd, pProductCd)

Dim vAvgRating
Dim v1Cnt
Dim v0Cnt
Dim vHalfCnt
Dim vTotalCnt
Dim vReview
Dim RSv
Dim i

'---- Select ���i���r���[ ���ρC���� �擾
' 2012/01/23 GV Mod Start
'wSQL = ""
'wSQL = wSQL & "SELECT SUM(a.�]��) AS �]�����v"
'wSQL = wSQL & "     , COUNT(a.ID) AS ���r���[��"
'wSQL = wSQL & "  FROM ���i���r���[ a WITH (NOLOCK) "				' 2012/01/19 GV Mod (NOLOCK �I�v�V�����t�^)
'wSQL = wSQL & " WHERE a.���[�J�[�R�[�h = '" & pMakerCd & "'"
'wSQL = wSQL & "   AND a.���i�R�[�h = '" & pProductCd & "'"
'
''@@@@response.write(wSQL)
'
'Set RSv = Server.CreateObject("ADODB.Recordset")
'RSv.Open wSQL, Connection, adOpenStatic
'
'if RSv("���r���[��") = 0 then
'	CreateReviewImg = ""
'	exit function
'end if
'
'vAvgRating = Round(RSv("�]�����v")/RSv("���r���[��"), 1)
'v1Cnt = Fix(vAvgRating)
'if (vAvgRating - v1Cnt) >= 0.5 then
'	vHalfCnt = 1
'else
'	vHalfCnt = 0
'end if
'v0Cnt = 5 - v1Cnt - vHalfCnt
'
'vTotalCnt = RSv("���r���[��")
'Rsv.Close

CreateReviewImg = ""

'---- Select ���i���r���[ ���ρC���� �擾
wSQL = ""
wSQL = wSQL & "SELECT "
wSQL = wSQL & "      a.���r���[�]������ "
wSQL = wSQL & "    , a.���r���[���� "
wSQL = wSQL & "FROM "
wSQL = wSQL & "    ���i���r���[�W�v a WITH (NOLOCK) "
wSQL = wSQL & "WHERE "
wSQL = wSQL & "        a.���[�J�[�R�[�h = '" & pMakerCd & "' "
wSQL = wSQL & "    AND a.���i�R�[�h     = '" & Replace(pProductCd, "'", "''") & "' "

'@@@@response.write(wSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic

If RSv.EOF Then
	RSv.Close
	Set RSv = Nothing
	Exit Function
End If

vAvgRating = Round(RSv("���r���[�]������"), 1)
vTotalCnt = RSv("���r���[����")

Rsv.Close
Set RSv = Nothing

If vTotalCnt = 0 Then
	Exit Function
End If

v1Cnt = Fix(vAvgRating)
If (vAvgRating - v1Cnt) >= 0.5 Then
	vHalfCnt = 1
Else
	vHalfCnt = 0
End If
v0Cnt = 5 - v1Cnt - vHalfCnt
' 2012/01/23 GV Mod End

'--- �����]���ҏW
For i = 1 to v1Cnt
	vReview = vReview & "<img src='images/review_icon10.png' alt=''>"
Next
If vHalfcnt = 1 Then
	vReview = vReview & "<img src='images/review_icon05.png' alt=''>"
End If
For i=1 to v0Cnt
	vReview = vReview & "<img src='images/review_icon00.png' alt=''>"
Next

CreateReviewImg = vReview

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
<title>�����L���O�b�T�E���h�n�E�X</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/Ranking.css?20140401a" type="text/css">
<script type="text/javascript">
//
// ====== 	Function:	order_onClick
//
function order_onClick(pForm){
	if (pForm.qt.value == ""){
		pForm.qt.value = 0;
	}else{
		if (numericCheck(pForm.qt.value) == false){
			pForm.qt.value = 0;
		}
	}
	if (pForm.qt.value > 0){
		return true;
	}else{
		return false;
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
<!--
  <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
    <p class="home"><a href="../"><img src="../images/icon_home.gif" alt="HOME"></a></p>
    <ul id="path">
      <li class="now">�����L�[���[�h</li>
    </ul>
  </div></div></div>
  <h1 class="title">�����L�[���[�h</h1>
-->

<!-- Mainpage START -->
<div id="ranking_key_main_flame">
  <div id="shukei">�i�W�v�F<%=Left(wYYYYMM,4)%>�N<%=right(wYYYYMM,2)%>���j</div>
<!-- Menu START -->
  <div id="ranking_key_top_menu">
    <div class="top_menu_parts">
      <a href="BestSellerList.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image15','','images/ranking/ts_btn_on.jpg',1)"><img src="images/ranking/ts_btn_off.jpg" alt="" name="Image15" width="114" height="80" />
      </a>
    </div>
    <div class="top_menu_parts">
      <a href="RankingSearchWord.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image163','','images/ranking/sk_btn_on.jpg',1)">
        <img src="images/ranking/sk_btn_off.jpg" alt="" name="Image163" width="114" height="80" />
      </a>
    </div>
    <!--
    <div class="top_menu_parts">
      <a href="RankingAccess.asp?RankType=<%=Server.URLEncode("���i�r���[")%>" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image12','','images/ranking/noc_btn_on.jpg',1)">
        <img src="images/ranking/<% if RankType="���i�r���[" then%>noc_btn_on.jpg<% else %>noc_btn_off.jpg<% end if%>" alt="" name="Image12" width="114" height="80" />
      </a>
    </div>
    -->
    <div class="top_menu_parts">
      <a href="RankingAccess.asp?RankType=<%=Server.URLEncode("�F�B�ɂ�����")%>" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image13','','images/ranking/rtaf_btn_on.jpg',1)"><img src="images/ranking/<% if RankType="�F�B�ɂ�����" then%>rtaf_btn_on.jpg<% else %>rtaf_btn_off.jpg<% end if%>" alt="" name="Image13" width="114" height="80" />
      </a>
    </div>
    <div class="top_menu_parts">
      <a href="RankingAccess.asp?RankType=<%=Server.URLEncode("�~�������̃��X�g")%>" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image14','','images/ranking/wl_btn_on.jpg',1)"><img src="images/ranking/<% if RankType="�~�������̃��X�g" then%>wl_btn_on.jpg<% else %>wl_btn_off.jpg<% end if%>" alt="" name="Image14" width="114" height="80" />
      </a>
    </div>
    <div class="top_menu_parts">
      <a href="RankingReview.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image16','','images/ranking/nor_btn_on.jpg',1)"><img src="images/ranking/nor_btn_off.jpg" alt="" name="Image16" width="114" height="80" />
      </a>
    </div>
    <div class="top_menu_parts">
      <a href="RankingReviewPoint.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image17','','images/ranking/rr_btn_on.jpg',1)"><img src="images/ranking/rr_btn_off.jpg" alt="" name="Image17" width="113" height="80" /></a>
    </div>
  </div>

<!-- Menu END -->
<!--  container START  -->
  <div id="container">

<%=wTop3HTML%>
<%=wUnder4HTML%>

  </div>
<!-- container END -->
</div>

  </div>
  <div id="globalSide">
    <!--#include file="../Navi/NaviSide.inc"-->
  </div>
</div>
<!--#include file="../Navi/Navibottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>