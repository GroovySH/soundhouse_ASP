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
'	���i�ڍ׃y�[�W
'�X�V����
'2009/09/08 an �V�K�쐬
'2011/04/14 hn SessionID�֘A�ύX
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2012/01/20 an SELECT����LAC�N�G���[�Ă�K�p
'2012/02/20 na �u���i�A�N�Z�X���v���X�|���X�΍�̂��ߒ�~
'2014/03/19 GV ����ő��łɔ���2�d�\���Ή�

'========================================================================

On Error Resume Next

Dim wUserID

Dim maker_cd
Dim product_cd

Dim item
Dim item_list()
Dim item_cnt

Dim wMakerName
Dim wProductName
Dim wMakerCode
Dim wCategoryCode
Dim wKoukeiMakerCd
Dim wKoukeiProductCd
Dim wTokucho

Dim wLogoHTML
Dim wProductHTML
Dim wTokuchoHTML
Dim wPictureHTML
Dim wSpecHTML
Dim wOthersHTML
Dim wCartHTML

Dim Connection
Dim RS

Dim wMinimumPrice
Dim wMidCategoryCd
Dim wSalesTaxRate
Dim wProdTermFl
Dim wPrice
Dim wPriceNoTax			'2014/03/19 GV add
Dim wNoData

Dim wItemChar1
Dim wItemChar2
Dim wItemNum1
Dim wItemNum2
Dim wItemDate1
Dim wItemDate2

Dim wSQL
Dim wHTML
Dim wMSG
Dim wErrDesc   '2011/08/01 an add

'========================================================================

Response.buffer = true

wUserID = Session("UserID")

'---- Get input data
item = ReplaceInput(Trim(Request("item")))

'���[�J�[�R�[�h�A���i�R�[�h�ɕ���
if item <> "" then 
	item_cnt = cf_unstring(item, item_list, "^")
	maker_cd = item_list(0)
	product_cd = item_list(1)
end if

'---- Execute main
call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "PremiumGuitarsDetail.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

if Err.Description <> "" then
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

call close_db()

'---- �Y�����i�Ȃ��̂Ƃ�
if wNoData = "Y" then
	Response.Redirect "SearchNotFound.asp"
end if

'---- ��p�@�킪����ꍇ�͂��̏��i��\��
if wKoukeiMakerCd <> "" then
	Response.Redirect "SearchList.asp?i_type=successor&s_maker_cd=" & wKoukeiMakerCd & "&s_product_cd=" & Server.URLEncode(wKoukeiProductCd)
end if


'========================================================================
'
'	Function	Connect database
'
'========================================================================
'
Function connect_db()
Dim i

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

Dim vPointer

'---- �ΏۃJ�e�S���[�R�[�h�A�Œ�P����o��
call getCntlMst("���i","PuremiumGuitar","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)
wMinimumPrice = Clng(wItemNum1)
wMidCategoryCd = wItemChar1

'---- ����ŗ���o��
call getCntlMst("����","����ŗ�","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)
wSalesTaxRate = Clng(wItemNum1)

'---- ���i�����o��
call GetProduct()
if wMSG <> "" OR wKoukeiMakerCd <> "" OR wNoData = "Y" then
	exit function
end if

'---- ���[�J�[���S�A���i���A���i�摜��
call CreateLogoHTML()

'---- ���[�J�[/���i��
Call CreateProductHTML()

'---- ����
call CreateTokuchoHTML()

'---- ���i�摜��
call CreatePictureHTML()

'---- �X�y�b�N
call CreateSpecificationHTML()

'---- ���̑��̏��i������
call CreateOthersHTML()

'----- �J�[�g���HTML�쐬�i���ʊ֐��j
wCartHTML = fCreateCartHtml()

'----- ���i�A�N�Z�X�����o�^ '2012/02/20 na ���X�|���X�΍�̂��ߒ�~
'call SetAccessCount()

RS.Close

End function

'========================================================================
'
'	Function	���i�����o��
'
'========================================================================
'
Function GetProduct()

Dim vInventoryCd
Dim vInventoryImage
Dim vSetCount
Dim vRowspan
Dim vWidth
Dim vHeight

Dim vProdPic(4)

'---- ���i�����o��
wSQL = ""
wSQL = wSQL & "SELECT DISTINCT a.���i�R�[�h"
wSQL = wSQL & "     , a.���i��"
wSQL = wSQL & "     , a.�̔��P��"
wSQL = wSQL & "     , a.������P��"
wSQL = wSQL & "     , a.�����萔��"
wSQL = wSQL & "     , a.������󒍍ϐ���"
wSQL = wSQL & "     , a.���i���l"
wSQL = wSQL & "     , a.���i�T��Web"
wSQL = wSQL & "     , a.���i�摜�t�@�C����_��"
wSQL = wSQL & "     , a.���i�摜�t�@�C����_��2"
wSQL = wSQL & "     , a.���i�摜�t�@�C����_��3"
wSQL = wSQL & "     , a.���i�摜�t�@�C����_��4"
wSQL = wSQL & "     , a.ASK���i�t���O"
wSQL = wSQL & "     , a.�I����"
wSQL = wSQL & "     , a.�戵���~��"
wSQL = wSQL & "     , a.�p�ԓ�"
wSQL = wSQL & "     , a.Web���i�t���O"
wSQL = wSQL & "     , a.�J�e�S���[�R�[�h"
wSQL = wSQL & "     , a.���[�J�[�R�[�h"
wSQL = wSQL & "     , a.�󏭐���"
wSQL = wSQL & "     , a.�Z�b�g���i�t���O"
wSQL = wSQL & "     , a.���[�J�[�������敪"
wSQL = wSQL & "     , a.���A���i�t���O"
wSQL = wSQL & "     , a.�����ߏ��i�R�����g"
wSQL = wSQL & "     , a.�֘A�L���^�C�g��1"
wSQL = wSQL & "     , a.�֘A�L��URL1"
wSQL = wSQL & "     , a.�֘A�L���^�C�g��2"
wSQL = wSQL & "     , a.�֘A�L��URL2"
wSQL = wSQL & "     , a.�֘A�L���^�C�g��3"
wSQL = wSQL & "     , a.�֘A�L��URL3"
wSQL = wSQL & "     , a.�֘A�L���^�C�g��4"
wSQL = wSQL & "     , a.�֘A�L��URL4"
wSQL = wSQL & "     , a.Web�[����\���t���O"
wSQL = wSQL & "     , a.��p�@�탁�[�J�[�R�[�h"
wSQL = wSQL & "     , a.��p�@�폤�i�R�[�h"
wSQL = wSQL & "     , a.���ח\�薢��t���O"
wSQL = wSQL & "     , a.���i�X�y�b�N�g�p�s�t���O"
wSQL = wSQL & "     , a.B�i�P��"
wSQL = wSQL & "     , a.������"
wSQL = wSQL & "     , a.B�i�t���O"
wSQL = wSQL & "     , a.���i���l�C���T�[�gURL1"
wSQL = wSQL & "     , a.���i���l�C���T�[�gURL2"
wSQL = wSQL & "     , a.���i���l�C���T�[�g�T�C�YW1"
wSQL = wSQL & "     , a.���i���l�C���T�[�g�T�C�YH1"
wSQL = wSQL & "     , a.���i���l�C���T�[�g�T�C�YW2"
wSQL = wSQL & "     , a.���i���l�C���T�[�g�T�C�YH2"
wSQL = wSQL & "     , b.���[�J�[��"
wSQL = wSQL & "     , b.���[�J�[���J�i"
wSQL = wSQL & "     , b.���[�J�[���S�t�@�C����"
wSQL = wSQL & "     , b.���[�J�[�z�[���y�[�WURL"
wSQL = wSQL & "     , b.�ڍ׏��^�C�g��1"
wSQL = wSQL & "     , b.�ڍ׏��URL1"
wSQL = wSQL & "     , b.�ڍ׏��Web�\���t���O1"
wSQL = wSQL & "     , b.�ڍ׏��^�C�g��2"
wSQL = wSQL & "     , b.�ڍ׏��URL2"
wSQL = wSQL & "     , b.�ڍ׏��Web�\���t���O2"
wSQL = wSQL & "     , b.�ڍ׏��^�C�g��3"
wSQL = wSQL & "     , b.�ڍ׏��URL3"
wSQL = wSQL & "     , b.�ڍ׏��Web�\���t���O3"
wSQL = wSQL & "     , b.�ڍ׏��^�C�g��4"
wSQL = wSQL & "     , b.�ڍ׏��URL4"
wSQL = wSQL & "     , b.�ڍ׏��Web�\���t���O4"
wSQL = wSQL & "     , c.�J�e�S���[��"
wSQL = wSQL & "     , d.���J�e�S���[�R�[�h"
wSQL = wSQL & "     , d.���J�e�S���[�����{��"
wSQL = wSQL & "     , e.��J�e�S���[�R�[�h"
wSQL = wSQL & "     , e.��J�e�S���[��"
wSQL = wSQL & "     , f.�F"
wSQL = wSQL & "     , f.�K�i"
wSQL = wSQL & "     , f.�����\����"
wSQL = wSQL & "     , f.�����\���ח\���"
wSQL = wSQL & "     , f.B�i�����\����"
wSQL = wSQL & "     , f.�F�K�i���i�摜�t�@�C����1"
wSQL = wSQL & "     , f.�F�K�i���i�摜�t�@�C����2"
wSQL = wSQL & "     , f.�F�K�i���i�摜�t�@�C����3"
wSQL = wSQL & "     , f.�F�K�i���i�摜�t�@�C����4"

'wSQL = wSQL & "  FROM Web���i a WITH (NOLOCK)"     '2012/01/20 an mod s
'wSQL = wSQL & "     , ���[�J�[ b WITH (NOLOCK)"
'wSQL = wSQL & "     , �J�e�S���[ c WITH (NOLOCK)"
'wSQL = wSQL & "     , ���J�e�S���[ d WITH (NOLOCK) "
'wSQL = wSQL & "     , ��J�e�S���[ e WITH (NOLOCK) "
'wSQL = wSQL & "     , Web�F�K�i�ʍ݌� f WITH (NOLOCK)"
'wSQL = wSQL & " WHERE b.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'wSQL = wSQL & "   AND c.�J�e�S���[�R�[�h = a.�J�e�S���[�R�[�h"
'wSQL = wSQL & "   AND d.���J�e�S���[�R�[�h = c.���J�e�S���[�R�[�h"
'wSQL = wSQL & "   AND e.��J�e�S���[�R�[�h = d.��J�e�S���[�R�[�h"
'wSQL = wSQL & "   AND a.Web���i�t���O = 'Y'"
'wSQL = wSQL & "   AND a.���[�J�[�R�[�h = '" & maker_cd & "'"
'wSQL = wSQL & "   AND a.���i�R�[�h = '" & product_cd & "'"
'wSQL = wSQL & "   AND f.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'wSQL = wSQL & "   AND f.���i�R�[�h = a.���i�R�[�h"
'wSQL = wSQL & "   AND f.�F = ''"
'wSQL = wSQL & "   AND f.�K�i = ''"
'wSQL = wSQL & "   AND f.�I���� IS NULL"

wSQL = wSQL & "  FROM Web���i                   a WITH (NOLOCK)"
wSQL = wSQL & "      INNER JOIN ���[�J�[        b WITH (NOLOCK)"
wSQL = wSQL & "        ON    b.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
wSQL = wSQL & "      INNER JOIN �J�e�S���[      c WITH (NOLOCK)"
wSQL = wSQL & "        ON     c.�J�e�S���[�R�[�h = a.�J�e�S���[�R�[�h"
wSQL = wSQL & "      INNER JOIN ���J�e�S���[    d WITH (NOLOCK) "
wSQL = wSQL & "        ON     d.���J�e�S���[�R�[�h = c.���J�e�S���[�R�[�h"
wSQL = wSQL & "      INNER JOIN ��J�e�S���[    e WITH (NOLOCK) "
wSQL = wSQL & "        ON     e.��J�e�S���[�R�[�h = d.��J�e�S���[�R�[�h"
wSQL = wSQL & "      INNER JOIN Web�F�K�i�ʍ݌� f WITH (NOLOCK)"
wSQL = wSQL & "        ON     f.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
wSQL = wSQL & "          AND  f.���i�R�[�h = a.���i�R�[�h"
wSQL = wSQL & "      LEFT JOIN ( SELECT 'Y' AS 'ShohinWebY' ) t1 "
wSQL = wSQL & "        ON     a.Web���i�t���O      = t1.ShohinWebY "
wSQL = wSQL & "      LEFT JOIN ( SELECT ''  AS 'Iro' )        t2 "
wSQL = wSQL & "        ON     f.�F               = t2.Iro "
wSQL = wSQL & "      LEFT JOIN ( SELECT ''  AS 'Kikaku' )     t3 "
wSQL = wSQL & "        ON     f.�K�i             = t3.Kikaku "
wSQL = wSQL & " WHERE "
wSQL = wSQL & "        t1.ShohinWebY   IS NOT NULL "
wSQL = wSQL & "    AND t2.Iro          IS NOT NULL "
wSQL = wSQL & "    AND t3.Kikaku       IS NOT NULL "
wSQL = wSQL & "    AND a.���[�J�[�R�[�h = '" & maker_cd & "'"
wSQL = wSQL & "    AND a.���i�R�[�h = '" & product_cd & "'"
wSQL = wSQL & "    AND f.�I���� IS NULL"    '2012/01/20 an mod e

'@@@@@@@@@@@response.write(wSQL)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

if RS.EOF = true then
	wPictureHTML = "<p class='error'>�Y�����i�͂���܂���B</p>"
	wMSG = "no data"
	wNoData = "Y"
	exit function
end if

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

'---- �戵���~�܂��́A�p�ԍ݌ɖ����Ō�p�@��`�F�b�N ����Ό�p�@���\��
if wProdTermFl = "Y" AND RS("��p�@�탁�[�J�[�R�[�h") <> "" then
	wKoukeiMakerCd = RS("��p�@�탁�[�J�[�R�[�h")
	wKoukeiProductCd = RS("��p�@�폤�i�R�[�h")
	exit function
end if

'----- ���[�J�[���A���i��� �^�C�g��
wMakerName = RS("���[�J�[��")
wProductName = RS("���i��")
wMakerCode = RS("���[�J�[�R�[�h")

End Function

'========================================================================
'
'	Function	���[�J�[���S�A���i���A���i�摜�� HTML�쐬
'
'========================================================================
'
Function CreateLogoHTML()

wLogoHTML = ""
wLogoHTML = wLogoHTML & "<div id='pgProductNameBox'>" & vbNewLine
If RS("���[�J�[���S�t�@�C����") <> "" Then
	wLogoHTML = wLogoHTML & "  <img src='maker_img/" & RS("���[�J�[���S�t�@�C����") & "'>" & vbNewLine
End If
wLogoHTML = wLogoHTML & "  <h1><span>" & wMakerName & "</span>" & wProductName  & "</h1>" & vbNewLine
wLogoHTML = wLogoHTML & "  <a href='javascript:void(0);' onclick='history.back();' class='tipBtn'>BACK</a>" & vbNewLine
wLogoHTML = wLogoHTML & "</div>" & vbNewLine
wLogoHTML = wLogoHTML & "<div id='pgLargeImage'><img name='LargeImage' src='prod_img/"

if Trim(RS("�F�K�i���i�摜�t�@�C����1")) <> "" then
	wLogoHTML = wLogoHTML & RS("�F�K�i���i�摜�t�@�C����1") & "' alt=''></div>" & vbNewLine
else
	if RS("���i�摜�t�@�C����_��") <> "" then
		wLogoHTML = wLogoHTML & RS("���i�摜�t�@�C����_��") & "' alt=''></div>" & vbNewLine
	else
		wLogoHTML = wLogoHTML & "n/nopict.jpg' alt=''></div>" & vbNewLine
	end if
end if

End Function

'========================================================================
'
'	Function	���i��� HTML�쐬
'
'========================================================================
'
Function CreateProductHTML()

Dim vInventoryCd
Dim vInventoryImage

'----- �݌ɏ�
vInventoryCd = GetInventoryStatus(RS("���[�J�[�R�[�h"),RS("���i�R�[�h"),RS("�F"),RS("�K�i"),RS("�����\����"),RS("�󏭐���"),RS("�Z�b�g���i�t���O"),RS("���[�J�[�������敪"),RS("�����\���ח\���"),wProdTermFl)

'---- �݌ɏ󋵁A�F���ŏI�Z�b�g
call GetInventoryStatus2(RS("�����\����"), RS("Web�[����\���t���O"), RS("���ח\�薢��t���O"), RS("�p�ԓ�"), RS("B�i�t���O"), RS("B�i�����\����"), RS("�����萔��"), RS("������󒍍ϐ���"), wProdTermFl, vInventoryCd, vInventoryImage)

wProductHTML = ""
wProductHTML = wProductHTML & "        <h2>���[�J�[/���i��</h2>" & vbNewLine
wProductHTML = wProductHTML & "        <dl class='pgDetailBox'>" & vbNewLine
wProductHTML = wProductHTML & "          <dt>���[�J�[</dt>" & vbNewLine
wProductHTML = wProductHTML & "          <dd>" &  wMakerName & " ( " & RS("���[�J�[���J�i") & " )</dd>" & vbNewLine
wProductHTML = wProductHTML & "          <dt>���i��</dt>" & vbNewLine
wProductHTML = wProductHTML & "          <dd>" & wProductName & "</dd>" & vbNewLine
wProductHTML = wProductHTML & "          <dt>�J�e�S���[</dt>" & vbNewLine
wProductHTML = wProductHTML & "          <dd>" & RS("�J�e�S���[��") & "</dd>" & vbNewLine
wProductHTML = wProductHTML & "          <dt>�̔����i</dt>" & vbNewLine
wProductHTML = wProductHTML & "          <dd>"

if RS("�����萔��") > RS("������󒍍ϐ���") AND RS("�����萔��") > 0 then
	wPrice = calcPrice(RS("������P��"), wSalesTaxRate)
	wPriceNoTax = RS("������P��")						'2014/03/19 GV add
else
	wPrice = calcPrice(RS("�̔��P��"), wSalesTaxRate)
	wPriceNoTax = RS("�̔��P��")						'2014/03/19 GV add
end if

'2014/03/19 GV mod start ---->
'wProductHTML = wProductHTML & FormatNumber(wPrice,0) & "�~(�ō�)</dd>" & vbNewLine
wProductHTML = wProductHTML & FormatNumber(wPriceNoTax,0) & "�~(�Ŕ�)</dd>" & vbNewLine

wProductHTML = wProductHTML & "          <dt>&nbsp;</dt>" & vbNewLine
wProductHTML = wProductHTML & "<dd>" & FormatNumber(wPrice,0) & "�~(�ō�)</dd>" & vbNewLine
'2014/03/19 GV mod end   <----
wProductHTML = wProductHTML & "          <dt>�݌ɏ�</dt>" & vbNewLine

wProductHTML = wProductHTML & "          <dd><img src='images/" & vInventoryImage & "' width='10' height='10' style='vertical-align:baseline;'> " & vInventoryCd & "</dd>" & vbNewLine
wProductHTML = wProductHTML & "          <dt></dt>" & vbNewLine
wProductHTML = wProductHTML & "          " & vbNewLine
wProductHTML = wProductHTML & "          <dd id='cartBox'><form name='f_data' method='post' action='OrderPreInsert.asp' onSubmit='return order_onClick(this);'><input type='text' name='qt' size='2' maxsize='3' value='1'>" & vbNewLine

if (IsNull(RS("�戵���~��")) = false) OR (IsNull(RS("������")) = false) OR (RS("B�i�t���O") = "Y" AND RS("B�i�����\����") <= 0) OR (IsNull(RS("�p�ԓ�")) = false AND RS("�����\����") <= 0) then
	wProductHTML = wProductHTML & "<img src='images/Kanbai2.jpg' alt='����'>" & vbNewLine
else
	wProductHTML = wProductHTML & "            <input type='image' src='images/PremiumGuitars/grey_cart.jpg' width='80' height='23' alt=''>" & vbNewLine
	wProductHTML = wProductHTML & "            <input type='hidden' name='Item' value='" & RS("���[�J�[�R�[�h") & "^" & RS("���i�R�[�h") & "'>" & vbNewLine
end if

wProductHTML = wProductHTML & "          </form>" & vbNewLine
wProductHTML = wProductHTML & "          </dd>" & vbNewLine
wProductHTML = wProductHTML & "        </dl>" & vbNewLine

End Function

'========================================================================
'
'	Function	���� HTML�쐬
'
'========================================================================
'
Function CreateTokuchoHTML()

wHTML = ""

'---- ����, ���A���i�\��
If RS("�����ߏ��i�R�����g") <> "" Or RS("���A���i�t���O") = "Y" Then
	wTokuchoHTML = wTokuchoHTML & "        <h2>����</h2>" & vbNewLine
	wTokuchoHTML = wTokuchoHTML & "        <div class='pgDetailBox'>" & vbNewLine
	if RS("�����ߏ��i�R�����g") <> "" then
		wTokuchoHTML = wTokuchoHTML & RS("�����ߏ��i�R�����g") & "<br>" & vbNewLine

		'---- meta description�p�f�[�^�擾
		wTokucho = fDeleteHTMLTag(RS("�����ߏ��i�R�����g")) 'HTML�^�O�폜
		wTokucho = replace(replace(replace(wTokucho, vbCr, ""), vbLf, ""), vbTab, "") '���s�ATab�̍폜

		if Len(wTokucho) > 97 then  '�����ꍇ��100�����ɏȗ�
			wTokucho = Left(wTokucho,97) & "..."
		end if

	end if
	if RS("���A���i�t���O") = "Y" then
		wTokuchoHTML = wTokuchoHTML & "<a href='../information/direct_import.asp'>[���A���i]</a>" & vbNewLine
	end if
	wTokuchoHTML = wTokuchoHTML & "        </div>" & vbNewLine
End If

End Function

'========================================================================
'
'	Function	���i���摜 HTML�쐬
'
'========================================================================
'
Function CreatePictureHTML()

Dim vProdPic(4)
Dim i

'�F�K�i���i�摜�t�@�C����������ꍇ�͂������D��
if Trim(RS("�F�K�i���i�摜�t�@�C����1")) <> "" then
	vProdPic(1) = RS("�F�K�i���i�摜�t�@�C����1")
	vProdPic(2) = RS("�F�K�i���i�摜�t�@�C����2")
	vProdPic(3) = RS("�F�K�i���i�摜�t�@�C����3")
	vProdPic(4) = RS("�F�K�i���i�摜�t�@�C����4")
else
	vProdPic(1) = RS("���i�摜�t�@�C����_��")
	vProdPic(2) = RS("���i�摜�t�@�C����_��2")
	vProdPic(3) = RS("���i�摜�t�@�C����_��3")
	vProdPic(4) = RS("���i�摜�t�@�C����_��4")
end if

wPictureHTML = ""
wPictureHTML = wPictureHTML & "      <ul id='pgSmallImage'>" & vbNewLine

for i=1 to 4
	wPictureHTML = wPictureHTML & "        <li><img src='prod_img/"
	if vProdPic(i) <> "" then
		wPictureHTML = wPictureHTML & vProdPic(i) & "' alt='" & wMakerName & " / " & wProductname & " �摜" & i & "' onMouseOver='SmallImage_onMouseOver(""prod_img/" & vProdPic(i) & """);'>"
	else
		'���摜���Ȃ��ꍇ�͑�։摜��\��
		wPictureHTML = wPictureHTML & "p/pg_photo_s.jpg' alt='" & wMakerName & " / " & wProductname & " �摜" & i & "'>"
	end if
	wPictureHTML = wPictureHTML & "        </li>" & vbNewLine
Next

wPictureHTML = wPictureHTML & "      </ul>" & vbNewLine

End Function

'========================================================================
'
'	Function	�X�y�b�N HTML�쐬
'
'========================================================================
'
Function CreateSpecificationHTML()

Dim vWidth
Dim vHeight

wSpecHTML = ""

'---- �X�y�b�N
wSpecHTML = wSpecHTML & "        <h2>�X�y�b�N</h2>" & vbNewLine
wSpecHTML = wSpecHTML & "        <div class='pgDetailBox'>" & vbNewLine

if RS("���i���l�C���T�[�gURL1") <> "" then
	if RS("���i���l�C���T�[�g�T�C�YW1") <> 0 then
		vWidth = RS("���i���l�C���T�[�g�T�C�YW1")
		if vWidth > 600 then
			vWidth = 600
		end if
	else
		vWidth = 600
	end if
	if RS("���i���l�C���T�[�g�T�C�YH1") <> 0 then
		vHeight = RS("���i���l�C���T�[�g�T�C�YH1")
		if vHeight > 290 then
			vHeight = 290
		end if
	else
		vHeight = 290
	end if
'	wSpecHTML = wSpecHTML & "<iframe marginwidth='0' marginheight='0' src='" & RS("���i���l�C���T�[�gURL1") & "' frameborder='0' scrolling='no' style='PADDING-RIGHT: 0px; PADDING-LEFT: 0px; PADDING-BOTTOM: 0px; MARGIN: 0px; WIDTH: " & vWidth & "px; PADDING-TOP: 0px; HEIGHT: " & vHeight & "px'> </iframe>" & vbNewLine
end if

wSpecHTML = wSpecHTML & CreateSpecHTML(RS("�J�e�S���[�R�[�h"),RS("���[�J�[�R�[�h"),RS("���i�R�[�h"),RS("���i���l"),RS("���i�X�y�b�N�g�p�s�t���O")) & vbNewLine

if RS("���i���l�C���T�[�gURL2") <> "" then
	if RS("���i���l�C���T�[�g�T�C�YW2") <> 0 then
		vWidth = RS("���i���l�C���T�[�g�T�C�YW2")
	else
		vWidth = 600
	end if
	if RS("���i���l�C���T�[�g�T�C�YH2") <> 0 then
		vHeight = RS("���i���l�C���T�[�g�T�C�YH2")
	else
		vHeight = 300
	end if
'	wSpecHTML = wSpecHTML & "<iframe marginwidth='0' marginheight='0' src='" & RS("���i���l�C���T�[�gURL2") & "' frameborder='0' scrolling='no' style='PADDING-RIGHT: 0px; PADDING-LEFT: 0px; PADDING-BOTTOM: 0px; MARGIN: 0px; WIDTH: " & vWidth & "px; PADDING-TOP: 0px; HEIGHT: " & vHeight & "px'> </iframe>" & vbNewLine
end if

wSpecHTML = wSpecHTML & "        </div>" & vbNewLine

End Function

'========================================================================
'
'	Function	���̑��̏��i������ HTML�쐬
'
'========================================================================
'
Function CreateOthersHTML

Dim RSv

'----���ꃁ�[�J�[�̏��15�i�i����o�^�����j ���o��
wSQL = ""
wSQL = wSQL & "SELECT DISTINCT TOP 15"
wSQL = wSQL & "   a.���i��"
wSQL = wSQL & " , a.���i�R�[�h"
wSQL = wSQL & " , a.����o�^��"

'wSQL = wSQL & " FROM  Web���i  a WITH (NOLOCK)"    '2012/01/20 an mod s
'wSQL = wSQL & "     , �J�e�S���[���J�e�S���[ b WITH (NOLOCK)"
'wSQL = wSQL & " WHERE (SELECT CASE"
'wSQL = wSQL & "                   WHEN x.�����萔�� > x.������󒍍ϐ��� THEN (x.������P�� * (100 + " & wSalesTaxRate & " )/100)"
'wSQL = wSQL & "                   ELSE (x.�̔��P�� * (100 + " & wSalesTaxRate & " )/100)"
'wSQL = wSQL & "               END"
'wSQL = wSQL & "        FROM web���i x "
'wSQL = wSQL & "        WHERE x.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'wSQL = wSQL & "          AND x.���i�R�[�h = a.���i�R�[�h"
'wSQL = wSQL & "        ) > " & wMinimumPrice
'wSQL = wSQL & "    AND a.�J�e�S���[�R�[�h = b.�J�e�S���[�R�[�h"
'wSQL = wSQL & "    AND b.���J�e�S���[�R�[�h IN (" & wMidCategoryCd & ")"
'wSQL = wSQL & "    AND a.Web���i�t���O = 'Y'"
'wSQL = wSQL & "    AND a.���[�J�[�R�[�h =" & RS("���[�J�[�R�[�h")

wSQL = wSQL & " FROM  Web���i                          a WITH (NOLOCK)"
wSQL = wSQL & "      INNER JOIN �J�e�S���[���J�e�S���[ b WITH (NOLOCK)"
wSQL = wSQL & "        ON     b.�J�e�S���[�R�[�h = a.�J�e�S���[�R�[�h"
wSQL = wSQL & "      LEFT JOIN ( SELECT 'Y' AS 'ShohinWebY' ) t1 "
wSQL = wSQL & "        ON     a.Web���i�t���O      = t1.ShohinWebY "
wSQL = wSQL & " WHERE "
wSQL = wSQL & "        t1.ShohinWebY   IS NOT NULL "
wSQL = wSQL & "    AND (SELECT CASE"
wSQL = wSQL & "                   WHEN x.�����萔�� > x.������󒍍ϐ��� THEN (x.������P�� * (100 + " & wSalesTaxRate & " )/100)"
wSQL = wSQL & "                   ELSE (x.�̔��P�� * (100 + " & wSalesTaxRate & " )/100)"
wSQL = wSQL & "               END"
wSQL = wSQL & "        FROM web���i x WITH (NOLOCK)"
wSQL = wSQL & "        WHERE x.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
wSQL = wSQL & "          AND x.���i�R�[�h = a.���i�R�[�h"
wSQL = wSQL & "        ) > " & wMinimumPrice
wSQL = wSQL & "    AND b.���J�e�S���[�R�[�h IN (" & wMidCategoryCd & ")"
wSQL = wSQL & "    AND a.���[�J�[�R�[�h =" & RS("���[�J�[�R�[�h")
wSQL = wSQL & " ORDER BY a.����o�^�� DESC"    '2012/01/20 an mod e

'@@@@@@@@@@response.write(wSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

	wOthersHTML = ""
	wOthersHTML = wOthersHTML & "<div id='pgOtherProduct'>" & vbNewLine
	wOthersHTML = wOthersHTML & "  <h2>���̑��̏��i������</h2>" & vbNewLine
	wOthersHTML = wOthersHTML & "  <ul>" & vbNewLine
	Do Until RSv.EOF = true
		'--- �ڍו\�����̏��i�̓��X�g�ɕ\�����Ȃ�
		if RSv("���i�R�[�h") <> product_cd then
			wOthersHTML = wOthersHTML & "    <li><a href='PremiumGuitarsDetail.asp?Item=" & RS("���[�J�[�R�[�h") & "^" & RSv("���i�R�[�h") & "' style='text-decoration:none; color:#cccccc;'>" & RSv("���i��") & "</a></li>" & vbNewLine
		end if
		RSv.MoveNext
	Loop
	wOthersHTML = wOthersHTML & "  </ul>" & vbNewLine
	wOthersHTML = wOthersHTML & "  <div class='more'><li><a href='PremiumGuitarsList.asp?MakerCd=" & RS("���[�J�[�R�[�h") & "'>" &  wMakerName & "<br>�v���~�A���M�^�[�ꗗ</a></li></div>" & vbNewLine
	wOthersHTML = wOthersHTML & "</div>" & vbNewLine
	
RSv.Close

End Function
'========================================================================
'
'	Function	�ŋ߃`�F�b�N�������i�ɒǉ�
'
'========================================================================
'
Function AddViewdProduct()

Dim RSv

wSQL = ""
wSQL = wSQL & "SELECT *"
wSQL = wSQL & "  FROM �ŋ߃`�F�b�N�������i"
wSQL = wSQL & " WHERE �ڋq�ԍ� = " & wUserID
wSQL = wSQL & "   AND ���[�J�[�R�[�h = '" & maker_cd & "'"
wSQL = wSQL & "   AND ���i�R�[�h = '" & product_cd & "'"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RSv.EOF = true then
	RSv.AddNew

	RSv("�ڋq�ԍ�") = wUserID
	RSv("���[�J�[�R�[�h") = maker_cd
	RSv("���i�R�[�h") = product_cd
end if

RSv("�`�F�b�N��") = Now()

RSv.Update
RSv.close

End function

'========================================================================
'
'	Function	���i�A�N�Z�X�J�E���g�o�^�i�y�[�W�r���[�j
'
'========================================================================
'
Function SetAccessCount()

Dim vYYYYMM
Dim RSv

'---- ����Z�b�V������1��ڂ��ǂ����`�F�b�N
	wSQL = ""
	wSQL = wSQL & "SELECT *"
	wSQL = wSQL & "  FROM �Z�b�V�����f�[�^"
	wSQL = wSQL & " WHERE SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod
	wSQL = wSQL & "   AND ���ږ� = '" & maker_cd & "^" & product_cd & "'"

	Set RSv = Server.CreateObject("ADODB.Recordset")
	RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

	if RSv.EOF = true then
		'---- �Z�b�V�����f�[�^�o�^
		RSv.AddNew

		RSv("SessionID") = gSessionID		'2011/04/14 hn mod
		RSv("���ږ�") = maker_cd & "^" & product_cd
		RSv("���e") = "�y�[�W�r���[�`�F�b�N�p"
		RSv("�ŏI�X�V��") = Now()

		RSv.Update
		RSv.close

		'---- �y�[�W�r���[�o�^
		vYYYYMM = Year(Now()) & Right("0" & Month(Now()),2)

		wSQL = ""
		wSQL = wSQL & "SELECT *"
		wSQL = wSQL & "  FROM ���i�A�N�Z�X����"
		wSQL = wSQL & " WHERE ���[�J�[�R�[�h = '" & maker_cd & "'"
		wSQL = wSQL & "   AND ���i�R�[�h = '" & product_cd & "'"
		wSQL = wSQL & "   AND �N�� = '" & vYYYYMM & "'"

		Set RSv = Server.CreateObject("ADODB.Recordset")
		RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

		if RSv.EOF = true then
			RSv.AddNew

			RSv("���[�J�[�R�[�h") = maker_cd
			RSv("���i�R�[�h") = product_cd
			RSv("�N��") = vYYYYMM
			RSv("�y�[�W�r���[����") = 1
		else
			RSv("�y�[�W�r���[����") = RSv("�y�[�W�r���[����") + 1
		end if

		RSv.Update
		RSv.close
	end if

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
<meta name="robots" content="noindex,nofollow">
<title>�v���~�A���M�^�[ <%=wMakerName%> / <%=wProductName%>�b�T�E���h�n�E�X</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link href='http://fonts.googleapis.com/css?family=Ovo' rel='stylesheet' type='text/css'>
<link rel="stylesheet" href="style/PremiumGuitars.css" type="text/css">
<% if wTokucho <> "" then%>
<meta name="description" content="<%=wTokucho%>">
<% end if %>
<meta name="keywords" content="<%=wLargeCategoryName%>,<%=wMidCategoryName%>,<%=wCategoryName%>,<%=wMakerName%>,<%=wProductName%>">
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
	if (pForm.qt.value <= 0){
		alert("���ʂ���͂��Ă���J�[�g�{�^���������Ă��������B");
		return false;
	}
	return true;
}
//
// ====== 	Function:	SmallImage_onMouseOver
//
function SmallImage_onMouseOver(pFile){
	document.images["LargeImage"].src = pFile;
}
</script>
<style type="text/css">
#globalContents ul.sns {
	overflow: hidden;
	padding: 5px;
}

#globalContents ul.sns li {
	float: right;
	width: 100px;
	height: 20px;
}
</style>
</head>
<body>
<!--#include file="../Navi/Navitop.inc"-->

<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="��������{���ł�"></a></span>

<!-- �R���e���cstart -->
<div id="globalContents">
    <div id='path_box'><div id='path_box_inner01'><div id='path_box_inner02'>
    <p class='home'><a href="<%=g_HTTP%>"><img src="<%=g_RelLink%>images/icon_home.gif" alt="HOME"></a></p>
    <ul id='path'>
      <li><a href="<%=g_HTTP%>material/">SPECIAL SELECTION�ꗗ</a></li>
      <li><a href="PremiumGuitars.asp">�v���~�A���M�^�[</a></li>
      <li><a href="PremiumGuitarsList.asp?MakerCd=<%=wMakerCode%>"><%=wMakerName%></a></li>
      <li class="now"><%=wProductName%></li>
    </ul>
  </div></div></div>
    <ul class="sns">
          <li><a href="https://twitter.com/share" class="twitter-share-button" data-lang="ja">�c�C�[�g</a><script>!function(d,s,id){var js,fjs=d.getElementsByTagName(s)[0];if(!d.getElementById(id)){js=d.createElement(s);js.id=id;js.src="//platform.twitter.com/widgets.js";fjs.parentNode.insertBefore(js,fjs);}}(document,"script","twitter-wjs");</script></li>
          <li><iframe src="//www.facebook.com/plugins/like.php?href=http%3A%2F%2Fwww.soundhouse.co.jp%2Fshop%2FPremiumGuitars.asp&amp;send=false&amp;layout=button_count&amp;width=100&amp;show_faces=false&amp;action=like&amp;colorscheme=light&amp;font&amp;height=21&amp;appId=191447484218062" scrolling="no" frameborder="0" style="border:none; overflow:hidden; width:100px; height:21px;" allowTransparency="true"></iframe></li>
        </ul>

  <div id="pgContainer">
<!-- �g�b�v�摜 START -->
<div id="pgHeader">
  <div class="topbox">
    <div class="left"></div>
    <div class="right"></div>
  </div>
</div>
<!-- �g�b�v�摜 END -->

<!-- ���[�J�[���A���i���A���i�摜�� -->
<%=wLogoHTML%>

<!-- ���i�摜�� START -->
<%=wPictureHTML%>
<!-- ���i�摜�� END -->

<div id="pgInfoBox">
<div class="left">
    
<!-- ���i��� START -->
<%=wProductHTML%>
<!-- ���i��� END -->

<!-- ���� START -->
<%=wTokuchoHTML%>
<!-- ���� END -->
<!-- �X�y�b�N START -->
<%=wSpecHTML%>
<!-- �X�y�b�N END -->
</div>
<!-- ���̑��̏��i START -->
<%=wOthersHTML%>
<!-- ���̑��̏��i END -->

</div>

  <p class="arrow"><a href="#site_title"><img src="images/PremiumGuitars/white_arrow_up.gif" alt=""></a></p>
</div>

  </div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>