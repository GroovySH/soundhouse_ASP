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
<!--#include file="../common/SalesCommon.inc"-->
<!--#include file="../common/SearchListCommon.inc"-->
<%
'========================================================================
'
'    ���ʌ���o�[�Q��
'
'�X�V����
'2009/09/28 an SearchList.asp���番�����ĐV�K�쐬
'2010/03/18 an ���ёւ���ʕύX�i�V���i�A�������ߕ]�����̒ǉ��A�s�v���ёւ��̍폜�j
'2010/05/13 an ���r���[�]�����ς̌v�Z�����u�V���b�v�R�����g IS NULL�v���폜
'              ���r���[�]�����όv�Z����CAST(decimal)��ǉ����A�����_�ȉ����l��
'2010/06/08 an ��NAVI�̍i�荞�ݏ����ύX�Ή�
'2010/06/29 an �����̑�J�e�S���[�ɏ������鏤�i���d���\�������s����C��
'2010/07/09 st ���r���[�摜�쐬�֐� CreateReviewImg ���폜
'2010/07/12 st ���ёւ������ύX�E�ǉ��i�������ߕ]���˕]�����ɕύX�A�]���������̒ǉ��j
'2010/07/16 an �݌ɗL�\�[�g�̕s��C���iB�i�A������̏�������Ɂj
'2011/02/23 GV(dy) #826 �������S�����\���̑Ή�
'2011/05/25 hn �F�K�i����ł��݌ɏ���\������悤�ɕύX
'2011/06/09 hn �p�Ԃō݌ɂȂ��{�����Ȃ��@�̎��Ɋ����Ƃ���悤�ɕύX
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2012/01/10 GV Web�Z�[�����i�e�[�u�����쓮�\�Ƃ���Ή�
'2012/07/11 ok ���j���[�A���V�f�U�C���ύX
'2012/10/22 ok ���ёւ��ɂ������ǉ�
'2012/11/09 ok �i�荞�ݗp�\�[�g����\�����ɕύX
'2014/03/19 GV ����ő��łɔ���2�d�\���Ή�
'
'========================================================================

On Error Resume Next


'Dim s_mid_category_cd       '2010/06/08 an del
'Dim s_category_cd           '2010/06/08 an del
Dim s_maker_cd
Dim s_product_cd
Dim sPriceFrom
Dim sPriceTo
Dim sSeriesCd
Dim i_page
Dim i_sort
Dim i_page_size
Dim i_ListType

Dim wSalesTaxRate
Dim wHikaku
Dim wTemp
Dim LargeCategoryCd          '2010/06/08 an add
Dim MidCategoryCd            '2010/06/08 an add
Dim CategoryCd               '2010/06/08 an add

Dim wListHTML
Dim wCountHTML
Dim wMakerHTML
Dim wNaviMakerHTML
Dim wNaviCategoryHTML
Dim wNaviLargeCategoryHTML   '2010/06/08 an add
Dim wNaviMidCategoryHTML     '2010/06/08 an add
Dim wNaviPricerangeHTML
Dim wMakerInfoHTML
Dim wNoDataHTML              '2010/06/08 an add

Dim wItemChar1
Dim wItemChar2
Dim wItemNum1
Dim wItemNum2
Dim wItemDate1
Dim wItemDate2

Dim Connection
Dim RS

Dim wSQL
Dim wSQL2              '2010/06/08 an add
Dim wSQLMaker
Dim wSQLCategory
Dim wSQLMidCategory    '2010/06/08 an add
Dim wSQLLargeCategory  '2010/06/08 an add
Dim wSQLPricerange

Dim wNoData
Dim wHTML
Dim wFootprintHTML   '2010/06/08 an add
Dim wErrDesc   '2011/08/01 an add
Dim wTitle   '2012/07/11 ok add


'========================================================================

Response.buffer = true

'---- Get input data
s_product_cd = ReplaceInput(Trim(Request("s_product_cd")))
s_maker_cd = ReplaceInput(Trim(Request("s_maker_cd")))
LargeCategoryCd = ReplaceInput(Trim(Request("s_large_category_cd")))   '2010/06/08 an add
MidCategoryCd = ReplaceInput(Trim(Request("s_mid_category_cd")))       '2010/06/08 an add
CategoryCd = ReplaceInput(Trim(Request("s_category_cd")))              '2010/06/08 an mod
sPriceFrom = ReplaceInput(Trim(Request("sPriceFrom")))
sPriceTo = ReplaceInput(Trim(Request("sPriceTo")))
sSeriesCd = ReplaceInput(Trim(Request("sSeriesCd")))
i_page = ReplaceInput(Trim(Request("i_page")))
i_sort = ReplaceInput(Trim(Request("i_sort")))
i_page_size = ReplaceInput(Trim(Request("i_page_size")))
i_ListType = ReplaceInput(Trim(Request("i_ListType")))

if ISNumeric(sPriceFrom) = false then
    sPriceFrom = 0
end if
if ISNumeric(sPriceTo) = false then
    sPriceTo = 9999999
end if

sPriceFrom = CCur(sPriceFrom)
sPriceTo = CCur(sPriceTo)

if sPriceTo < sPriceFrom then
    wTemp = sPriceFrom
    sPriceFrom = sPriceTo
    sPriceTo = wTemp
end if

'---- ��rCookie���o��
wHikaku = Session("compare")

'---- �\���^�C�v���o��    09/05/26
if i_ListType = "" then
    i_ListType = Session("ListType")
    if i_ListType = "" then
        i_ListType = "type1"
        Session("ListType") = i_ListType
    end if
else
    Session("ListType") = i_ListType
end if

'---- �y�[�W�T�C�Y�ݒ�
if i_page = "" then
    i_page = 1
else
    i_page = Clng(i_page)
end if

if i_page_size = "" then
    i_page_size = Session("PageSize")
end if

if i_page_size = "" then
    i_page_size = g_page_size
else
    i_page_size = Clng(i_page_size)
end if

if i_ListType = "type1" then
    if i_page_size = 12 then i_page_size = 10
    if i_page_size = 32 then i_page_size = 30
    if i_page_size = 52 then i_page_size = 50
else
    if i_page_size = 10 then i_page_size = 12
    if i_page_size = 30 then i_page_size = 32
    if i_page_size = 50 then i_page_size = 52
end if

Session("PageSize") = i_page_size

'---- Execute main
call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "BargainSale.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

call close_db()

if Err.Description <> "" then
    Response.Redirect g_HTTP & "shop/Error.asp"
end if

'========================================================================
'
'    Function    Connect database
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
'    Function    Main
'
'========================================================================
'
Function main()

Dim vPointer

'---- ����ŗ���o��
call getCntlMst("����","����ŗ�","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)            '����ŗ�
wSalesTaxRate = Clng(wItemNum1)

'---- �Y�����i���o��
call GetProducts()

'---- �p�������X�g�쐬
call fCreateFootprintHTML("BargainSale.asp","�Ռ������i")  '2010/06/08 an add

'-----
if RS.EOF = true then
	wNoData = "Y"
    wNoDataHTML = wNoDataHTML & "�Y�����鏤�i��������܂���B" & vbNewLine
else

	'----- ListHTML�쐬
	call fCreateSearchListHTML(RS, i_page_size, i_page, i_ListType, wSalesTaxRate, wListHTML)

	'----- �������ݒ�	'2012/10/22 ok Add
	wListHTML = Replace(wListHTML,"�I�����Ă�������</option>","�I�����Ă�������</option>" & vbNewLIne & "<option value='Nesage_DESC'>������</option>")

	'---- ���[�J�[���쐬
	if s_maker_cd <> "" then
	    call CreateMakerInfo()
	end if

	'---- NAVI Left Sale�pHTML�쐬     '2010/06/08 an mod s
	if LargeCategoryCd <> "" then
		'----- ��J�e�S���[�R�[�h�w�莞�A���J�e�S���[�ꗗ�쐬�@NAVI�p
		call fCreateSalesNAVIMidCategoryHTML(wSQLMidCategory, wNaviMidCategoryHTML)
	else
		if MidCategoryCd <> "" then
			'----- ���J�e�S���[�R�[�h�w�莞�A�J�e�S���[�ꗗ�쐬�@NAVI�p
			call fCreateSalesNAVICategoryHTML(wSQLCategory, wNaviCategoryHTML)
		else
			if CategoryCd <> "" then
				'----- �J�e�S���[�R�[�h�w�莞�A���[�J�[�ꗗ�쐬
				call fCreateSalesNAVIMakerHTML(wSQLMaker, wNaviMakerHTML)
			else
				'---- �w��Ȃ��̏ꍇ�A��J�e�S���[�ꗗ�쐬
				call fCreateSalesNAVILargeCategoryHTML(wSQLLargeCategory, wNaviLargeCategoryHTML)
			end if
		end if
	end if

	'----- ���i�ёI���쐬�@NAVI�p
	call fCreateSalesNAVIPriceRangeHTML(wSQLPriceRange, wSalesTaxRate, wNaviPriceRangeHTML)    '2010/06/08 an mod e


end if

RS.close

End Function


'========================================================================
'
'    Function    �Y�����i���o��
'
'========================================================================
'
Function GetProducts()

Dim v_order

'---- SQL�쐬
wSQL = ""
' 20120110 GV Mod Start
'wSQL = wSQL & "SELECT DISTINCT"    '2005/07/19
'wSQL = wSQL & "       a.���[�J�[�R�[�h"            '2005/07/19
'wSQL = wSQL & "     , a.���i�R�[�h"
'wSQL = wSQL & "     , a.���i��"
'wSQL = wSQL & "     , a.���i�T��Web"
'wSQL = wSQL & "     , a.�����敪"
'wSQL = wSQL & "     , a.���菤�i��"
'wSQL = wSQL & "     , a.�d�ʏ��i����"
'wSQL = wSQL & "     , a.���i�摜�t�@�C����_��"
'wSQL = wSQL & "     , a.���i���l"
'wSQL = wSQL & "     , a.�W���P��"
'wSQL = wSQL & "     , a.�̔��P��"
'wSQL = wSQL & "     , CASE"
'wSQL = wSQL & "         WHEN a.B�i�t���O = 'Y' THEN a.B�i�P��"   '2010/07/16 an mod
'wSQL = wSQL & "         WHEN a.�����萔�� > a.������󒍍ϐ��� THEN a.������P��"   '2010/07/16 an mod
'wSQL = wSQL & "         ELSE a.�̔��P��"
'wSQL = wSQL & "       END AS ���̔��P��"
'wSQL = wSQL & "     , a.������P��"
'wSQL = wSQL & "     , a.�����萔��"
'wSQL = wSQL & "     , a.������󒍍ϐ���"
'wSQL = wSQL & "     , a.�I�[�v�����i�t���O"
'wSQL = wSQL & "     , a.���[�J�[�������敪"
'wSQL = wSQL & "     , a.ASK���i�t���O"
'wSQL = wSQL & "     , a.�戵���~��"
'wSQL = wSQL & "     , a.�p�ԓ�"
'wSQL = wSQL & "     , a.�I����"
'wSQL = wSQL & "     , a.�󏭐���"
'wSQL = wSQL & "     , a.�Z�b�g���i�t���O"
'wSQL = wSQL & "     , a.�J�e�S���[�R�[�h"
'wSQL = wSQL & "     , a.���A���i�t���O"
'wSQL = wSQL & "     , a.�����t���O"
'wSQL = wSQL & "     , a.����URL"
'wSQL = wSQL & "     , a.����t���O"
'wSQL = wSQL & "     , a.����URL"
'wSQL = wSQL & "     , a.Web�[����\���t���O"
'wSQL = wSQL & "     , a.���ח\�薢��t���O"
'wSQL = wSQL & "     , a.���i�X�y�b�N�g�p�s�t���O"
'wSQL = wSQL & "     , a.B�i�P��"
'wSQL = wSQL & "     , a.������"
'wSQL = wSQL & "     , a.������"
'wSQL = wSQL & "     , a.�O��P���ύX��"
'wSQL = wSQL & "     , a.�O��̔��P��"
'wSQL = wSQL & "     , a.B�i�t���O"
'wSQL = wSQL & "     , a.����o�^��"   '2010/03/18 an add
'wSQL = wSQL & "     , a.�������S�������i�t���O"				' 2011/02/23 GV Add
'wSQL = wSQL & "     , CASE"
'wSQL = wSQL & "         WHEN (a.�����ߏ��i�\���� = 0) THEN 999999"
'wSQL = wSQL & "         ELSE a.�����ߏ��i�\����"
'wSQL = wSQL & "       END AS �����ߏ��i�\����"
'wSQL = wSQL & "     , b.�F"
'wSQL = wSQL & "     , b.�K�i"
'wSQL = wSQL & "     , b.�����\����"
'wSQL = wSQL & "     , b.��������"								'2011/06/09 hn add
'wSQL = wSQL & "     , b.�����\���ח\���"
'wSQL = wSQL & "     , b.B�i�����\����"
'wSQL = wSQL & "     , b.���iID"
'wSQL = wSQL & "     , b.�K���݌ɐ���"
'wSQL = wSQL & "     , CASE"
'wSQL = wSQL & "         WHEN a.B�i�t���O = 'Y' THEN b.B�i�����\����"    '2010/07/16 an mod
'wSQL = wSQL & "         WHEN a.�����萔�� > 0 AND (a.�����萔��-a.������󒍍ϐ���) > 0 THEN a.�����萔��-a.������󒍍ϐ���"    '2010/07/16 an mod
'wSQL = wSQL & "         WHEN b.�����\���� <= 0 THEN -1"                                     '2010/03/18 an mod s
'wSQL = wSQL & "         WHEN b.�����\���� > 0 AND b.�����\���� <= a.�󏭐��� THEN 0"
'wSQL = wSQL & "         WHEN b.�����\���� > 0 AND b.�����\���� > a.�󏭐��� THEN 9999"
'wSQL = wSQL & "         ELSE 99999"
'wSQL = wSQL & "       END AS �݌ɗL��"                                                        '2010/03/18 an mod e
'wSQL = wSQL & "     , CASE"
'wSQL = wSQL & "         WHEN a.������ IS NULL AND a.�戵���~�� IS NULL THEN 1"
'wSQL = wSQL & "         ELSE 2"
'wSQL = wSQL & "       END AS �����敪"
'wSQL = wSQL & "     , c.���[�J�[��"
'
'wSQL = wSQL & "     , (SELECT COUNT(*)"
'wSQL = wSQL & "          FROM ���i�X�y�b�N s WITH (NOLOCK)"
'wSQL = wSQL & "         WHERE s.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'wSQL = wSQL & "           AND s.���i�R�[�h = a.���i�R�[�h"
'wSQL = wSQL & "       ) AS ���i�X�y�b�NCNT"
'
''�F�K�i�����邩�ǂ��� 2007/05/30
'wSQL = wSQL & "     , (SELECT COUNT(*)"
'wSQL = wSQL & "          FROM Web�F�K�i�ʍ݌� t WITH (NOLOCK)"
'wSQL = wSQL & "         WHERE t.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'wSQL = wSQL & "           AND t.���i�R�[�h = a.���i�R�[�h"
'wSQL = wSQL & "           AND (t.�F != '' OR t.�K�i != '')"
'wSQL = wSQL & "           AND t.�I���� IS NULL"
'wSQL = wSQL & "       ) AS �F�K�iCNT"
'
''---- �F�K�i�̍��v�����\����	2011/06/09 hn mod
'wSQL = wSQL & "     , ISNULL((SELECT SUM(�����\����)"
'wSQL = wSQL & "                 FROM Web�F�K�i�ʍ݌� u WITH (NOLOCK)"
'wSQL = wSQL & "                WHERE u.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'wSQL = wSQL & "                  AND u.���i�R�[�h = a.���i�R�[�h"
'wSQL = wSQL & "                  AND u.�I���� IS NULL"
'wSQL = wSQL & "                  AND u.�����\���� > 0),0)"
'wSQL = wSQL & "       AS �F�K�i���v�����\����"
'
''---- �F�K�i�̍��v��������	2011/06/09 hn add
'wSQL = wSQL & "     , ISNULL((SELECT SUM(��������)"
'wSQL = wSQL & "                FROM Web�F�K�i�ʍ݌� w WITH (NOLOCK)"
'wSQL = wSQL & "               WHERE w.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'wSQL = wSQL & "                 AND w.���i�R�[�h = a.���i�R�[�h"
'wSQL = wSQL & "                 AND w.�I���� IS NULL"
'wSQL = wSQL & "                 AND w.�������� > 0),0)"
'wSQL = wSQL & "       AS �F�K�i���v��������"
'
''���r���[�]���̕���  '2010/03/18 an add
'wSQL = wSQL & "     , (SELECT CAST(AVG(CAST(ISNULL(v.�]��,0) AS decimal(1,0))) AS decimal(2,1)) "   '2010/05/13 an changed
'wSQL = wSQL & "          FROM ���i���r���[ v WITH (NOLOCK)"
'wSQL = wSQL & "         WHERE v.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'wSQL = wSQL & "           AND v.���i�R�[�h = a.���i�R�[�h"
''''wSQL = wSQL & "           AND v.�V���b�v�R�����g�� IS NULL"   '2010/05/13 an del
'wSQL = wSQL & "       ) AS ���r���[�]������"
'
''���r���[�]������  '2010/07/12 st add
'wSQL = wSQL & "     , (SELECT COUNT(*) "
'wSQL = wSQL & "          FROM ���i���r���[ w WITH (NOLOCK)"
'wSQL = wSQL & "         WHERE w.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'wSQL = wSQL & "           AND w.���i�R�[�h = a.���i�R�[�h"
'wSQL = wSQL & "       ) AS ���r���[��"
'
''---- FROM
'wSQL2 = ""
'wSQL2 = wSQL2 & "  FROM Web���i a WITH (NOLOCK)"
'wSQL2 = wSQL2 & "     , Web�F�K�i�ʍ݌� b WITH (NOLOCK)"
'wSQL2 = wSQL2 & "     , ���[�J�[ c WITH (NOLOCK)"
'wSQL2 = wSQL2 & "     , �J�e�S���[ d WITH (NOLOCK)"
'wSQL2 = wSQL2 & "     , ���J�e�S���[ e WITH (NOLOCK)"
'wSQL2 = wSQL2 & "     , ��J�e�S���[ h WITH (NOLOCK)"    '2010/06/08 an add
'wSQL2 = wSQL2 & "     , ���i�J�e�S���[ f WITH (NOLOCK)"        '2005/07/19
'
''---- WHERE
'wSQL2 = wSQL2 & " WHERE a.Web���i�t���O = 'Y'"
'wSQL2 = wSQL2 & "   AND b.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'wSQL2 = wSQL2 & "   AND b.���i�R�[�h = a.���i�R�[�h"
'wSQL2 = wSQL2 & "   AND b.�F = ''"    '2007/05/30 add
'wSQL2 = wSQL2 & "   AND b.�K�i = ''"    '2007/05/30 add
'wSQL2 = wSQL2 & "   AND c.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'
'wSQL2 = wSQL2 & "   AND d.�J�e�S���[�R�[�h = f.�J�e�S���[�R�[�h "        '2005/07/19
'wSQL2 = wSQL2 & "   AND e.���J�e�S���[�R�[�h = d.���J�e�S���[�R�[�h"        '2005/07/19
'wSQL2 = wSQL2 & "   AND f.���[�J�[�R�[�h = a.���[�J�[�R�[�h"        '2005/07/19
'wSQL2 = wSQL2 & "   AND f.���i�R�[�h = a.���i�R�[�h "        '2005/07/19
'wSQL2 = wSQL2 & "   AND h.��J�e�S���[�R�[�h = e.��J�e�S���[�R�[�h"   '2010/06/08 an add
'
''---- �J�e�S���[���w�肵�či�荞��
'if CategoryCd <> "" then   '2010/06/17 an mod
'    wSQL2 = wSQL2 & "  AND f.�J�e�S���[�R�[�h = '" & CategoryCd & "'"    '2005/07/19 '2010/06/17 an mod
'end if
'
''---- ���J�e�S���[���w�肵�či�荞��
'if MidCategoryCd <> "" then                                                                          '2010/06/08 an add s
'    wSQL2 = wSQL2 & "   AND e.���J�e�S���[�R�[�h = '" & MidCategoryCd & "'"
'end if
'
''---- ��J�e�S���[���w�肵�či�荞��
'if LargeCategoryCd <> "" then
'    wSQL2 = wSQL2 & "   AND e.��J�e�S���[�R�[�h = '" & LargeCategoryCd & "'"
'end if                                                                                                '2010/06/08 an add e
'
'wSQL2 = wSQL2 & "   AND ((a.�����萔�� > a.������󒍍ϐ��� AND a.�����萔�� > 0) OR a.�p�ԓ� IS NOT NULL)"
'
'if s_maker_cd <> "" then
'    wSQL2 = wSQL2 & "   AND c.���[�J�[�R�[�h = '" & s_maker_cd & "'"
'end if
'
'if trim(Request("sPriceFrom")) = "" AND Trim(Request("sPriceTo")) = "" then
'else
'    wSQL2 = wSQL2 & "   AND (a.�̔��P�� * (" & wSalesTaxRate & " + 100) / 100) BETWEEN " & sPriceFrom & " AND " & sPriceTo
'end if
'
'if s_product_cd <> "" then
'    if instr(s_product_cd, "%") > 0 then
'        wSQL2 = wSQL2 & "   AND a.���i�R�[�h LIKE '" & s_product_cd & "'"
'    else
'        wSQL2 = wSQL2 & "   AND a.���i�R�[�h = '" & s_product_cd & "'"
'    end if
'end if
'
'if sSeriesCd <> "" then
'    wSQL2 = wSQL2 & "   AND a.�V���[�Y�R�[�h = '" & sSeriesCd & "'"
'end if
'
''---- ORDER BY  2010/03/18 an mod
'v_order = ""
'
'Select Case i_sort
'    Case "Price_ASC"
'    	v_order = v_order & " ORDER BY ���̔��P��, c.���[�J�[��, a.���i��"
'    Case "Price_DESC"
'    	v_order = v_order & " ORDER BY ���̔��P�� DESC, c.���[�J�[��, a.���i��"
'    Case "MakerName_ASC"
'    	v_order = v_order & " ORDER BY c.���[�J�[��, a.���i��"
'    Case "ProductName_ASC"
'    	v_order = v_order & " ORDER BY a.���i��"
'    Case "NewArrivals"
'    	v_order = v_order & " ORDER BY a.������ DESC, c.���[�J�[��, a.���i��"
'    Case "Reviews"
'    	v_order = v_order & " ORDER BY ���r���[�]������ DESC, c.���[�J�[��, a.���i��"
'    Case "ReviewCount"
'        v_order = " ORDER BY ���r���[�� DESC, c.���[�J�[��, a.���i��"        '2010/07/12 st add
'    Case "Zaiko_DESC"
'    	v_order = v_order & " ORDER BY �����敪, �݌ɗL�� DESC, c.���[�J�[��, a.���i��"
'    Case Else
'    	v_order = v_order & " ORDER BY �����ߏ��i�\����, c.���[�J�[��, a.���i��"
'End Select
'
''---- �Y�����i�ꗗSQL
'wSQL = wSQL & wSQL2 & v_order
'
''---- �i�荞�ݗp���[�J�[SQL
'wSQLMaker = "SELECT c.���[�J�[��, c.���[�J�[�R�[�h, COUNT(DISTINCT a.���i�R�[�h) AS ���i����"
'wSQLMaker = wSQLMaker & wSQL2     '2010/06/08 an mod
'wSQLMaker = wSQLMaker & " GROUP BY c.���[�J�[��, c.���[�J�[�R�[�h"
'wSQLMaker = wSQLMaker & " ORDER BY 1"
'
''---- �i�荞�ݗp�J�e�S���[SQL
'wSQLCategory = "SELECT e.���J�e�S���[�����{��, d.�J�e�S���[��, d.�J�e�S���[�R�[�h, COUNT(DISTINCT a.���i�R�[�h) AS ���i����"
'wSQLCategory = wSQLCategory & wSQL2
'wSQLCategory = wSQLCategory & " GROUP BY e.���J�e�S���[�����{��, d.�J�e�S���[��, d.�J�e�S���[�R�[�h"
'wSQLCategory = wSQLCategory & " ORDER BY 1"
'
''---- �i�荞�ݗp���J�e�S���[SQL                                                                  '2010/06/08 an add s
'wSQLMidCategory = ""
'wSQLMidCategory = "SELECT h.��J�e�S���[��, e.���J�e�S���[�����{��, e.���J�e�S���[�R�[�h, COUNT(DISTINCT a.���i�R�[�h) AS ���i����"
'wSQLMidCategory = wSQLMidCategory & wSQL2
'wSQLMidCategory = wSQLMidCategory & " GROUP BY h.��J�e�S���[��, e.���J�e�S���[�����{��, e.���J�e�S���[�R�[�h"
'wSQLMidCategory = wSQLMidCategory & " ORDER BY 1"
'
''---- �i�荞�ݗp��J�e�S���[SQL
'wSQLLargeCategory = ""
'wSQLLargeCategory = "SELECT h.��J�e�S���[��, h.��J�e�S���[�R�[�h, COUNT(DISTINCT a.���i�R�[�h) AS ���i����"
'wSQLLargeCategory = wSQLLargeCategory & wSQL2
'wSQLLargeCategory = wSQLLargeCategory & " GROUP BY h.��J�e�S���[��, h.��J�e�S���[�R�[�h"
'wSQLLargeCategory = wSQLLargeCategory & " ORDER BY 1"                                            '2010/06/08 an add e
'
''---- �i�荞�ݗp���i��SQL
'wSQLPricerange = "SELECT MAX(a.�̔��P��) AS MAX�̔��P��, MIN(a.�̔��P��) AS MIN�̔��P��"
'wSQLPricerange = wSQLPricerange & wSQL2      '2010/06/08 an mod
'--- Web�Z�[�����i �e�[�u���Ή� �������� ---
wSQL = wSQL & "SELECT DISTINCT "
wSQL = wSQL & "      a.���[�J�[�R�[�h "
wSQL = wSQL & "    , a.���i�R�[�h "
wSQL = wSQL & "    , a.���i�� "
wSQL = wSQL & "    , a.���i�T��Web "
wSQL = wSQL & "    , a.�����敪 "
wSQL = wSQL & "    , a.���菤�i�� "
wSQL = wSQL & "    , a.�d�ʏ��i���� "
wSQL = wSQL & "    , a.���i�摜�t�@�C����_�� "
wSQL = wSQL & "    , a.���i���l "
wSQL = wSQL & "    , a.�W���P�� "
wSQL = wSQL & "    , a.�̔��P�� "
wSQL = wSQL & "    , a.������P��               AS ���̔��P�� "
wSQL = wSQL & "    , a.���̔��P��                 AS ������P�� "
wSQL = wSQL & "    , a.�����萔�� "
wSQL = wSQL & "    , a.������󒍍ϐ��� "
wSQL = wSQL & "    , a.�I�[�v�����i�t���O "
wSQL = wSQL & "    , a.���[�J�[�������敪 "
wSQL = wSQL & "    , a.ASK���i�t���O "
wSQL = wSQL & "    , a.�戵���~�� "
wSQL = wSQL & "    , a.�p�ԓ� "
wSQL = wSQL & "    , a.�I���� "
wSQL = wSQL & "    , a.�󏭐��� "
wSQL = wSQL & "    , a.�Z�b�g���i�t���O "
'wSQL = wSQL & "    , b.�J�e�S���[�R�[�h "
wSQL = wSQL & "    , a.�J�e�S���[�R�[�h "	'2011/01/14 na mod
wSQL = wSQL & "    , a.���A���i�t���O "
wSQL = wSQL & "    , a.�����t���O "
wSQL = wSQL & "    , a.����URL "
wSQL = wSQL & "    , a.����t���O "
wSQL = wSQL & "    , a.����URL "
wSQL = wSQL & "    , a.Web�[����\���t���O "
wSQL = wSQL & "    , a.���ח\�薢��t���O "
wSQL = wSQL & "    , a.���i�X�y�b�N�g�p�s�t���O "
wSQL = wSQL & "    , a.B�i�P�� "
wSQL = wSQL & "    , a.������ "
wSQL = wSQL & "    , a.������ "
wSQL = wSQL & "    , a.�O��P���ύX�� "
wSQL = wSQL & "    , a.�O��̔��P�� "
wSQL = wSQL & "    , a.B�i�t���O "
wSQL = wSQL & "    , a.����o�^�� "
wSQL = wSQL & "    , a.�������S�������i�t���O "
wSQL = wSQL & "    , a.�����\����                 AS �����ߏ��i�\���� "
wSQL = wSQL & "    , a.�F "
wSQL = wSQL & "    , a.�K�i "
wSQL = wSQL & "    , a.�����\���� "
wSQL = wSQL & "    , a.�������� "
wSQL = wSQL & "    , a.�����\���ח\��� "
wSQL = wSQL & "    , a.B�i�����\���� "
wSQL = wSQL & "    , a.���iID "
wSQL = wSQL & "    , a.�K���݌ɐ��� "
wSQL = wSQL & "    , a.�݌ɗL�� "
wSQL = wSQL & "    , a.�����敪 "
wSQL = wSQL & "    , a.���[�J�[�� "
wSQL = wSQL & "    , a.���i�X�y�b�NCNT "
wSQL = wSQL & "    , a.�F�K�iCNT "
wSQL = wSQL & "    , a.�F�K�i���v�����\���� "
wSQL = wSQL & "    , a.�F�K�i���v�������� "
wSQL = wSQL & "    , a.���r���[�]������ "
wSQL = wSQL & "    , a.���r���[�� "
wSQL = wSQL & "    , CASE "							'2012/10/22 ok Add
wSQL = wSQL & "        WHEN a.ASK���i�t���O != 'Y' THEN "
wSQL = wSQL & "         CASE "
wSQL = wSQL & "           WHEN a.�����萔�� > a.������󒍍ϐ���         AND (a.�̔��P�� - a.������P��) > 0 THEN (a.�̔��P�� - a.������P��) / a.�̔��P�� "
wSQL = wSQL & "           WHEN a.B�i�t���O = 'Y'                             AND (a.�̔��P�� - a.B�i�P��) > 0 THEN (a.�̔��P�� - a.B�i�P��) / a.�̔��P�� "
wSQL = wSQL & "           WHEN DATEADD(d, 60, a.�O��P���ύX��) >= GETDATE() AND (a.�O��̔��P�� - a.�̔��P��) > 0 THEN (a.�O��̔��P�� - a.�̔��P��) / a.�O��̔��P�� "
wSQL = wSQL & "           ELSE 0 "
wSQL = wSQL & "         END "
wSQL = wSQL & "        ELSE 0 "
wSQL = wSQL & "      END AS �l������ "

'--- FROM
wSQL2 = ""
wSQL2 = wSQL2 & "FROM "
wSQL2 = wSQL2 & "      Web�Z�[�����i       a WITH (NOLOCK) "

'wSQL2 = wSQL2 & "        LEFT JOIN Web���i b WITH (NOLOCK) "				'2011/01/14 na del
'wSQL2 = wSQL2 & "          ON     b.���[�J�[�R�[�h = a.���[�J�[�R�[�h "
'wSQL2 = wSQL2 & "             AND b.���i�R�[�h     = a.���i�R�[�h "

'--- WHERE
wSQL2 = wSQL2 & "WHERE "
wSQL2 = wSQL2 & "       a.�Z�[���敪�ԍ� = 3 "						' ���ʌ��� (BargainSalse) �Z�[���敪�ԍ� : 3

'--- �J�e�S���[���w�肵�či�荞��
If CategoryCd <> "" Then
    wSQL2 = wSQL2 & "    AND a.�J�e�S���[�R�[�h = '" & CategoryCd & "' "
End If

'--- ���J�e�S���[���w�肵�či�荞��
If MidCategoryCd <> "" Then
    wSQL2 = wSQL2 & "    AND a.���J�e�S���[�R�[�h = '" & MidCategoryCd & "' "
End If

'--- ��J�e�S���[���w�肵�či�荞��
If LargeCategoryCd <> "" Then
    wSQL2 = wSQL2 & "    AND a.��J�e�S���[�R�[�h = '" & LargeCategoryCd & "' "
End If

If s_maker_cd <> "" Then
    wSQL2 = wSQL2 & "    AND a.���[�J�[�R�[�h = '" & s_maker_cd & "' "
End If

If Trim(Request("sPriceFrom")) = "" And Trim(Request("sPriceTo")) = "" Then
Else
'2014/03/19 GV mod start ---->
'���i�т̌���������Ŕ���
'    wSQL2 = wSQL2 & "    AND (a.�̔��P�� * (" & wSalesTaxRate & " + 100) / 100) BETWEEN " & sPriceFrom & " AND " & sPriceTo & " "
    wSQL2 = wSQL2 & "    AND a.�̔��P�� BETWEEN " & sPriceFrom & " AND " & sPriceTo & " "
'2014/03/19 GV mod end   <----
End If

If s_product_cd <> "" Then
    If Instr(s_product_cd, "%") > 0 Then
        wSQL2 = wSQL2 & "    AND a.���i�R�[�h LIKE '" & s_product_cd & "' "
    Else
        wSQL2 = wSQL2 & "    AND a.���i�R�[�h = '" & s_product_cd & "' "
    End If
End If

If sSeriesCd <> "" Then
    wSQL2 = wSQL2 & "    AND a.�V���[�Y�R�[�h = '" & sSeriesCd & "'"
End If

Select Case i_sort
    Case "Price_ASC"
        v_order = " ORDER BY ���̔��P��, a.���[�J�[��, a.���i�� "
    Case "Price_DESC"
        v_order = " ORDER BY ���̔��P�� DESC, a.���[�J�[��, a.���i�� "
    Case "MakerName_ASC"
        v_order = " ORDER BY a.���[�J�[��, a.���i�� "
    Case "ProductName_ASC"
        v_order = " ORDER BY a.���i�� "
    Case "NewArrivals"
        v_order = " ORDER BY a.������ DESC, a.���[�J�[��, a.���i�� "
    Case "Reviews"
        v_order = " ORDER BY a.���r���[�]������ DESC, a.���[�J�[��, a.���i�� "
    Case "ReviewCount"
        v_order = " ORDER BY a.���r���[�� DESC, a.���[�J�[��, a.���i�� "
    Case "Zaiko_DESC"
        v_order = " ORDER BY a.�����敪, a.�݌ɗL�� DESC, a.���[�J�[��, a.���i�� "
	Case "Nesage_DESC"								'2012/10/22 ok Add
		v_order = " ORDER BY �l������ DESC, a.���[�J�[��, a.���i�� "
    Case Else
        v_order = " ORDER BY a.�����\����, a.���[�J�[��, a.���i�� "
End Select

'--- �Y�����i�ꗗSQL
wSQL = wSQL & wSQL2 & v_order

'--- �i�荞�ݗp���[�J�[SQL
wSQLMaker = "SELECT a.���[�J�[��, a.���[�J�[�R�[�h, COUNT(DISTINCT a.���i�R�[�h) AS ���i���� " _
          & wSQL2 & " "_
          & "GROUP BY a.���[�J�[��, a.���[�J�[�R�[�h " _
          & "ORDER BY 1"

'--- �i�荞�ݗp�J�e�S���[SQL	'2012/11/09 ok Mod
wSQLCategory = "SELECT a.���J�e�S���[�����{��, a.�J�e�S���[��, a.�J�e�S���[�R�[�h, COUNT(DISTINCT a.���i�R�[�h) AS ���i���� " _
             & wSQL2 & " " _
             & "GROUP BY a.���J�e�S���[�����{��, a.�J�e�S���[��, a.�J�e�S���[�R�[�h "
'             & "ORDER BY 1"
wSQLCategory = "SELECT a.* FROM (" & wSQLCategory & ") a INNER JOIN �J�e�S���[ b WITH (NOLOCK) ON " _
             & "a.�J�e�S���[�R�[�h = b.�J�e�S���[�R�[�h ORDER BY b.�\���� "

'--- �i�荞�ݗp���J�e�S���[SQL	'2010/06/08 an add s	'2012/11/09 ok Mod
wSQLMidCategory = "SELECT a.��J�e�S���[��, a.���J�e�S���[�����{��, a.���J�e�S���[�R�[�h, COUNT(DISTINCT a.���i�R�[�h) AS ���i���� " _
                & wSQL2 & " " _
                & "GROUP BY a.��J�e�S���[��, a.���J�e�S���[�����{��, a.���J�e�S���[�R�[�h "
'                & "ORDER BY 1"
wSQLMidCategory = "SELECT a.* FROM (" & wSQLMidCategory & ") a INNER JOIN ���J�e�S���[ b WITH (NOLOCK) ON " _
                & "a.���J�e�S���[�R�[�h = b.���J�e�S���[�R�[�h ORDER BY b.�\���� "

'--- �i�荞�ݗp��J�e�S���[SQL	'2012/11/09 ok Mod
wSQLLargeCategory = "SELECT a.��J�e�S���[��, a.��J�e�S���[�R�[�h, COUNT(DISTINCT a.���i�R�[�h) AS ���i���� " _
                  & wSQL2 & " " _
                  & "GROUP BY a.��J�e�S���[��, a.��J�e�S���[�R�[�h "
'                  & "ORDER BY 1"
wSQLLargeCategory = "SELECT a.* FROM (" & wSQLLargeCategory & ") a INNER JOIN ��J�e�S���[ b WITH (NOLOCK) ON " _
                  & "a.��J�e�S���[�R�[�h = b.��J�e�S���[�R�[�h ORDER BY b.�\���� "

'--- �i�荞�ݗp���i��SQL
wSQLPricerange = "SELECT MAX(a.�̔��P��) AS MAX�̔��P��, MIN(a.�̔��P��) AS MIN�̔��P�� " _
               & wSQL2
' 20120110 GV Mod End

'@@@@response.write("<br>s_maker_cd=" & s_maker_cd & "<br>s_category_cd=" & s_category_cd & "<br>" & w_where)
'@@@@response.write(wSQL)
'@@@@response.write("<br><br>" & wSQLMaker)
'@@@@response.write("<br><br>" & wSQLCategory)
'@@@@response.write("<br><br>" & wSQLMidCategory)
'@@@@response.write("<br><br>" & wSQLLargeCategory)

Set RS = Server.CreateObject("ADODB.Recordset")

RS.Open wSQL, Connection, adOpenStatic

End Function

'========================================================================
'
'    Function    ���[�J�[���쐬
'
'========================================================================
'
Function CreateMakerInfo()

Dim RSv
Dim i

'---- ���[�J�[�����o��
wSQL = ""
wSQL = wSQL & "SELECT a.���[�J�[��"
wSQL = wSQL & "     , a.���[�J�[�z�[���y�[�WURL"
wSQL = wSQL & "     , a.���[�J�[���S�t�@�C����"
wSQL = wSQL & "     , a.���[�J�[�Љ�"
wSQL = wSQL & "  FROM ���[�J�[ a WITH (NOLOCK)"
wSQL = wSQL & " WHERE a.���[�J�[�R�[�h = '" & s_maker_cd & "'"

'@@@@@@@@@@response.write(wSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic

wHTML = ""
'---- ���[�J�[���
wHTML = wHTML & "    <div id='category_box'>" & vbNewLine
wHTML = wHTML & "      <p class='logo'>"

If RSv("���[�J�[���S�t�@�C����") <> "" Then
    wHTML = wHTML & "<img src='maker_img/" & RSv("���[�J�[���S�t�@�C����") & "' alt='" & RSv("���[�J�[��") & "'>"
end if
wHTML = wHTML & "</p>" & vbNewLine
wHTML = wHTML & "      <p class='txt'>" & Replace(RSv("���[�J�[�Љ�"), vbNewLine, "<br>") & "</p>" & vbNewLine
wHTML = wHTML & "    </div>" & vbNewLine

RSv.Close

'2012/07/11 ok Del Start
'---- ���[�J�[����؃����L���O ���o��
'wSQL = ""
'wSQL = wSQL & "SELECT TOP 5"
'wSQL = wSQL & "       a.���[�J�[�R�[�h"
'wSQL = wSQL & "     , a.���i�R�[�h"
'wSQL = wSQL & "     , b.���i��"
'wSQL = wSQL & "     , c.���[�J�[��"
'wSQL = wSQL & "  FROM "
'wSQL = wSQL & "       ���؏��i a WITH (NOLOCK)"
'wSQL = wSQL & "     , Web���i b WITH (NOLOCK)"
'wSQL = wSQL & "     , ���[�J�[ c WITH (NOLOCK)"
'wSQL = wSQL & "     , �J�e�S���[ d WITH (NOLOCK)"
'wSQL = wSQL & "     , Web�F�K�i�ʍ݌� g WITH (NOLOCK)"
'wSQL = wSQL & " WHERE "
'wSQL = wSQL & "       b.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'wSQL = wSQL & "   AND b.���i�R�[�h = a.���i�R�[�h"
'wSQL = wSQL & "   AND c.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'wSQL = wSQL & "   AND d.�J�e�S���[�R�[�h = a.�J�e�S���[�R�[�h"
'wSQL = wSQL & "   AND g.���[�J�[�R�[�h = b.���[�J�[�R�[�h"
'wSQL = wSQL & "   AND g.���i�R�[�h = b.���i�R�[�h"
'wSQL = wSQL & "   AND b.�I���� IS NULL"
'wSQL = wSQL & "   AND g.�I���� IS NULL"
'wSQL = wSQL & "   AND b.Web���i�t���O = 'Y'"
'wSQL = wSQL & "   AND d.����؃����L���O�\���t���O = 'Y'"
'wSQL = wSQL & "   AND a.���[�J�[�R�[�h = '" & s_maker_cd & "'"
'wSQL = wSQL & "   AND a.�N�� = (SELECT MAX(�N��) FROM ���؏��i)"
'wSQL = wSQL & " ORDER BY"
'wSQL = wSQL & "       a.�󒍐��� DESC"
'
''@@@@@@@@@@response.write(wSQL)
'
'Set RSv = Server.CreateObject("ADODB.Recordset")
'RSv.Open wSQL, Connection, adOpenStatic
'
'wHTML = wHTML & "    <td valign=top>" & vbNewLine
'
'if RSv.EOF = false then
'    wHTML = wHTML & "      <table width=215 border=0 cellspacing=0 cellpadding=0>" & vbNewLine
'    wHTML = wHTML & "        <tr>" & vbNewLine
'    wHTML = wHTML & "          <td align='center' nowrap='nowrap' bgcolor='#ccccff' class='honbun'>����</td>" & vbNewLine
'    wHTML = wHTML & "          <td bgcolor='#eeeeee' class='honbun'><h3 style='font-size:100%;font-weight:normal;margin: 0px 0px 0px 0px'>" & RSv("���[�J�[��") & "</h3></td>" & vbNewLine
'    wHTML = wHTML & "        </tr>" & vbNewLine
'
'    i = 0
'    '----�����L���O�쐬
'    Do until RSv.EOF = true
'        i = i + 1
'        wHTML = wHTML & "        <tr>" & vbNewLine
'        wHTML = wHTML & "          <td align='center' nowrap='nowrap' class='honbun'>" & i & ".</td>" & vbNewLine
'        wHTML = wHTML & "          <td><a href='ProductDetail.asp?item=" & Server.URLEncode(RSv("���[�J�[�R�[�h") & "^" & RSv("���i�R�[�h")) & "' class='link'>" & RSv("���i��") & "</a></td>" & vbNewLine
'        wHTML = wHTML & "        </tr>" & vbNewLine
'
'        RSv.MoveNext
'    Loop
'
'    RSv.Close
'
'    wHTML = wHTML & "      </table>" & vbNewLine
'end if
'
'wHTML = wHTML & "    </td>" & vbNewLine
'wHTML = wHTML & "  </tr>" & vbNewLine
'wHTML = wHTML & "</table>" & vbNewLine
'wHTML = wHTML & "<div style='padding: 10px 0px 0px 0px;'></div>" & vbNewLine
'2012/07/11 ok Del End

wMakerInfoHTML = wHTML

End Function

'========================================================================
'
'    Function    Close database
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
<title>�Ռ������i <%=wTitle%> �ꗗ�b�T�E���h�n�E�X</title>
<meta name="Description" content="�T�E���h�n�E�X���������߂���Ƃ��Ă����������ȏ��i���Ռ������ł��񋟁B���ʌ���̂��߂������͂����߂ɁI">
<meta name="keywords" content="�Ռ������i,�A�E�g���b�g,���ʌ���o�[�Q��,�����s��,�v���C�X�_�E�����">
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/shop.css?20121108" type="text/css">
<link rel="stylesheet" href="Style/searchlist.css?20121201" type="text/css">
<link rel="stylesheet" href="style/ask.css?20140401a" type="text/css">
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
<%=wFootprintHTML%>
      </ul>
    </div></div></div>

<!-- �^�u�o�i�[�����̋L�q START -->
    <div class="tab" id="bargainsale">
      <ul>
        <li><a href="SpecialPriceSale.asp"><img src="images/tab_bargainsalea.png" alt="�����s��"></a></li>
        <li><a href="Outlet.asp"><img src="images/tab_outlet.png" alt="�킯����s��"></a></li>
        <li><a href="PriceDownSale.asp"><img src="images/tab_pricedown.png" alt="�v���C�X�_�E��"></a></li>
      </ul>
    </div>
<!-- �^�u�o�i�[�����̋L�q END -->

<!-- ���[�J�[��� (���[�J�[�Ō�����) -->
<%=wMakerInfoHTML%>

<!-- ���i�ꗗ -->
<% = wListHTML %>

  <!--/#contents --></div>

<!-- �i�������pForm -->
<form name="f_search" method="get" action="BargainSale.asp">
  <input type="hidden" name="s_maker_cd" value="<%=s_maker_cd%>">
  <input type="hidden" name="s_category_cd" value="<%=CategoryCd%>">
  <input type="hidden" name="s_mid_category_cd" value="<%=MidCategoryCd%>">
  <input type="hidden" name="s_large_category_cd" value="<%=LargeCategoryCd%>">
  <input type="hidden" name="s_product_cd" value="<%=s_product_cd%>">
  <input type="hidden" name="sSeriesCd" value="<%=sSeriesCd%>">
  <input type="hidden" name="sPriceFrom" value="<%=sPriceFrom%>">
  <input type="hidden" name="sPriceTo" value="<%=sPriceTo%>">
  <input type="hidden" name="i_page" value="1">
  <input type="hidden" name="i_sort" value="<%=i_sort%>">
  <input type="hidden" name="i_page_size" value="<%=i_page_size%>">
  <input type="hidden" name="i_ListType" value="<%=i_ListType%>">
</form>
      <div id="globalSide">
<%
'----NAVI�p�p�����[�^�Z�b�g
NAVISearchListMakerListHTML = wNaviMakerHTML
NAVISearchListCategoryListHTML = wNaviCategoryHTML
NAVISearchListPriceRangeListHTML = wNaviPriceRangeHTML
NAVISearchListLargeCategoryListHTML = wNaviLargeCategoryHTML
NAVISearchListMidCategoryListHTML = wNaviMidCategoryHTML
%>
	<!--#include file="../Navi/NaviSideSale.inc"-->
	<!--#include file="../Navi/NaviSide.inc"-->
	<!--/#globalSide --></div>
<!--/#main --></div>
<!--#include file="../Navi/Navibottom.inc"-->
<div class="tooltip"><p>ASK</p></div>
<!--#include file="../Navi/NaviScript.inc"-->
<script type="text/javascript" src="jslib/ask.js?20140401a"></script>
<script type="text/javascript" src="jslib/SearchList.js?20121108" charset="Shift_JIS"></script>
<script type="text/javascript" src="../jslib/jquery.tinyscrollbar.min.js"></script>
<script type="text/javascript">
$(function(){
    $('#scrollbar1').tinyscrollbar();
});
<% if wNoData <> "Y" then%>

    preset_values();

<% end if %>


//
//    Search onClick
//
function Search_onClick(pMakerCd, pCategoryCd, pMidCategoryCd, pLargeCategoryCd, pPriceFrom, pPriceTo){
    document.f_search.s_maker_cd.value = pMakerCd;
    document.f_search.s_category_cd.value = pCategoryCd;
    document.f_search.s_mid_category_cd.value = pMidCategoryCd;
    document.f_search.s_large_category_cd.value = pLargeCategoryCd;
    document.f_search.sPriceFrom.value = pPriceFrom;
    document.f_search.sPriceTo.value = pPriceTo;
    document.f_search.submit();
}
</script>
</body>
</html>
