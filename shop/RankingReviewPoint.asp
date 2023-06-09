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
'	���r���[�|�C���g�y�[�W
'
'�X�V����
'2010/04/26 an �V�K�쐬
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2012/01/20 GV �f�[�^�擾 SELECT���� LAC�N�G���[�Ă�K�p
'2012/01/23 GV �u���i���r���[�v�e�[�u������u���i���r���[�W�v�v�e�[�u���g�p�ɕύX (CreateReviewPointHTML()�v���V�[�W��)
'2012/08/08 if-web ���j���[�A�����C�A�E�g����
'2014/03/19 GV ����ő��łɔ���2�d�\���Ή�
'
'========================================================================

On Error Resume Next

Dim LargeCategoryCd

Dim wSalesTaxRate
Dim wLargeCategoryName
Dim wMidCategoryName
Dim wNoData

Dim wLargeCategoryHTML
Dim wReviewPointHTML

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
Dim wErrDesc   '2011/08/01 an add

'========================================================================

Response.buffer = true

'---- Get input data
LargeCategoryCd = ReplaceInput(Trim(Request("LargeCategoryCd")))

'---- ��J�e�S���[�R�[�h�̎w�肪�Ȃ��ꍇ
if LargeCategoryCd = "" then
	LargeCategoryCd = "1"
end if

'---- Execute main
call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "RankingReviewPoint.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc)
end if                                           '2011/08/01 an add e

call close_db()

'---- �z��O�̑�J�e�S���[�R�[�h���w�肳�ꂽ�ꍇ���G���[
if wNoData = "Y" OR Err.Description <> "" then
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

'---- ����ŗ���o��
call getCntlMst("����","����ŗ�","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)			'����ŗ�
wSalesTaxRate = Clng(wItemNum1)

'---- ��J�e�S���[�ꗗ�쐬
call CreateLargeCategoryHTML()

if wNoData <> "Y" then  '�z��O�̑�J�e�S���[���w�肳���NoData�̏ꍇ�̓G���[
	'---- ���r���[�|�C���g�����L���O�쐬
	call CreateReviewPointHTML()
end if

End Function

'========================================================================
'
'	Function	��J�e�S���[�ꗗ�\��
'
'========================================================================
'
Function CreateLargeCategoryHTML()

Dim vCount

'---- �S��J�e�S���[�����o��
wSQL = ""
wSQL = wSQL & "SELECT a.��J�e�S���[�R�[�h"
wSQL = wSQL & "     , a.��J�e�S���[��"
wSQL = wSQL & "  FROM ��J�e�S���[ a WITH (NOLOCK)"
wSQL = wSQL & " WHERE a.Web��J�e�S���[�t���O = 'Y'"
wSQL = wSQL & " ORDER BY a.�\����"

'@@@@@@@@@@response.write(wSQL)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

wHTML = ""
vCount = 0
wHTML = wHTML & "  <p id='cat_select'>"  & vbNewLine

Do Until RS.EOF = true

	vCount = vCount + 1
	wHTML = wHTML & "<a href='RankingReviewPoint.asp?LargeCategoryCd=" & RS("��J�e�S���[�R�[�h") & "'>" & RS("��J�e�S���[��") & "</a>"

	if RS("��J�e�S���[�R�[�h") = LargeCategoryCd then
		wLargeCategoryName = RS("��J�e�S���[��")  '���r���[�|�C���g�ꗗ�̃^�C�g���Ŏg�p
	end if

	RS.MoveNext

	'���Ƀf�[�^������Ύd�؂����\��
	if RS.EOF = false then
		wHTML = wHTML & "�b"

		if vCount = 8 then
			wHTML = wHTML & "<br>"  & vbNewLine
		end if
	end if

Loop

if wLargeCategoryName = "" then
	wNoData = "Y" '�z��O�̑�J�e�S���[���w�肳�ꂽ�ꍇ
else
	wHTML = wHTML & vbNewLine
	wHTML = wHTML & "  </p>"  & vbNewLine

	wLargeCategoryHTML = wHTML
end if

RS.close

End Function

'========================================================================
'
'	Function	���r���[�|�C���g�����L���O�ꗗ
'
'========================================================================
'
Function CreateReviewPointHTML()

Dim RSv
Dim vPrice
Dim vPriceNoTax				'2014/03/19 GV add
Dim vItem
Dim vRank

Dim vMakerProduct
Dim vProductName
Dim vProdTermFl '�̔��I�����i�t���O
Dim vInventoryCd
Dim vInventoryImage

Dim vBGColor

'---- ��J�e�S���[���Ƃ̃��r���[�|�C���gTOP25
wSQL = ""
' 2012/01/20 GV Mod Start
'wSQL = wSQL & "SELECT DISTINCT TOP 25"
'wSQL = wSQL & "     (SELECT CAST(AVG(CAST(ISNULL(h.�]��,0) AS decimal(1,0))) AS decimal(2,1))"
'wSQL = wSQL & "        FROM ���i���r���[ h WITH (NOLOCK)"
'wSQL = wSQL & "       WHERE h.���[�J�[�R�[�h = b.���[�J�[�R�[�h"
'wSQL = wSQL & "         AND h.���i�R�[�h = b.���i�R�[�h"
'wSQL = wSQL & "     ) AS ���r���[�]������"
'wSQL = wSQL & "     ,(SELECT COUNT(*)"
'wSQL = wSQL & "         FROM ���i���r���[ i WITH (NOLOCK)"
'wSQL = wSQL & "        WHERE i.���[�J�[�R�[�h = b.���[�J�[�R�[�h"
'wSQL = wSQL & "          AND i.���i�R�[�h = b.���i�R�[�h"
'wSQL = wSQL & "     ) AS ���r���[�R�����g��"
'wSQL = wSQL & "     ,(SELECT SUM(j.�]��)"
'wSQL = wSQL & "         FROM ���i���r���[ j WITH (NOLOCK)"
'wSQL = wSQL & "        WHERE j.���[�J�[�R�[�h = b.���[�J�[�R�[�h"
'wSQL = wSQL & "          AND j.���i�R�[�h = b.���i�R�[�h"
'wSQL = wSQL & "     ) AS ���r���[�]�����v"
'wSQL = wSQL & "     , b.���[�J�[�R�[�h"
'wSQL = wSQL & "     , b.�J�e�S���[�R�[�h"
'wSQL = wSQL & "     , b.���i�R�[�h"
'wSQL = wSQL & "     , b.���i��"
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
'wSQL = wSQL & "     , c.�F"
'wSQL = wSQL & "     , c.�K�i"
'wSQL = wSQL & "     , c.B�i�����\����"
'wSQL = wSQL & "     , c.�����\����"
'wSQL = wSQL & "     , c.�����\���ח\���"
'wSQL = wSQL & "     , c.���iID"
'wSQL = wSQL & "     , CASE"
'wSQL = wSQL & "         WHEN b.�����萔�� > b.������󒍍ϐ��� THEN b.������P��"
'wSQL = wSQL & "         WHEN b.B�i�t���O = 'Y' THEN b.B�i�P��"
'wSQL = wSQL & "       ELSE b.�̔��P��"
'wSQL = wSQL & "       END AS ���̔��P��"
'wSQL = wSQL & "     , d.���[�J�[��"
'wSQL = wSQL & "     , e.�J�e�S���[��"
'wSQL = wSQL & "     , g.��J�e�S���[��"
'
''�F�K�i�����邩�ǂ��� 2007/05/30
'wSQL = wSQL & "     , (SELECT COUNT(*)"
'wSQL = wSQL & "          FROM Web�F�K�i�ʍ݌� k WITH (NOLOCK)"
'wSQL = wSQL & "         WHERE k.���[�J�[�R�[�h = b.���[�J�[�R�[�h"
'wSQL = wSQL & "           AND k.���i�R�[�h = b.���i�R�[�h"
'wSQL = wSQL & "           AND (k.�F != '' OR k.�K�i != '')"
'wSQL = wSQL & "           AND k.�I���� IS NULL"
'wSQL = wSQL & "       ) AS �F�K�iCNT"
'
'wSQL = wSQL & " FROM ���i���r���[ a WITH (NOLOCK)"
'wSQL = wSQL & "    , Web���i b WITH (NOLOCK) "
'wSQL = wSQL & "    , Web�F�K�i�ʍ݌� c WITH (NOLOCK) "
'wSQL = wSQL & "    , ���[�J�[ d WITH (NOLOCK)  "
'wSQL = wSQL & "    , �J�e�S���[ e WITH (NOLOCK)"
'wSQL = wSQL & "    , ���J�e�S���[ f WITH (NOLOCK)  "
'wSQL = wSQL & "    , ��J�e�S���[ g WITH (NOLOCK) "
'wSQL = wSQL & " WHERE b.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'wSQL = wSQL & "   AND b.���i�R�[�h = a.���i�R�[�h"
'wSQL = wSQL & "   AND c.���[�J�[�R�[�h = b.���[�J�[�R�[�h"
'wSQL = wSQL & "   AND c.���i�R�[�h = b.���i�R�[�h"
'wSQL = wSQL & "   AND d.���[�J�[�R�[�h = b.���[�J�[�R�[�h"
'wSQL = wSQL & "   AND e.�J�e�S���[�R�[�h = b.�J�e�S���[�R�[�h"
'wSQL = wSQL & "   AND f.���J�e�S���[�R�[�h = e.���J�e�S���[�R�[�h"
'wSQL = wSQL & "   AND g.��J�e�S���[�R�[�h = f.��J�e�S���[�R�[�h"
'wSQL = wSQL & "   AND g.��J�e�S���[�R�[�h = '" & LargeCategoryCd & "'"
'wSQL = wSQL & "   AND b.Web���i�t���O = 'Y'"
'wSQL = wSQL & "   AND c.�F = ''"
'wSQL = wSQL & "   AND c.�K�i = ''"
'wSQL = wSQL & " ORDER BY"
'wSQL = wSQL & "    ���r���[�]������ DESC"
'wSQL = wSQL & "  , ���r���[�R�����g�� DESC"
'wSQL = wSQL & "  , d.���[�J�[��"
'wSQL = wSQL & "  , b.���i��"
' 2012/01/23 GV Mod Start
'wSQL = wSQL & "SELECT DISTINCT TOP 25 "
'wSQL = wSQL & "     (SELECT CAST(AVG(CAST(ISNULL(h.�]��, 0) AS DECIMAL(1, 0))) AS DECIMAL(2, 1)) "
'wSQL = wSQL & "        FROM ���i���r���[ h WITH (NOLOCK) "
'wSQL = wSQL & "       WHERE     h.���[�J�[�R�[�h = b.���[�J�[�R�[�h "
'wSQL = wSQL & "             AND h.���i�R�[�h = b.���i�R�[�h "
'wSQL = wSQL & "     ) AS ���r���[�]������ "
'wSQL = wSQL & "     ,(SELECT COUNT(i.ID) "
'wSQL = wSQL & "         FROM ���i���r���[ i WITH (NOLOCK) "
'wSQL = wSQL & "        WHERE     i.���[�J�[�R�[�h = b.���[�J�[�R�[�h "
'wSQL = wSQL & "              AND i.���i�R�[�h = b.���i�R�[�h "
'wSQL = wSQL & "     ) AS ���r���[�R�����g�� "
'wSQL = wSQL & "     ,(SELECT SUM(j.�]��) "
'wSQL = wSQL & "         FROM ���i���r���[ j WITH (NOLOCK) "
'wSQL = wSQL & "        WHERE     j.���[�J�[�R�[�h = b.���[�J�[�R�[�h "
'wSQL = wSQL & "              AND j.���i�R�[�h = b.���i�R�[�h "
'wSQL = wSQL & "     ) AS ���r���[�]�����v "
'wSQL = wSQL & "     , b.���[�J�[�R�[�h "
'wSQL = wSQL & "     , b.�J�e�S���[�R�[�h "
'wSQL = wSQL & "     , b.���i�R�[�h "
'wSQL = wSQL & "     , b.���i�� "
'wSQL = wSQL & "     , b.�戵���~�� "
'wSQL = wSQL & "     , b.������ "
'wSQL = wSQL & "     , b.�p�ԓ� "
'wSQL = wSQL & "     , b.B�i�t���O "
'wSQL = wSQL & "     , b.ASK���i�t���O "
'wSQL = wSQL & "     , b.�󏭐��� "
'wSQL = wSQL & "     , b.�����萔�� "
'wSQL = wSQL & "     , b.������󒍍ϐ��� "
'wSQL = wSQL & "     , b.�Z�b�g���i�t���O "
'wSQL = wSQL & "     , b.���[�J�[�������敪 "
'wSQL = wSQL & "     , b.���ח\�薢��t���O "
'wSQL = wSQL & "     , b.Web�[����\���t���O "
'wSQL = wSQL & "     , c.�F "
'wSQL = wSQL & "     , c.�K�i "
'wSQL = wSQL & "     , c.B�i�����\���� "
'wSQL = wSQL & "     , c.�����\���� "
'wSQL = wSQL & "     , c.�����\���ח\��� "
'wSQL = wSQL & "     , c.���iID "
'wSQL = wSQL & "     , CASE "
'wSQL = wSQL & "         WHEN b.�����萔�� > b.������󒍍ϐ��� THEN b.������P�� "
'wSQL = wSQL & "         WHEN b.B�i�t���O = 'Y' THEN b.B�i�P�� "
'wSQL = wSQL & "       ELSE b.�̔��P�� "
'wSQL = wSQL & "       END AS ���̔��P�� "
'wSQL = wSQL & "     , d.���[�J�[�� "
'wSQL = wSQL & "     , e.�J�e�S���[�� "
'wSQL = wSQL & "     , g.��J�e�S���[�� "
'wSQL = wSQL & "     , (SELECT COUNT(k.���i�R�[�h) "
'wSQL = wSQL & "          FROM Web�F�K�i�ʍ݌� k WITH (NOLOCK) "
'wSQL = wSQL & "         WHERE     k.���[�J�[�R�[�h = b.���[�J�[�R�[�h "
'wSQL = wSQL & "               AND k.���i�R�[�h = b.���i�R�[�h "
'wSQL = wSQL & "               AND (k.�F != '' OR k.�K�i != '') "
'wSQL = wSQL & "               AND k.�I���� IS NULL "
'wSQL = wSQL & "       ) AS �F�K�iCNT "
'wSQL = wSQL & "FROM "
'wSQL = wSQL & "    ���i���r���[                 a WITH (NOLOCK) "
'wSQL = wSQL & "      INNER JOIN Web���i         b WITH (NOLOCK) "
'wSQL = wSQL & "        ON     b.���[�J�[�R�[�h     = a.���[�J�[�R�[�h "
'wSQL = wSQL & "           AND b.���i�R�[�h         = a.���i�R�[�h "
'wSQL = wSQL & "      INNER JOIN Web�F�K�i�ʍ݌� c WITH (NOLOCK) "
'wSQL = wSQL & "        ON     c.���[�J�[�R�[�h     = b.���[�J�[�R�[�h "
'wSQL = wSQL & "           AND c.���i�R�[�h         = b.���i�R�[�h "
'wSQL = wSQL & "      INNER JOIN ���[�J�[        d WITH (NOLOCK) "
'wSQL = wSQL & "        ON     d.���[�J�[�R�[�h     = b.���[�J�[�R�[�h "
'wSQL = wSQL & "      INNER JOIN �J�e�S���[      e WITH (NOLOCK) "
'wSQL = wSQL & "        ON     e.�J�e�S���[�R�[�h   = b.�J�e�S���[�R�[�h "
'wSQL = wSQL & "      INNER JOIN ���J�e�S���[    f WITH (NOLOCK) "
'wSQL = wSQL & "        ON     f.���J�e�S���[�R�[�h = e.���J�e�S���[�R�[�h "
'wSQL = wSQL & "      INNER JOIN ��J�e�S���[    g WITH (NOLOCK) "
'wSQL = wSQL & "        ON     g.��J�e�S���[�R�[�h = f.��J�e�S���[�R�[�h "
'wSQL = wSQL & "      LEFT JOIN ( SELECT 'Y' AS 'ShohinWebY' ) t1 "
'wSQL = wSQL & "        ON     b.Web���i�t���O      = t1.ShohinWebY "
'wSQL = wSQL & "      LEFT JOIN ( SELECT ''  AS 'Iro' )        t2 "
'wSQL = wSQL & "        ON     c.�F               = t2.Iro "
'wSQL = wSQL & "      LEFT JOIN ( SELECT ''  AS 'Kikaku' )     t3 "
'wSQL = wSQL & "        ON     c.�K�i             = t3.Kikaku "
'wSQL = wSQL & "WHERE "
'wSQL = wSQL & "        t1.ShohinWebY IS NOT NULL "
'wSQL = wSQL & "    AND t2.Iro        IS NOT NULL "
'wSQL = wSQL & "    AND t3.Kikaku     IS NOT NULL "
'wSQL = wSQL & "    AND g.��J�e�S���[�R�[�h = '" & LargeCategoryCd & "' "
'wSQL = wSQL & "ORDER BY "
'wSQL = wSQL & "      ���r���[�]������ DESC "
'wSQL = wSQL & "    , ���r���[�R�����g�� DESC "
'wSQL = wSQL & "    , d.���[�J�[�� "
'wSQL = wSQL & "    , b.���i�� "

wSQL = wSQL & "SELECT DISTINCT TOP 25 "
wSQL = wSQL & "       a.���r���[�]������ "
wSQL = wSQL & "     , a.���r���[���� "
wSQL = wSQL & "     , b.���[�J�[�R�[�h "
wSQL = wSQL & "     , b.�J�e�S���[�R�[�h "
wSQL = wSQL & "     , b.���i�R�[�h "
wSQL = wSQL & "     , b.���i�� "
wSQL = wSQL & "     , b.�戵���~�� "
wSQL = wSQL & "     , b.������ "
wSQL = wSQL & "     , b.�p�ԓ� "
wSQL = wSQL & "     , b.B�i�t���O "
wSQL = wSQL & "     , b.ASK���i�t���O "
wSQL = wSQL & "     , b.�󏭐��� "
wSQL = wSQL & "     , b.�����萔�� "
wSQL = wSQL & "     , b.������󒍍ϐ��� "
wSQL = wSQL & "     , b.�Z�b�g���i�t���O "
wSQL = wSQL & "     , b.���[�J�[�������敪 "
wSQL = wSQL & "     , b.���ח\�薢��t���O "
wSQL = wSQL & "     , b.Web�[����\���t���O "
wSQL = wSQL & "     , c.�F "
wSQL = wSQL & "     , c.�K�i "
wSQL = wSQL & "     , c.B�i�����\���� "
wSQL = wSQL & "     , c.�����\���� "
wSQL = wSQL & "     , c.�����\���ח\��� "
wSQL = wSQL & "     , c.���iID "
wSQL = wSQL & "     , CASE "
wSQL = wSQL & "         WHEN b.�����萔�� > b.������󒍍ϐ��� THEN b.������P�� "
wSQL = wSQL & "         WHEN b.B�i�t���O = 'Y' THEN b.B�i�P�� "
wSQL = wSQL & "       ELSE b.�̔��P�� "
wSQL = wSQL & "       END AS ���̔��P�� "
wSQL = wSQL & "     , d.���[�J�[�� "
wSQL = wSQL & "     , e.�J�e�S���[�� "
wSQL = wSQL & "     , g.��J�e�S���[�� "
wSQL = wSQL & "     , (SELECT COUNT(k.���i�R�[�h) "
wSQL = wSQL & "          FROM Web�F�K�i�ʍ݌� k WITH (NOLOCK) "
wSQL = wSQL & "         WHERE     k.���[�J�[�R�[�h = b.���[�J�[�R�[�h "
wSQL = wSQL & "               AND k.���i�R�[�h = b.���i�R�[�h "
wSQL = wSQL & "               AND (k.�F != '' OR k.�K�i != '') "
wSQL = wSQL & "               AND k.�I���� IS NULL "
wSQL = wSQL & "       ) AS �F�K�iCNT "
wSQL = wSQL & "FROM "
wSQL = wSQL & "    ���i���r���[�W�v             a WITH (NOLOCK) "
wSQL = wSQL & "      INNER JOIN Web���i         b WITH (NOLOCK) "
wSQL = wSQL & "        ON     b.���[�J�[�R�[�h     = a.���[�J�[�R�[�h "
wSQL = wSQL & "           AND b.���i�R�[�h         = a.���i�R�[�h "
wSQL = wSQL & "      INNER JOIN Web�F�K�i�ʍ݌� c WITH (NOLOCK) "
wSQL = wSQL & "        ON     c.���[�J�[�R�[�h     = b.���[�J�[�R�[�h "
wSQL = wSQL & "           AND c.���i�R�[�h         = b.���i�R�[�h "
wSQL = wSQL & "      INNER JOIN ���[�J�[        d WITH (NOLOCK) "
wSQL = wSQL & "        ON     d.���[�J�[�R�[�h     = b.���[�J�[�R�[�h "
wSQL = wSQL & "      INNER JOIN �J�e�S���[      e WITH (NOLOCK) "
wSQL = wSQL & "        ON     e.�J�e�S���[�R�[�h   = b.�J�e�S���[�R�[�h "
wSQL = wSQL & "      INNER JOIN ��J�e�S���[    g WITH (NOLOCK) "
wSQL = wSQL & "        ON     g.��J�e�S���[�R�[�h = a.��J�e�S���[�R�[�h "
wSQL = wSQL & "      LEFT JOIN ( SELECT 'Y' AS 'ShohinWebY' ) t1 "
wSQL = wSQL & "        ON     b.Web���i�t���O      = t1.ShohinWebY "
wSQL = wSQL & "      LEFT JOIN ( SELECT ''  AS 'Iro' )        t2 "
wSQL = wSQL & "        ON     c.�F               = t2.Iro "
wSQL = wSQL & "      LEFT JOIN ( SELECT ''  AS 'Kikaku' )     t3 "
wSQL = wSQL & "        ON     c.�K�i             = t3.Kikaku "
wSQL = wSQL & "WHERE "
wSQL = wSQL & "        t1.ShohinWebY IS NOT NULL "
wSQL = wSQL & "    AND t2.Iro        IS NOT NULL "
wSQL = wSQL & "    AND t3.Kikaku     IS NOT NULL "
wSQL = wSQL & "    AND a.��J�e�S���[�R�[�h = '" & LargeCategoryCd & "' "
wSQL = wSQL & "ORDER BY "
wSQL = wSQL & "      a.���r���[�]������ DESC "
wSQL = wSQL & "    , a.���r���[���� DESC "
wSQL = wSQL & "    , d.���[�J�[�� "
wSQL = wSQL & "    , b.���i�� "
' 2012/01/23 GV Mod End
' 2012/01/20 GV Mod End

'@@@@response.write(wSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic

wHTML = ""
wHTML = wHTML & "<!--  container START  -->" & vbNewLine
wHTML = wHTML & "  <div id='container'>" & vbNewLine
wHTML = wHTML & "    <h1>" & wLargeCategoryName & "</h1>" & vbNewLine
wHTML = wHTML & "    <!-- s_box TH START -->" & vbNewLine
wHTML = wHTML & "    <div id='s_box_th_rp'>" & vbNewLine
wHTML = wHTML & "      <div id='th_no'>����</div>" & vbNewLine
wHTML = wHTML & "      <div id='th_prod'>���[�J�[�@���i</div>" & vbNewLine
wHTML = wHTML & "      <div id='th_cat'>�J�e�S���[</div>" & vbNewLine
wHTML = wHTML & "      <div id='th_point'>���r���[�|�C���g</div>" & vbNewLine
wHTML = wHTML & "      <div id='th_stock'>�݌ɏ�</div>" & vbNewLine
wHTML = wHTML & "      <div id='th_cart'>�J�[�g</div>" & vbNewLine
wHTML = wHTML & "    </div>" & vbNewLine
wHTML = wHTML & "    <!-- s_box TH END -->" & vbNewLine

if RSv.EOF = true then
	wHTML = wHTML & "<p>���r���[�����e����Ă��܂���B</p>" & vbNewLine   '���i���r���[�f�[�^���S���Ȃ��ꍇ
	wHTML = wHTML & "</div>"
else

	vRank = 0    '���ʂ̃J�E���^

	Do Until RSv.EOF = true

		vPrice = FormatNumber(calcPrice(RSv("���̔��P��"), wSalesTaxRate),0)
		vPriceNoTax = FormatNumber(RSv("���̔��P��"),0)								'2014/03/19 GV add
		vItem = Server.URLEncode(RSv("���[�J�[�R�[�h") & "^" & RSv("���i�R�[�h") & "^" & "^")
		vRank = vRank + 1  '����

		'---- �����Ɗ�Ŕw�i�F��ύX
		if vRank Mod 2 <> 0 then
			vBGColor = "s_box1"
		else
			vBGColor = "s_box2"
		end if

		'---- �݌ɏ󋵕\���̂��߁A�I���`�F�b�N
		vProdTermFl = "N"
		if isNull(RSv("�戵���~��")) = false then		'�戵���~
			vProdTermFl = "Y"
		end if
		if isNull(RSv("�p�ԓ�")) = false AND RSv("�����\����") <= 0 then		'�p�Ԃō݌ɖ���
			vProdTermFl = "Y"
		end if
		if isNull(RSv("������")) = false then		'�������i
			vProdTermFl = "Y"
		end if
		if RSv("B�i�t���O") = "Y" AND RSv("B�i�����\����") <= 0 then    'B�i�ō݌ɂȂ�
			vProdTermFl = "Y"
		end if

		'---- �݌ɏ󋵍쐬
		if RSv("�F�K�iCNT") = 0 then
			vInventoryCd = GetInventoryStatus(RSv("���[�J�[�R�[�h"),RSv("���i�R�[�h"),RSv("�F"),RSv("�K�i"),RSv("�����\����"),RSv("�󏭐���"),RSv("�Z�b�g���i�t���O"),RSv("���[�J�[�������敪"),RSv("�����\���ח\���"),vProdTermFl)

			'---- �݌ɏ󋵁A�F���ŏI�Z�b�g
			call GetInventoryStatus2(RSv("�����\����"), RSv("Web�[����\���t���O"), RSv("���ח\�薢��t���O"), RSv("�p�ԓ�"), RSv("B�i�t���O"), RSv("B�i�����\����"), RSv("�����萔��"), RSv("������󒍍ϐ���"), vProdTermFl, vInventoryCd, vInventoryImage)

		end if

		wHTML = wHTML & "    <!-- s_box START -->" & vbNewLine
		wHTML = wHTML & "    <div class='" & vBGColor & "'>" & vbNewLine
		wHTML = wHTML & "      <div class='s_box_height'>" & vbNewLine
		wHTML = wHTML & "        <div class='rp_no'>" & vbNewLine

		'---- 1�`3�ʂ͉����\��
		if vRank <= 3 then
			wHTML = wHTML & "          <div class='crown_pad'>" & vbNewLine
			wHTML = wHTML & "            <img height='30' src='images/ranking/ico_no" & vRank & "crown.gif' alt='' width='41'>" & vbNewLine
			wHTML = wHTML & "          </div>" & vbNewLine
		'---- 4�`25�ʂ͏��ʕ\��
		else
			wHTML = wHTML & "          " & vRank & vbNewLine
		end if

		wHTML = wHTML & "        </div>" & vbNewLine
		wHTML = wHTML & "        <div class='rp_prod'>" & vbNewLine
		wHTML = wHTML & "          <a href='ProductDetail.asp?Item=" & vItem & "'>" & vbNewLine

		'--- ���[�J�[���{���i����������2�s�ɂȂ�ꍇ��"..."�ŏȗ�
		vMakerProduct = RSv("���[�J�[��") & " " & RSv("���i��")
		if Len(vMakerProduct) > 33 then
			vProductName = Left(RSv("���i��"), 30-Len(RSv("���[�J�[��"))) &  "..."
		else
			vProductName = RSv("���i��")
		end if

		wHTML = wHTML & "            <span class='txt_maker'>" & RSv("���[�J�[��") & "</span>&nbsp;" & vbNewLine
		wHTML = wHTML & "            <span class='txt_product'>" & vProductName & "</span><br>" & vbNewLine
		wHTML = wHTML & "            <span class='txt_price_h'>�Ռ�����		�F</span>"  & vbNewLine
		wHTML = wHTML & "            <span class='txt_price_d'>"

		'---- ASK���i��ASK�\����<a>�����q�ɂł��Ȃ��̂Ń����N�͂Ȃ�
		if RSv("ASK���i�t���O") = "Y" then
			wHTML = wHTML & "ASK"
		else
'2014/03/19 GV mod start ---->
'			wHTML = wHTML & FormatNumber(vPrice,0) & "�~(�ō�)"
			wHTML = wHTML & FormatNumber(vPriceNoTax,0) & "�~(�Ŕ�)</span>"
			wHTML = wHTML & "�@<span class='txt_price_t'>(�ō�&nbsp;" & FormatNumber(vPrice,0) & "�~)</span>"
		end if

'		wHTML = wHTML & "</span>�@" & vbNewLine
'2014/03/19 GV mod end   <----
		wHTML = wHTML & "          </a>" & vbNewLine
		wHTML = wHTML & "        </div>" & vbNewLine
		wHTML = wHTML & "        <div class='rp_cat'>" & vbNewLine
		wHTML = wHTML & "          <a href='SearchList.asp?i_type=c&s_category_cd=" & RSv("�J�e�S���[�R�[�h")  & "'><strong>" & RSv("�J�e�S���[��") & "</strong></a>" & vbNewLine
		wHTML = wHTML & "        </div>" & vbNewLine
		wHTML = wHTML & "        <div class='rp_point'>" & vbNewLine
		wHTML = wHTML & "          <div class='pt8'>" & vbNewLine
' 2012/01/23 GV Mod Start
'		wHTML = wHTML & "            " & CreateReviewImg(RSv("���r���[�R�����g��"),RSv("���r���[�]�����v")) &  "<br>" & vbNewLine
		wHTML = wHTML & "            " & CreateReviewImg(RSv("���r���[����"), RSv("���r���[�]������")) &  "<br>" & vbNewLine
' 2012/01/23 GV Mod End
		wHTML = wHTML & "          </div>" & vbNewLine
		wHTML = wHTML & "        </div>" & vbNewLine
		wHTML = wHTML & "        <div class='rp_stock'>" & vbNewLine
		wHTML = wHTML & "          <div class='pt12'>" & vbNewLine

		if RSv("�F�K�iCNT") = 0 then
			wHTML = wHTML & "            <img height='10' src='images/" & vInventoryImage & "' width='10' alt=''> " & vInventoryCd & vbNewLine
		end if

		wHTML = wHTML & "          </div>" & vbNewLine
		wHTML = wHTML & "        </div>" & vbNewLine
		wHTML = wHTML & "        <div class='rp_cart'>" & vbNewLine
		wHTML = wHTML & "        <form name='f_item' method='post' action='OrderPreInsert.asp' onSubmit='return order_onClick(this);'>" & vbNewLine

		'---- �F�K�i�Ȃ�
		if RSv("�F�K�iCNT") = 0 then
			if vProdTermFl = "Y" then
				wHTML = wHTML & "            <img src='images/icon_sold.gif' alt='����'><br>" & vbNewLine
			else
				wHTML = wHTML & "            <input type='hidden' name='qt' value='1'>" & vbNewLine
				wHTML = wHTML & "            <input type='hidden' name='maker_cd' value='" & RSv("���[�J�[�R�[�h") & "'>" & vbNewLine
				wHTML = wHTML & "            <input type='hidden' name='product_cd' value='" & RSv("���i�R�[�h") & "'>" & vbNewLine
				wHTML = wHTML & "            <input type='hidden' name='category_cd' value='" & RSv("�J�e�S���[�R�[�h") & "'>" & vbNewLine
				wHTML = wHTML & "            <input type='image' src='images/btn_cart.png' style='vertical-align:middle' alt='�J�[�g��'><br>" & vbNewLine
			end if

		'----�F�K�i����
		else
			if vProdTermFl = "Y" then
				wHTML = wHTML & "            <img src='images/icon_sold.gif' alt='����'><br>" & vbNewLine
			else
				wHTML = wHTML & "            <input type='hidden' name='qt' value='0'>" & vbNewLine
				wHTML = wHTML & "            <a href='ProductDetail.asp?Item=" & vItem & "'><img src='images/btn_detail.png' alt='�ڍׂ�����'></a><br>" & vbNewLine
			end if
		end if

		wHTML = wHTML & "              ���iID:" & RSv("���iID") & vbNewLine

		wHTML = wHTML & "        </form>" & vbNewLine
		wHTML = wHTML & "        </div>" & vbNewLine
		wHTML = wHTML & "      </div>" & vbNewLine
		wHTML = wHTML & "    </div>" & vbNewLine
		wHTML = wHTML & "    <!-- s_box END -->" & vbNewLine



		RSv.MoveNext

	Loop

	wHTML = wHTML & "  </div>" & vbNewLine
	wHTML = wHTML & "  <!-- container END -->" & vbNewLine

end if

wReviewPointHTML = wHTML

RSv.close

End Function

'========================================================================
'
'	Function	���i���r���[�摜�쐬
'
'========================================================================
'
' 2012/01/23 GV Mod Start
'Function CreateReviewImg(pRatingNum,pRatingSum)
Function CreateReviewImg(pRatingNum, pAvgRating)
' 2012/01/23 GV Mod End

Dim vAvgRating
Dim v1Cnt
Dim v0Cnt
Dim vHalfCnt
Dim vReview
Dim i

vReview = ""

if pRatingNum > 0 then
	'---- ���r���[�]���̕��ς��v�Z
' 2012/01/23 GV Mod Start
	vAvgRating = Round(pAvgRating, 1)
' 2012/01/23 GV Mod End
	v1Cnt = Fix(vAvgRating)
	if (vAvgRating - v1Cnt) >= 0.5 then
		vHalfCnt = 1
	else
		vHalfCnt = 0
	end if
	v0Cnt = 5 - v1Cnt - vHalfCnt

	'---- �����]���\��
	For i=1 to v1Cnt
		vReview = vReview & "<img src='images/review_icon10.png' alt=''>"
	Next
	if vHalfcnt = 1 then
		vReview = vReview & "<img src='images/review_icon05.png' alt=''>"
	end if
	For i=1 to v0Cnt
		vReview = vReview & "<img src='images/review_icon00.png' alt=''>"
	Next
end if

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
<title>���r���[�|�C���g�b�T�E���h�n�E�X</title>
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
      <li class="now">���r���[�|�C���g</li>
    </ul>
  </div></div></div>
  <h1 class="title">���r���[�|�C���g</h1>
-->

<!-- Mainpage START -->
<div id="ranking_key_main_flame">

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
        <img src="images/ranking/noc_btn_off.jpg" alt="" name="Image12" width="114" height="80" />
      </a>
    </div>
    -->
    <div class="top_menu_parts">
      <a href="RankingAccess.asp?RankType=<%=Server.URLEncode("�F�B�ɂ�����")%>" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image13','','images/ranking/rtaf_btn_on.jpg',1)"><img src="images/ranking/rtaf_btn_off.jpg" alt="" name="Image13" width="114" height="80" />
      </a>
    </div>
    <div class="top_menu_parts">
      <a href="RankingAccess.asp?RankType=<%=Server.URLEncode("�~�������̃��X�g")%>" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image14','','images/ranking/wl_btn_on.jpg',1)"><img src="images/ranking/wl_btn_off.jpg" alt="" name="Image14" width="114" height="80" />
      </a>
    </div>
    <div class="top_menu_parts">
      <a href="RankingReview.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image16','','images/ranking/nor_btn_on.jpg',1)"><img src="images/ranking/nor_btn_off.jpg" alt="" name="Image16" width="114" height="80" />
      </a>
    </div>
    <div class="top_menu_parts">
      <a href="RankingReviewPoint.asp">
      <img src="images/ranking/rr_btn_on.jpg" alt="" name="Image17" width="113" height="80" />
      </a>
    </div>
  </div>
<!-- ��J�e�S���[�ꗗ -->
<%=wLargeCategoryHTML%>

<!-- ���r���[�|�C���gTOP 25 �ꗗ -->
<%=wReviewPointHTML%>

</div>
<!-- Mainpage END -->

  </div>
  <div id="globalSide">
    <!--#include file="../Navi/NaviSide.inc"-->
  </div>
</div>
<!--#include file="../Navi/Navibottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>