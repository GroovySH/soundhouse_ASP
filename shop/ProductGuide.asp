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
<!--#include file="../common/SearchListCommon.inc"-->

<%
'========================================================================
'
'	�������ߏ��i�y�[�W��������
'
'	�X�V����
'2005/02/16 �����萔�ʒP�����o�����̏��������@�����萔�ʁ�0��ǉ�
'2005/02/21 �V���i�̒��o�����ɔ��������g�p����悤�ύX
'2005/03/02 �^�C�g���ɂ��J�e�S���[�R�[�h��\������悤�ɕύX
'2005/03/15 �V���[�Y���i�\���ǉ�
'2005/03/18 ASK���i�P���\���ύX
'2005/07/21 �V���[�Y���ו\�����烁�[�J�[�����폜�i�V���[�Y���Ń��[�J�����\������Ă��邽��)
'2005/08/10 �F�K�i������ꍇ�́A�V���[�Y�Ɠ����`���̕\�����s���
'2005/09/07 ���i���o���T�u�J�e�S���[���l��
'2005/09/08 ���C���J�e�S���[���i���ɁA�T�u�J�e�S���[���i����ɕ\������悤�ύX
'2005/11/01 �����EMovie�|�b�v�A�b�v�y�[�W�ւ̃����N�ǉ�
'2006/01/10 �����A����̃����N��http���܂܂�Ă���ꍇ�͊O�������N�Ƃ���
'2006/01/23 �V���[�Y�\�����̃\�[�g����̔��P���ɕύX�A�V���[�Y���i���ɕ\��
'2006/04/05 �p�������ǉ�
'2006/12/08 �^�C�g���ύX
'2007/01/22 �݌ɏ󋵕\���ǉ�
'2007/03/15 �p�����[�^�ɑ΂���ReplaceInput��ǉ�
'2007/05/30 �F�K�i���菤�i�Ή�
'2008/10/23 �F�K�i�Ȃ��������ߏ��i�������\�Ȃ̂ɕ\������Ȃ������s��Ή�
'2008/10/23 (�ύX�˗�#503)�����萔�ʂ̕\�������̂悤�ɕύX
'						4�ȉ� ���s�ǂ���/5-9 ����5��/10-14 ����10��/15-19 ����15��/20�ȏ�A����20��
'2008/12/24 �݌ɏ󋵃Z�b�g�֐���
'2010/02/12 hn ASK���i�p�����[�^��Server.URLEncode���s�Ȃ�
'2010/11/10 an �V���[�Y���i�̃\�[�g�����C�����\�����w�肪�Ȃ��Ƃ��ɕ\��������Ȃ��悤��
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2011/10/19 hn 1063 ASK�\�����@�ύX
'2012/01/19 GV �f�[�^�擾 SELECT���� LAC�N�G���[�Ă�K�p
'2012/07/14 nt ���j���[�A���p�Ƀf�[�^�擾 SELECT�������asp��ʏo�͂��C��
'2012/07/23 nt ���݂��Ȃ��J�e�S���[�R�[�h�w�莞�̃G���[��ʃ��_�C���N�g��ǉ�
'
'========================================================================

On Error Resume Next

Dim CategoryCd

Dim wCategoryName
Dim wMidCategoryCd
Dim wMidCategoryName
Dim wLargeCategoryCd
Dim wLargeCategoryName
Dim wCategoryComment

Dim wSalesTaxRate
Dim wPrice

Dim wItemChar1
Dim wItemChar2
Dim wItemNum1
Dim wItemNum2
Dim wItemDate1
Dim wItemDate2

Dim Connection
Dim RS

Dim wRedirectURL
Dim wProductList
Dim wProductList2
Dim wSaleItemHTML
Dim wNewItemHTML
Dim wKanrenCtegoryLinkHTML

Dim wSQL
Dim wHTML
Dim wMSG
Dim wErrDesc   '2011/08/01 an add
Dim wNoData '2012/07/23 nt add

'========================================================================

Response.buffer = true

'---- ���M�f�[�^�[�̎��o��
CategoryCd = ReplaceInput(Request("CategoryCd"))
Response.Status="301 Moved Permanently" 
Response.AddHeader "Location", "http://www.soundhouse.co.jp/products/guide/?s_category_cd=" & CategoryCd

'---- Execute main
call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "ProductGuide.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc)
end if                                           '2011/08/01 an add e

call close_db()

'2012/07/23 nt add
If wNoData = "Y" Or Err.Description <> "" Then
	Response.Redirect g_HTTP & "shop/Error.asp"
End If

'2012/07/23 nt del
'if Err.Description <> "" then
'	Response.Redirect g_HTTP & "shop/Error.asp"
'end if

if wRedirectURL <> "" then
	Response.Redirect wRedirectURL
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
Dim v_item

'---- ����ŗ���o��
call getCntlMst("����","����ŗ�","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)			'����ŗ�
wSalesTaxRate = Clng(wItemNum1)

'---- �J�e�S���[�����o��
wSQL = ""
' 2012/01/19 GV Mod Start
'wSQL = wSQL & "SELECT a.�J�e�S���[��"
'wSQL = wSQL & "     , a.�����߃J�e�S���[�R�����g"
'wSQL = wSQL & "     , a.�����߃J�e�S���[URL"
'wSQL = wSQL & "     , b.���J�e�S���[�R�[�h"
'wSQL = wSQL & "     , b.���J�e�S���[�����{��"
'wSQL = wSQL & "     , c.��J�e�S���[�R�[�h"
'wSQL = wSQL & "     , c.��J�e�S���[��"
'wSQL = wSQL & "  FROM �J�e�S���[ a WITH (NOLOCK)"
'wSQL = wSQL & "     , ���J�e�S���[ b WITH (NOLOCK)"
'wSQL = wSQL & "     , ��J�e�S���[ c WITH (NOLOCK)"
'wSQL = wSQL & " WHERE b.���J�e�S���[�R�[�h = a.���J�e�S���[�R�[�h"
'wSQL = wSQL & "   AND c.��J�e�S���[�R�[�h = b.��J�e�S���[�R�[�h"
'wSQL = wSQL & "   AND a.�J�e�S���[�R�[�h = '" & CategoryCd & "'"
wSQL = wSQL & "SELECT "
wSQL = wSQL & "      a.�J�e�S���[�� "
wSQL = wSQL & "    , a.�����߃J�e�S���[�R�����g "
wSQL = wSQL & "    , a.�����߃J�e�S���[URL "
wSQL = wSQL & "    , b.���J�e�S���[�R�[�h "
wSQL = wSQL & "    , b.���J�e�S���[�����{�� "
wSQL = wSQL & "    , c.��J�e�S���[�R�[�h "
wSQL = wSQL & "    , c.��J�e�S���[�� "
wSQL = wSQL & "FROM "
wSQL = wSQL & "    �J�e�S���[                a WITH (NOLOCK) "
wSQL = wSQL & "      INNER JOIN ���J�e�S���[ b WITH (NOLOCK) "
wSQL = wSQL & "        ON     b.���J�e�S���[�R�[�h = a.���J�e�S���[�R�[�h "
wSQL = wSQL & "      INNER JOIN ��J�e�S���[ c WITH (NOLOCK) "
wSQL = wSQL & "        ON     c.��J�e�S���[�R�[�h = b.��J�e�S���[�R�[�h "
wSQL = wSQL & "WHERE "
wSQL = wSQL & "        a.�J�e�S���[�R�[�h = '" & CategoryCd & "' "
' 2012/01/19 GV Mod End

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

wRedirectURL = ""

'2012/07/23 nt add
if RS.EOF = True then
	wNoData = "Y"
	RS.close	'2012/07/23 nt add
else
'if RS.EOF = false then	'2012/07/23 nt del
	if RS("�����߃J�e�S���[URL") <> "" then
		wRedirectURL = RS("�����߃J�e�S���[URL")
		exit function
	else
		if RS("�����߃J�e�S���[�R�����g") <> "" then
' 2012/07/18 nt Mod Start
			wCategoryComment = wCategoryComment & Replace(RS("�����߃J�e�S���[�R�����g"), vbNewLine, "<br>") & vbnewline
'			wCategoryComment = "            <table width='610' border='1' cellspacing='0' cellpadding='2' bordercolor='#999999' bordercolorlight='#999999' bordercolordark='#ffffff' >" & vbnewline
'			wCategoryComment = wCategoryComment & "              <tr align='center' valign='top'>" & vbnewline
'			wCategoryComment = wCategoryComment & "                <td align='left' bgcolor='#ffffee' class='honbun' >" & vbnewline
'			wCategoryComment = wCategoryComment & Replace(RS("�����߃J�e�S���[�R�����g"), vbNewLine, "<br>") & vbnewline
'			wCategoryComment = wCategoryComment & "                </td>" & vbnewline
'			wCategoryComment = wCategoryComment & "              </tr>" & vbnewline
'			wCategoryComment = wCategoryComment & "            </table>" & vbnewline
' 2012/07/18 nt Mod End
		else
			wCategoryComment = ""
		end if
	end if
	wCategoryName = RS("�J�e�S���[��")
	wMidCategoryCd = RS("���J�e�S���[�R�[�h")
	wMidCategoryName = RS("���J�e�S���[�����{��")
	wLargeCategoryCd = RS("��J�e�S���[�R�[�h")
	wLargeCategoryName = RS("��J�e�S���[��")

	RS.close	'2012/07/23 nt add

	'---- �����ߏ��i�����o��
	call CreateProductList()	'2012/07/23 nt add

	'---- �����ߏ��i�����o��2�i�V���[�Y���i)	'2005/03/14
	call CreateProductList2()	'2012/07/23 nt add

end if

'RS.close	'2012/07/23 nt del

'---- �����ߏ��i�����o��
'call CreateProductList()	'2012/07/23 nt del

'---- �����ߏ��i�����o��2�i�V���[�Y���i)	'2005/03/14
'call CreateProductList2()	'2012/07/23 nt del

'---- �Ռ������i�����o��
'call CreateSaleItemHTML	'2012/07/18 nt del

'---- �V���i�����o��
'Call CreateNewItemHTML	'2012/07/18 nt del

'---- �֘A�J�e�S���[�����N�쐬
'Call CreateCategoryLinkHTML()	'2012/07/18 nt del

End Function

'========================================================================
'
'	Function	�����ߏ��i���ҏW
'
'========================================================================
'
Function CreateProductList()

'---- ���iRecordset�쐬
wSQL = ""
' 2012/01/19 GV Mod Start
'wSQL = wSQL & "SELECT a.���[�J�[�R�[�h"
'wSQL = wSQL & "     , a.���i�R�[�h"
'wSQL = wSQL & "     , b.�F"
'wSQL = wSQL & "     , b.�K�i"
'wSQL = wSQL & "     , a.���i��"
'wSQL = wSQL & "     , a.���i�摜�t�@�C����_��"
'wSQL = wSQL & "     , a.�����ߏ��i�R�����g"
'wSQL = wSQL & "     , a.ASK���i�t���O"
'wSQL = wSQL & "     , a.�����t���O"
'wSQL = wSQL & "     , a.����URL"
'wSQL = wSQL & "     , a.����t���O"
'wSQL = wSQL & "     , a.����URL"
'wSQL = wSQL & "     , CASE"
'wSQL = wSQL & "         WHEN (a.�����萔�� > a.������󒍍ϐ��� AND a.�����萔�� > 0) THEN a.������P��"
'wSQL = wSQL & "         ELSE a.�̔��P��"
'wSQL = wSQL & "       END AS �̔��P��"
'wSQL = wSQL & "     , a.�����萔��"
'wSQL = wSQL & "     , a.������󒍍ϐ���"
'wSQL = wSQL & "     , CASE"
'wSQL = wSQL & "         WHEN (a.�����萔�� > a.������󒍍ϐ��� AND a.�����萔�� > 0) THEN 'Y'"
'wSQL = wSQL & "         ELSE 'N'"
'wSQL = wSQL & "       END AS ������P���t���O"
'wSQL = wSQL & "     , a.�󏭐���"
'wSQL = wSQL & "     , a.�Z�b�g���i�t���O"
'wSQL = wSQL & "     , a.���[�J�[�������敪"
'wSQL = wSQL & "     , a.Web�[����\���t���O"
'wSQL = wSQL & "     , a.�p�ԓ�"
'wSQL = wSQL & "     , a.B�i�t���O"
'wSQL = wSQL & "     , a.���ח\�薢��t���O"
'wSQL = wSQL & "     , b.�����\���ח\���"
'wSQL = wSQL & "     , b.�����\����"
'wSQL = wSQL & "     , b.B�i�����\����"
'wSQL = wSQL & "     , c.���[�J�[��"
'wSQL = wSQL & "  FROM Web���i a WITH (NOLOCK)"
'wSQL = wSQL & "     , Web�F�K�i�ʍ݌� b WITH (NOLOCK)"
'wSQL = wSQL & "     , ���[�J�[ c WITH (NOLOCK)"
'wSQL = wSQL & "     , ���i�J�e�S���[ d WITH (NOLOCK)"
'wSQL = wSQL & " WHERE b.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'wSQL = wSQL & "   AND b.���i�R�[�h = a.���i�R�[�h"
'wSQL = wSQL & "   AND c.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'wSQL = wSQL & "   AND d.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'wSQL = wSQL & "   AND d.���i�R�[�h = a.���i�R�[�h"
'wSQL = wSQL & "   AND d.�J�e�S���[�R�[�h = '" & CategoryCd & "'"
'wSQL = wSQL & "   AND a.�����ߏ��i�t���O = 'Y'"
'wSQL = wSQL & "   AND a.�戵���~�� IS NULL"
'wSQL = wSQL & "   AND ((a.�p�ԓ� IS NULL AND b.�I���� IS NULL) OR (a.�p�ԓ� IS NOT NULL AND b.�����\���� > 0)) "
'wSQL = wSQL & "   AND a.Web���i�t���O = 'Y'"
'wSQL = wSQL & "   AND b.�I���� IS NULL"
'wSQL = wSQL & " ORDER BY"
'wSQL = wSQL & "       d.�J�e�S���[�敪"
'wSQL = wSQL & "     , CASE"
'wSQL = wSQL & "       	WHEN a.�����ߏ��i�\���� = 0 THEN 99999"
'wSQL = wSQL & "       	ELSE a.�����ߏ��i�\����"
'wSQL = wSQL & "       END"
'wSQL = wSQL & "     , c.���[�J�[��"
'wSQL = wSQL & "     , a.���i��"
'wSQL = wSQL & "     , b.�F"
'wSQL = wSQL & "     , b.�K�i"
wSQL = wSQL & "SELECT "
wSQL = wSQL & "      a.���[�J�[�R�[�h "
wSQL = wSQL & "    , a.���i�R�[�h "
'wSQL = wSQL & "    , b.�F " 2012/07/17 natori del
'wSQL = wSQL & "    , b.�K�i " 2012/07/17 natori del
wSQL = wSQL & "    , a.���i�� "
wSQL = wSQL & "    , a.���i�摜�t�@�C����_�� "
wSQL = wSQL & "    , a.�����ߏ��i�R�����g "
'wSQL = wSQL & "    , a.ASK���i�t���O " 2012/07/17 natori del
'wSQL = wSQL & "    , a.�����t���O " 2012/07/17 natori del
'wSQL = wSQL & "    , a.����URL " 2012/07/17 natori del
'wSQL = wSQL & "    , a.����t���O " 2012/07/17 natori del
'wSQL = wSQL & "    , a.����URL " 2012/07/17 natori del
'wSQL = wSQL & "    , CASE "
'wSQL = wSQL & "        WHEN (a.�����萔�� > a.������󒍍ϐ��� AND a.�����萔�� > 0) THEN a.������P�� "
'wSQL = wSQL & "        ELSE a.�̔��P�� "
'wSQL = wSQL & "      END AS �̔��P�� " 2012/07/17 natori del
'wSQL = wSQL & "    , a.�����萔�� " 2012/07/17 natori del
'wSQL = wSQL & "    , a.������󒍍ϐ��� " 2012/07/17 natori del
'wSQL = wSQL & "    , CASE "
'wSQL = wSQL & "        WHEN (a.�����萔�� > a.������󒍍ϐ��� AND a.�����萔�� > 0) THEN 'Y' "
'wSQL = wSQL & "        ELSE 'N' "
'wSQL = wSQL & "      END AS ������P���t���O " 2012/07/17 natori del
'wSQL = wSQL & "    , a.�󏭐��� " 2012/07/17 natori del
'wSQL = wSQL & "    , a.�Z�b�g���i�t���O " 2012/07/17 natori del
'wSQL = wSQL & "    , a.���[�J�[�������敪 " 2012/07/17 natori del
'wSQL = wSQL & "    , a.Web�[����\���t���O " 2012/07/17 natori del
'wSQL = wSQL & "    , a.�p�ԓ� " 2012/07/17 natori del
'wSQL = wSQL & "    , a.B�i�t���O " 2012/07/17 natori del
'wSQL = wSQL & "    , a.���ח\�薢��t���O " 2012/07/17 natori del
'wSQL = wSQL & "    , b.�����\���ח\��� " 2012/07/17 natori del
'wSQL = wSQL & "    , b.�����\���� " 2012/07/17 natori del
'wSQL = wSQL & "    , b.B�i�����\���� " 2012/07/17 natori del
wSQL = wSQL & "    , c.���[�J�[�� "
wSQL = wSQL & "    , COUNT(b.�F) AS �F�K�i�� "
wSQL = wSQL & "FROM "
wSQL = wSQL & "    Web���i                      a WITH (NOLOCK) "
wSQL = wSQL & "      INNER JOIN Web�F�K�i�ʍ݌� b WITH (NOLOCK) "
wSQL = wSQL & "        ON     b.���[�J�[�R�[�h = a.���[�J�[�R�[�h "
wSQL = wSQL & "           AND b.���i�R�[�h     = a.���i�R�[�h "
wSQL = wSQL & "      INNER JOIN ���[�J�[        c WITH (NOLOCK) "
wSQL = wSQL & "        ON     c.���[�J�[�R�[�h = a.���[�J�[�R�[�h "
wSQL = wSQL & "      INNER JOIN ���i�J�e�S���[  d WITH (NOLOCK) "
wSQL = wSQL & "        ON     d.���[�J�[�R�[�h = a.���[�J�[�R�[�h "
wSQL = wSQL & "           AND d.���i�R�[�h     = a.���i�R�[�h "
wSQL = wSQL & "      LEFT JOIN ( SELECT 'Y' AS 'ShohinWebY' )   t1 "
wSQL = wSQL & "        ON     a.Web���i�t���O    = t1.ShohinWebY "
wSQL = wSQL & "      LEFT JOIN ( SELECT 'Y' AS 'RecommendY' )   t2 "
wSQL = wSQL & "        ON     a.�����ߏ��i�t���O = t2.RecommendY "
wSQL = wSQL & "WHERE "
wSQL = wSQL & "        t1.ShohinWebY  IS NOT NULL "
wSQL = wSQL & "    AND t2.RecommendY  IS NOT NULL "
wSQL = wSQL & "    AND a.�戵���~�� IS NULL "
wSQL = wSQL & "    AND (    (    a.�p�ԓ� IS NULL "
wSQL = wSQL & "              AND b.�I���� IS NULL) "
wSQL = wSQL & "         OR  (    a.�p�ԓ� IS NOT NULL "
wSQL = wSQL & "              AND b.�����\���� > 0)) "
wSQL = wSQL & "    AND b.�I���� IS NULL "
wSQL = wSQL & "    AND d.�J�e�S���[�R�[�h = '" & CategoryCd & "' "
'---- �F�E�K�i�Ȃǃf�[�^�s�v�̂��߁AGROUP BY���ǉ�(2012/7/17 natori add)
wSQL = wSQL & "GROUP BY "
wSQL = wSQL & "      a.���[�J�[�R�[�h "
wSQL = wSQL & "    , a.���i�R�[�h "
wSQL = wSQL & "    , a.���i�� "
wSQL = wSQL & "    , a.���i�摜�t�@�C����_�� "
wSQL = wSQL & "    , a.�����ߏ��i�R�����g "
wSQL = wSQL & "    , c.���[�J�[�� "
wSQL = wSQL & "    , d.�J�e�S���[�敪 "
wSQL = wSQL & "    , a.�����ߏ��i�\���� "
wSQL = wSQL & "ORDER BY "
wSQL = wSQL & "      d.�J�e�S���[�敪 "
wSQL = wSQL & "    , CASE "
wSQL = wSQL & "        WHEN a.�����ߏ��i�\���� = 0 THEN 99999 "
wSQL = wSQL & "        ELSE                             a.�����ߏ��i�\���� "
wSQL = wSQL & "      END "
wSQL = wSQL & "    , c.���[�J�[�� "
wSQL = wSQL & "    , a.���i�� "
'wSQL = wSQL & "    , b.�F " 2012/07/17 natori del
'wSQL = wSQL & "    , b.�K�i " 2012/07/17 natori del
' 2012/01/19 GV Mod End

'@@@@@@@@@@response.write(wSQL)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

wHTML = ""
'wHTML = wHTML & "<table border='0' cellspacing='1' cellpadding='0'>" & vbNewLine	2012/07/17 natori del

Do Until RS.EOF = true
	wHTML = wHTML & "<ul class='item_list gridtype productguide'>" & vbNewLine

	'---- �� ���i
	Call CreateProductHTML

	if RS.EOF = false then
	'---- ������ ���i
		call CreateProductHTML()

		if RS.EOF = false then
		'---- �����E ���i
			call CreateProductHTML()

			if RS.EOF = false then
			'---- �E ���i
				call CreateProductHTML()
			end if
		end if
	end if

	wHTML = wHTML & "</ul>"
Loop

wProductList = wHTML

RS.Close
End function

'========================================================================
'
'	Function	�ʏ��iHTML�쐬
'
'========================================================================
'
Function CreateProductHTML()
Dim vComment
Dim vOldProductCd
'2012/07/17 natori del Start
'Dim vIroKikakuFl
'Dim vWidth1
'Dim vWidth2
'Dim vItemCnt
'Dim vItemList()
'Dim vSoundMovie
'Dim vInventoryCd
'Dim vInventoryImage

'vIroKikakuFl = false
'vWidth1 = 110
'vWidth2 = 85

'if Trim(RS("�F")) <> "" OR Trim(RS("�K�i")) <> "" then
'	vIroKikakuFl = true
'	vWidth1 = 160
'	vWidth2 = 35
'end if

'wHTML = wHTML & "    <td>" & vbNewLine
'wHTML = wHTML & "      <table width='200' border='1' cellspacing='0' cellpadding='0' bordercolor='#999999' bordercolorlight='#999999' bordercolordark='#ffffff'>" & vbNewLine

'---- ���[�J�[���C���i�ԍ�
'wHTML = wHTML & "        <tr>" & vbNewLine
'wHTML = wHTML & "          <td height='39' colspan='2' bgcolor='#eeeeee'>" & vbNewLine
'wHTML = wHTML & "            <span class='honbun'><b>" & RS("���[�J�[��") & "</b></span> "
'if vIroKikakuFl = false then
'	wHTML = wHTML & "<a href='ProductDetail.asp?item=" & RS("���[�J�[�R�[�h") & "^" & Server.URLEncode(RS("���i�R�[�h")) & "' class='link'><b>" & RS("���i��") & "</b></a>" & vbNewLine
'else
'	wHTML = wHTML & "<span class='honbun'><b>" & RS("���i��") & "</b></span>" & vbNewLine
'end if
'wHTML = wHTML & "          </td>" & vbNewLine
'wHTML = wHTML & "        </tr>" & vbNewLine

'---- ���i�摜
'wHTML = wHTML & "        <tr align='center'>" & vbNewLine
'wHTML = wHTML & "          <td height='100' colspan='2'>" & vbNewLine
'if vIroKikakuFl = false then
'	wHTML = wHTML & "            <a href='ProductDetail.asp?item=" & RS("���[�J�[�R�[�h") & "^" & Server.URLEncode(RS("���i�R�[�h")) & "' class='link'>"
'end if

'if RS("���i�摜�t�@�C����_��") <> "" then
'	wHTML = wHTML & "<img src='prod_img/" & RS("���i�摜�t�@�C����_��") & "' width='198' height='99' border='0'>"
'else
'	wHTML = wHTML & "<img src='images/blank.gif' width='198' height='99' border='0'>"
'end if
'if vIroKikakuFl = false then
'	wHTML = wHTML & "</a>" & vbNewLine
'end if

'wHTML = wHTML & "          </td>" & vbNewLine
'wHTML = wHTML & "        </tr>" & vbNewLine
'2012/07/17 natori del End
'---- �����ߏ��i����
if Trim(RS("�����ߏ��i�R�����g")) <> "" Then
	vComment = Replace(RS("�����ߏ��i�R�����g"), vbNewLine, "<br>")
End If
if vComment = "" then
	vComment = "&nbsp;"
end if

wHTML = wHTML & " <li>" & vbNewLine
wHTML = wHTML & "  <div class='photo'><a href='ProductDetail.asp?item=" & RS("���[�J�[�R�[�h") & "^" & Server.URLEncode(RS("���i�R�[�h")) & "'>"
If RS("���i�摜�t�@�C����_��") <> "" Then
	wHTML = wHTML & "<img src='prod_img/" & RS("���i�摜�t�@�C����_��") & "' alt='" & Replace(RS("���[�J�[��") & " / " & RS("���i��"),"'","&#39;") & "' class='opover'>"
End If
wHTML = wHTML & "</a></div>" & vbNewLine
wHTML = wHTML & "  <ul class='detail'>" & vbNewLine
wHTML = wHTML & "   <li><strong>" & RS("���[�J�[��") & "</strong></li>" & vbNewLine
wHTML = wHTML & "   <li class='prod_name'><strong><a href='ProductDetail.asp?item=" & RS("���[�J�[�R�[�h") & "^" & Server.URLEncode(RS("���i�R�[�h")) & "'>" & RS("���i��") & "</a></strong></li>" & vbNewLine
wHTML = wHTML & "   <li>" & vComment & "</li>"
wHTML = wHTML & "  </ul>" & vbNewLine
wHTML = wHTML & "  <div class='other_detail'>" & vbNewLine
wHTML = wHTML & "   <a href='ProductDetail.asp?item=" & RS("���[�J�[�R�[�h") & "^" & Server.URLEncode(RS("���i�R�[�h")) & "'><img src='images/btn_detail.png' alt='�ڍ�' class='opover'></a>" & vbNewLine
If RS("�F�K�i��") = 1 Then
	wHTML = wHTML & "   <a href='OrderPreInsert.asp?maker_cd=" & RS("���[�J�[�R�[�h") & "&amp;product_cd=" & Server.URLEncode(RS("���i�R�[�h")) & "&amp;qt=1'><img src='images/btn_cart.png' alt='�J�[�g�ɓ����' class='opover'></a>" & vbNewLine
End If
wHTML = wHTML & "  </div>" & vbNewLine
wHTML = wHTML & " </li>" & vbNewLine

RS.MoveNext

'2012/07/17 natori del Start
'wHTML = wHTML & "        <tr align='left' valign='top'>" & vbNewLine
'wHTML = wHTML & "          <td height='75' colspan='2' class='honbun'>" & vbNewLine
'wHTML = wHTML & vComment & vbNewLine
'wHTML = wHTML & "          </td>" & vbNewLine
'wHTML = wHTML & "        </tr>" & vbNewLine

'----���������N
'vSoundMovie = ""
'if RS("�����t���O") = "Y" AND RS("����URL") <> "" then
'	vItemCnt = cf_unstring(RS("����URL"), vItemList, ",")
'	if vItemCnt > 1 then
'		vSoundMovie = vSoundMovie & "<a href='JavaScript:void(0);' onClick=""window.open('SoundMoviePopUp.asp?item=" & RS("���[�J�[�R�[�h") & "^" & Server.URLEncode(RS("���i�R�[�h")) & "','SoundMovie', 'width=201 height=200 scrollbars=false memubar=false toolbar=false statusbar=false personalbar=false locationbar=false');"" class='link'><img src='images/Shichou.gif' width='18' height='18' border='0' alt='��������'></a>&nbsp;"
'	else
'		if InStr(LCase(RS("����URL")), "http://") > 0 then
'			vSoundMovie = "<a href='" & RS("����URL") & "' target='_blank'><img src='images/Shichou.gif' width='18' height='18' border='0' alt='��������'></a>&nbsp;&nbsp;"
'		else
'			vSoundMovie = "<a href='" & g_HTTP & RS("����URL") & "' target='_blank'><img src='images/Shichou.gif' width='18' height='18' border='0' alt='��������'></a>&nbsp;&nbsp;"
'		end if
'	end if
'end if

'----���惊���N
'if RS("����t���O") = "Y" AND RS("����URL") <> "" then
'	vItemCnt = cf_unstring(RS("����URL"), vItemList, ",")
'	if vItemCnt > 1 then
'		vSoundMovie = vSoundMovie & "<a href='JavaScript:void(0);' onClick=""window.open('SoundMoviePopUp.asp?item=" & RS("���[�J�[�R�[�h") & "^" & Server.URLEncode(RS("���i�R�[�h")) & "','SoundMovie', 'width=201 height=200 scrollbars=false memubar=false toolbar=false statusbar=false personalbar=false locationbar=false');"" class='link'><img src='images/Movie.jpg' width='18' height='18' border='0' alt='���������'></a>&nbsp;"
'	else
'		if InStr(LCase(RS("����URL")), "http://") > 0 then
'			vSoundMovie = vSoundMovie & "<a href='" & RS("����URL") & "' target='_blank'><img src='images/Movie.jpg' width='18' height='18' border='0' alt='���������'></a>&nbsp;"
'		else
'			vSoundMovie = vSoundMovie & "<a href='" & g_HTTP & RS("����URL") & "' target='_blank'><img src='images/Movie.jpg' width='18' height='18' border='0' alt='���������'></a>&nbsp;"
'		end if
'	end if
'end if

'if vSoundMovie <> "" then
'	wHTML = wHTML & "        <tr align='left' valign='middle'>" & vbNewLine
'	wHTML = wHTML & "          <td height='25' colspan='2' class='honbun'>" & vbNewLine
'	wHTML = wHTML & "            �T���v���F " & vSoundMovie & vbNewLine
'	wHTML = wHTML & "          </td>" & vbNewLine
'	wHTML = wHTML & "        </tr>" & vbNewLine
'end if

'vOldProductCd = RS("���[�J�[�R�[�h") & "+" & RS("���i�R�[�h")

'---- ���ꏤ�i�I���܂ŌJ��Ԃ� (�F�K�i������ꍇ�̂݌J��Ԃ�)
'Do until vOldProductCd <> RS("���[�J�[�R�[�h") & "+" & RS("���i�R�[�h")
	'---- �F�K�i
'	wHTML = wHTML & "        <tr>" & vbNewLine
'	wHTML = wHTML & "          <td  height='25' align='right' nowrap colspan='2'>" & vbNewLine
'	if Trim(RS("�F")) <> "" OR Trim(RS("�K�i")) <> "" then
'		wHTML = wHTML & "            <a href='ProductDetail.asp?item=" & RS("���[�J�[�R�[�h") & "^" & Server.URLEncode(RS("���i�R�[�h")) & "^" & RS("�F") & "^" & RS("�K�i") & "' class='link'>"
'		if RS("�F") <> "" then
'			wHTML = wHTML & RS("�F") & " "
'		end if
'		if RS("�K�i") <> "" then
'			wHTML = wHTML & RS("�K�i")
'		end if
'		wHTML = wHTML & "</a>" & vbNewLine
'	end if

	'---- �P��
'	wPrice = calcPrice(RS("�̔��P��"), wSalesTaxRate)

'	if RS("ASK���i�t���O") = "Y" then
'2011/10/19 hn mod s
'		wHTML = wHTML & "         <span class='honbun'>�Ռ������F<a href='JavaScript:void(0);' onClick=""askWin=window.open('AskPrice.asp?MakerName=" & Server.URLEncode(RS("���[�J�[��")) & "&ProductName=" & Server.URLEncode(RS("���i��")) & "&Price=" & wPrice & "' ,'ask', 'width=250 height=80 scrollbars=false memubar=false toolbar=false statusbar=false personalbar=false locationbar=false');"" class='link'>ASK</a></span>" & vbNewLine

'		wHTML = wHTML & "            <span class='honbun'>�Ռ������F<a class='tip'>ASK<span>" & FormatNumber(wPrice,0) & "�~(�ō�)</span></a></span>" & vbNewLine
'2011/10/19 hn mod e

'	else
'		wHTML = wHTML & "            <span class='honbun'>�Ռ������F<b>" & FormatNumber(wPrice,0) & "�~(�ō�)</b></span>" & vbNewLine

'	end if

'	wHTML = wHTML & "          </td>" & vbNewLine
'	wHTML = wHTML & "        </tr>" & vbNewLine

'----- �݌ɏ�
'	vInventoryCd = GetInventoryStatus(RS("���[�J�[�R�[�h"),RS("���i�R�[�h"),RS("�F"),RS("�K�i"),RS("�����\����"),RS("�󏭐���"),RS("�Z�b�g���i�t���O"),RS("���[�J�[�������敪"),RS("�����\���ח\���"),"N")

	'---- �݌ɏ󋵁A�F���ŏI�Z�b�g
'	call GetInventoryStatus2(RS("�����\����"), RS("Web�[����\���t���O"), RS("���ח\�薢��t���O"), RS("�p�ԓ�"), RS("B�i�t���O"), RS("B�i�����\����"), RS("�����萔��"), RS("������󒍍ϐ���"), "N", vInventoryCd, vInventoryImage)

'	wHTML = wHTML & "        <tr>" & vbNewLine
'	wHTML = wHTML & "          <td width='" & vWidth1 & "' height='25' align='right' nowrap class='honbun'><img src='images/" & vInventoryImage & "' width=10 height=10> " & vInventoryCd & "</td>" & vbNewLine

	'---- �ڍ׃{�^���C�J�[�g�{�^��
'	wHTML = wHTML & "          <td width='" & vWidth2 & "' nowrap height='25' align='center' valign='middle'>" & vbNewLine
'	if vIroKikakuFl = false then
'		wHTML = wHTML & "            <a href='ProductDetail.asp?item=" & RS("���[�J�[�R�[�h") & "^" & Server.URLEncode(RS("���i�R�[�h")) & "^" & RS("�F") & "^" & RS("�K�i") & "'><img src='images/Shousai.gif' width='50' height='19' border='0' align='middle'></a>" & vbNewLine
'	end if
'	wHTML = wHTML & "            <a href='OrderPreInsert.asp?maker_cd=" & RS("���[�J�[�R�[�h") & "&product_cd=" & Server.URLEncode(RS("���i�R�[�h")) & "&iro=" & RS("�F") & "&kikaku=" & RS("�K�i") & "&qt=1'><img src='images/CartBlue.gif' width='30' height='19' border='0' align='middle'></a>" & vbNewLine
'	wHTML = wHTML & "          </td>" & vbNewLine
'	wHTML = wHTML & "        </tr>" & vbNewLine

'	RS.MoveNext
'	if RS.EOF = true then
'		Exit Do
'	end if
'Loop

'wHTML = wHTML & "      </table>" & vbNewLine
'wHTML = wHTML & "    </td>" & vbNewLine
'2012/07/17 natori del End

End function

'========================================================================
'
'	Function	�����ߏ��i���ҏW2 �i�V���[�Y���i)
'
'========================================================================
'
Function CreateProductList2()

'---- ���iRecordset�쐬
wSQL = ""
' 2012/01/19 GV Mod Start
'wSQL = wSQL & "SELECT a.���[�J�[�R�[�h"
'wSQL = wSQL & "     , a.���i�R�[�h"
'wSQL = wSQL & "     , a.���i��"
'wSQL = wSQL & "     , a.���i�摜�t�@�C����_��"
'wSQL = wSQL & "     , a.�����ߏ��i�R�����g"
'wSQL = wSQL & "     , a.ASK���i�t���O"
'wSQL = wSQL & "     , CASE"
'wSQL = wSQL & "         WHEN (a.�����萔�� > a.������󒍍ϐ��� AND a.�����萔�� > 0) THEN a.������P��"
'wSQL = wSQL & "         ELSE a.�̔��P��"
'wSQL = wSQL & "       END AS �̔��P��"
'wSQL = wSQL & "     , c.���[�J�[��"
'wSQL = wSQL & "     , e.�V���[�Y�R�[�h"
'wSQL = wSQL & "     , e.�V���[�Y��"
'wSQL = wSQL & "     , e.�V���[�Y�摜�t�@�C����"
'wSQL = wSQL & "     , e.�����߃V���[�Y���l"
'wSQL = wSQL & "     , e.�����߃V���[�Y�\����"
'wSQL = wSQL & "  FROM Web���i a WITH (NOLOCK)"
'wSQL = wSQL & "     , Web�F�K�i�ʍ݌� b WITH (NOLOCK)"
'wSQL = wSQL & "     , ���[�J�[ c WITH (NOLOCK)"
'wSQL = wSQL & "     , ���i�J�e�S���[ d WITH (NOLOCK)"
'wSQL = wSQL & "     , �V���[�Y e WITH (NOLOCK)"
'wSQL = wSQL & " WHERE b.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'wSQL = wSQL & "   AND b.���i�R�[�h = a.���i�R�[�h"
'wSQL = wSQL & "   AND c.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'wSQL = wSQL & "   AND d.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'wSQL = wSQL & "   AND d.���i�R�[�h = a.���i�R�[�h"
'wSQL = wSQL & "   AND e.�V���[�Y�R�[�h = a.�V���[�Y�R�[�h"
'wSQL = wSQL & "   AND d.�J�e�S���[�R�[�h = '" & CategoryCd & "'"
'wSQL = wSQL & "   AND e.�����߃V���[�Y�t���O = 'Y'"
'wSQL = wSQL & "   AND a.�戵���~�� IS NULL"
'wSQL = wSQL & "   AND ((a.�p�ԓ� IS NULL) OR (a.�p�ԓ� IS NOT NULL AND b.�����\���� > 0)) "
'wSQL = wSQL & "   AND a.Web���i�t���O = 'Y'"
'wSQL = wSQL & " ORDER BY"
''wSQL = wSQL & "       d.�J�e�S���[�敪"       '2010/11/10 an del
'wSQL = wSQL & "       e.�����߃V���[�Y�\����"
'wSQL = wSQL & "     , e.�V���[�Y�R�[�h"        '2010/11/10 an add
'wSQL = wSQL & "     , �̔��P��"
wSQL = wSQL & "SELECT "
wSQL = wSQL & "      a.���[�J�[�R�[�h "
wSQL = wSQL & "    , a.���i�R�[�h "
wSQL = wSQL & "    , a.���i�� "
wSQL = wSQL & "    , a.���i�摜�t�@�C����_�� "
wSQL = wSQL & "    , a.�����ߏ��i�R�����g "
wSQL = wSQL & "    , a.ASK���i�t���O "
wSQL = wSQL & "    , CASE "
wSQL = wSQL & "        WHEN (a.�����萔�� > a.������󒍍ϐ��� AND a.�����萔�� > 0) THEN a.������P�� "
wSQL = wSQL & "        ELSE a.�̔��P�� "
wSQL = wSQL & "      END AS �̔��P�� "
wSQL = wSQL & "    , c.���[�J�[�� "
wSQL = wSQL & "    , e.�V���[�Y�R�[�h "
wSQL = wSQL & "    , e.�V���[�Y�� "
wSQL = wSQL & "    , e.�V���[�Y�摜�t�@�C���� "
wSQL = wSQL & "    , e.�����߃V���[�Y���l "
wSQL = wSQL & "    , e.�����߃V���[�Y�\���� "
wSQL = wSQL & "FROM "
wSQL = wSQL & "    Web���i                      a WITH (NOLOCK) "
wSQL = wSQL & "      INNER JOIN Web�F�K�i�ʍ݌� b WITH (NOLOCK) "
wSQL = wSQL & "        ON     b.���[�J�[�R�[�h = a.���[�J�[�R�[�h "
wSQL = wSQL & "           AND b.���i�R�[�h     = a.���i�R�[�h "
wSQL = wSQL & "      INNER JOIN ���[�J�[        c WITH (NOLOCK) "
wSQL = wSQL & "        ON     c.���[�J�[�R�[�h = a.���[�J�[�R�[�h "
wSQL = wSQL & "      INNER JOIN ���i�J�e�S���[  d WITH (NOLOCK) "
wSQL = wSQL & "        ON     d.���[�J�[�R�[�h = a.���[�J�[�R�[�h "
wSQL = wSQL & "           AND d.���i�R�[�h     = a.���i�R�[�h "
wSQL = wSQL & "      INNER JOIN �V���[�Y        e WITH (NOLOCK) "
wSQL = wSQL & "        ON     e.�V���[�Y�R�[�h = a.�V���[�Y�R�[�h "
wSQL = wSQL & "      LEFT JOIN ( SELECT 'Y' AS 'ShohinWebY' )         t1 "
wSQL = wSQL & "        ON     a.Web���i�t���O    = t1.ShohinWebY "
wSQL = wSQL & "      LEFT JOIN ( SELECT 'Y' AS 'RecommendSeriesY' )   t2 "
wSQL = wSQL & "        ON     e.�����߃V���[�Y�t���O = t2.RecommendSeriesY "
wSQL = wSQL & "WHERE"
wSQL = wSQL & "        t1.ShohinWebY       IS NOT NULL "
wSQL = wSQL & "    AND t2.RecommendSeriesY IS NOT NULL "
wSQL = wSQL & "    AND a.�戵���~�� IS NULL "
wSQL = wSQL & "    AND (    (    a.�p�ԓ� IS NULL) "
wSQL = wSQL & "         OR  (    a.�p�ԓ� IS NOT NULL "
wSQL = wSQL & "              AND b.�����\���� > 0)) "
wSQL = wSQL & "    AND d.�J�e�S���[�R�[�h = '" & CategoryCd & "' "
wSQL = wSQL & "ORDER BY "
wSQL = wSQL & "      e.�����߃V���[�Y�\���� "
wSQL = wSQL & "    , e.�V���[�Y�R�[�h "
wSQL = wSQL & "    , �̔��P�� "
' 2012/01/19 GV Mod End

'@@@@@response.write(wSQL)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

wHTML = ""
'wHTML = wHTML & "<table border='0' cellspacing='1' cellpadding='0'>" & vbNewLine

Do Until RS.EOF = true
	wHTML = wHTML & "<ul class='item_list gridtype productguide'>" & vbNewLine

	'---- �� ���i
	Call CreateProduct2HTML

	if RS.EOF = false then
	'---- ������ ���i
		call CreateProduct2HTML()

		if RS.EOF = false then
		'---- �����E ���i
			call CreateProduct2HTML()

			if RS.EOF = false then
			'---- �E ���i
				call CreateProduct2HTML()
			end if
		end if
	end if

	wHTML = wHTML & "</ul>" & vbNewLine
Loop

wProductList2 = wHTML

RS.Close

End function

'========================================================================
'
'	Function	�ʏ��iHTML�쐬2 �i�V���[�Y���i�j
'
'========================================================================
'
Function CreateProduct2HTML()
Dim vComment
Dim vOldSeriesCd

'---- �����߃V���[�Y����
vComment = Replace(RS("�����߃V���[�Y���l"), vbNewLine, "<br>")
if vComment = "" then
	vComment = "&nbsp;"
end if

wHTML = wHTML & "<li>" & vbNewLine
wHTML = wHTML & " <div class='photo'><a href='SearchList.asp?i_type=se&amp;sSeriesCd=" & RS("�V���[�Y�R�[�h") & "'>"
If RS("�V���[�Y�摜�t�@�C����") <> "" Then
	wHTML = wHTML & "<img src='prod_img/" & RS("�V���[�Y�摜�t�@�C����") & "' alt='" & Replace(RS("���[�J�[��") & " / " & RS("�V���[�Y��"),"'","&#39;") & "'" & " class='opover'>"
End If
wHTML = wHTML & "</a></div>" & vbNewLine
wHTML = wHTML & " <ul class='detail'>" & vbNewLine
wHTML = wHTML & "  <li><strong>" & RS("���[�J�[��") & "</strong></li>" & vbNewLine
wHTML = wHTML & "  <li class='prod_name'><strong><a href='SearchList.asp?i_type=se&amp;sSeriesCd=" & RS("�V���[�Y�R�[�h") & "'>" & RS("�V���[�Y��") & "</a>" & vbNewLine & "</strong></li>" & vbNewLine
wHTML = wHTML & "  <li>" & vComment & "</li>" & vbNewLine
wHTML = wHTML & " </ul>" & vbNewLine
wHTML = wHTML & " <div class='other_detail'><a href='SearchList.asp?i_type=se&amp;sSeriesCd=" & RS("�V���[�Y�R�[�h") & "'><img src='images/btn_alllist.png' alt='�ꗗ' class='opover'></a></div>" & vbNewLine
wHTML = wHTML & "</li>" & vbNewLine

'---- ����V���[�Y�I���܂ŌJ��Ԃ��i����V���[�Y�W��j
vOldSeriesCd = RS("�V���[�Y�R�[�h")
Do until vOldSeriesCd <> RS("�V���[�Y�R�[�h")
	RS.MoveNext
	if RS.EOF = true then
		Exit Do
	end if
Loop

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
<meta name="description" content="<%=wCategoryComment%>">
<meta name="keywords" content="<%=wLargeCategoryName%>,<%=wMidCategoryName%>,<%=wCategoryName%>">
<title><%=wCategoryName%>�̂������ߏ��i�b�T�E���h�n�E�X</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/shop.css" type="text/css">
<link rel="stylesheet" href="Style/searchlist.css?20120811" type="text/css">
<link rel="stylesheet" href="style/ask.css" type="text/css">
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
				<li><span itemscope itemtype='http://data-vocabulary.org/Breadcrumb'><a href='LargeCategoryList.asp?LargeCategoryCd=<%=wLargeCategoryCd%>' itemprop='url'><span itemprop='title'><%=wLargeCategoryName%></span></a></span></li>
				<li><span itemscope itemtype='http://data-vocabulary.org/Breadcrumb'><a href='MidCategoryList.asp?MidCategoryCd=<%=wMidCategoryCd%>' itemprop='url'><span itemprop='title'><%=wMidCategoryName%></span></a></span></li>
				<li><span itemscope itemtype='http://data-vocabulary.org/Breadcrumb'><a href='SearchList.asp?i_type=c&amp;s_category_cd=<%=CategoryCd%>' itemprop='url'><span itemprop='title'><%=wCategoryName%></span></a></span></li>
				<li class="now"><span itemscope itemtype='http://data-vocabulary.org/Breadcrumb'><span itemprop='title'>�������ߏ��i</span></span></li>
			</ul>
			</div></div></div>

			<h1 class="title"><%=wCategoryName%> �̂������ߏ��i</h1>
			<p><%=wCategoryComment%></p>

			<!-- �����߃V���[�Y���i�ꗗ-->
			<%=wProductList2%>

			<!-- �����ߏ��i�ꗗ-->
			<%=wProductList%>
		<!--/#contents -->
		</div>
		<div id="globalSide">
			<!--#include file="../Navi/NaviSide.inc"-->
		<!--/#globalSide -->
		</div>
	<!--/#main -->
	</div>
	<!--#include file="../Navi/Navibottom.inc"-->
	<!--#include file="../Navi/NaviScript.inc"-->
	<div class="tooltip"><p>ASK</p></div>
	<script type="text/javascript" src="jslib/ask.js"></script>
	<script type="text/javascript" src="jslib/SearchList.js?20120321" charset="Shift_JIS"></script>
	<script type="text/javascript" src="../jslib/jquery.tinyscrollbar.min.js"></script>
	<script type="text/javascript">
		$(function(){
		    $('#scrollbar1').tinyscrollbar();
		});
	</script>
</body>
</html>