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
'	�x�X�g�Z���[���i(�J�e�S���[��)
'
'	�X�V����
'2007/04/13 �̔��P���Ɍ�����P�����l��
'2007/05/08 �n�b�J�[�Z�[�t�Ή�
'2007/05/25 �V���[�Y�Ή�
'2007/05/29 ���ʕ\����1�ʁA2�ʁB�B�B�ɕύX
'2009/04/30 �G���[����error.asp�ֈړ�
'2010/02/18 an ASK���i�p�����[�^��Server.URLEncode���s�Ȃ�
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2011/10/19 hn 1063 ASK�\�����@�ύX
'2012/01/19 GV �f�[�^�擾 SELECT���� LAC�N�G���[�Ă�K�p
'2012/01/20 GV �f�[�^�擾 SELECT�����甄�؏��i�e�[�u���̍ŐV���̃f�[�^�̂ݒ��o����������폜
'2012/07/11 ok ���j���[�A���V�f�U�C���ύX
'2012/07/23 nt ���݂��Ȃ��J�e�S���[�R�[�h�w�莞�̃G���[��ʃ��_�C���N�g��ǉ�
'2014/03/19 GV ����ő��łɔ���2�d�\���Ή�
'
'========================================================================

On Error Resume Next

Dim CategoryCd

Dim wCategoryName
Dim wMidCategoryCd
Dim wMidCategoryName
Dim wLargeCategoryCd
Dim wLargeCategoryName

Dim wSalesTaxRate
Dim wPrice
Dim wRank

Dim wItemChar1
Dim wItemChar2
Dim wItemNum1
Dim wItemNum2
Dim wItemDate1
Dim wItemDate2

Dim Connection
Dim RS

Dim wProductList

Dim wSQL
Dim wHTML
Dim wMSG
Dim wErrDesc   '2011/08/01 an add
Dim wNoData    '2012/07/23 nt add

'========================================================================

'---- ���M�f�[�^�[�̎��o��
CategoryCd = ReplaceInput(Request("CategoryCd"))
Response.Status="301 Moved Permanently" 
Response.AddHeader "Location", "http://www.soundhouse.co.jp/best_seller/category/" & CategoryCd

'---- Execute main
call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "BestSellerListByCategory.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
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

'---- �x�X�g�Z���[���i���o��
wSQL = ""
' 2012/01/19 GV Mod Start
'wSQL = wSQL & "SELECT a.���[�J�[�R�[�h"
'wSQL = wSQL & "     , a.���i�R�[�h"
'wSQL = wSQL & "     , '' AS �V���[�Y�R�[�h"
'wSQL = wSQL & "     , b.�F"
'wSQL = wSQL & "     , b.�K�i"
'wSQL = wSQL & "     , a.���i��"
'wSQL = wSQL & "     , a.���i�摜�t�@�C����_��"
'wSQL = wSQL & "     , a.�����ߏ��i�R�����g"
'wSQL = wSQL & "     , a.���i�T��Web"
'wSQL = wSQL & "     , a.ASK���i�t���O"
'wSQL = wSQL & "     , a.�����t���O"
'wSQL = wSQL & "     , a.����URL"
'wSQL = wSQL & "     , a.����t���O"
'wSQL = wSQL & "     , a.����URL"
'wSQL = wSQL & "     , CASE"
'wSQL = wSQL & "         WHEN (a.�����萔�� > a.������󒍍ϐ��� AND a.�����萔�� > 0) THEN a.������P��"
'wSQL = wSQL & "         ELSE a.�̔��P��"
'wSQL = wSQL & "       END AS �̔��P��"
'wSQL = wSQL & "     , c.���[�J�[��"
'wSQL = wSQL & "     , e.����"
'wSQL = wSQL & "     , f.�J�e�S���[��"
'wSQL = wSQL & "     , g.���J�e�S���[�R�[�h"
'wSQL = wSQL & "     , g.���J�e�S���[�����{��"
'wSQL = wSQL & "     , h.��J�e�S���[�R�[�h"
'wSQL = wSQL & "     , h.��J�e�S���[��"
'wSQL = wSQL & "  FROM Web���i a WITH (NOLOCK)"
'wSQL = wSQL & "     , Web�F�K�i�ʍ݌� b WITH (NOLOCK)"
'wSQL = wSQL & "     , ���[�J�[ c WITH (NOLOCK)"
'wSQL = wSQL & "     , ���i�J�e�S���[ d WITH (NOLOCK)"
'wSQL = wSQL & "     , ���؏��i e WITH (NOLOCK)"
'wSQL = wSQL & "     , �J�e�S���[ f WITH (NOLOCK)"
'wSQL = wSQL & "     , ���J�e�S���[ g WITH (NOLOCK)"
'wSQL = wSQL & "     , ��J�e�S���[ h WITH (NOLOCK)"
'wSQL = wSQL & " WHERE b.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'wSQL = wSQL & "   AND b.���i�R�[�h = a.���i�R�[�h"
'wSQL = wSQL & "   AND c.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'wSQL = wSQL & "   AND d.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'wSQL = wSQL & "   AND d.���i�R�[�h = a.���i�R�[�h"
'wSQL = wSQL & "   AND e.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'wSQL = wSQL & "   AND e.���i�R�[�h = a.���i�R�[�h"
'wSQL = wSQL & "   AND f.�J�e�S���[�R�[�h = d.�J�e�S���[�R�[�h"
'wSQL = wSQL & "   AND g.���J�e�S���[�R�[�h = f.���J�e�S���[�R�[�h"
'wSQL = wSQL & "   AND h.��J�e�S���[�R�[�h = g.��J�e�S���[�R�[�h"
'wSQL = wSQL & "   AND d.�J�e�S���[�R�[�h = '" & CategoryCd & "'"
'wSQL = wSQL & "   AND a.�戵���~�� IS NULL"
'wSQL = wSQL & "   AND ((a.�p�ԓ� IS NULL) OR (a.�p�ԓ� IS NOT NULL AND b.�����\���� > 0)) "
'wSQL = wSQL & "   AND a.Web���i�t���O = 'Y'"
'wSQL = wSQL & "   AND b.�I���� IS NULL"
'wSQL = wSQL & "   AND e.�N�� = (SELECT MAX(�N��) FROM ���؏��i)"
'
'wSQL = wSQL & " UNION "
'
'wSQL = wSQL & "SELECT a.���[�J�[�R�[�h"
'wSQL = wSQL & "     , '' AS ���i�R�[�h"
'wSQL = wSQL & "     , a.�V���[�Y�R�[�h"
'wSQL = wSQL & "     , '' AS �F"
'wSQL = wSQL & "     , '' AS �K�i"
'wSQL = wSQL & "     , a.�V���[�Y�� AS ���i��"
'wSQL = wSQL & "     , a.�V���[�Y�摜�t�@�C���� AS ���i�摜�t�@�C����_��"
'wSQL = wSQL & "     , a.�V���[�Y���l AS �����ߏ��i�R�����g"
'wSQL = wSQL & "     , '' AS ���i�T��Web"
'wSQL = wSQL & "     , '' AS ASK���i�t���O"
'wSQL = wSQL & "     , '' AS �����t���O"
'wSQL = wSQL & "     , '' AS ����URL"
'wSQL = wSQL & "     , '' AS ����t���O"
'wSQL = wSQL & "     , '' AS ����URL"
'wSQL = wSQL & "     , '' AS �̔��P��"
'wSQL = wSQL & "     , c.���[�J�[��"
'wSQL = wSQL & "     , e.����"
'wSQL = wSQL & "     , f.�J�e�S���[��"
'wSQL = wSQL & "     , g.���J�e�S���[�R�[�h"
'wSQL = wSQL & "     , g.���J�e�S���[�����{��"
'wSQL = wSQL & "     , h.��J�e�S���[�R�[�h"
'wSQL = wSQL & "     , h.��J�e�S���[��"
'wSQL = wSQL & "  FROM �V���[�Y a WITH (NOLOCK)"
'wSQL = wSQL & "     , ���[�J�[ c WITH (NOLOCK)"
'wSQL = wSQL & "     , ���؏��i e WITH (NOLOCK)"
'wSQL = wSQL & "     , �J�e�S���[ f WITH (NOLOCK)"
'wSQL = wSQL & "     , ���J�e�S���[ g WITH (NOLOCK)"
'wSQL = wSQL & "     , ��J�e�S���[ h WITH (NOLOCK)"
'wSQL = wSQL & " WHERE c.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'wSQL = wSQL & "   AND e.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'wSQL = wSQL & "   AND e.�V���[�Y�R�[�h = a.�V���[�Y�R�[�h"
'wSQL = wSQL & "   AND f.�J�e�S���[�R�[�h = a.�J�e�S���[�R�[�h"
'wSQL = wSQL & "   AND g.���J�e�S���[�R�[�h = f.���J�e�S���[�R�[�h"
'wSQL = wSQL & "   AND h.��J�e�S���[�R�[�h = g.��J�e�S���[�R�[�h"
'wSQL = wSQL & "   AND e.�J�e�S���[�R�[�h = '" & CategoryCd & "'"
'wSQL = wSQL & "   AND e.�N�� = (SELECT MAX(�N��) FROM ���؏��i)"
'
'wSQL = wSQL & " ORDER BY"
'wSQL = wSQL & "       e.����"
'wSQL = wSQL & "     , ���[�J�[��"
'wSQL = wSQL & "     , ���i��"
'wSQL = wSQL & "     , �F"
'wSQL = wSQL & "     , �K�i"
wSQL = wSQL & "SELECT "
wSQL = wSQL & "      a.���[�J�[�R�[�h "
wSQL = wSQL & "    , a.���i�R�[�h "
wSQL = wSQL & "    , '' AS �V���[�Y�R�[�h "
'wSQL = wSQL & "    , b.�F "  2012/07/11 nt del
'wSQL = wSQL & "    , b.�K�i "  2012/07/11 nt del
wSQL = wSQL & "    , a.���i�� "
wSQL = wSQL & "    , a.���i�摜�t�@�C����_�� "
wSQL = wSQL & "    , a.�����ߏ��i�R�����g "
wSQL = wSQL & "    , a.���i�T��Web "
wSQL = wSQL & "    , a.ASK���i�t���O "
wSQL = wSQL & "    , a.�����t���O "
wSQL = wSQL & "    , a.����URL "
wSQL = wSQL & "    , a.����t���O "
wSQL = wSQL & "    , a.����URL "
'wSQL = wSQL & "    , CASE "
'wSQL = wSQL & "        WHEN (a.�����萔�� > a.������󒍍ϐ��� AND a.�����萔�� > 0) THEN a.������P�� "
'wSQL = wSQL & "        ELSE a.�̔��P�� "
'wSQL = wSQL & "      END AS �̔��P�� "
wSQL = wSQL & "    , a.�̔��P�� "
wSQL = wSQL & "    , a.�����萔�� "
wSQL = wSQL & "    , a.������󒍍ϐ��� "
wSQL = wSQL & "    , a.������P�� "
wSQL = wSQL & "    , a.B�i�t���O "
wSQL = wSQL & "    , a.B�i�P�� "
wSQL = wSQL & "    , c.���[�J�[�� "
wSQL = wSQL & "    , e.���� "
wSQL = wSQL & "    , f.�J�e�S���[�� "
wSQL = wSQL & "    , g.���J�e�S���[�R�[�h "
wSQL = wSQL & "    , g.���J�e�S���[�����{�� "
wSQL = wSQL & "    , h.��J�e�S���[�R�[�h "
wSQL = wSQL & "    , h.��J�e�S���[�� "
wSQL = wSQL & "FROM "
wSQL = wSQL & "    Web���i                      a WITH (NOLOCK) "
wSQL = wSQL & "      INNER JOIN Web�F�K�i�ʍ݌� b WITH (NOLOCK) "
wSQL = wSQL & "        ON     b.���[�J�[�R�[�h     = a.���[�J�[�R�[�h "
wSQL = wSQL & "           AND b.���i�R�[�h         = a.���i�R�[�h "
wSQL = wSQL & "      INNER JOIN ���[�J�[        c WITH (NOLOCK) "
wSQL = wSQL & "        ON     c.���[�J�[�R�[�h     = a.���[�J�[�R�[�h "
wSQL = wSQL & "      INNER JOIN ���i�J�e�S���[  d WITH (NOLOCK) "
wSQL = wSQL & "        ON     d.���[�J�[�R�[�h     = a.���[�J�[�R�[�h "
wSQL = wSQL & "           AND d.���i�R�[�h         = a.���i�R�[�h "
wSQL = wSQL & "      INNER JOIN ���؏��i        e WITH (NOLOCK) "
wSQL = wSQL & "        ON     e.���[�J�[�R�[�h     = a.���[�J�[�R�[�h "
wSQL = wSQL & "           AND e.���i�R�[�h         = a.���i�R�[�h "
wSQL = wSQL & "           AND e.�J�e�S���[�R�[�h   = d.�J�e�S���[�R�[�h"	'2012/7/11 ok add
wSQL = wSQL & "      INNER JOIN �J�e�S���[      f WITH (NOLOCK) "
wSQL = wSQL & "        ON     f.�J�e�S���[�R�[�h   = d.�J�e�S���[�R�[�h "
wSQL = wSQL & "      INNER JOIN ���J�e�S���[    g WITH (NOLOCK) "
wSQL = wSQL & "        ON     g.���J�e�S���[�R�[�h = f.���J�e�S���[�R�[�h "
wSQL = wSQL & "      INNER JOIN ��J�e�S���[    h WITH (NOLOCK) "
wSQL = wSQL & "        ON     h.��J�e�S���[�R�[�h = g.��J�e�S���[�R�[�h "
wSQL = wSQL & "      LEFT JOIN ( SELECT 'Y' AS 'ShohinWebY' )   t1 "
wSQL = wSQL & "        ON     a.Web���i�t���O    = t1.ShohinWebY "
wSQL = wSQL & "WHERE "
wSQL = wSQL & "        t1.ShohinWebY  IS NOT NULL "
wSQL = wSQL & "    AND a.�戵���~��   IS NULL "
wSQL = wSQL & "    AND (   (a.�p�ԓ� IS NULL) "
wSQL = wSQL & "         OR (    a.�p�ԓ� IS NOT NULL "
wSQL = wSQL & "             AND b.�����\���� > 0)) "
wSQL = wSQL & "    AND b.�I���� IS NULL "
'wSQL = wSQL & "    AND e.�N�� = (SELECT MAX(�N��) FROM ���؏��i) "				' 2012/01/20 GV Del
wSQL = wSQL & "    AND d.�J�e�S���[�R�[�h = '" & CategoryCd & "' "

'---- �F�E�K�i�f�[�^�s�v�̂��߁AGROUP BY���ǉ�(2012/07/11 nt add)
wSQL = wSQL & "GROUP BY "
wSQL = wSQL & "           a.���[�J�[�R�[�h "
wSQL = wSQL & "         , a.���i�R�[�h "
wSQL = wSQL & "         , a.���i�� "
wSQL = wSQL & "         , a.���i�摜�t�@�C����_�� "
wSQL = wSQL & "         , a.�����ߏ��i�R�����g "
wSQL = wSQL & "         , a.���i�T��Web "
wSQL = wSQL & "         , a.ASK���i�t���O "
wSQL = wSQL & "         , a.�����t���O "
wSQL = wSQL & "         , a.����URL "
wSQL = wSQL & "         , a.����t���O "
wSQL = wSQL & "         , a.����URL "
wSQL = wSQL & "         , a.�����萔��"
wSQL = wSQL & "         , a.������󒍍ϐ���"
wSQL = wSQL & "         , a.������P�� "
wSQL = wSQL & "         , a.�̔��P�� "
wSQL = wSQL & "         , a.B�i�t���O "
wSQL = wSQL & "         , a.B�i�P�� "
wSQL = wSQL & "         , c.���[�J�[�� "
wSQL = wSQL & "         , e.���� "
wSQL = wSQL & "         , f.�J�e�S���[�� "
wSQL = wSQL & "         , g.���J�e�S���[�R�[�h "
wSQL = wSQL & "         , g.���J�e�S���[�����{�� "
wSQL = wSQL & "         , h.��J�e�S���[�R�[�h "
wSQL = wSQL & "         , h.��J�e�S���[�� "

wSQL = wSQL & "UNION "

wSQL = wSQL & "SELECT "
wSQL = wSQL & "      a.���[�J�[�R�[�h "
wSQL = wSQL & "    , '' AS ���i�R�[�h "
wSQL = wSQL & "    , a.�V���[�Y�R�[�h "
'wSQL = wSQL & "    , '' AS �F "  2012/07/11 nt del
'wSQL = wSQL & "    , '' AS �K�i "  2012/07/11 nt del
wSQL = wSQL & "    , a.�V���[�Y�� AS ���i�� "
wSQL = wSQL & "    , a.�V���[�Y�摜�t�@�C���� AS ���i�摜�t�@�C����_�� "
wSQL = wSQL & "    , a.�V���[�Y���l AS �����ߏ��i�R�����g "
wSQL = wSQL & "    , '' AS ���i�T��Web "
wSQL = wSQL & "    , '' AS ASK���i�t���O "
wSQL = wSQL & "    , '' AS �����t���O "
wSQL = wSQL & "    , '' AS ����URL "
wSQL = wSQL & "    , '' AS ����t���O "
wSQL = wSQL & "    , '' AS ����URL "
wSQL = wSQL & "    , '' AS �̔��P�� "
wSQL = wSQL & "    , '' AS �����萔�� "
wSQL = wSQL & "    , '' AS ������󒍍ϐ��� "
wSQL = wSQL & "    , '' AS ������P�� "
wSQL = wSQL & "    , '' AS B�i�t���O "
wSQL = wSQL & "    , '' AS B�i�P�� "
wSQL = wSQL & "    , c.���[�J�[�� "
wSQL = wSQL & "    , e.���� "
wSQL = wSQL & "    , f.�J�e�S���[�� "
wSQL = wSQL & "    , g.���J�e�S���[�R�[�h "
wSQL = wSQL & "    , g.���J�e�S���[�����{�� "
wSQL = wSQL & "    , h.��J�e�S���[�R�[�h "
wSQL = wSQL & "    , h.��J�e�S���[�� "
wSQL = wSQL & "FROM "
wSQL = wSQL & "    �V���[�Y                  a WITH (NOLOCK) "
wSQL = wSQL & "      INNER JOIN ���[�J�[     c WITH (NOLOCK) "
wSQL = wSQL & "        ON     c.���[�J�[�R�[�h     = a.���[�J�[�R�[�h "
wSQL = wSQL & "      INNER JOIN ���؏��i     e WITH (NOLOCK) "
wSQL = wSQL & "        ON     e.���[�J�[�R�[�h     = a.���[�J�[�R�[�h "
wSQL = wSQL & "           AND e.�V���[�Y�R�[�h     = a.�V���[�Y�R�[�h "
wSQL = wSQL & "      INNER JOIN �J�e�S���[   f WITH (NOLOCK) "
wSQL = wSQL & "        ON     f.�J�e�S���[�R�[�h   = a.�J�e�S���[�R�[�h "
wSQL = wSQL & "      INNER JOIN ���J�e�S���[ g WITH (NOLOCK) "
wSQL = wSQL & "        ON     g.���J�e�S���[�R�[�h = f.���J�e�S���[�R�[�h "
wSQL = wSQL & "      INNER JOIN ��J�e�S���[ h WITH (NOLOCK) "
wSQL = wSQL & "        ON     h.��J�e�S���[�R�[�h = g.��J�e�S���[�R�[�h "
wSQL = wSQL & "WHERE "
'wSQL = wSQL & "        e.�N�� = (SELECT MAX(�N��) FROM ���؏��i) "				' 2012/01/20 GV Del
wSQL = wSQL & "        e.�J�e�S���[�R�[�h = '" & CategoryCd & "' "

'---- �F�E�K�i�f�[�^�s�v�̂��߁AGROUP BY���ǉ�(2012/07/11 nt add)
wSQL = wSQL & "GROUP BY "
wSQL = wSQL & "           a.���[�J�[�R�[�h "
wSQL = wSQL & "         , a.�V���[�Y�R�[�h  "
wSQL = wSQL & "         , a.�V���[�Y��"
wSQL = wSQL & "         , a.�V���[�Y�摜�t�@�C����"
wSQL = wSQL & "         , a.�V���[�Y���l"
wSQL = wSQL & "         , c.���[�J�[�� "
wSQL = wSQL & "         , e.���� "
wSQL = wSQL & "         , f.�J�e�S���[�� "
wSQL = wSQL & "         , g.���J�e�S���[�R�[�h "
wSQL = wSQL & "         , g.���J�e�S���[�����{�� "
wSQL = wSQL & "         , h.��J�e�S���[�R�[�h "
wSQL = wSQL & "         , h.��J�e�S���[�� "

wSQL = wSQL & "ORDER BY "
wSQL = wSQL & "      e.���� "
wSQL = wSQL & "    , ���[�J�[�� "
wSQL = wSQL & "    , ���i�� "
'wSQL = wSQL & "    , �F " 2012/07/11 nt del
'wSQL = wSQL & "    , �K�i " 2012/07/11 nt del
' 2012/01/19 GV Mod End

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

'@@@@@		response.write(wSQL)

'2012/07/23 nt add
if RS.EOF = True then
	wNoData = "Y"
else
'if RS.EOF = false then 2012/07/23 nt del
	wCategoryName = RS("�J�e�S���[��")
	wMidCategoryCd = RS("���J�e�S���[�R�[�h")
	wMidCategoryName = RS("���J�e�S���[�����{��")
	wLargeCategoryCd = RS("��J�e�S���[�R�[�h")
	wLargeCategoryName = RS("��J�e�S���[��")
end if

wRank = 0
wHTML = ""
wHTML = wHTML & "    <ul class='item_list listtype bestseller'>" & vbNewLine

Do Until RS.EOF = true
'2012/07/23 ok Mod Start
'	wHTML = wHTML & "  <tr valign='top'>" & vbNewLine

	'---- ���i
	call CreateProductHTML()

'	if RS.EOF = false then
	'---- ���� ���i
'		call CreateProductHTML()

'		if RS.EOF = false then
		'---- �E ���i
'			call CreateProductHTML()
'		end if
'	end if

'	wHTML = wHTML & "  </tr>" & vbNewLine
	RS.MoveNext
'	if RS.EOF = true then
'		Exit Do
'	end if
Loop

'wHTML = wHTML & "</table>" & vbNewLine
wHTML = wHTML & "    </ul>" & vbNewLine
'2012/07/23 ok Mod End
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
Dim vIroKikakuFl
Dim vWidth1
Dim vWidth2
Dim vItemCnt
Dim vItemList()
Dim vSoundMovie

wRank = wRank + 1

'2012/07/11 ok Mod Start
If wRank < 4 Then
	wHTML = wHTML & "      <li class='toprank'>" & vbNewLine
Else
	wHTML = wHTML & "      <li>" & vbNewLine
End If

wPrice = calcPrice(RS("�̔��P��"), wSalesTaxRate)

wHTML = wHTML & "        <ul class='detail'>" & vbNewLine
wHTML = wHTML & "          <li class='rank'><img src='../top_images/ranking/rank" & Right("0" & Cstr(wRank),2) & ".png' alt='" & wRank & "��'></li>" & vbNewLine

'---- ���[�J�[���C���i�ԍ��C�摜
wHTML = wHTML & "          <li class='prod_name'><strong>" & RS("���[�J�[��") & " / <a href='ProductDetail.asp?item=" & RS("���[�J�[�R�[�h") & "^" & Server.URLEncode(RS("���i�R�[�h")) & "'>" & RS("���i��") & "</a></strong></li>" & vbNewLine

'wHTML = wHTML & "          <li>�̔����i�F" & FormatNumber(wPrice,0) & "�~�i�ō��j</li>" & vbNewLine

wHTML = wHTML & "                        <li class='price'>"
If RS("ASK���i�t���O") <> "Y" Then
	'---- B�i�P��
	If RS("B�i�t���O") = "Y" Then
		wPrice = calcPrice(RS("B�i�P��"), wSalesTaxRate)
'2014/03/19 GV mod start ---->
'		wHTML = wHTML & "�킯����i�����F<span>" & FormatNumber(wPrice,0) & "�~</span>(�ō�)"
		wHTML = wHTML & "�킯����i�����F<span>" & FormatNumber(RS("B�i�P��"),0) & "�~</span>(�Ŕ�)<br>"
		wHTML = wHTML & "(�ō�&nbsp;<span>" & FormatNumber(wPrice,0) & "�~</span>)"
'2014/03/19 GV mod end   <----
	'---- ������P��
	ElseIf RS("�����萔��") > RS("������󒍍ϐ���") AND RS("�����萔��") > 0 Then
		wPrice = calcPrice(RS("������P��"), wSalesTaxRate)
'2014/03/19 GV mod start ---->
'		wHTML = wHTML & "��������F<span>" & FormatNumber(wPrice,0) & "�~</span>(�ō�)"
		wHTML = wHTML & "��������F<span>" & FormatNumber(RS("������P��"),0) & "�~</span>(�Ŕ�)<br>"
		wHTML = wHTML & "(�ō�&nbsp;<span>" & FormatNumber(wPrice,0) & "�~</span>)"
'2014/03/19 GV mod end   <----
	'---- �ʏ폤�i
	Else
'2014/03/19 GV mod start ---->
'		wHTML = wHTML & "�Ռ������F<span>" & FormatNumber(wPrice,0) & "�~</span>(�ō�)"
		wHTML = wHTML & "�Ռ������F<span>" & FormatNumber(RS("�̔��P��"),0) & "�~</span>(�Ŕ�)<br>"
		wHTML = wHTML & "(�ō�&nbsp;<span>" & FormatNumber(wPrice,0) & "�~</span>)"
'2014/03/19 GV mod end   <----
	End If
Else
	'---- B�i�P��
	If RS("B�i�t���O") = "Y" Then
		wPrice = calcPrice(RS("B�i�P��"), wSalesTaxRate)
'2014/03/19 GV mod start ---->
'		wHTML = wHTML & "�킯����i�����F<a class='tip'>ASK<span>" & FormatNumber(wPrice,0) & "�~(�ō�)</span></a>"
		wHTML = wHTML & "�킯����i�����F<a class='tip'>ASK<span class='exc-tax'>" & FormatNumber(RS("B�i�P��"),0) & "�~(�Ŕ�)</span><br>"
		wHTML = wHTML & "<span class='inc-tax'>(�ō�&nbsp;" & FormatNumber(wPrice,0) & "�~)</span></a>"
'2014/03/19 GV mod end   <----
	'---- ������P��
	ElseIf RS("�����萔��") > RS("������󒍍ϐ���") AND RS("�����萔��") > 0 Then
		wPrice = calcPrice(RS("������P��"), wSalesTaxRate)
'2014/03/19 GV mod start ---->
'		wHTML = wHTML & "��������F<a class='tip'>ASK<span>" & FormatNumber(wPrice,0) & "�~(�ō�)</span></a>"
		wHTML = wHTML & "��������F<a class='tip'>ASK<span class='exc-tax'>" & FormatNumber(RS("������P��"),0) & "�~(�Ŕ�)</span><br>"
		wHTML = wHTML & "<span class='inc-tax'>(�ō�&nbsp;" & FormatNumber(wPrice,0) & "�~)</span></a>"
'2014/03/19 GV mod end   <----
	'---- �ʏ폤�i
	Else
'2014/03/19 GV mod start ---->
'		wHTML = wHTML & "�Ռ������F<a class='tip'>ASK<span>" & FormatNumber(wPrice,0) & "�~(�ō�)</span></a>"
		wHTML = wHTML & "�Ռ������F<a class='tip'>ASK<span class='exc-tax'>" & FormatNumber(RS("�̔��P��"),0) & "�~(�Ŕ�)</span><br>"
		wHTML = wHTML & "<span class='inc-tax'>(�ō�&nbsp;" & FormatNumber(wPrice,0) & "�~)</span></a>"
'2014/03/19 GV mod end   <----
	End If
End If
wHTML = wHTML & "</li>" & vbNewLine

wHTML = wHTML & "          <li class='photo'><a href='ProductDetail.asp?item=" & RS("���[�J�[�R�[�h") & "^" & Server.URLEncode(RS("���i�R�[�h")) & "'>"
If RS("���i�摜�t�@�C����_��") <> "" Then
	wHTML = wHTML & "<img src='prod_img/" & RS("���i�摜�t�@�C����_��") & "' alt='" & Replace(RS("���[�J�[��") & " / " & RS("���i��"),"'","&#39;") & "' class='opover'>"
End If
wHTML = wHTML & "</a></li>" & vbNewLine
'---- ���i����
If RS("�����ߏ��i�R�����g") <> "" Then
	wHTML = wHTML & "        <li>" & Replace(RS("�����ߏ��i�R�����g"), vbNewLine, "<br>") & "</li>" & vbNewLine
Else
	wHTML = wHTML & "        <li>" & Replace(RS("���i�T��Web"), vbNewLine, "<br>") & "</li>" & vbNewLine
End If
wHTML = wHTML & "        </ul>" & vbNewLine

wHTML = wHTML & "        <div class='other_detail'>" & vbNewLine
wHTML = wHTML & "          <ul>" & vbNewLine
wHTML = wHTML & "            <li><a href='ProductDetail.asp?item=" & RS("���[�J�[�R�[�h") & "^" & Server.URLEncode(RS("���i�R�[�h")) &"'><img src='images/btn_detail.png' alt='�ڍׂ�����' class='opover'></a></li>"
wHTML = wHTML & "          </ul>" & vbNewLine
wHTML = wHTML & "        </div>" & vbNewLine
wHTML = wHTML & "      </li>" & vbNewLine
'2012/07/11 ok End Start

'2012/07/11 nt del Start
''----���������N
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
'
''----���惊���N
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
'
'if vSoundMovie <> "" then
'	wHTML = wHTML & "        <tr align='left' valign='middle'>" & vbNewLine
'	wHTML = wHTML & "          <td height='25' colspan='2' class='honbun'>" & vbNewLine
'	wHTML = wHTML & "            �T���v���F " & vSoundMovie & vbNewLine
'	wHTML = wHTML & "          </td>" & vbNewLine
'	wHTML = wHTML & "        </tr>" & vbNewLine
'end if
'
'vOldProductCd = RS("���[�J�[�R�[�h") & "+" & RS("���i�R�[�h")
'
''---- ���ꏤ�i�I���܂ŌJ��Ԃ� (�F�K�i������ꍇ�̂݌J��Ԃ�)
'Do until vOldProductCd <> RS("���[�J�[�R�[�h") & "+" & RS("���i�R�[�h")
'	'---- �F�K�i, �P��
'	wHTML = wHTML & "        <tr>" & vbNewLine
'	wHTML = wHTML & "          <td width='" & vWidth1 & "' height='25' align='right' nowrap>" & vbNewLine
'	if RS("�F") <> "" OR RS("�K�i") <> "" then
'		wHTML = wHTML & "            <a href='ProductDetail.asp?item=" & RS("���[�J�[�R�[�h") & "^" & Server.URLEncode(RS("���i�R�[�h")) & "^" & RS("�F") & "^" & RS("�K�i") & "' class='link'>"
'		if RS("�F") <> "" then
'			wHTML = wHTML & RS("�F") & " "
'		end if
'		if RS("�K�i") <> "" then
'			wHTML = wHTML & RS("�K�i")
'		end if
'		wHTML = wHTML & "</a>"
'	end if
'
'	wPrice = calcPrice(RS("�̔��P��"), wSalesTaxRate)
'
'	if RS("���i�R�[�h") <> "" then
'		if RS("ASK���i�t���O") = "Y" then
'
''2011/10/19 hn  mod s
''			wHTML = wHTML & "         <a href='JavaScript:void(0);' onClick=""askWin=window.open('AskPrice.asp?MakerName=" & Server.URLEncode(RS("���[�J�[��")) & "&ProductName=" & Server.URLEncode(RS("���i��")) & "&Price=" & wPrice & "' ,'ask', 'width=250 height=80 scrollbars=false memubar=false toolbar=false statusbar=false personalbar=false locationbar=false');"" class='link'>ASK</a>" & vbNewLine
'			wHTML = wHTML & "                        <a class='tip'>ASK<span>" & FormatNumber(wPrice,0) & "�~(�ō�)</span></a>" & vbNewLine
''2011/10/19 hn mod e
'
'		else
'			wHTML = wHTML & "            <span class='honbun'>" & FormatNumber(wPrice,0) & "�~(�ō�)</span>" & vbNewLine
'		end if
'	else
'		wHTML = wHTML & "            &nbsp;" & vbNewLine
'	end if
'
'	wHTML = wHTML & "          </td>" & vbNewLine
'
'	'---- �ڍ׃{�^���C�J�[�g�{�^��
'	wHTML = wHTML & "          <td width='" & vWidth2 & "' nowrap height='25' align='center' valign='middle'>" & vbNewLine
'
'	' �ʏ폤�i
'	if RS("���i�R�[�h") <> "" then
'		if vIroKikakuFl = false then
'			wHTML = wHTML & "            <a href='ProductDetail.asp?item=" & RS("���[�J�[�R�[�h") & "^" & Server.URLEncode(RS("���i�R�[�h")) & "^" & RS("�F") & "^" & RS("�K�i") & "'><img src='images/Shousai.gif' width='50' height='19' border='0' align='middle'></a>" & vbNewLine
'		end if
'		wHTML = wHTML & "            <a href='OrderPreInsert.asp?maker_cd=" & RS("���[�J�[�R�[�h") & "&product_cd=" & Server.URLEncode(RS("���i�R�[�h")) & "&iro=" & RS("�F") & "&kikaku=" & RS("�K�i") & "&qt=1'><img src='images/CartBlue.gif' width='30' height='19' border='0' align='middle'></a>" & vbNewLine
'
'	' �V���[�Y
'	else
'		wHTML = wHTML & "            <a href='SearchList.asp?i_type=se&sSeriesCd=" & RS("�V���[�Y�R�[�h") & "'><img src='images/Shousai.gif' width='50' height='19' border='0' align='middle'></a>"
'		wHTML = wHTML & "            <img src='images/blank.gif' width='30' height='19' border='0' align='middle'>" & vbNewLine
'	end if
'
'	wHTML = wHTML & "          </td>" & vbNewLine
'	wHTML = wHTML & "        </tr>" & vbNewLine
'
'	RS.MoveNext
'	if RS.EOF = true then
'		Exit Do
'	end if
'Loop
'
'wHTML = wHTML & "      </table>" & vbNewLine
'wHTML = wHTML & "    </td>" & vbNewLine
'2012/07/11 nt del End

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
<meta name="Description" content="�T�E���h�n�E�X�u<%=wCategoryName%>�v�̔���؁i�x�X�g�Z���[�jTOP10���i�����ē����܂��B">
<meta name="keyword" content="<%=wCategoryName%>">
<title><%=wCategoryName%>�̔����TOP10�b�T�E���h�n�E�X</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/shop.css" type="text/css">
<link rel="stylesheet" href="Style/searchlist.css" type="text/css">
<link rel="stylesheet" href="style/ask.css?20140401a" type="text/css">
</head>
<body>
<!--#include file="../Navi/NaviTop.inc"-->
<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="��������{���ł�"></a></span>
  <!-- �R���e���cstart -->
  <div id="globalContents">
    <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
      <p class="home"><a href="../"><img src="../images/icon_home.gif" alt="HOME"></a></p>
      <ul id="path">
        <li><span itemscope itemtype='http://data-vocabulary.org/Breadcrumb'><a href='LargeCategoryList.asp?LargeCategoryCd=<%=wLargeCategoryCd%>' itemprop='url'><span itemprop='title'><%=wLargeCategoryName%></span></a></span></li>
        <li><span itemscope itemtype='http://data-vocabulary.org/Breadcrumb'><a href='MidCategoryList.asp?MidCategoryCd=<%=wMidCategoryCd%>' itemprop='url'><span itemprop='title'><%=wMidCategoryName%></span></a></span></li>
        <li><span itemscope itemtype='http://data-vocabulary.org/Breadcrumb'><a href='SearchList.asp?i_type=c&s_category_cd=<%=CategoryCd%>' itemprop='url'><span itemprop='title'><%=wCategoryName%></span></a></span></li>
        <li class="now"><span itemscope itemtype='http://data-vocabulary.org/Breadcrumb'><span itemprop='title'>�����TOP10</span></span></li>
      </ul>
    </div></div></div>

    <h1 class="title"><%=wCategoryName%>�̔����TOP10</h1>

<!-- �x�X�g�Z���[���i�ꗗ-->
<%=wProductList%>

    <!--/#contents --></div>
  <div id="globalSide">
  <!--#include file="../Navi/NaviSide.inc"-->
  <!--/#globalSide --></div>
<!--/#main --></div>
<!--#include file="../Navi/Navibottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
<div class="tooltip"><p>ASK</p></div>
<script type="text/javascript" src="jslib/ask.js?20140401a"></script>
<script type="text/javascript" src="jslib/SearchList.js?20120321" charset="Shift_JIS"></script>
<script type="text/javascript" src="../jslib/jquery.tinyscrollbar.min.js"></script>
<script type="text/javascript">
$(function(){
    $('#scrollbar1').tinyscrollbar();
});
</script>
</body>
</html>