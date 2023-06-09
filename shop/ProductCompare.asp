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
'	���i��r�y�[�W
'
'	�X�V����
'2005/07/01 �[���\�������悻�̓����ɕύX
'2007/05/30 �F�K�i���菤�i��1�ɂ܂Ƃ߁A�J�[�g�{�^���̑���ɁA�ڍו\���{�^����\���B���i�ʉ�ʂ�\��
'2008/12/24 �݌ɏ󋵃Z�b�g�֐���
'2010/02/18 an ASK���i�p�����[�^��Server.URLEncode���s�Ȃ�
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2011/10/19 hn 1063 ASK�\�����@�ύX
'2012/01/20 GV �f�[�^�擾 SELECT���� LAC�N�G���[�Ă�K�p
'2012/07/18 nt ���j���[�A���p��asp��ʏo�͂��C��
'2014/03/19 GV ����ő��łɔ���2�d�\���Ή�
'
'========================================================================

On Error Resume Next

Dim wTitleWithLink
Dim wNaveWithLink

Dim wHikaku
Dim CategoryCd(5)
Dim MakerCd(5)
Dim ProductCd(5)
Dim Iro(5)
Dim Kikaku(5)
Dim MakerName(5)
Dim ProductName(5)
Dim Price(5)
Dim ImageFile(5)
Dim Chokusou(5)
Dim ASKfl(5)
Dim HikiateKanouSuu(5)
Dim HikiateKanouNyuukaYoteibi(5)
Dim KisyouSuu(5)
Dim Setfl(5)
Dim IroKikakuCnt(5)
Dim	WebNoukiHihyoujiFl(5)
Dim	NyukayoteiMiteiFl(5)
Dim	Haibanbi(5)
Dim	BhinFl(5)
Dim	BhinHikiateKanouQt(5)
Dim	KosuuGenteiQt(5)
Dim	KosuuGenteiJyuchuuQt(5)

Dim wRecCnt

Dim SpecNo(100)			'�\�����ɏ��i�X�y�b�N���ڔԍ�
Dim SpecName(100)		'�\�����ɏ��i�X�y�b�N��
Dim SpecComment(5,100)	'���i,�\�������Y���� �X�y�b�N���e

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
Dim RS_Template

Dim wHTML

Dim wSQL
Dim wMsg
Dim wErrDesc   '2011/08/01 an add

'========================================================================

Response.buffer = true

'---- Execute main
call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "ProductCompare.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
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
'	�X�V����
'2008/05/07 ��؂蕶���ύX
'
'========================================================================
'
Function main()

Dim i
Dim vTemp

'---- ���M�f�[�^�[�̎��o��
wHikaku = Split(ReplaceInput(Request("item")), "$")
wRecCnt = Ubound(wHikaku)

For i=1 to wRecCnt
	vTemp = Split(wHikaku(i), "^")
	CategoryCd(i) = Trim(vTemp(0))
	MakerCd(i) = Trim(vTemp(1))
	ProductCd(i) = Trim(vTemp(2))
	Iro(i) = Trim(vTemp(3))
	Kikaku(i) = Trim(vTemp(4))
Next

'---- ����ŗ���o��
call getCntlMst("����","����ŗ�","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)			'����ŗ�
wSalesTaxRate = Clng(wItemNum1)

'---- �i�r�Q�[�V�����Z�b�g
call SetNavi()

'---- �^�C�g���Z�b�g
call SetTitle()

'---- ���i�X�y�b�N�e���v���[�g���o��
call GetTemplate()

'---- ��r���i�f�[�^���o��
call getCompareProduct()

'---- ��r���i�ꗗ�쐬
call CreateCompareList()

End Function

'========================================================================
'
'	Function	�i�r�Q�[�V�����Z�b�g
'
'========================================================================
'
Function SetNavi()

Dim RSv

'---- �i�r�Q�[�V�����Z�b�g
wSQL = ""
' 2012/01/20 GV Mod Start
'wSQL = wSQL & "SELECT a.��J�e�S���[�R�[�h"
'wSQL = wSQL & "     , a.��J�e�S���[��"
'wSQL = wSQL & "     , b.���J�e�S���[�R�[�h"
'wSQL = wSQL & "     , b.���J�e�S���[�����{��"
'wSQL = wSQL & "     , c.�J�e�S���[�R�[�h"
'wSQL = wSQL & "     , c.�J�e�S���[��"
'wSQL = wSQL & "     , c.�����߃J�e�S���[�t���O"
'wSQL = wSQL & "  FROM ��J�e�S���[ a"
'wSQL = wSQL & "     , ���J�e�S���[ b"
'wSQL = wSQL & "     , �J�e�S���[ c"
'wSQL = wSQL & " WHERE b.��J�e�S���[�R�[�h = a.��J�e�S���[�R�[�h"
'wSQL = wSQL & "   AND c.���J�e�S���[�R�[�h = b.���J�e�S���[�R�[�h"
'wSQL = wSQL & "   AND c.�J�e�S���[�R�[�h = '" & CategoryCd(i) & "'"
wSQL = wSQL & "SELECT "
wSQL = wSQL & "       a.��J�e�S���[�R�[�h "
wSQL = wSQL & "     , a.��J�e�S���[�� "
wSQL = wSQL & "     , b.���J�e�S���[�R�[�h "
wSQL = wSQL & "     , b.���J�e�S���[�����{�� "
wSQL = wSQL & "     , c.�J�e�S���[�R�[�h "
wSQL = wSQL & "     , c.�J�e�S���[�� "
wSQL = wSQL & "     , c.�����߃J�e�S���[�t���O "
wSQL = wSQL & "FROM "
wSQL = wSQL & "    ��J�e�S���[              a WITH (NOLOCK) "
wSQL = wSQL & "      INNER JOIN ���J�e�S���[ b WITH (NOLOCK) "
wSQL = wSQL & "        ON     b.��J�e�S���[�R�[�h = a.��J�e�S���[�R�[�h "
wSQL = wSQL & "      INNER JOIN �J�e�S���[   c WITH (NOLOCK) "
wSQL = wSQL & "        ON     c.���J�e�S���[�R�[�h = b.���J�e�S���[�R�[�h "
wSQL = wSQL & "WHERE "
wSQL = wSQL & "         c.�J�e�S���[�R�[�h = '" & CategoryCd(1) & "' "
' 2012/01/20 GV Mod End

'@@@@@@@@@@response.write(wSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic

wNaveWithLink = ""
wNaveWithLink = wNaveWithLink & "<div id='path_box'><div id='path_box_inner01'><div id='path_box_inner02'>" & vbNewLine
wNaveWithLink = wNaveWithLink & " <p class='home'><a href='../'><img src='../images/icon_home.gif' alt='HOME'></a></p>" & vbNewLine
wNaveWithLink = wNaveWithLink & " <ul id='path'>" & vbNewLine
wNaveWithLink = wNaveWithLink & "  <li><a href='LargeCategoryList.asp?LargeCategoryCd=" & RSv("��J�e�S���[�R�[�h") & "'>" & RSv("��J�e�S���[��") & "</a></li>" & vbNewLine
wNaveWithLink = wNaveWithLink & "  <li><a href='MidCategoryList.asp?MidCategoryCd=" & RSv("���J�e�S���[�R�[�h") & "'>" & RSv("���J�e�S���[�����{��") & "</a></li>" & vbNewLine
wNaveWithLink = wNaveWithLink & "  <li><a href='SearchList.asp?i_type=c&s_category_cd=" & RSv("�J�e�S���[�R�[�h") & "'>" &  RSv("�J�e�S���[��") & "</a></li>" & vbNewLine
wNaveWithLink = wNaveWithLink & "  <li class='now'>���i��r</li>" & vbNewLine
wNaveWithLink = wNaveWithLink & "  </ul>" & vbNewLine
wNaveWithLink = wNaveWithLink & "</div></div></div>" & vbNewLine
RSv.close

End Function

'========================================================================
'
'	Function	�^�C�g���Z�b�g
'
'========================================================================
'
Function SetTitle()

Dim RSv

'---- �^�C�g���Z�b�g
wSQL = ""
' 2012/01/20 GV Mod Start
'wSQL = wSQL & "SELECT a.��J�e�S���[�R�[�h"
'wSQL = wSQL & "     , a.��J�e�S���[��"
'wSQL = wSQL & "     , b.���J�e�S���[�R�[�h"
'wSQL = wSQL & "     , b.���J�e�S���[�����{��"
'wSQL = wSQL & "     , c.�J�e�S���[�R�[�h"
'wSQL = wSQL & "     , c.�J�e�S���[��"
'wSQL = wSQL & "     , c.�����߃J�e�S���[�t���O"
'wSQL = wSQL & "  FROM ��J�e�S���[ a"
'wSQL = wSQL & "     , ���J�e�S���[ b"
'wSQL = wSQL & "     , �J�e�S���[ c"
'wSQL = wSQL & " WHERE b.��J�e�S���[�R�[�h = a.��J�e�S���[�R�[�h"
'wSQL = wSQL & "   AND c.���J�e�S���[�R�[�h = b.���J�e�S���[�R�[�h"
'wSQL = wSQL & "   AND c.�J�e�S���[�R�[�h = '" & CategoryCd(1) & "'"
wSQL = wSQL & "SELECT "
wSQL = wSQL & "       a.��J�e�S���[�R�[�h "
wSQL = wSQL & "     , a.��J�e�S���[�� "
wSQL = wSQL & "     , b.���J�e�S���[�R�[�h "
wSQL = wSQL & "     , b.���J�e�S���[�����{�� "
wSQL = wSQL & "     , c.�J�e�S���[�R�[�h "
wSQL = wSQL & "     , c.�J�e�S���[�� "
wSQL = wSQL & "     , c.�����߃J�e�S���[�t���O "
wSQL = wSQL & "FROM "
wSQL = wSQL & "    ��J�e�S���[              a WITH (NOLOCK) "
wSQL = wSQL & "      INNER JOIN ���J�e�S���[ b WITH (NOLOCK) "
wSQL = wSQL & "        ON     b.��J�e�S���[�R�[�h = a.��J�e�S���[�R�[�h "
wSQL = wSQL & "      INNER JOIN �J�e�S���[   c WITH (NOLOCK) "
wSQL = wSQL & "        On     c.���J�e�S���[�R�[�h = b.���J�e�S���[�R�[�h "
wSQL = wSQL & "WHERE "
wSQL = wSQL & "         c.�J�e�S���[�R�[�h = '" & CategoryCd(1) & "' "
' 2012/01/20 GV Mod Start

'@@@@@@@@@@response.write(wSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic

'2012/7/19 nt add
wTitleWithLink = ""
wTitleWithLink = wTitleWithLink & "<h1 class='title'>" & RSv("�J�e�S���[��") & " ���i��r</h1>" & vbNewLine

if RSv("�����߃J�e�S���[�t���O") = "Y" then
	wTitleWithLink = wTitleWithLink & "<div class='btn_box'>" & vbNewLine
	wTitleWithLink = wTitleWithLink & "<a href='ProductGuide.asp?CategoryCd=" & RSv("�J�e�S���[�R�[�h") & "'><img src='images/btn_recommend.png' alt='�������ߏ��i' class='opover'></a>" & vbNewLine
	wTitleWithLink = wTitleWithLink & "</div>" & vbNewLine
end if

'2012/7/19 nt del
'wTitleWithLink = "<b><font color='#696684'><a href='LargeCategoryList.asp?LargeCategoryCd=" & RSv("��J�e�S���[�R�[�h") & "' class='link'>" & RSv("��J�e�S���[��") & "</a>/<a href='MidCategoryList.asp?MidCategoryCd=" & RSv("���J�e�S���[�R�[�h") & "' class='link'>" & RSv("���J�e�S���[�����{��") & "</a>/<a href='SearchList.asp?i_type=c&s_category_cd=" & RSv("�J�e�S���[�R�[�h") & "' class='link'>" & RSv("�J�e�S���[��") & "</a></font></b>"

'if RSv("�����߃J�e�S���[�t���O") = "Y" then
'	wTitleWithLink = wTitleWithLink & "�@�@<a href='ProductGuide.asp?CategoryCd=" & RSv("�J�e�S���[�R�[�h") & "' class='link'>>>�������ߏ��i�͂�����</a>" & vbNewLine
'end if

RSv.close

End Function

'========================================================================
'
'	Function	���i�X�y�b�N�e���v���[�g���o��
'
'========================================================================
'
Function GetTemplate()

Dim i

'---- ���i�X�y�b�N�e���v���[�g���o��
wSQL = ""
wSQL = wSQL & "SELECT ���i�X�y�b�N���ڔԍ�"
wSQL = wSQL & "     , ���i�X�y�b�N���ږ�"
wSQL = wSQL & "  FROM ���i�X�y�b�N�e���v���[�g WITH (NOLOCK)"
wSQL = wSQL & " WHERE �J�e�S���[�R�[�h = '" & CategoryCd(1) & "'"
wSQL = wSQL & " ORDER BY �\����"

Set RS_Template = Server.CreateObject("ADODB.Recordset")
RS_Template.Open wSQL, Connection, adOpenStatic

i = 1
Do until RS_Template.EOF = true
	SpecNo(i) = RS_Template("���i�X�y�b�N���ڔԍ�")
	SpecName(i) = RS_Template("���i�X�y�b�N���ږ�")
	RS_Template.Movenext
	i = i + 1
Loop

RS_Template.close

End Function

'========================================================================
'
'	Function	��r���i�f�[�^���o��
'
'========================================================================
'
Function getCompareProduct()

Dim i
Dim j

For i=1 to wRecCnt
	'---- ���iRecordset�쐬
	wSQL = ""
' 2012/01/20 GV Mod Start
'	wSQL = wSQL & "SELECT b.���[�J�[�R�[�h"
'	wSQL = wSQL & "     , b.���i�R�[�h"
'	wSQL = wSQL & "     , b.�F"
'	wSQL = wSQL & "     , b.�K�i"
'	wSQL = wSQL & "     , a.���i��"
'	wSQL = wSQL & "     , a.���i�摜�t�@�C����_��"
'	wSQL = wSQL & "     , a.���[�J�[�������敪"
'	wSQL = wSQL & "     , a.ASK���i�t���O"
'	wSQL = wSQL & "     , a.�󏭐���"
'	wSQL = wSQL & "     , a.�Z�b�g���i�t���O"
'	wSQL = wSQL & "     , a.�����萔��"
'	wSQL = wSQL & "     , a.������󒍍ϐ���"
'	wSQL = wSQL & "     , a.Web�[����\���t���O"
'	wSQL = wSQL & "     , a.���ח\�薢��t���O"
'	wSQL = wSQL & "     , a.�p�ԓ�"
'	wSQL = wSQL & "     , a.B�i�t���O"
'	wSQL = wSQL & "     , CASE"
'	wSQL = wSQL & "         WHEN (a.�����萔�� > a.������󒍍ϐ��� AND a.�����萔�� > 0) THEN a.������P��"
'	wSQL = wSQL & "         ELSE a.�̔��P��"
'	wSQL = wSQL & "       END AS �̔��P��"
'	wSQL = wSQL & "     , b.�����\����"
'	wSQL = wSQL & "     , b.�����\���ח\���"
'	wSQL = wSQL & "     , b.B�i�����\����"
'	wSQL = wSQL & "     , c.���[�J�[��"
'	wSQL = wSQL & "     , d.���i�X�y�b�N���ڔԍ�"
'	wSQL = wSQL & "     , d.���i�X�y�b�N���e"
'
'		'�F�K�i�����邩�ǂ��� 2007/05/30
'	wSQL = wSQL & "     , (SELECT COUNT(*)"
'	wSQL = wSQL & "          FROM Web�F�K�i�ʍ݌� t"
'	wSQL = wSQL & "         WHERE t.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'	wSQL = wSQL & "           AND t.���i�R�[�h = a.���i�R�[�h"
'	wSQL = wSQL & "           AND (t.�F != '' OR t.�K�i != '')"
'	wSQL = wSQL & "           AND t.�I���� IS NULL"
'	wSQL = wSQL & "       ) AS �F�K�iCNT"
'
'	wSQL = wSQL & "  FROM Web���i a WITH (NOLOCK)"
'	wSQL = wSQL & "     , Web�F�K�i�ʍ݌� b WITH (NOLOCK)"
'	wSQL = wSQL & "     , ���[�J�[ c WITH (NOLOCK)"
'	wSQL = wSQL & "     , ���i�X�y�b�N d WITH (NOLOCK)"
'	wSQL = wSQL & " WHERE b.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'	wSQL = wSQL & "   AND b.���i�R�[�h = a.���i�R�[�h"
'	wSQL = wSQL & "   AND c.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'	wSQL = wSQL & "   AND d.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'	wSQL = wSQL & "   AND d.���i�R�[�h = a.���i�R�[�h"
'	wSQL = wSQL & "   AND b.���[�J�[�R�[�h = '" & MakerCd(i) & "'"
'	wSQL = wSQL & "   AND b.���i�R�[�h = '" & ProductCd(i) & "'"
'	wSQL = wSQL & "   AND b.�F = '" & Iro(i) & "'"
'	wSQL = wSQL & "   AND b.�K�i = '" & Kikaku(i) & "'"
'	wSQL = wSQL & " ORDER BY"
'	wSQL = wSQL & "       c.���[�J�[��"
'	wSQL = wSQL & "     , a.���i��"
'	wSQL = wSQL & "     , d.���i�X�y�b�N���ڔԍ�"

	wSQL = wSQL & "SELECT "
	wSQL = wSQL & "      b.���[�J�[�R�[�h "
	wSQL = wSQL & "    , b.���i�R�[�h "
	wSQL = wSQL & "    , b.�F "
	wSQL = wSQL & "    , b.�K�i "
	wSQL = wSQL & "    , a.���i�� "
	wSQL = wSQL & "    , a.���i�摜�t�@�C����_�� "
	wSQL = wSQL & "    , a.���[�J�[�������敪 "
	wSQL = wSQL & "    , a.ASK���i�t���O "
	wSQL = wSQL & "    , a.�󏭐��� "
	wSQL = wSQL & "    , a.�Z�b�g���i�t���O "
	wSQL = wSQL & "    , a.�����萔�� "
	wSQL = wSQL & "    , a.������󒍍ϐ��� "
	wSQL = wSQL & "    , a.Web�[����\���t���O "
	wSQL = wSQL & "    , a.���ח\�薢��t���O "
	wSQL = wSQL & "    , a.�p�ԓ� "
	wSQL = wSQL & "    , a.B�i�t���O "
	wSQL = wSQL & "    , CASE "
	wSQL = wSQL & "        WHEN (a.�����萔�� > a.������󒍍ϐ��� AND a.�����萔�� > 0) THEN a.������P�� "
	wSQL = wSQL & "        ELSE a.�̔��P�� "
	wSQL = wSQL & "      END AS �̔��P�� "
	wSQL = wSQL & "    , b.�����\���� "
	wSQL = wSQL & "    , b.�����\���ח\��� "
	wSQL = wSQL & "    , b.B�i�����\���� "
	wSQL = wSQL & "    , c.���[�J�[�� "
	wSQL = wSQL & "    , d.���i�X�y�b�N���ڔԍ� "
	wSQL = wSQL & "    , d.���i�X�y�b�N���e "
	wSQL = wSQL & "    , (SELECT COUNT(t.���i�R�[�h) "
	wSQL = wSQL & "         FROM Web�F�K�i�ʍ݌� t "
	wSQL = wSQL & "        WHERE     t.���[�J�[�R�[�h = a.���[�J�[�R�[�h "
	wSQL = wSQL & "              AND t.���i�R�[�h = a.���i�R�[�h "
	wSQL = wSQL & "              AND (t.�F != '' OR t.�K�i != '') "
	wSQL = wSQL & "              AND t.�I���� IS NULL "
	wSQL = wSQL & "      ) AS �F�K�iCNT "
	wSQL = wSQL & "FROM "
	wSQL = wSQL & "    Web���i                      a WITH (NOLOCK) "
	wSQL = wSQL & "      INNER JOIN Web�F�K�i�ʍ݌� b WITH (NOLOCK) "
	wSQL = wSQL & "        ON     b.���[�J�[�R�[�h = a.���[�J�[�R�[�h "
	wSQL = wSQL & "           AND b.���i�R�[�h     = a.���i�R�[�h "
	wSQL = wSQL & "      INNER JOIN ���[�J�[        c WITH (NOLOCK) "
	wSQL = wSQL & "        ON     c.���[�J�[�R�[�h = a.���[�J�[�R�[�h "
	wSQL = wSQL & "      INNER JOIN ���i�X�y�b�N    d WITH (NOLOCK) "
	wSQL = wSQL & "        ON     d.���[�J�[�R�[�h = a.���[�J�[�R�[�h "
	wSQL = wSQL & "           AND d.���i�R�[�h     = a.���i�R�[�h "
	wSQL = wSQL & "WHERE "
	wSQL = wSQL & "        b.���[�J�[�R�[�h = '" & MakerCd(i) & "' "
	wSQL = wSQL & "    AND b.���i�R�[�h     = '" & Replace(ProductCd(i), "'", "''") & "' "
	wSQL = wSQL & "    AND b.�F             = '" & Iro(i) & "' "
	wSQL = wSQL & "    AND b.�K�i           = '" & Kikaku(i) & "' "
	wSQL = wSQL & "ORDER BY "
	wSQL = wSQL & "      c.���[�J�[�� "
	wSQL = wSQL & "    , a.���i�� "
	wSQL = wSQL & "    , d.���i�X�y�b�N���ڔԍ� "
' 2012/01/20 GV Mod Start

	Set RS = Server.CreateObject("ADODB.Recordset")
	RS.Open wSQL, Connection, adOpenStatic

	Do Until RS.EOF = true
		MakerName(i) = RS("���[�J�[��")
		ProductName(i) = RS("���i��")
		Price(i) = RS("�̔��P��")
		ImageFile(i) = RS("���i�摜�t�@�C����_��")
		Chokusou(i) = RS("���[�J�[�������敪")
		ASKfl(i) = RS("ASK���i�t���O")
		KisyouSuu(i) = RS("�󏭐���")
		Setfl(i) = RS("�Z�b�g���i�t���O")
		HikiateKanouSuu(i) = RS("�����\����")
		HikiateKanouNyuukaYoteibi(i) = RS("�����\���ח\���")
		IroKikakuCnt(i) = RS("�F�K�iCNT")

		WebNoukiHihyoujiFl(i) = RS("Web�[����\���t���O")
		NyukayoteiMiteiFl(i) = RS("���ח\�薢��t���O")
		Haibanbi(i) = RS("�p�ԓ�")
		BhinFl(i) = RS("B�i�t���O")
		BhinHikiateKanouQt(i) = RS("B�i�����\����")
		KosuuGenteiQt(i) = RS("�����萔��")
		KosuuGenteiJyuchuuQt(i) = RS("������󒍍ϐ���")

		For j=1 to 100
			if SpecNo(j) = RS("���i�X�y�b�N���ڔԍ�") then
				SpecComment(i, j) = RS("���i�X�y�b�N���e")
				Exit for
			end if
		Next

		RS.MoveNext
	Loop
Next

RS.Close

End function

'========================================================================
'
'	Function	��r���i�ꗗ�쐬
'
'========================================================================
'
Function createCompareList()

Dim i
Dim j
Dim vLine
Dim vPrice
Dim vProductName
Dim vInventoryCd
Dim vInventoryImage
Dim vWidth
Dim vBgColor

'2012/7/19 nt add
'---- ���w��
if wRecCnt = 1 then
	vWidth = "200"
elseif wRecCnt = 2 then
	vWidth = "150"
elseif wRecCnt = 3 then
	vWidth = "120"
elseif wRecCnt = 4 then
	vWidth = "100"
elseif wRecCnt = 5 then
	vWidth = "80"
else
	vWidth = "200"
end if

'2012/7/19 nt add
wHTML = ""
wHTML = wHTML & "<table class='productcompare'>" & vbNewLine
wHTML = wHTML & " <tbody>" & vbNewLine

'2012/7/19 nt del
'---- ��؂��
'vLine = ""
'vLine = vLine & "  <tr>" & vbNewLine
'For i=0 to wRecCnt
'	vLine = vLine & "    <td width='100' height='1' bgcolor='#6699cc'><img src='images/blank.gif' width=1 height=1></td>" & vbNewLine
'Next
'vLine = vLine & "  </tr>"

'vWidth = (795 - 110) / wRecCnt

'----
'wHTML = ""
'wHTML = wHTML & "<table border='0' cellspacing='1' cellpadding='0'>" & vbNewLine
'wHTML = wHTML & vLine

'2012/7/19 nt add
'---- ���i�ʐ^
wHTML = wHTML & "<tr id='prod_img'>" & vbNewLine
wHTML = wHTML & " <th width='" & vWidth & "'>���i�ʐ^</th>" & vbNewLine
For i=1 to wRecCnt
	wHTML = wHTML & " <td><img src='prod_img/" & ImageFile(i) & "' alt='" & MakerName(i) & " / " & ProductName(i) & "'></td>" & vbNewLine
Next
wHTML = wHTML & "</tr>" & vbNewLine

'2012/7/19 nt del
'---- ���i�ʐ^
'wHTML = wHTML & "  <tr>"
'wHTML = wHTML & "    <td width='100' align='center' bgcolor='#eeeeee' nowrap class='honbun'>���i�ʐ^</td>" & vbNewLine

'For i=1 to wRecCnt
'	wHTML = wHTML & "    <td width='" & vWidth & "' align='center' bgcolor='#ffffff'><img src='prod_img/" & ImageFile(i) & "' width='124' height='62'></td>" & vbNewLine
'Next

'wHTML = wHTML & "  </tr>" & vbNewLine
'wHTML = wHTML & vLine

'2012/7/19 nt add
'---- ���[�J�[
wHTML = wHTML & "<tr>" & vbNewLine
wHTML = wHTML & " <th>���[�J�[</th>" & vbNewLine
For i=1 to wRecCnt
	wHTML = wHTML & " <td>" & MakerName(i) & "</td>" & vbNewLine
Next
wHTML = wHTML & "</tr>" & vbNewLine

'2012/7/19 nt del
'---- ���[�J�[
'wHTML = wHTML & "  <tr>"
'wHTML = wHTML & "    <td width='100' align='center' bgcolor='#eeeeee' nowrap class='honbun'>���[�J�[</td>" & vbNewLine

'For i=1 to wRecCnt
'	if i Mod 2 = 0 then
'		vBgColor = "#eeeeee"
'	else
'		vBgColor = "#ffffff"
'	end if
'	wHTML = wHTML & "    <td width='" & vWidth & "' align='center' bgcolor='" & vBgColor & "'class='honbun'>" & MakerName(i) & "</td>" & vbNewLine & vbNewLine
'Next

'wHTML = wHTML & "  </tr>" & vbNewLine
'wHTML = wHTML & vLine

'2012/7/19 nt add
'---- ���i��/�F/�K�i
wHTML = wHTML & "<tr>" & vbNewLine
wHTML = wHTML & " <th>���i��</th>" & vbNewLine
For i=1 to wRecCnt
	vProductName = ProductName(i)
	if Trim(Iro(i)) <> "" then
		vProductName = vProductName & "/" & Trim(Iro(i))
	end if
	if Trim(Kikaku(i)) <> "" then
		vProductName = vProductName & "/" & Trim(Kikaku(i))
	end if

	wHTML = wHTML & " <td><a href='ProductDetail.asp?item=" & MakerCd(i) & "^" & ProductCd(i) & "^" & Iro(i) & "^" & Kikaku(i) & "'>" & vProductName & "</a></td>" & vbNewLine
Next
wHTML = wHTML & "</tr>"

'2012/7/19 nt del
'---- ���i��/�F/�K�i
'wHTML = wHTML & "  <tr>"
'wHTML = wHTML & "    <td width='100' align='center' bgcolor='#eeeeee' nowrap class='honbun'>���i��</td>" & vbNewLine

'For i=1 to wRecCnt
'	if i Mod 2 = 0 then
'		vBgColor = "#eeeeee"
'	else
'		vBgColor = "#ffffff"
'	end if

'	vProductName = ProductName(i)
'	if Trim(Iro(i)) <> "" then
'		vProductName = vProductName & "/" & Trim(Iro(i))
'	end if
'	if Trim(Kikaku(i)) <> "" then
'		vProductName = vProductName & "/" & Trim(Kikaku(i))
'	end if

'	wHTML = wHTML & "    <td width='" & vWidth & "' align='center' bgcolor='" & vBgColor & "'><a href='ProductDetail.asp?item=" & MakerCd(i) & "^" & ProductCd(i) & "^" & Iro(i) & "^" & Kikaku(i) & "' class='link'>" & vProductName & "</a></td>" & vbNewLine
'Next

'wHTML = wHTML & "  </tr>" & vbNewLine
'wHTML = wHTML & vLine

'2012/7/19 nt add
'---- �Ռ�����
wHTML = wHTML & "<tr>" & vbNewLine
wHTML = wHTML & " <th>�Ռ�����</th>" & vbNewLine
For i=1 to wRecCnt
	vPrice = calcPrice(Price(i), wSalesTaxRate)
	wHTML = wHTML & " <td>"
	if ASKfl(i) = "Y" then
'2014/03/19 GV mod start ---->
'		wHTML = wHTML & "<span class='honbun'><a class='tip'>ASK<span>" & FormatNumber(vPrice,0) & "�~(�ō�)</span></a></span>" & vbNewLine
		wHTML = wHTML & "<span class='honbun'><a class='tip'>ASK<span class='exc-tax'>" & FormatNumber(Price(i),0) & "�~(�Ŕ�)</span><br>"
		wHTML = wHTML & "<span class='inc-tax'>(�ō�&nbsp;" & FormatNumber(vPrice,0) & "�~)</span></a>" & vbNewLine
'2014/03/19 GV mod end   <----
	else
'2014/03/19 GV mod start ---->
'		wHTML = wHTML & FormatNumber(vPrice,0) & "�~(�ō�)" & vbNewLine
		wHTML = wHTML & FormatNumber(Price(i),0) & "�~(�Ŕ�)<br>" & vbNewLine
		wHTML = wHTML & "(�ō�&nbsp;" & FormatNumber(vPrice,0) & "�~)" & vbNewLine
'2014/03/19 GV mod end   <----
	end if
	wHTML = wHTML & " </td>" & vbNewLine
Next
wHTML = wHTML & "</tr>" & vbNewLine

'2012/7/19 nt del
'---- �Ռ�����
'wHTML = wHTML & "  <tr>" & vbNewLine
'wHTML = wHTML & "    <td width='100' align='center' bgcolor='#eeeeee' nowrap class='honbun'>�Ռ�����</td>" & vbNewLine

'For i=1 to wRecCnt
'	if i Mod 2 = 0 then
'		vBgColor = "#eeeeee"
'	else
'		vBgColor = "#ffffff"
'	end if
'	vPrice = calcPrice(Price(i), wSalesTaxRate)
'	wHTML = wHTML & "    <td width='" & vWidth & "' align='center' bgcolor='" & vBgColor & "'>"
'	if ASKfl(i) = "Y" then

'2011/10/19 hn mod s
'		wHTML = wHTML & "<a href='JavaScript:void(0);' onClick=""askWin=window.open('AskPrice.asp?MakerName=" & Server.URLEncode(MakerName(i)) & "&ProductName=" & Server.URLEncode(ProductName(i)) & "&Price=" & vPrice & "' ,'ask', 'width=250 height=80 scrollbars=false memubar=false toolbar=false statusbar=false personalbar=false locationbar=false');"" class='link'>ASK</a>"

'		wHTML = wHTML & "<span class='honbun'><a class='tip'>ASK<span>" & FormatNumber(vPrice,0) & "�~(�ō�)</span></a></span>"

'2011/10/19 hn mod e

'	else
'		wHTML = wHTML & "<span class='honbun'>" & FormatNumber(vPrice,0) & "�~(�ō�)</span>"
'	end if
'	wHTML = wHTML & "</td>" & vbNewLine
'Next

'wHTML = wHTML & "  </tr>" & vbNewLine
'wHTML = wHTML & vLine

'2012/7/19 nt add
'---- �݌ɏ�
wHTML = wHTML & "<tr>" & vbNewLine
wHTML = wHTML & " <th>�݌ɏ�</th>" & vbNewLine
For i=1 to wRecCnt
	if IroKikakuCnt(i) = 0 then	'2007/05/30
		vInventoryCd = GetInventoryStatus(Makercd(i),ProductCd(i),Iro(i),Kikaku(i),HikiateKanouSuu(i),KisyouSuu(i),Setfl(i),Chokusou(i),HikiateKanouNyuukaYoteibi(i),"N")

		'---- �݌ɏ󋵁A�F���ŏI�Z�b�g
		call GetInventoryStatus2(HikiateKanouSuu(i), WebNoukiHihyoujiFl(i), NyukayoteiMiteiFl(i), Haibanbi(i), BhinFl(i), BhinHikiateKanouQt(i), KosuuGenteiQt(i), KosuuGenteiJyuchuuQt(i), "N", vInventoryCd, vInventoryImage)
		wHTML = wHTML & " <td class='stock'><img src='images/" & vInventoryImage & "'>" & vInventoryCd & "</td>" & vbNewLine

	else
		wHTML = wHTML & " <td class='stock'></td>" & vbNewLine
	end if
Next
wHTML = wHTML & "</tr>" & vbNewLine

'2012/7/19 nt del
'---- �݌ɏ�
'wHTML = wHTML & "  <tr>"
'wHTML = wHTML & "    <td width='100' align='center' bgcolor='#eeeeee' nowrap class='honbun'>�݌ɏ�</td>" & vbNewLine

'For i=1 to wRecCnt
'	if i Mod 2 = 0 then
'		vBgColor = "#eeeeee"
'	else
'		vBgColor = "#ffffff"
'	end if

'	if IroKikakuCnt(i) = 0 then	'2007/05/30
'		vInventoryCd = GetInventoryStatus(Makercd(i),ProductCd(i),Iro(i),Kikaku(i),HikiateKanouSuu(i),KisyouSuu(i),Setfl(i),Chokusou(i),HikiateKanouNyuukaYoteibi(i),"N")

		'---- �݌ɏ󋵁A�F���ŏI�Z�b�g
'		call GetInventoryStatus2(HikiateKanouSuu(i), WebNoukiHihyoujiFl(i), NyukayoteiMiteiFl(i), Haibanbi(i), BhinFl(i), BhinHikiateKanouQt(i), KosuuGenteiQt(i), KosuuGenteiJyuchuuQt(i), "N", vInventoryCd, vInventoryImage)

'		wHTML = wHTML & "    <td width='" & vWidth & "' align='center' bgcolor='" & vBgColor & "' class='honbun'><img src='images/" & vInventoryImage & "' width=10 height=10> " & vInventoryCd & "</td>" & vbNewLine

'	else
'		wHTML = wHTML & "    <td width='" & vWidth & "' align='center' bgcolor='" & vBgColor & "' class='honbun'></td>" & vbNewLine
'	end if
'Next

'wHTML = wHTML & "  </tr>" & vbNewLine

'wHTML = wHTML & vLine

'2012/7/19 nt add
'---- �X�y�b�N
wHTML = wHTML & "<tr id='spec'>" & vbNewLine
wHTML = wHTML & " <th colspan='6'>�X�y�b�N</th>" & vbNewLine
wHTML = wHTML & "</tr>" & vbNewLine
For j=1 to 100
	if SpecNo(j) = "" then
		exit for
	end if

	wHTML = wHTML & "<tr>" & vbNewLine
	wHTML = wHTML & " <th>" & SpecName(j) & "</th>" & vbNewLine

	For i=1 to wRecCnt
		wHTML = wHTML & " <td>" & vbNewLine
		if Trim(SpecComment(i, j)) = "" OR IsNull(SpecComment(i, j)) = true then
			wHTML = wHTML & "-" & vbNewLine
		else
			wHTML = wHTML & SpecComment(i, j) & vbNewLine
		end if
		wHTML = wHTML & " </td>" & vbNewLine
	Next

	wHTML = wHTML & "  </tr>" & vbNewLine
Next

'2012/7/19 nt del
'---- �X�y�b�N
'wHTML = wHTML & "  <tr align='left' valign='bottom'>" & vbNewLine
'wHTML = wHTML & "    <td align='center' height='30' class='honbun'><b>�X�y�b�N</b></td>" & vbNewLine
'wHTML = wHTML & "  </tr>" & vbNewLine
'wHTML = wHTML & vLine

'For j=1 to 100
'	if SpecNo(j) = "" then
'		exit for
'	end if

'	wHTML = wHTML & "  <tr class='honbun'>"
'	wHTML = wHTML & "    <td width='100' align='center' bgcolor='#eeeeee' nowrap>" & SpecName(j) & "</td>" & vbNewLine

'	For i=1 to wRecCnt
'		if i Mod 2 = 0 then
'			vBgColor = "#eeeeee"
'		else
'			vBgColor = "#ffffff"
'		end if
'		wHTML = wHTML & "    <td width='" & vWidth & "' align='center' valign='top' bgcolor='" & vBgColor & "'>"
'		if Trim(SpecComment(i, j)) = "" OR IsNull(SpecComment(i, j)) = true then
'			wHTML = wHTML & "-"
'		else
'			wHTML = wHTML & SpecComment(i, j)
'		end if
'		wHTML = wHTML & "</td>" & vbNewLine
'	Next

'	wHTML = wHTML & "  </tr>" & vbNewLine
'	wHTML = wHTML & vLine
'Next

'2012/7/19 nt add
'---- �J�[�g
wHTML = wHTML & "<tr id='cart'>" & vbNewLine
wHTML = wHTML & " <th>�J�[�g��</th>" & vbNewLine

For i=1 to wRecCnt
	if IroKikakuCnt(i) = 0 then
		wHTML = wHTML & " <form name='f_item' method='post' action='OrderPreInsert.asp' onSubmit='return order_onClick(this);'>" & vbNewLine
		wHTML = wHTML & "  <td nowrap>" & vbNewLine
		wHTML = wHTML & "   <input type='text' name='qt' value='1'>" & vbNewLine
		wHTML = wHTML & "   <input type='image' src='images/btn_cart.png' alt='�J�[�g�ɓ����' class='opover'>" & vbNewLine
		wHTML = wHTML & "   <input type='hidden' name='maker_cd' value='" & MakerCd(i) & "'>" & vbNewLine
		wHTML = wHTML & "   <input type='hidden' name='product_cd' value='" & ProductCd(i) & "'>" & vbNewLine
		wHTML = wHTML & "   <input type='hidden' name='iro' value='" & Iro(i) & "'>" & vbNewLine
		wHTML = wHTML & "   <input type='hidden' name='kikaku' value='" & Kikaku(i) & "'>" & vbNewLine
		wHTML = wHTML & "   <input type='hidden' name='category_cd' value='" & CategoryCd(i) & "'>" & vbNewLine
		wHTML = wHTML & "  </td>" & vbNewLine
		wHTML = wHTML & " </form>" & vbNewLine
	else
		wHTML = wHTML & " <td>" & vbNewLine
		wHTML = wHTML & "  <a href='ProductDetail.asp?Item=" & MakerCd(i) & "^" & ProductCd(i) & "'>" & vbNewLine
		wHTML = wHTML & "   <img src='images/Shousai.gif'>" & vbNewLine
		wHTML = wHTML & "  </a>" & vbNewLine
		wHTML = wHTML & " </td>" & vbNewLine
	end if
Next
wHTML = wHTML & "</tr>" & vbNewLine

'2012/7/19 nt del
'---- �J�[�g
'wHTML = wHTML & "  <tr>"
'wHTML = wHTML & "    <td width='100' align='center' bgcolor='#ffffff' nowrap class='honbun'>�J�[�g��</td>" & vbNewLine

'For i=1 to wRecCnt
'	if IroKikakuCnt(i) = 0 then	'2007/05/30
'		wHTML = wHTML & "    <form name='f_item' method='post' action='OrderPreInsert.asp' onSubmit='return order_onClick(this);'>" & vbNewLine
'		wHTML = wHTML & "    <td width='" & vWidth & "' align='center' bgcolor='#ffffff'class='honbun'>" & vbNewLine
'		wHTML = wHTML & "      <input type='text' name='qt' size='3' value='1'>" & vbNewLine
'		wHTML = wHTML & "      <input type='image' src='images/CartSmall.jpg' width='22' height='18'>" & vbNewLine
'		wHTML = wHTML & "      <input type='hidden' name='maker_cd' value='" & MakerCd(i) & "'>" & vbNewLine
'		wHTML = wHTML & "      <input type='hidden' name='product_cd' value='" & ProductCd(i) & "'>" & vbNewLine
'		wHTML = wHTML & "      <input type='hidden' name='iro' value='" & Iro(i) & "'>" & vbNewLine
'		wHTML = wHTML & "      <input type='hidden' name='kikaku' value='" & Kikaku(i) & "'>" & vbNewLine
'		wHTML = wHTML & "      <input type='hidden' name='category_cd' value='" & CategoryCd(i) & "'>" & vbNewLine
'		wHTML = wHTML & "    </td>" & vbNewLine
'		wHTML = wHTML & "    </form>" & vbNewLine

'	else
'		wHTML = wHTML & "    <td width='" & vWidth & "' align='center' bgcolor='#ffffff'class='honbun'>" & vbNewLine
'		wHTML = wHTML & "      <a href='ProductDetail.asp?Item=" & MakerCd(i) & "^" & ProductCd(i) & "'>"
'		wHTML = wHTML & "      <img src='images/Shousai.gif' border='0'></a>" & vbNewLine
'		wHTML = wHTML & "    </td>" & vbNewLine
'	end if
'Next

'wHTML = wHTML & "  </tr>" & vbNewLine

'wHTML = wHTML & vLine

wHTML = wHTML & "</tbody>" '2012/7/19 nt add
wHTML = wHTML & "</table>"

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
<title>���i��r�b�T�E���h�n�E�X</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="Style/shop.css" type="text/css">
<link rel="stylesheet" href="Style/productcompare.css" type="text/css">
<link rel="stylesheet" href="style/ask.css?20140401a" type="text/css">

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
		alert("���ʂ���͂��Ă���J�[�g�{�^���������Ă��������B");
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
		<!-- �i�r�Q�[�V���� -->
		<%=wNaveWithLink%>
		<!-- �^�C�g�� -->	
		<%=wTitleWithLink%>
		<!-- ��r���X�g -->
		<%=wHTML%>
	</div>
	<div id="globalSide">
		<!--#include file="../Navi/NaviSide.inc"-->
	</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
<div class="tooltip"><p>ASK</p></div>
<script type="text/javascript" src="jslib/ask.js?20140401a"></script>
</body>
</html>
