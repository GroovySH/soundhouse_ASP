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
'	�x�X�g�Z���[�ꗗ�y�[�W
'
'�X�V����
'2005/03/31 �F�K�i�����鏤�i�ւ̃����N�͏��i�ʂł͂Ȃ����i�ꗗ�ɕύX
'2006/04/05 �W�v�P�ʂ𒆃J�e�S���[�ʂɕύX
'2006/11/08 ���[�J�[��+���i����25�����ŃJ�b�g
'2007/05/25 �V���[�Y�Ή�
'2009/04/30 �G���[����error.asp�ֈړ�
'2010/05/29 �����L���O�y�[�W���j���[�A���Ή�
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2012/01/19 GV �f�[�^�擾 SELECT���� LAC�N�G���[�Ă�K�p
'2012/01/20 GV �f�[�^�擾 SELECT�����甄�؏��i�e�[�u���̍ŐV���̃f�[�^�̂ݒ��o����������폜
'2012/08/07 if-web ���j���[�A�����C�A�E�g����
'
'========================================================================

On Error Resume Next

'----2010/05/29 st add
Dim LargeCategoryCd

Dim wLargeCategoryHTML
Dim wLargeCategoryName
Dim wNoData
'----2010/05/29

Dim wYYYYMM

Dim Connection
Dim RS

Dim wSQL
Dim wHTML
Dim w_error_msg
Dim wErrDesc   '2011/08/01 an add

'========================================================================

'---- Get input data 2010/05/29 st add
LargeCategoryCd = ReplaceInput(Trim(Request("LargeCategoryCd")))

'---- ��J�e�S���[�R�[�h�̎w�肪�Ȃ��ꍇ
if LargeCategoryCd = "" then
	LargeCategoryCd = "1"
end if
'----  2010/05/29

'---- Execute main
call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "BestSellerList.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc)
end if                                           '2011/08/01 an add e

call close_db()

'---- 2010/05/29 st mod
if wNoData = "Y" OR Err.Description <> "" then
	Response.Redirect g_HTTP & "shop/Error.asp"
end if
'----  2010/05/29

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

'----2010/05/29 st del
'Dim vLCatHTML
'Dim vCatHTML
'Dim vProdHTML(10)
'Dim vBreakKey2
'Dim vBreakNextKey2
'Dim i
'----2010/05/29

'----2010/05/29 st add
Dim vRank
Dim vBGColor
'----2010/05/29

Dim vCatCount
Dim vProdCount
Dim vMakerProduct
Dim vYYYYMM
Dim vBreakKey1
Dim vBreakNextKey1

vYYYYMM = Left(fFormatDate(Date()), 7)


'---- ��J�e�S���[�ꗗ�쐬
call CreateLargeCategoryHTML()
if wNoData <> "Y" then  '�z��O�̑�J�e�S���[���w�肳���NoData�̏ꍇ�̓G���[


'---- ����؃����L���O ���o��
wSQL = ""
' 2012/01/19 GV Mod Start
'wSQL = wSQL & "SELECT"
'wSQL = wSQL & "       a.���[�J�[�R�[�h"
'wSQL = wSQL & "     , a.���i�R�[�h"
'wSQL = wSQL & "     , '' AS �V���[�Y�R�[�h"
'wSQL = wSQL & "     , a.�󒍐���"
'wSQL = wSQL & "     , b.���i��"
'wSQL = wSQL & "     , c.���[�J�[��"
'wSQL = wSQL & "     , e.���J�e�S���[�R�[�h"
'wSQL = wSQL & "     , e.���J�e�S���[�����{��"
'wSQL = wSQL & "     , e.�\���� AS ���J�e�S���[�\����"
'wSQL = wSQL & "     , f.��J�e�S���[�R�[�h"
'wSQL = wSQL & "     , f.��J�e�S���[��"
'wSQL = wSQL & "     , COUNT(g.���i�R�[�h) AS �F�K�i�ʍ݌Ɍ���"
'wSQL = wSQL & "  FROM "
'wSQL = wSQL & "       ���؏��i a WITH (NOLOCK)"
'wSQL = wSQL & "     , Web���i b WITH (NOLOCK)"
'wSQL = wSQL & "     , ���[�J�[ c WITH (NOLOCK)"
'wSQL = wSQL & "     , �J�e�S���[ d WITH (NOLOCK)"
'wSQL = wSQL & "     , ���J�e�S���[ e WITH (NOLOCK)"
'wSQL = wSQL & "     , ��J�e�S���[ f WITH (NOLOCK)"
'wSQL = wSQL & "     , Web�F�K�i�ʍ݌� g WITH (NOLOCK)"
'wSQL = wSQL & " WHERE "
'wSQL = wSQL & "       b.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'wSQL = wSQL & "   AND b.���i�R�[�h = a.���i�R�[�h"
'wSQL = wSQL & "   AND b.�J�e�S���[�R�[�h = a.�J�e�S���[�R�[�h"
'wSQL = wSQL & "   AND c.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'wSQL = wSQL & "   AND d.�J�e�S���[�R�[�h = a.�J�e�S���[�R�[�h"
'wSQL = wSQL & "   AND e.���J�e�S���[�R�[�h = d.���J�e�S���[�R�[�h"
'wSQL = wSQL & "   AND f.��J�e�S���[�R�[�h = e.��J�e�S���[�R�[�h"
'wSQL = wSQL & "   AND g.���[�J�[�R�[�h = b.���[�J�[�R�[�h"
'wSQL = wSQL & "   AND g.���i�R�[�h = b.���i�R�[�h"
'wSQL = wSQL & "   AND b.�I���� IS NULL"
'wSQL = wSQL & "   AND g.�I���� IS NULL"
'wSQL = wSQL & "   AND b.Web���i�t���O = 'Y'"
'wSQL = wSQL & "   AND d.����؃����L���O�\���t���O = 'Y'"
'wSQL = wSQL & "   AND a.�N�� = (SELECT MAX(�N��) FROM ���؏��i)"
'wSQL = wSQL & "   AND f.��J�e�S���[�R�[�h = '" & LargeCategoryCd & "'" '2010/05/29 st add
'wSQL = wSQL & " GROUP BY"
'wSQL = wSQL & "       a.���[�J�[�R�[�h"
'wSQL = wSQL & "     , a.���i�R�[�h"
'wSQL = wSQL & "     , a.�󒍐���"
'wSQL = wSQL & "     , b.���i��"
'wSQL = wSQL & "     , c.���[�J�[��"
'wSQL = wSQL & "     , e.���J�e�S���[�R�[�h"
'wSQL = wSQL & "     , e.���J�e�S���[�����{��"
'wSQL = wSQL & "     , e.�\����"
'wSQL = wSQL & "     , f.��J�e�S���[�R�[�h"
'wSQL = wSQL & "     , f.��J�e�S���[��"
'
'wSQL = wSQL & " UNION "
'
'wSQL = wSQL & "SELECT"
'wSQL = wSQL & "       a.���[�J�[�R�[�h"
'wSQL = wSQL & "     , '' AS ���i�R�[�h"
'wSQL = wSQL & "     , a.�V���[�Y�R�[�h"
'wSQL = wSQL & "     , a.�󒍐���"
'wSQL = wSQL & "     , b.�V���[�Y�� AS ���i��"
'wSQL = wSQL & "     , c.���[�J�[��"
'wSQL = wSQL & "     , e.���J�e�S���[�R�[�h"
'wSQL = wSQL & "     , e.���J�e�S���[�����{��"
'wSQL = wSQL & "     , e.�\���� AS ���J�e�S���[�\����"
'wSQL = wSQL & "     , f.��J�e�S���[�R�[�h"
'wSQL = wSQL & "     , f.��J�e�S���[��"
'wSQL = wSQL & "     , 2 AS �F�K�i�ʍ݌Ɍ���"
'wSQL = wSQL & "  FROM "
'wSQL = wSQL & "       ���؏��i a WITH (NOLOCK)"
'wSQL = wSQL & "     , �V���[�Y b WITH (NOLOCK)"
'wSQL = wSQL & "     , ���[�J�[ c WITH (NOLOCK)"
'wSQL = wSQL & "     , �J�e�S���[ d WITH (NOLOCK)"
'wSQL = wSQL & "     , ���J�e�S���[ e WITH (NOLOCK)"
'wSQL = wSQL & "     , ��J�e�S���[ f WITH (NOLOCK)"
'wSQL = wSQL & " WHERE "
'wSQL = wSQL & "       b.�V���[�Y�R�[�h = a.�V���[�Y�R�[�h"
'wSQL = wSQL & "   AND b.�J�e�S���[�R�[�h = a.�J�e�S���[�R�[�h"
'wSQL = wSQL & "   AND c.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'wSQL = wSQL & "   AND d.�J�e�S���[�R�[�h = a.�J�e�S���[�R�[�h"
'wSQL = wSQL & "   AND e.���J�e�S���[�R�[�h = d.���J�e�S���[�R�[�h"
'wSQL = wSQL & "   AND f.��J�e�S���[�R�[�h = e.��J�e�S���[�R�[�h"
'wSQL = wSQL & "   AND d.����؃����L���O�\���t���O = 'Y'"
'wSQL = wSQL & "   AND a.�N�� = (SELECT MAX(�N��) FROM ���؏��i)"
'wSQL = wSQL & "   AND f.��J�e�S���[�R�[�h = '" & LargeCategoryCd & "'" '2010/05/29 st add
'
'wSQL = wSQL & " ORDER BY"
'wSQL = wSQL & "       f.��J�e�S���[�R�[�h"
'wSQL = wSQL & "     , e.�\����"		'���J�e�S���[�\����
'wSQL = wSQL & "     , a.�󒍐��� DESC"
wSQL = wSQL & "SELECT "
wSQL = wSQL & "      a.���[�J�[�R�[�h "
wSQL = wSQL & "    , a.���i�R�[�h "
wSQL = wSQL & "    , '' AS �V���[�Y�R�[�h "
wSQL = wSQL & "    , a.�󒍐��� "
wSQL = wSQL & "    , b.���i�� "
wSQL = wSQL & "    , c.���[�J�[�� "
wSQL = wSQL & "    , e.���J�e�S���[�R�[�h "
wSQL = wSQL & "    , e.���J�e�S���[�����{�� "
wSQL = wSQL & "    , e.�\���� AS ���J�e�S���[�\���� "
wSQL = wSQL & "    , f.��J�e�S���[�R�[�h "
wSQL = wSQL & "    , f.��J�e�S���[�� "
wSQL = wSQL & "    , COUNT(g.���i�R�[�h) AS �F�K�i�ʍ݌Ɍ��� "
wSQL = wSQL & "FROM "
wSQL = wSQL & "    ���؏��i                     a WITH (NOLOCK) "
wSQL = wSQL & "      INNER JOIN Web���i         b WITH (NOLOCK) "
wSQL = wSQL & "        ON     b.���[�J�[�R�[�h     = a.���[�J�[�R�[�h "
wSQL = wSQL & "           AND b.���i�R�[�h         = a.���i�R�[�h "
wSQL = wSQL & "           AND b.�J�e�S���[�R�[�h   = a.�J�e�S���[�R�[�h "
wSQL = wSQL & "      INNER JOIN ���[�J�[        c WITH (NOLOCK) "
wSQL = wSQL & "        ON     c.���[�J�[�R�[�h     = a.���[�J�[�R�[�h "
wSQL = wSQL & "      INNER JOIN �J�e�S���[      d WITH (NOLOCK) "
wSQL = wSQL & "        ON     d.�J�e�S���[�R�[�h   = a.�J�e�S���[�R�[�h "
wSQL = wSQL & "      INNER JOIN ���J�e�S���[    e WITH (NOLOCK) "
wSQL = wSQL & "        ON     e.���J�e�S���[�R�[�h = d.���J�e�S���[�R�[�h "
wSQL = wSQL & "      INNER JOIN ��J�e�S���[    f WITH (NOLOCK) "
wSQL = wSQL & "        ON     f.��J�e�S���[�R�[�h = e.��J�e�S���[�R�[�h "
wSQL = wSQL & "      INNER JOIN Web�F�K�i�ʍ݌� g WITH (NOLOCK) "
wSQL = wSQL & "        ON     g.���[�J�[�R�[�h     = b.���[�J�[�R�[�h "
wSQL = wSQL & "           AND g.���i�R�[�h         = b.���i�R�[�h "
wSQL = wSQL & "      LEFT JOIN ( SELECT 'Y' AS 'ShohinWebY' )   t1 "
wSQL = wSQL & "        ON     b.Web���i�t���O    = t1.ShohinWebY "
wSQL = wSQL & "      LEFT JOIN ( SELECT 'Y' AS 'HotSellingY' )  t2 "
wSQL = wSQL & "        ON     d.����؃����L���O�\���t���O = t2.HotSellingY "
wSQL = wSQL & "WHERE "
wSQL = wSQL & "        t1.ShohinWebY  IS NOT NULL "
wSQL = wSQL & "    AND t2.HotSellingY IS NOT NULL "
wSQL = wSQL & "    AND b.�I����       IS NULL "
wSQL = wSQL & "    AND g.�I����       IS NULL "
'wSQL = wSQL & "    AND a.�N�� = (SELECT MAX(�N��) FROM ���؏��i) "				' 2012/01/20 GV Del
wSQL = wSQL & "    AND f.��J�e�S���[�R�[�h = '" & LargeCategoryCd & "' "
wSQL = wSQL & "GROUP BY "
wSQL = wSQL & "      a.���[�J�[�R�[�h "
wSQL = wSQL & "    , a.���i�R�[�h "
wSQL = wSQL & "    , a.�󒍐��� "
wSQL = wSQL & "    , b.���i�� "
wSQL = wSQL & "    , c.���[�J�[�� "
wSQL = wSQL & "    , e.���J�e�S���[�R�[�h "
wSQL = wSQL & "    , e.���J�e�S���[�����{�� "
wSQL = wSQL & "    , e.�\���� "
wSQL = wSQL & "    , f.��J�e�S���[�R�[�h "
wSQL = wSQL & "    , f.��J�e�S���[�� "

wSQL = wSQL & "UNION "

wSQL = wSQL & "SELECT "
wSQL = wSQL & "      a.���[�J�[�R�[�h "
wSQL = wSQL & "    , '' AS ���i�R�[�h "
wSQL = wSQL & "    , a.�V���[�Y�R�[�h "
wSQL = wSQL & "    , a.�󒍐��� "
wSQL = wSQL & "    , b.�V���[�Y�� AS ���i�� "
wSQL = wSQL & "    , c.���[�J�[�� "
wSQL = wSQL & "    , e.���J�e�S���[�R�[�h "
wSQL = wSQL & "    , e.���J�e�S���[�����{�� "
wSQL = wSQL & "    , e.�\���� AS ���J�e�S���[�\���� "
wSQL = wSQL & "    , f.��J�e�S���[�R�[�h "
wSQL = wSQL & "    , f.��J�e�S���[�� "
wSQL = wSQL & "    , 2 AS �F�K�i�ʍ݌Ɍ��� "
wSQL = wSQL & "FROM "
wSQL = wSQL & "    ���؏��i                  a WITH (NOLOCK) "
wSQL = wSQL & "      INNER JOIN �V���[�Y     b WITH (NOLOCK) "
wSQL = wSQL & "        ON     b.�V���[�Y�R�[�h     = a.�V���[�Y�R�[�h "
wSQL = wSQL & "           AND b.�J�e�S���[�R�[�h   = a.�J�e�S���[�R�[�h "
wSQL = wSQL & "      INNER JOIN ���[�J�[     c WITH (NOLOCK) "
wSQL = wSQL & "        ON     c.���[�J�[�R�[�h     = a.���[�J�[�R�[�h "
wSQL = wSQL & "      INNER JOIN �J�e�S���[   d WITH (NOLOCK) "
wSQL = wSQL & "        ON     d.�J�e�S���[�R�[�h   = a.�J�e�S���[�R�[�h "
wSQL = wSQL & "      INNER JOIN ���J�e�S���[ e WITH (NOLOCK) "
wSQL = wSQL & "        ON     e.���J�e�S���[�R�[�h = d.���J�e�S���[�R�[�h "
wSQL = wSQL & "      INNER JOIN ��J�e�S���[ f WITH (NOLOCK) "
wSQL = wSQL & "        ON     f.��J�e�S���[�R�[�h = e.��J�e�S���[�R�[�h "
wSQL = wSQL & "      LEFT JOIN ( SELECT 'Y' AS 'HotSellingY' )  t1  "
wSQL = wSQL & "        ON     d.����؃����L���O�\���t���O = t1.HotSellingY "
wSQL = wSQL & "WHERE "
wSQL = wSQL & "        t1.HotSellingY IS NOT NULL "
'wSQL = wSQL & "    AND a.�N�� = (SELECT MAX(�N��) FROM ���؏��i) "				' 2012/01/20 GV Del
wSQL = wSQL & "    AND f.��J�e�S���[�R�[�h = '" & LargeCategoryCd & "' "

wSQL = wSQL & "ORDER BY"
wSQL = wSQL & "      f.��J�e�S���[�R�[�h "
wSQL = wSQL & "    , e.�\���� "
wSQL = wSQL & "    , a.�󒍐��� DESC "
' 2012/01/19 GV Mod End
'@@@@response.write(wSQL)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

'----- ����؃����L���OHTML�ҏW
'---- Break Key initialize
If RS.EOF = True Then
    vBreakNextKey1 = "@EOF"
'    vBreakNextKey2 = "@EOF"  '2010/05/29 st del
Else
    vBreakNextKey1 = RS("���J�e�S���[�R�[�h")
'    vBreakNextKey2 = RS("��J�e�S���[�R�[�h")  '2010/05/29 st del
End If

'---- ��J�e�S���[���o���ҏW
wHTML = ""
wHTML = wHTML & "<!--  container START  -->" & vbNewLine
wHTML = wHTML & "  <div id='container'>" & vbNewLine
wHTML = wHTML & "  <h1>" & RS("��J�e�S���[��") & "</h1>" & vbNewLine


Do Until (RS.EOF = True)

'---- ��J�e�S���[���o���ҏW 2010/05/29 st del
'	if vBreakKey2 <> vBreakNextKey2 then
'		wHTML = wHTML & "  <h1>" & RS("��J�e�S���[��") & "</h1>" & vbNewLine
'		vCatCount = 0
'	end if
'2010/05/29 st del

	vBreakKey1 = vBreakNextKey1
    'vBreakKey2 = vBreakNextKey2  '2010/05/29 st del

	'---- ���J�e�S���[�ʏ��iHTML �쐬
	If vCatCount Mod 3 = 0 Then
		'---- ���ʃ^�C�g��
		wHTML = wHTML & "<!-- ranking_container START -->" & vbNewLine
		wHTML = wHTML & "    <div class='ranking_container'>" & vbNewLine
		wHTML = wHTML & "      <div class='juni'>" & vbNewLine
		wHTML = wHTML & "       <div class='juni_th'>����</div>" & vbNewLine

		'---- ���i���� 1�`10�ݒ�
		for vRank=1 to 10

			'---- �����Ɗ�Ŕw�i�F��ύX
			if vRank Mod 2 <> 0 then
				vBGColor = "juni_td bg_color1"
			else
				vBGColor = "juni_td bg_color2"
			end if

			wHTML = wHTML & "       <div class='" & vBGColor & "'>" & vRank & "��</div>" & vbNewLine

		Next

		wHTML = wHTML & "      </div>" & vbNewLine
	End If

	'---- �J�e�S���[�w�b�_
	wHTML = wHTML & "      <div class='rank_cat_box'>" & vbNewLine
	wHTML = wHTML & "        <div class='rank_cat_th'>" & vbNewLine
	wHTML = wHTML & "          <a href='MidCategoryList.asp?MidCategoryCd=" & RS("���J�e�S���[�R�[�h") & "'>" & RS("���J�e�S���[�����{��") & "</a>" & vbNewLine
	wHTML = wHTML & "        </div>" & vbNewLine

	vCatCount = vCatCount + 1
	vRank = 0

	'---- 1�`10�ʂ̏��i�쐬�@�J�e�S���[������؏��i
  Do Until (vBreakKey1 <> vBreakNextKey1)
    vRank = vRank + 1

		'----���[�J�[��+���i���Z�b�g
		vMakerProduct = RS("���[�J�[��") & ":" & RS("���i��")
		if Len(vMakerProduct) > 25 then
			vMakerProduct = Left(vMakerProduct, 22) & "..."
		end if

		'---- ���[�J�[���C���i��
		'�F�K�i�Ȃ��F���i�ʂփ����N
		'�F�K�i����F���i�ꗗ�փ����N

		'---- �����Ɗ�Ń^�O��ύX
		if vRank Mod 2 <> 0 then
			vBGColor = "rank_cat_td1"
		else
			vBGColor = "rank_cat_td2"
		end if

		wHTML = wHTML & "        <div class='" & vBGColor & "'>" & vbNewLine


		if RS("�F�K�i�ʍ݌Ɍ���") = 1 then
    		wHTML = wHTML & "          <a href='ProductDetail.asp?Item=" & Server.URLEncode(RS("���[�J�[�R�[�h") & "^" & RS("���i�R�[�h") & "^^") & "'>" & vMakerProduct & "</a>" & vbNewLine '2010/05/29 st mod
		else
			'---- �F�K�i���菤�i
			if RS("���i�R�[�h") <> "" then
	    		wHTML = wHTML & "          <a href='SearchList.asp?i_type=mp2&s_maker_cd=" & RS("���[�J�[�R�[�h") & "&s_product_cd=" & Server.URLEncode(RS("���i�R�[�h")) & "'>" & vMakerProduct & "</a>" & vbNewLine

			'---- �V���[�Y
			else
	    		wHTML = wHTML & "          <a href='SearchList.asp?i_type=se&sSeriesCd=" & RS("�V���[�Y�R�[�h") & "'>" & vMakerProduct & "</a>" & vbNewLine
			end if
		end if

		wHTML = wHTML & "        </div>" & vbNewLine

    If vRank = 10 Then
      '----���̃J�e�S���[�܂œǂݔ�΂�
      Do Until (vBreakKey1 <> vBreakNextKey1)
        RS.MoveNext
        If RS.EOF = true then
          vBreakNextKey1 = "@EOF"
'          vBreakNextKey2 = "@EOF" '2010/05/29 st del
        Else
          vBreakNextKey1 = RS("���J�e�S���[�R�[�h")
'          vBreakNextKey2 = RS("��J�e�S���[�R�[�h") '2010/05/29 st del
        End If
      Loop
    Else
      '���̃��R�[�h
      RS.MoveNext
      If RS.EOF = true then
        vBreakNextKey1 = "@EOF"
'        vBreakNextKey2 = "@EOF" '2010/05/29 st del
      Else
        vBreakNextKey1 = RS("���J�e�S���[�R�[�h")
'        vBreakNextKey2 = RS("��J�e�S���[�R�[�h")  '2010/05/29 st del
      End If
    End If
  Loop

	'---- 10�ʂ܂łȂ��ꍇ�A�󏤕i���Z�b�g
	for vRank = vRank + 1 to 10

		'---- �����Ɗ�Ń^�O��ύX
		if vRank Mod 2 <> 0 then
			vBGColor = "rank_cat_td1"
		else
			vBGColor = "rank_cat_td2"
		end if

		wHTML = wHTML & "       <div class='" & vBGColor & "'>" & vbNewLine

		wHTML = wHTML & "        </div>" & vbNewLine

	next

'---- ��J�e�S���[�u���[�N�i���J�e�S���[�^�C�g���A���i�Z�b�g)
'	if vBreakKey2 <> vBreakNextKey2 then
'		Do until vCatCount Mod 3 = 0
'			vCatHTML = vCatHTML & "<td width='225' bgcolor='#eeeeee'>�@</td>" & vbNewLine
'			for i=1 to 10
'				vProdHTML(i) = vProdHTML(i) & "<td>�@</td>" & vbNewLine
'			next
'			vCatCount = vCatCount + 1
'		Loop
'	end if

'---- 3���J�e�S���[�u���[�N�i���J�e�S���[�^�C�g���A���i�Z�b�g)
	if vCatCount Mod 3 = 0 Then
		wHTML = wHTML & "      </div>" & vbNewLine
		wHTML = wHTML & "    </div>" & vbNewLine
		wHTML = wHTML & "<!-- ranking_container END -->" & vbNewLine
	else
		wHTML = wHTML & "      </div>" & vbNewLine
	end if
Loop


if vCatCount Mod 3 <> 0 Then
wHTML = wHTML & "    </div>" & vbNewLine
wHTML = wHTML & "<!-- ranking_container END -->" & vbNewLine
end if
wHTML = wHTML & "  </div>" & vbNewLine
wHTML = wHTML & "<!-- container END -->" & vbNewLine
wHTML = wHTML & "</div>" & vbNewLine

RS.close

end if


End function

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

	wHTML = wHTML & "<a href='BestSellerList.asp?LargeCategoryCd=" & RS("��J�e�S���[�R�[�h") & "'>" & RS("��J�e�S���[��") & "</a>"

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
<title>����؏��i�b�T�E���h�n�E�X</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/Ranking.css?20120921" type="text/css">
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
      <li class="now">����؏��i</li>
    </ul>
  </div></div></div>
  <h1 class="title">����؏��i</h1>
-->

<!-- Mainpage START -->
<div id="ranking_key_main_flame">

<!-- Menu START -->
  <div id="ranking_key_top_menu">
    <div class="top_menu_parts">
      <a href="BestSellerList.asp">
      <img src="images/ranking/ts_btn_on.jpg" alt="" name="Image15" width="114" height="80" />
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
      <a href="RankingAccess.asp?RankType=<%=Server.URLEncode("�~�������̃��X�g")%>" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image14','','images/ranking/wl_btn_on.jpg',1)"><img src="images/ranking/wl_btn_off.jpg" alt="" name="Image14" width="114" height="80" /></a>
    </div>
    <div class="top_menu_parts">
      <a href="RankingReview.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image16','','images/ranking/nor_btn_on.jpg',1)">
        <img src="images/ranking/nor_btn_off.jpg" alt="" name="Image16" width="114" height="80" />
      </a>
    </div>
    <div class="top_menu_parts">
      <a href="RankingReviewPoint.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image17','','images/ranking/rr_btn_on.jpg',1)">
        <img src="images/ranking/rr_btn_off.jpg" alt="" name="Image17" width="113" height="80" />
      </a>
    </div>
  </div>
<!-- Menu END -->

<!-- ��J�e�S���[�ꗗ -->
<%=wLargeCategoryHTML%>

<!-- ����؃����L���O -->
<%=wHTML%>

  </div>
  <div id="globalSide">
    <!--#include file="../Navi/NaviSide.inc"-->
  </div>
</div>
<!--#include file="../Navi/Navibottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>