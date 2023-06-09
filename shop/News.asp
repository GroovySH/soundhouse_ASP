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
'	News(���i�L��)�\���y�[�W
'
'�X�V����
'2007/01/23 �摜�ƋL���{���𕪊�
'2007/04/20 �VNAVI�Ή�
'2007/04/27 ��J�e�S���[�ʂ̂Ƃ�MAX30���ɁA�S���{�^��������
'2008/05/07 ���̓f�[�^�`�F�b�N����
'2008/05/23 ���̓f�[�^�`�F�b�N����
'2008/08/18 �L���敪�Ƀv���X�����[�X�ǉ�
'2008/08/19 (�ύX�˗�#478)���i�L��.���J�e�S���[�p�~�����i�L�����J�e�S���[�e�[�u���ǉ�
'2009/04/30 �G���[����error.asp�ֈړ�
'2009/11/18 an ��J�e�S���[�ɑ΂��ď����ꂽ�L�����\������悤�ɏC��
'2010/08/20 an NewsNo�w��L����<title>�������\������悤�ɏC��
'2010/11/05 an �L��No�w�莞��meta keyword,description��ǉ�
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2012/03/22 if-web �\�[�V�����{�^��(twitter, facebook)�����\��
'2012/05/31 ok #1366 ���i�L�� ���o���̃\�[�g�������L�����t����L���ԍ��ɕύX
'2012/07/19 GV #1404 NEWS�y�[�W�f�U�C���ύX
'2012/07/27 GV #1404 �[�i��̌����ɂĐ������ύX
'2012/08/06 ok �f�U�C��������
'========================================================================

On Error Resume Next

Dim NewsNo
Dim NewsDate
Dim NewsDate0		'2012/07/19 GV Add
Dim LargeCategoryCd
Dim CalenderYYYYMM
Dim CalenderYYYYMM0	'2012/07/19 GV Add
'Dim ShowAll		'2012/07/19 GV Del
Dim NewsCategory       '2008/08/18
Dim PageNo		'2012/07/19 GV Add
Dim wPageNo		'2012/07/19 GV Add
Dim wNowPage		'2012/07/19 GV Add

Dim wTitle             '2010/08/20 an add
Dim wMetaKeyword       '2010/11/05 an add
Dim wMetaDescription   '2010/11/05 an add
Dim wLargeCategoryName	'2012/07/19 GV Add
Dim iCnt		'2012/07/19 GV Add
Dim wAddParameter	'2012/07/19 GV Add
Dim wImg

Dim Connection
Dim RS

Dim wSQL
Dim wHTML
Dim xHTML		'2012/07/19 GV Add
Dim wMSG
Dim wErrDesc   '2011/08/01 an add

'========================================================================

'---- Get input data
NewsNo = ReplaceInput(Trim(Request("NewsNo")))
NewsDate = ReplaceInput(Trim(Request("NewsDate")))
LargeCategoryCd = ReplaceInput(Trim(Request("LargeCategoryCd")))
CalenderYYYYMM = ReplaceInput(Request("NaviCalenderYYYYMM"))
'ShowAll = ReplaceInput(Trim(Request("ShowAll")))	'2012/07/19 GV Del
NewsCategory = ReplaceInput(Trim(Request("NewsCategory")))  	'2008/08/18
PageNo = ReplaceInput(Trim(Request("PageNo")))	'2012/07/19 GV Add
  
if NewsNo = "" OR IsNumeric(NewsNo) = false then
	NewsNo = ""
end if

'2012/07/27 GV Mod Start
'if NewsDate = "" OR IsDate(NewsDate) = false then
if NewsDate = "" OR IsNumeric(Replace(NewsDate, "/", "")) = false then
'2012/07/27 GV Mod End
	NewsDate = ""
end if

if CalenderYYYYMM = "" OR IsNumeric(Replace(CalenderYYYYMM, "/", "")) = false then
	CalenderYYYYMM = ""
end if

'2012/07/19 GV Add Start
if PageNo = "" OR IsNumeric(PageNo) = false then
	PageNo = ""
end if
'2012/07/19 GV Add End

'---- Execute main
call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "News.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

if Err.Description <> "" then	
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

if NewsNo <> "" then
	Response.Status="301 Moved Permanently" 
	Response.AddHeader "Location", "http://www.soundhouse.co.jp/news/detail?NewsNo=" & NewsNo
end if

if LargeCategoryCd <> "" then
	Response.Status="301 Moved Permanently" 
	Response.AddHeader "Location", "http://www.soundhouse.co.jp/news/index?LargeCategoryCd=" & LargeCategoryCd
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

Dim wSELECT
Dim wFROM
Dim wWHERE
Dim wWHERE2	'2012/07/19 GV Add
Dim wORDER
Dim vDelStrings
Dim vRegEx
Dim wRecNum	'2012/07/19 GV Add
Dim RSv		'2012/07/19 GV Add
Dim RSx		'2012/07/19 GV Add
Dim vSQL	'2012/07/19 GV Add
Dim xSQL	'2012/07/19 GV Add
Dim wFromPage	'2012/07/19 GV Add
Dim wToPage	'2012/07/19 GV Add
Dim wFromRec	'2012/07/19 GV Add
Dim wToRec	'2012/07/19 GV Add
Dim wPageNum	'2012/07/19 GV Add

wSQL = ""
xSQL = ""	'2012/07/19 GV Add

Const PAGE_COUNT = 7	'2012/07/19 GV Add
Const ITEM_COUNT = 10	'2012/07/19 GV Add


'�����J�E���g
'2012/07/19 GV Add Start
xSQL = "SELECT COUNT(DISTINCT  a.�L���ԍ�)"	'�����J�E���g�pSQL
xSQL = xSQL & " FROM     (���i�L�� a WITH (NOLOCK) "
xSQL = xSQL & " LEFT JOIN ���i�L�����J�e�S���[ b WITH (NOLOCK) ON a.�L���ԍ� = b.�L���ԍ�) "
xSQL = xSQL & " LEFT JOIN ���J�e�S���[ c WITH (NOLOCK) on c.���J�e�S���[�R�[�h = b.���J�e�S���[�R�[�h "
if NewsDate <> "" then
	xSQL = xSQL & "WHERE  Year(a.�L�����t) = " & Year(NewsDate) & " "
	xSQL = xSQL & "  AND Month(a.�L�����t) = " & Month(NewsDate) & " "
	xSQL = xSQL & "  AND   Day(a.�L�����t) = " & Day(NewsDate) & " "
end if
if LargeCategoryCd <> "" then
	xSQL = xSQL & "WHERE c.��J�e�S���[�R�[�h = '" & LargeCategoryCd & "' "
	xSQL = xSQL & "   OR a.��J�e�S���[�R�[�h = '" & LargeCategoryCd & "' " 
end if
if CalenderYYYYMM <> "" then
	CalenderYYYYMM = CalenderYYYYMM & "/1"
	xSQL = xSQL & "WHERE  Year(a.�L�����t) = " & Year(CalenderYYYYMM) & " "
	xSQL = xSQL & "  AND Month(a.�L�����t) = " & Month(CalenderYYYYMM) & " "
end if

'@@@@response.write(xSQL)
Set RSx = Server.CreateObject("ADODB.Recordset")
RSx.Open xSQL, Connection, adOpenStatic
wRecNum =  RSx("")	'�Y������
wPageNum = Round(wRecNum/10+0.5)	'�S�y�[�W��
RSx.Close
'2012/07/19 GV Add End

'---- ���i�L�� ���o��
'---- select��
'2012/07/19 GV Del Start
'wSELECT = "SELECT DISTINCT a.�L���ԍ� "

'if NewsNo = "" AND NewsDate = "" AND LargeCategoryCd = "" AND NewsCategory = "" then
'	wSELECT = "SELECT DISTINCT TOP 7 a.�L���ԍ� " 
'end if

'if (LargeCategoryCd <> "" OR NewsCategory <> "") AND ShowAll <> "Y" then
'	wSELECT = "SELECT DISTINCT TOP 30 a.�L���ԍ� "
'end if
'2012/07/19 GV Del End

'2012/07/19 GV Mod Start
'wSELECT = wSELECT & "             , a.�L�����t "
'wSELECT = wSELECT & "             , a.�L���^�C�g�� "
'wSELECT = wSELECT & "             , a.�L�����e "
'wSELECT = wSELECT & "             , a.���[�J�[�R�[�h "
'wSELECT = wSELECT & "             , a.���i�R�[�h "
'wSELECT = wSELECT & "             , a.�L���摜�t�@�C��URL "
if NewsNo <> "" then		'���o�pSQL(�L���ԍ��w��)
	wSELECT = "SELECT"
	wSELECT = wSELECT & "  h.�L���ԍ�,"
	wSELECT = wSELECT & "  h.�L�����t,"
	wSELECT = wSELECT & "  h.�L���^�C�g��,"
	wSELECT = wSELECT & "  h.�L�����e,"
	wSELECT = wSELECT & "  h.���[�J�[�R�[�h,"
	wSELECT = wSELECT & "  h.���i�R�[�h,"
	wSELECT = wSELECT & "  h.�L���摜�t�@�C��URL ,"
	wSELECT = wSELECT & "  h.�L��URL ,"
	wSELECT = wSELECT & "  h.���J�e�S���[�����{��,"
	wSELECT = wSELECT & "  h.��J�e�S���[��,"
	wSELECT = wSELECT & "  h.���[�J�[��,"
	wSELECT = wSELECT & "  h.���i��"
else				'���o�pSQL(�L���ԍ��w��ȊO)
	wSELECT = "SELECT"
	wSELECT = wSELECT & "  e.�L���ԍ�"
	wSELECT = wSELECT & ", e.�L�����t"
	wSELECT = wSELECT & ", e.�L���^�C�g��"
	wSELECT = wSELECT & ", e.�L�����e"
	wSELECT = wSELECT & ", e.���[�J�[�R�[�h"
	wSELECT = wSELECT & ", e.���i�R�[�h"
	wSELECT = wSELECT & ", e.�L���摜�t�@�C��URL"
	wSELECT = wSELECT & ", e.�L��URL"
end if
'2012/07/19 GV Mod End

'2012/07/19 GV Del Start
'if NewsNo <> "" then    '2010/11/05 an add s
'	wSELECT = wSELECT & "             , d.���J�e�S���[�����{��"
'	wSELECT = wSELECT & "             , e.��J�e�S���[�� " 
'	wSELECT = wSELECT & "             , f.���[�J�[�� "
'	wSELECT = wSELECT & "             , g.���i�� "
'end if                  '2010/11/05 an add e
'2012/07/19 GV Del End

'---- from��
'---- where��(�T�u�N�G���[)
'2012/07/19 GV Mod Start
'wFROM = wFROM & "         FROM (���i�L�� a WITH (NOLOCK) "
'wFROM = wFROM & "            LEFT JOIN ���i�L�����J�e�S���[ b WITH (NOLOCK) "
'wFROM = wFROM & "            ON a.�L���ԍ� = b.�L���ԍ�) "
if NewsNo <> "" then		'FROM��(�L���ԍ��w��)
	wFROM = " FROM"
	wFROM = wFROM & "  (SELECT *,ROW_NUMBER() OVER(ORDER BY �L�����t DESC , �L���ԍ� DESC) AS �s�ԍ�"
	wFROM = wFROM & "  FROM"
	wFROM = wFROM & "    (SELECT"
	wFROM = wFROM & "    DISTINCT a.�L���ԍ� ,"
	wFROM = wFROM & "      a.�L�����t ,"
	wFROM = wFROM & "      a.�L���^�C�g�� ,"
	wFROM = wFROM & "      a.�L�����e ,"
	wFROM = wFROM & "      a.���[�J�[�R�[�h ,"
	wFROM = wFROM & "      a.���i�R�[�h ,"
	wFROM = wFROM & "      a.�L���摜�t�@�C��URL,"
	wFROM = wFROM & "      a.�L��URL,"
	wFROM = wFROM & "      c.���J�e�S���[�����{��,"
	wFROM = wFROM & "      d.��J�e�S���[��,"
	wFROM = wFROM & "      e.���[�J�[��,"
	wFROM = wFROM & "      f.���i��"
	wFROM = wFROM & "        FROM ���i�L�� a WITH (NOLOCK)"
	wFROM = wFROM & "            LEFT JOIN ���i�L�����J�e�S���[ b WITH (NOLOCK) ON a.�L���ԍ� = b.�L���ԍ�"
	wFROM = wFROM & "            LEFT JOIN ���J�e�S���[ c WITH (NOLOCK) on c.���J�e�S���[�R�[�h = a.���J�e�S���[�R�[�h"
	wFROM = wFROM & "            LEFT JOIN ��J�e�S���[ d WITH (NOLOCK) on d.��J�e�S���[�R�[�h = a.��J�e�S���[�R�[�h"
	wFROM = wFROM & "            LEFT JOIN ���[�J�[ e WITH (NOLOCK) on e.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
	wFROM = wFROM & "            LEFT JOIN Web���i f WITH (NOLOCK) on f.���[�J�[�R�[�h = a.���[�J�[�R�[�h AND f.���i�R�[�h = a.���i�R�[�h"
	wFROM = wFROM & "        WHERE a.�L���ԍ� =" & NewsNo
else				'FROM��(�L���ԍ��w��ȊO)
	wFROM = " FROM "
	wFROM = wFROM & " (SELECT *,ROW_NUMBER() OVER(ORDER BY �L�����t DESC , �L���ԍ� DESC) AS �s�ԍ�"
	wFROM = wFROM & "   FROM "
	wFROM = wFROM & "    (SELECT DISTINCT  a.�L���ԍ� , a.�L�����t , a.�L���^�C�g�� , a.�L�����e , a.���[�J�[�R�[�h , a.���i�R�[�h , a.�L���摜�t�@�C��URL , a.�L��URL "
	wFROM = wFROM & "      FROM (���i�L�� a WITH (NOLOCK) "
	wFROM = wFROM & "        LEFT JOIN ���i�L�����J�e�S���[ b WITH (NOLOCK) ON a.�L���ԍ� = b.�L���ԍ�) "
end if

if NewsDate <> "" then		'WHERE��(�N�����w��)
	wWHERE = wWHERE & "          WHERE  Year(a.�L�����t) = " & Year(NewsDate) & " "
	wWHERE = wWHERE & "            AND Month(a.�L�����t) = " & Month(NewsDate) & " "
	wWHERE = wWHERE & "            AND  Day(a.�L�����t) = " & Day(NewsDate) & " "
end if

if LargeCategoryCd <> "" then	'FROM��WHERE��(��J�e�S���w��)
	wFROM = wFROM & "              LEFT JOIN ���J�e�S���[ c WITH (NOLOCK) on c.���J�e�S���[�R�[�h = b.���J�e�S���[�R�[�h "
	wWHERE = wWHERE & "          WHERE  c.��J�e�S���[�R�[�h = '" & LargeCategoryCd & "' "
	wWHERE = wWHERE & "             OR  a.��J�e�S���[�R�[�h = '" & LargeCategoryCd & "' "
end if

if CalenderYYYYMM <> "" then	'WHERE��(�N���w��)
	wWHERE = wWHERE & "          WHERE  Year(a.�L�����t) = " & Year(CalenderYYYYMM) & " "
	wWHERE = wWHERE & "            AND Month(a.�L�����t) = " & Month(CalenderYYYYMM) & " "
end if

'---- �L���敪=�v���X�����[�X�̏ꍇ 2008/08/18
if NewsCategory <> "" then
	wWHERE = wWHERE & "        WHERE  a.�L���敪 = '" & NewsCategory & "' "
end if

if NewsNo <> "" then		'FROM��(�L���ԍ��w��)
	wFROM = wFROM & "    ) g"
	wFROM = wFROM & "  ) h"
else				'WHERE��(�L���ԍ��w��ȊO)
	wWHERE = wWHERE & "  ) d "
	wWHERE = wWHERE & ") e "
end if
'2012/07/19 GV Mod End

'---- where��
'2012/07/19 GV Del Start
'---- �ʋL���w�莞�̓��^�^�O�쐬 2010/11/05 an add s
'if NewsNo <> "" then
'	wFROM = wFROM & "            LEFT JOIN ���J�e�S���[ d WITH (NOLOCK) on d.���J�e�S���[�R�[�h = a.���J�e�S���[�R�[�h"
'	wFROM = wFROM & "           	 LEFT JOIN ��J�e�S���[ e WITH (NOLOCK) on e.��J�e�S���[�R�[�h = a.��J�e�S���[�R�[�h"
'	wFROM = wFROM & "           	 	LEFT JOIN ���[�J�[ f WITH (NOLOCK) on f.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'	wFROM = wFROM & "           	 		LEFT JOIN Web���i g WITH (NOLOCK) on g.���[�J�[�R�[�h = a.���[�J�[�R�[�h AND g.���i�R�[�h = a.���i�R�[�h"  '2010/11/05 an add e
'	wWHERE = wWHERE & "        WHERE  a.�L���ԍ� = " & NewsNo & " "
'end if

'if NewsDate <> "" then
'	wWHERE = wWHERE & "        WHERE  Year(a.�L�����t) = " & Year(NewsDate) & " "
'	wWHERE = wWHERE & "          AND Month(a.�L�����t) = " & Month(NewsDate) & " "
'	wWHERE = wWHERE & "          AND Day(a.�L�����t) = " & Day(NewsDate) & " "
'end if

'if LargeCategoryCd <> "" then
'	wFROM = wFROM & "            LEFT JOIN ���J�e�S���[ c WITH (NOLOCK) on c.���J�e�S���[�R�[�h = b.���J�e�S���[�R�[�h " '2009/11/18 an �ύX
'	wWHERE = wWHERE & "        WHERE  c.��J�e�S���[�R�[�h = '" & LargeCategoryCd & "' "
'	wWHERE = wWHERE & "           OR  a.��J�e�S���[�R�[�h = '" & LargeCategoryCd & "'" '2009/11/18 an �ǉ�
'end if

'if CalenderYYYYMM <> "" then
'	CalenderYYYYMM = CalenderYYYYMM & "/1"
'	wWHERE = wWHERE & "        WHERE  Year(a.�L�����t) = " & Year(CalenderYYYYMM) & " "
'	wWHERE = wWHERE & "          AND Month(a.�L�����t) = " & Month(CalenderYYYYMM) & " "
'end if

'---- �L���敪=�v���X�����[�X�̏ꍇ 2008/08/18
'if NewsCategory <> "" then
'	wWHERE = wWHERE & "        WHERE  a.�L���敪 = '" & NewsCategory & "' "
'end if
'2012/07/19 GV Del End

'2012/07/19 GV Add Start
'---- where��(��N�G���[)
if NewsNo = "" then	'�L���ԍ��w��ȊO�̏ꍇ
	if PageNo = "" then
		wFromRec = 1
		wToRec = ITEM_COUNT
	else
		wToRec = PageNo * ITEM_COUNT
		wFromRec = wToRec - ( ITEM_COUNT - 1 )
	end if
	wWHERE2 = "WHERE e.�s�ԍ� BETWEEN " & wFromRec & " AND " & wToRec & " "
end if
'2012/07/19 GV Add End

'---- order��
'2012/07/19 GV Mod Start
'wORDER = wORDER & "     ORDER BY a.�L�����t DESC "				'2012/05/31 ok #1366 Mod
'wORDER = wORDER & "   ,          a.�L���ԍ� DESC "				'2012/05/31 ok #1366 Add
if NewsNo <> "" then	'�L���ԍ��w��̏ꍇ
	wORDER = wORDER & "     ORDER BY h.�L�����t DESC,"
	wORDER = wORDER & "              h.�L���ԍ� DESC"
else			'�L���ԍ��w��ȊO�̏ꍇ
	wORDER = wORDER & "     ORDER BY e.�L�����t DESC "				'2012/05/31 ok #1366 Mod
	wORDER = wORDER & "   ,          e.�L���ԍ� DESC "				'2012/05/31 ok #1366 Add
end if
'2012/07/19 GV Mod End

'---- ����
wSQL = wSELECT & wFROM
'2012/07/19 GV Mod Start
'if wWHERE <> "" then
if wWHERE <> "" or wWHERE2 <> "" then
'	wSQL = wSQL & wWHERE & wORDER
	wSQL = wSQL & wWHERE & wWHERE2 & wORDER
'2012/07/19 GV Mod End
Else
	wSQL = wSQL & wORDER
end if

'@@@@response.write(wSQL)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

wHTML = ""

if RS.EOF = true then
'2012/08/06 ok Mod Start
'	wHTML = wHTML & "  <tr>" & vbNewLine
'	wHTML = wHTML & "    <td class='honbun'>�Y��News��������܂���B</td>" & vbNewLine
'	wHTML = wHTML & "  </tr>" & vbNewLine
	wHTML = wHTML & "�Y��News��������܂���B" & vbNewLine
'2012/08/06 ok Mod End
	exit function
end if

'2012/07/19 GV Add Start
'---- News�J�e�S���[�擾
vSQL = ""
vSQL = vSQL & "SELECT c.��J�e�S���[�R�[�h"
vSQL = vSQL & "     , c.��J�e�S���[��"
vSQL = vSQL & "     , c.�\����"
vSQL = vSQL & "  FROM ���i�L�� a WITH (NOLOCK)"
vSQL = vSQL & "     , ���J�e�S���[ b WITH (NOLOCK)"
vSQL = vSQL & "     , ��J�e�S���[ c WITH (NOLOCK)"
vSQL = vSQL & "     , ���i�L�����J�e�S���[ d WITH (NOLOCK)"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "        d.�L���ԍ� = a.�L���ԍ�"
vSQL = vSQL & "   AND b.���J�e�S���[�R�[�h = d.���J�e�S���[�R�[�h"
vSQL = vSQL & "   AND c.��J�e�S���[�R�[�h = b.��J�e�S���[�R�[�h"
vSQL = vSQL & "   AND c.Web��J�e�S���[�t���O = 'Y'"
vSQL = vSQL & " GROUP BY"
vSQL = vSQL & "       c.��J�e�S���[�R�[�h"
vSQL = vSQL & "     , c.��J�e�S���[��"
vSQL = vSQL & "     , c.�\���� "
vSQL = vSQL & " ORDER BY c.�\����"

'@@@@@@@@@@response.write(vSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic

if RSv.EOF = true then
	Exit Function
end if

wLargeCategoryName = "�ŐV�̃j���[�X"
Do Until RSv.EOF = true
	if LargeCategoryCd = RSv("��J�e�S���[�R�[�h") then
		wLargeCategoryName = RSv("��J�e�S���[��") & "�̃j���[�X"
	end if
	RSv.MoveNext
Loop

RSv.close

'---- �t�b�^�̕ҏW�J�n ----
xHTML = ""

'---- �L���ԍ��w��̏ꍇ(�O�̋L���ԍ��Ǝ��̋L���ԍ����擾)
if NewsNo <> "" then
	xHTML = xHTML & "<ul class='newsnavi'>" & vbNewLine
	'�O�̋L���ԍ����o��
	xSQL = "SELECT TOP 1 �L���ԍ� FROM ���i�L�� WHERE (�L���ԍ� < " & NewsNo & " AND �L�����t = '" & RS("�L�����t") & "') OR (�L�����t < '" & RS("�L�����t") & "') ORDER BY �L�����t DESC,�L���ԍ� DESC"
	Set RSx = Server.CreateObject("ADODB.Recordset")
'	response.write(xSQL)
	RSx.Open xSQL, Connection, adOpenStatic
	if RSx.EOF = true then
		xHTML = xHTML & "  <li></li>" & vbNewLine
	else
		xHTML = xHTML & "  <li><a href='News.asp?NewsNo=" & RSx("�L���ԍ�") & "'>�O�̃j���[�X��</a></li>" & vbNewLine
	end if
	RSx.close
	'���̋L���ԍ����o��
	xSQL = "SELECT TOP 1 �L���ԍ� FROM ���i�L�� WHERE (�L���ԍ� > " & NewsNo & " AND �L�����t = '" & RS("�L�����t") & "') OR (�L�����t > '" & RS("�L�����t") & "') ORDER BY �L�����t,�L���ԍ�"
	Set RSx = Server.CreateObject("ADODB.Recordset")
'		response.write(xSQL)
	RSx.Open xSQL, Connection, adOpenStatic
	if RSx.EOF = true then
		xHTML = xHTML & "  <li></li>" & vbNewLine
	else
		xHTML = xHTML & "  <li><a href='News.asp?NewsNo=" & RSx("�L���ԍ�") & "'>���̃j���[�X��</a></li>" & vbNewLine
	end if
	RSx.close
	xHTML = xHTML & "</ul>" & vbNewLine
'---- ���̑��̏ꍇ
else
	'�P�y�[�W�\�������𒴂����ꍇ
	if wRecNum > ITEM_COUNT then
		xHTML = xHTML & "<div class='pagenavi_box'>" & vbNewLine
		xHTML = xHTML & "  <ol class='pagenavi'>" & vbNewLine
		'�y�[�W�ԍ����Ȃ��ꍇ
		if PageNo = "" then
			wFromPage = 1
			wNowPage = 1
			wToPage = PAGE_COUNT
		'�y�[�W�ԍ��������y�[�W����菬�����ꍇ
		elseIf Int(PageNo) <= Int(PAGE_COUNT / 2) then
			wFromPage = 1
			wNowPage = PageNo
			wToPage = PAGE_COUNT
		'�y�[�W�ԍ��������y�[�W�����傫���ꍇ
		elseIf (wPageNum - PageNo) <= Int(PAGE_COUNT / 2) then
			wFromPage = wPageNum - PAGE_COUNT + 1
			wNowPage = PageNo
			wToPage = wPageNum
		else
			wNowPage = PageNo
			wFromPage = PageNo - Int(PAGE_COUNT / 2)
			wToPage = PageNo + Int((PAGE_COUNT - 1) / 2)
		end if
		
		'�y�[�W�J�n�I���␳
		if wFromPage < 1 then
			wFromPage = 1
		end if
		if wToPage > wPageNum then
			wToPage = wPageNum
		end if

'		'ToPage�␳
'		if wRecNum mod c1Page > 0 then
'			if wToPage > INT( wRecNum / c1Page ) + 1 then
'				wToPage = INT( wRecNum / c1Page ) + 1
'			end if
'		else
'			if wToPage > INT( wRecNum / c1Page ) then
'				wToPage = INT( wRecNum / c1Page )
'			end if
'		end if
'		if wRecNum mod c1Page > 0 then
'			if wToPage > INT( wRecNum / c1Page ) + 1 then
'				wToPage = INT( wRecNum / c1Page ) + 1
'			end if
'		else
'			if wToPage > INT( wRecNum / c1Page ) then
'				wToPage = INT( wRecNum / c1Page )
'			end if
'		end if
		'�ǉ�����p�����[�^�̕ҏW
		wAddParameter = ""
		if NewsDate <> "" then
			wAddParameter = "&NewsDate=" & NewsDate
		elseif LargeCategoryCd <> "" then
			wAddParameter = "&LargeCategoryCd=" & LargeCategoryCd
		elseif CalenderYYYYMM <> "" then
			wAddParameter = "&NaviCalenderYYYYMM=" & CalenderYYYYMM
			if Right(wAddParameter,2) = "/1" then
				wAddParameter = Left(wAddParameter,Len(wAddParameter)-2)
			end if
		else
		end if
		'�O��
		if wNowPage <> 1 then
			xHTML = xHTML & "    <li class='back'><a href='News.Asp?PageNo=" & wNowPage - 1 & wAddParameter & "'>�O��</a></li>" & vbNewLine
		end if
		'�t�b�^�y�[�W�C���f�b�N�X�쐬
		for iCnt = wFromPage to wToPage
			if iCnt = INT( wNowPage ) then
				xHTML = xHTML & "    <li><span class='now'>" & iCnt & "</span></li>" & vbNewLine
			else
				xHTML = xHTML & "    <li><a href='News.Asp?PageNo=" & iCnt & wAddParameter & "'>" & iCnt & "</a></li>" & vbNewLine
			end if
		next
		'����
		if INT( wNowPage ) <> INT( wToPage ) then
			xHTML = xHTML & "    <li class='next'><a href='News.Asp?PageNo=" & wNowPage + 1 & wAddParameter & "'>����</a></li>" & vbNewLine
		end if
		xHTML = xHTML & "  </ol>" & vbNewLine
'		xHTML = xHTML & "<span class='page'>" & wPageNum & "�y�[�W��" & wNowPage & "�y�[�W</span>" & vbNewLine		'2012/08/06 ok Del
		xHTML = xHTML & "</div>" & vbNewLine
	end if
end if
'2012/07/19 GV Add End

'---- title,metatag�p�f�[�^�m��
if NewsNo <> "" then  '2010/08/20 an add s
	'---- title
	wTitle = RS("�L���^�C�g��")
	
	'---- keyword
	wMetaKeyword = RS("��J�e�S���[��")  '2010/11/05 an add s
	
	if RS("���J�e�S���[�����{��") <> "" then
		wMetaKeyword = wMetaKeyword & "," & RS("���J�e�S���[�����{��")
	end if
	
	if RS("���[�J�[��") <> "" then
		wMetaKeyword = wMetaKeyword & "," & RS("���[�J�[��")
	end if
	
	if RS("���i��") <> "" then
		wMetaKeyword = wMetaKeyword & "," & RS("���i��")
	end if
	'---- �]�v�Ȑ擪��","������΍폜
	if Left(wMetaKeyword,1) = "," then
		wMetaKeyword =  Mid(wMetaKeyword, 2)
	end if
	
	'---- description
	if RS("�L�����e") <> "" then
		wMetaDescription = fDeleteHTMLTag(RS("�L�����e")) 'HTML
		wMetaDescription = replace(replace(replace(wMetaDescription, vbCr, ""), vbLf, ""), vbTab, "") '���s�ATab�̍폜
			
		if Len(wMetaDescription) > 100 then
			wMetaDescription = Left(wMetaDescription, 97) & "..."
		else
			wMetaDescription = Left(wMetaDescription, 100)
		end if
	end if            '2010/11/05 an add  e
	
	wImg = ""
	If RS("�L���摜�t�@�C��URL") <> "" Then
		wImg = RS("�L���摜�t�@�C��URL")
		If InStr(wImg, "http") > 0 Then
		Else
			If InStr(wImg, "../") > 0 Then
				wImg = g_HTTP & Replace(wImg, "../", "")
			Else
				wImg = g_HTTP & wImg
			End If
		End If
	End If

'2012/07/19 GV Add Start
elseif NewsDate <> "" Then
	NewsDate0 = Year( NewsDate ) & "/"
	if Len( Month( NewsDate ) ) = 1 then	'�[���p�f�B���O
		NewsDate0 = NewsDate0 & "0" & Month( NewsDate ) & "/"
	else
		NewsDate0 = NewsDate0 & Month( NewsDate ) & "/"
	end if
	if Len( Day( NewsDate ) ) = 1 then	'�[���p�f�B���O
		NewsDate0 = NewsDate0 & "0" & Day( NewsDate )
	else
		NewsDate0 = NewsDate0 & Day( NewsDate )
	end if
	wTitle = NewsDate0 & "�̃j���[�X"
elseif CalenderYYYYMM <> "" then
	CalenderYYYYMM0 = Year( CalenderYYYYMM ) & "/"
	if Len( Month( CalenderYYYYMM ) ) = 1 then	'�[���p�f�B���O
		CalenderYYYYMM0 = CalenderYYYYMM0 & "0" & Month( CalenderYYYYMM )
	else
		CalenderYYYYMM0 = CalenderYYYYMM0 & Month( CalenderYYYYMM )
	end if
	wTitle = CalenderYYYYMM0 & "�̃j���[�X"
else
	wTitle = wLargeCategoryName
'2012/07/19 GV Add End
end if                '2010/08/20 an add e


'2012/07/19 GV Add Start
'�p�������X�g
wHTML = wHTML & "    <div id='path_box'><div id='path_box_inner01'><div id='path_box_inner02'>" & vbNewLine
wHTML = wHTML & "      <p class='home'><a href='../'><img src='../images/icon_home.gif' alt='HOME'></a></p>" & vbNewLine
wHTML = wHTML & "      <ul id='path'>" & vbNewLine
wHTML = wHTML & "        <li><a href='News.asp'>�j���[�X�L���ꗗ</a></li>" & vbNewLine
wHTML = wHTML & "        <li class='now'>" & wTitle & "</li>" & vbNewLine
wHTML = wHTML & "      </ul>" & vbNewLine
wHTML = wHTML & "    </div></div></div>" & vbNewLine
'2012/07/19 GV Add End

'2012/07/19 GV Add Start
'h1�^�C�g��
if NewsNo <> "" then			'�L���ԍ��w��̏ꍇ
elseif NewsDate <> "" then		'�N�����w��̏ꍇ
	wHTML = wHTML & "    <h1 class='title'>" & NewsDate0 & "�̃j���[�X</h1>" & vbNewLine
elseif CalenderYYYYMM <> "" then	'�N���w��̏ꍇ
	wHTML = wHTML & "    <h1 class='title'>" & CalenderYYYYMM0 & "�̃j���[�X</h1>" & vbNewLine
else					'�ŐV�̃j���[�X�A�܂��́A��J�e�S���w��̏ꍇ
	wHTML = wHTML & "    <h1 class='title'>" & wLargeCategoryName & "</h1>" & vbNewLine
end if
'2012/07/19 GV Add End

'�N���X��`
wHTML = wHTML & "    <ul class='article'>" & vbNewLine	'2012/07/19 GV Add

Do until RS.EOF = true

	'2012/07/19 GV Mod Start
'	if LargeCategoryCd = "" AND NewsCategory = "" then
'		wHTML = wHTML & "  <tr>" & vbNewLine
'		wHTML = wHTML & "    <td class='honbun'>" & vbNewLine
'		wHTML = wHTML & "      <h2>" & RS("�L���^�C�g��") & "</h2>�@" & fFormatDate(RS("�L�����t"))
'		wHTML = wHTML & "    </td>" & vbNewLine
'		wHTML = wHTML & "  </tr>" & vbNewLine
'		wHTML = wHTML & "  <tr>" & vbNewLine
'		wHTML = wHTML & "    <td class='honbun' style='padding:10px 0px'>" & vbNewLine
		wHTML = wHTML & "      <li>" & vbNewLine
		wHTML = wHTML & "        <h2 class='subject'><a href='News.asp?NewsNo=" & RS("�L���ԍ�") & "'>" & RS("�L���^�C�g��") & "</a></h2>" & vbNewLine
	'2012/07/19 GV Mod End
		if RS("�L���摜�t�@�C��URL") <> "" then
			'2013/05/22 if-web mod s
			If RS("�L��URL") <> "" Then
				wHTML = wHTML & "        <a href='" & RS("�L��URL") & "'><img src='" & RS("�L���摜�t�@�C��URL") & "' alt='" & RS("�L���^�C�g��") & "' class='opover'></a>" & vbNewLine
			Else
				'2012/07/19 GV Mod Start
	'			wHTML = wHTML & "<img src='" & RS("�L���摜�t�@�C��URL") & "' width='200' border='0' align='left' style='MARGIN: 0px 5px 5px 0px' alt='" & RS("�L���^�C�g��") & "'>"
				wHTML = wHTML & "        <img src='" & RS("�L���摜�t�@�C��URL") & "' alt='" & RS("�L���^�C�g��") & "'>" & vbNewLine
				'2012/07/19 GV Mod End
			End If
			'2013/05/22 if-web mod e
		end if
		wHTML = wHTML & "        <p class='date'>" & fFormatDate(RS("�L�����t")) & "</p>" & vbNewLine	'2012/07/19 GV Add

		if ISNULL(RS("�L�����e")) = false then
			'2012/07/19 GV Mod Start
'			wHTML = wHTML & Replace(RS("�L�����e"), vbNewLine, "<br>") & vbNewLine
			wHTML = wHTML & "        <p>" & Replace(RS("�L�����e"), vbNewLine, "<br>") & "</p>" & vbNewLine
			'2012/07/19 GV Mod End
'2012/03/22 if-web add start
			'2012/07/19 GV Mod Start
'			wHTML = wHTML & "      <ul class='news_smbtn'>" & vbNewLine
'			wHTML = wHTML & "        <li><a href='http://twitter.com/share' class='twitter-share-button' data-url='http://www.soundhouse.co.jp/shop/News.asp?NewsNo=" & RS("�L���ԍ�") & "' data-text='" & RS("�L���^�C�g��") & "' data-count='horizontal' data-via='soundhouse_jp' data-lang='ja'>Tweet</a></li>" & vbNewLine
'			wHTML = wHTML & "        <li><a name='fb_share' share_url='http://www.soundhouse.co.jp/shop/News.asp?NewsNo=" & RS("�L���ԍ�") & "'>�V�F�A����</a></li>" & vbNewLine
'			wHTML = wHTML & "      </ul>" & vbNewLine
			wHTML = wHTML & "        <ul class='sns'>" & vbNewLine
			wHTML = wHTML & "          <li><a href='http://twitter.com/share' class='twitter-share-button' data-url='http://www.soundhouse.co.jp/shop/News.asp?NewsNo=" & RS("�L���ԍ�") & "' data-text='" & RS("�L���^�C�g��") & "' data-count='horizontal' data-via='soundhouse_jp' data-lang='ja'>Tweet</a></li>" & vbNewLine
			wHTML = wHTML & "          <li><iframe src='//www.facebook.com/plugins/like.php?href=http%3A%2F%2Fwww.soundhouse.co.jp%2Fshop%2FNews.asp%3FNewsNo%3D" & RS("�L���ԍ�") & "&amp;send=false&amp;layout=button_count&amp;width=100&amp;show_faces=false&amp;action=like&amp;colorscheme=light&amp;font&amp;height=21&amp;appId=191447484218062' scrolling='no' frameborder='0' style='border:none; overflow:hidden; width:110px; height:21px;' allowTransparency='true'></iframe></li>" & vbNewLine
			wHTML = wHTML & "        </ul>" & vbNewLine
			'2012/07/19 GV Mod End
'2012/03/22 if-web add end
		end if

		'2012/07/19 GV Del Start
'		wHTML = wHTML & "    </td>" & vbNewLine
'		wHTML = wHTML & "  </tr>" & vbNewLine

'		wHTML = wHTML & "  <tr>" & vbNewLine
'		wHTML = wHTML & "    <td colSpan='5' height='5'><hr size='1'></td>" & vbNewLine
'		wHTML = wHTML & "  </tr>" & vbNewLine
'		'2012/07/19 GV Del End

	'2012/07/19 GV Del Start
'	else
'		wHTML = wHTML & "  <tr>" & vbNewLine
'		wHTML = wHTML & "    <td class='honbun'>" & vbNewLine
'		wHTML = wHTML & fFormatDate(RS("�L�����t")) & " <a href='News.asp?NewsNo=" & RS("�L���ԍ�") & "' class='link'>" & RS("�L���^�C�g��") & "</a>"
'		wHTML = wHTML & "    </td>" & vbNewLine
'		wHTML = wHTML & "  </tr>" & vbNewLine
'	end if
	'2012/07/19 GV Del End
	wHTML = wHTML & "      </li>" & vbNewLine
	RS.MoveNext
Loop
wHTML = wHTML & "    </ul>" & vbNewLine	'2012/07/19 GV Add

'---- �w���̃J�e�S���[�̋L����S�ĕ\������x��URL�쐬
'2012/07/19 GV Del Start
'if ShowAll <> "Y" then
	'---- ��J�e�S���[�R�[�h�̏ꍇ
'	if LargeCategoryCd <> "" then
'		wHTML = wHTML & "  <tr>" & vbNewLine
'		wHTML = wHTML & "    <td class='honbun'><br><a href='News.asp?LargeCategoryCd=" & LargeCategoryCd & "&ShowAll=Y' class='link'><b>���̃J�e�S���[�̋L����S�ĕ\������</b></a></td>" & vbNewLine
'		wHTML = wHTML & "  </tr>" & vbNewLine
'		exit function
'	end if
	'---- �L���敪����ʋL���A�ʋL���ȊO�̏ꍇ 2008/08/18
'	if NewsCategory <> "" then
'		wHTML = wHTML & "  <tr>" & vbNewLine
'		wHTML = wHTML & "    <td class='honbun'><br><a href='News.asp?NewsCategory=" & NewsCategory & "&ShowAll=Y' class='link'><b>���̃J�e�S���[�̋L����S�ĕ\������</b></a></td>" & vbNewLine
'		wHTML = wHTML & "  </tr>" & vbNewLine
'		exit function
'	end if
'end if
'2012/07/19 GV Del End

RS.Close

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
<html xmlns:og="http://ogp.me/ns#">
<head>
<meta charset="Shift_JIS">
<title><%=wTitle%>�b�T�E���h�n�E�X</title>
<% if wMetaDescription <> "" then %>
<meta name="description" content="<%=wMetaDescription%>">
<% end if%>
<% if wMetaKeyword <> "" then %>
<meta name="keywords" content="<%=wMetaKeyword%>">
<% end if%>
<% If NewsNo <> "" Then %>
<meta name="twitter:card" content="summary">
<meta name="twitter:site" content="@soundhouse_jp">
<meta property="og:title" content="<%=wTitle%>">
<meta property="og:type" content="article">
<meta property="og:description" content="<%=wMetaDescription%>">
<% If wImg <> "" Then %>
<meta property="og:image" content="<%=wImg%>">
<% End If %>
<meta property="og:url" content="<%=g_HTTP%>shop/News.asp?NewsNo=<%=NewsNo%>">
<% End If %>
<!--#include file="../Navi/NaviStyle.inc"-->
<link href="style/news.css?20140618" rel="stylesheet" type="text/css">

</head>

<body>
<!--#include file="../Navi/Navitop.inc"-->

<div id="globalMain">
	<span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="��������{���ł�"></a></span>
	<!-- �R���e���cstart -->
	<div id="globalContents">
<!-- �j���[�X -->
<%=wHTML%>
<%=xHTML%>
	<!--/#contents --></div>
	<div id="globalSide">
<!--#include file="../Navi/NaviLeftNews.inc"-->
<!--#include file="../Navi/NaviSide.inc"-->
	<!--/#globalSide --></div>
<!--/#main --></div>
<!--#include file="../Navi/NaviBottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
<script type="text/javascript" src="http://platform.twitter.com/widgets.js" charset="utf-8"></script>
</body>
</html>
<%
call close_db()
%>
