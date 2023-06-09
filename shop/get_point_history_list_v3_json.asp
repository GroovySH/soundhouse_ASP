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
<!--#include file="../3rdParty/aspJSON1.17.asp"-->
<%
'========================================================================
'
'	�w�������ꗗ�y�[�W�ɂ�����|�C���g�����擾
'
'
'�ύX����
'2016.03.11 GV �V�K�쐬�B(Web�����ύX�L�����Z���@�\)
'2016.07.14 GV �|�C���g��������ǉ��B
'2016.12.01 GV ������̏ꍇ�̒��o���������C�B
'2020.06.06 GV ������̏ꍇ�̒��o���������C�B
'2020.06.25 GV ������ȊO���u�������v��\���B(#2458)
'
'========================================================================
'On Error Resume Next

'Const PAGE_SIZE = 20			' �w����������1�y�[�W������̕\���s��

Dim Connection
Dim ConnectionEmax

Dim wErrMsg						' �G���[���b�Z�[�W (���̃y�[�W����n����郁�b�Z�[�W)
Dim wDispMsg					' �ʏ탁�b�Z�[�W(�G���[�ȊO) (���̃y�[�W����n����郁�b�Z�[�W)
Dim wErrDesc
Dim wMsg						' �G���[���b�Z�[�W (�{�y�[�W�ō쐬���郁�b�Z�[�W)
Dim wCustomerNo					' �ڋq�ԍ�
Dim wOrderNo					' �󒍔ԍ�
Dim oJSON						' JSON�I�u�W�F�N�g
Dim wPage						' �\������y�[�W�ʒu (�p�����[�^)
Dim wPageSize					'�w����������1�y�[�W������̕\���s��

'=======================================================================
'	�󂯓n�������o�� & �����ݒ�
'=======================================================================
' Get�p�����[�^
wCustomerNo = ReplaceInput(Trim(Request("cno")))
wOrderNo = ReplaceInput(Trim(Request("ono")))
wPage = ReplaceInput(Trim(Request("page")))	' �y�[�W�ʒu
wPageSize = ReplaceInput(Trim(Request("page_size")))

'�y�[�W�ԍ�
If wPage = "" Or IsNumeric(wPage) = False Then
	wPage = 1
Else
	wPage = CLng(wPage)
End If

'�y�[�W�T�C�Y
If wPageSize = "" Or IsNumeric(wPageSize) = False Then
	wPageSize = 10
Else
	wPageSize = CLng(wPageSize)
End If


'=======================================================================
'	Execute main
'=======================================================================
Call connect_db()

Call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   ' member�n�̑��̃y�[�W�����ɂȂ炤
If Err.Description <> "" Then
'	wErrDesc = THIS_PAGE_NAME & " " & Replace(Replace(Err.Description, vbCR, " "), vbLF, " ")
'	Call fSetSessionData(gSessionID, "ErrDesc", wErrDesc)
End If

Call close_db()

If Err.Description <> "" Then

End If


'========================================================================
'
'	Function	Connect database
'
'========================================================================
Function connect_db()

Set Connection = Server.CreateObject("ADODB.Connection")
Connection.Open g_connection

Set ConnectionEmax = Server.CreateObject("ADODB.Connection")
ConnectionEmax.Open g_connectionEmax

End function

'========================================================================
'
'	Function	Close database
'
'========================================================================
Function close_db()

Connection.close
Set Connection= Nothing

ConnectionEmax.close
Set ConnectionEmax= Nothing

End function

'========================================================================
'
'	Function	Main
'
'========================================================================
Function main()

Dim vSQL
Dim i '���R�[�h�Z�b�g�̃��[�v�C�e���[�^
Dim j 'JSON�̃C�e���[�^
Dim vRS
'Dim vPaymentMethod
Dim vKubun
Dim vOrderDate
Dim vPointDate
Dim vPoint
'Dim vOrderType
Dim vOrderNo
Dim vPointZan
Dim vAllPage
Dim vOffset
Dim vAdjust
Dim vPointExpire
Dim vOnoFlg
Dim maxLoop '2016.07.14 GV add
Dim addCnt  '2016.07.14 GV add

Set oJSON = New aspJSON

' ������
i         = 0
j         = 0
vPointZan = 0
vAllPage  = 0
vOffset   = 0
vAdjust   = 1
vOnoFlg   = False
maxLoop   = 0 '2016.07.14 GV add
addCnt    = 0 '2016.07.14 GV add

'--- �Y���ڋq�̃|�C���g���ׂ̎��o��
vSQL = ""
' OLD
'vSQL = vSQL & "SELECT "
'vSQL = vSQL & "  T1.�|�C���g���t AS ��������, "
'vSQL = vSQL & "  T1.�󒍔ԍ� AS �������ԍ�, "
'vSQL = vSQL & "  MAX(T2.�󒍌`��) AS ���������@, "
'vSQL = vSQL & "  MAX(T2.�x�����@) AS ���x�����@, "
'vSQL = vSQL & "  T1.�|�C���g�敪 AS �|�C���g���p�l��, "
'vSQL = vSQL & "  T1.�|�C���g���t AS �|�C���g�l����, "
'vSQL = vSQL & "  SUM(T1.�|�C���g) AS �|�C���g "
'vSQL = vSQL & "FROM "
'vSQL = vSQL & "  �|�C���g���� T1 WITH (NOLOCK) "
'vSQL = vSQL & "  LEFT JOIN �� T2 WITH (NOLOCK) "
'vSQL = vSQL & "    ON (T2.�󒍔ԍ� = T1.�󒍔ԍ� "
'vSQL = vSQL & "    AND T2.�ڋq�ԍ� = T1.�ڋq�ԍ�) "
'vSQL = vSQL & "WHERE "
'vSQL = vSQL & "  T1.�ڋq�ԍ�=" & wCustomerNo
'vSQL = vSQL & " AND T1.�|�C���g���t IS NOT NULL "
'vSQL = vSQL & "GROUP BY "
'vSQL = vSQL & "  T1.�|�C���g�敪, "
'vSQL = vSQL & "  T1.�|�C���g���t, "
'vSQL = vSQL & "  T1.�󒍔ԍ� "
'vSQL = vSQL & "HAVING "
'vSQL = vSQL & "  SUM(T1.�|�C���g) <> '0' "
'vSQL = vSQL & "ORDER BY "
'vSQL = vSQL & "  ��������, "
'vSQL = vSQL & "  �������ԍ�, "
'vSQL = vSQL & "  �|�C���g�l����, "
'vSQL = vSQL & "  �|�C���g���p�l��  DESC"

'--------------------------------------------------------
' �R�����g�A�E�g �������� 2020.06.25 GV
'--------------------------------------------------------
'��������
'vSQL = ""
'vSQL = vSQL & "SELECT "
'' 2017.07.14 GV mod start
''vSQL = vSQL & "  CONVERT(NVARCHAR, T1.�|�C���g���t, 111) AS ��������, "

''vSQL = vSQL & "  CASE "
''vSQL = vSQL & "     WHEN T1.�|�C���g���t IS NOT NULL "
''vSQL = vSQL & "       THEN CONVERT(NVARCHAR, T1.�|�C���g���t, 111) "
''vSQL = vSQL & "     WHEN T1.�|�C���g���t IS NULL "
''vSQL = vSQL & "       THEN CONVERT(NVARCHAR, T2.�o�׊�����, 111) "
''vSQL = vSQL & "  END AS ��������, "
'vSQL = vSQL & "CASE "
''  --�|�C���g���t���ݒ肳��Ă���A�l���̏ꍇ
'vSQL = vSQL & "  WHEN (T1.�|�C���g���t IS NOT NULL) AND (T1.�|�C���g�敪 = '�l��') "
'vSQL = vSQL & "    THEN "
'vSQL = vSQL & "      CASE "
'vSQL = vSQL & "        WHEN T2.�x�����@ = '�����' "
'vSQL = vSQL & "          THEN t2.�o�׊����� "
'vSQL = vSQL & "        ELSE T1.�|�C���g���t "
'vSQL = vSQL & "      END "
''  -- �|�C���g���t���ݒ肳��Ă���A�l���ȊO�̏ꍇ
'vSQL = vSQL & "  WHEN T1.�|�C���g���t IS NOT NULL AND T1.�|�C���g�敪 <> '�l��' "
'vSQL = vSQL & "    THEN T1.�|�C���g���t "
'vSQL = vSQL & "  ELSE "
''    -- ��L�ȊO�i�|�C���g���t��NULL�A�o�׊��������ݒ�j
'vSQL = vSQL & "    CASE "
'vSQL = vSQL & "      WHEN T2.�x�����@ = '�����' " '������
'vSQL = vSQL & "        THEN CONVERT(NVARCHAR, T2.�o�׊�����, 111) "
'vSQL = vSQL & "      ELSE  "
'vSQL = vSQL & "        CONVERT(NVARCHAR, T1.�|�C���g���t, 111) "
'vSQL = vSQL & "      END "
'vSQL = vSQL & "  end as ��������, "
'' 2017.07.14 GV mod end


'vSQL = vSQL & "  T1.�󒍔ԍ� AS �������ԍ�, "
''vSQL = vSQL & "  MAX(T2.�󒍌`��) AS ���������@, "
''vSQL = vSQL & "  MAX(T2.�x�����@) AS ���x�����@, "

'vSQL = vSQL & "  T1.�|�C���g�敪 AS �|�C���g���p�l��, "
'' 2017.07.14 GV add start
'vSQL = vSQL & "  CASE "
'vSQL = vSQL & "     WHEN T1.�|�C���g���t IS NULL THEN '������' "
'vSQL = vSQL & "     ELSE '������' "
'vSQL = vSQL & "  END AS �����敪, "
'' 2017.07.14 GV add end

'vSQL = vSQL & "  T1.�|�C���g���t AS �|�C���g�l����, "
'vSQL = vSQL & "  T1.�|�C���g����, "
'vSQL = vSQL & "  SUM(T1.�|�C���g) AS �|�C���g, "
'vSQL = vSQL & "  CASE T1.�|�C���g�敪 "
'vSQL = vSQL & "    WHEN '���p' THEN 0 "
'vSQL = vSQL & "    ELSE 1 "
'vSQL = vSQL & "  END AS POINT_SORT "
'vSQL = vSQL & "FROM "
'vSQL = vSQL & "  �|�C���g���� T1 WITH (NOLOCK) "

''2016.07.14 GV mod start
'vSQL = vSQL & "  LEFT JOIN �� T2 WITH (NOLOCK) "
'vSQL = vSQL & "    ON T2.�󒍔ԍ� = T1.�󒍔ԍ� "
'vSQL = vSQL & "    AND T2.�ڋq�ԍ� = T1.�ڋq�ԍ� "
''2016.07.14 GV mod end

'vSQL = vSQL & "WHERE "
'vSQL = vSQL & "  T1.�ڋq�ԍ�= " & wCustomerNo
''vSQL = vSQL & " AND T1.�|�C���g���t IS NOT NULL " ' 2016.07.14 GV mod
'' 2016.07.14 GV add start
'vSQL = vSQL & "  AND T2.�폜�� IS NULL "
'vSQL = vSQL & " AND ((T1.�|�C���g���t IS NOT NULL) "
''2016.12.01 GV mod start
''vSQL = vSQL & "   OR (T1.�|�C���g���t IS NULL AND T2.�x�����@ = '�����' AND T2.�o�׊����� IS NOT NULL)) "
''vSQL = vSQL & "   OR (T1.�|�C���g���t IS NULL AND T2.�x�����@ = '�����' AND T2.�o�׊����� IS NOT NULL AND T2.�ŏI������ IS NULL)) " ' 2020.06.06 GV mod
'vSQL = vSQL & "   OR (T1.�|�C���g���t IS NULL AND T2.�x�����@ = '�����' AND T2.�o�׊����� IS NOT NULL)) " ' 2020.06.06 GV add
''2016.12.01 GV mod end
'' 2016.07.14 GV add end

'vSQL = vSQL & "GROUP BY "
'vSQL = vSQL & "  T1.�|�C���g�敪, "
'vSQL = vSQL & "  T1.�|�C���g���t, "
'vSQL = vSQL & "  T1.�󒍔ԍ� "
'vSQL = vSQL & "  ,T1.�|�C���g���� "
'vSQL = vSQL & "  ,T2.�o�׊����� " ' 2016.07.14 GV add
'vSQL = vSQL & "  ,T2.�x�����@ " ' 2016.07.14 GV add
'vSQL = vSQL & "HAVING "
'vSQL = vSQL & "  SUM(T1.�|�C���g) <> '0' "
'vSQL = vSQL & "ORDER BY "
'vSQL = vSQL & "  ��������, "
'vSQL = vSQL & "  �������ԍ�, "
'vSQL = vSQL & "  POINT_SORT ASC"
'--------------------------------------------------------
' �R�����g�A�E�g �����܂� 2020.06.25 GV
'--------------------------------------------------------

'--------------------------------------------------------
' �V�K �������� 2020.06.25 GV
'--------------------------------------------------------
vSQL = vSQL & "SELECT "
vSQL = vSQL & "  x.�������� "
vSQL = vSQL & ",x.�������ԍ� "
vSQL = vSQL & ",x.�|�C���g���p�l�� "
vSQL = vSQL & ",x.�����敪 "
vSQL = vSQL & ",x.�|�C���g�l���� "
vSQL = vSQL & ",x.�|�C���g���� "
vSQL = vSQL & ",sum(x.�|�C���g) AS �|�C���g "
vSQL = vSQL & ",x.POINT_SORT "
vSQL = vSQL & " FROM ( "
' X�̃N�G��
vSQL = vSQL & "SELECT "
vSQL = vSQL & " CASE "
vSQL = vSQL & "   WHEN (T1.�|�C���g���t IS NOT NULL) AND (T1.�|�C���g�敪 = '�l��') "
'2020.11.16 takeuchi MOD start
'vSQL = vSQL & "     THEN "
'vSQL = vSQL & "       CASE "
'vSQL = vSQL & "         WHEN T2.�x�����@ = '�����' THEN t2.�o�׊����� "
'vSQL = vSQL & "         ELSE T1.�|�C���g���t "
'vSQL = vSQL & "       END "
vSQL = vSQL & "     THEN T1.�|�C���g���t "
'2020.11.16 takeuchi MOD end
vSQL = vSQL & "   WHEN T1.�|�C���g���t IS NOT NULL AND T1.�|�C���g�敪 <> '�l��' "
vSQL = vSQL & "     THEN T1.�|�C���g���t "
vSQL = vSQL & "   ELSE "
vSQL = vSQL & "     CASE "
vSQL = vSQL & "       WHEN T2.�x�����@ = '�����' "
vSQL = vSQL & "         THEN CONVERT(NVARCHAR, T2.�o�׊�����, 111) "
vSQL = vSQL & "       ELSE "
vSQL = vSQL & "         CONVERT(NVARCHAR, T1.�|�C���g���t, 111) "
vSQL = vSQL & "     END "
vSQL = vSQL & " END AS �������� "
vSQL = vSQL & ", T1.�󒍔ԍ� AS �������ԍ� "
vSQL = vSQL & ", T1.�|�C���g�敪 AS �|�C���g���p�l�� "
'vSQL = vSQL & ", CASE WHEN T1.�|�C���g���t IS NULL THEN '������' ELSE '������' END AS �����敪 "

vSQL = vSQL & ",T1.�g�p�󒍔ԍ� "
vSQL = vSQL & ",CASE "
vSQL = vSQL & "   WHEN T1.�|�C���g���t IS NULL THEN "
vSQL = vSQL & "     CASE "
vSQL = vSQL & "       WHEN LEFT(T1.�g�p�󒍔ԍ�, 2) = 'RA' THEN �|�C���g�敪 "
vSQL = vSQL & "       ELSE '������' "
vSQL = vSQL & "     END "
vSQL = vSQL & "   ELSE '������' END AS �����敪 "

vSQL = vSQL & ", T1.�|�C���g���t AS �|�C���g�l���� "
vSQL = vSQL & ", T1.�|�C���g���� "
vSQL = vSQL & ", SUM(T1.�|�C���g) AS �|�C���g "
vSQL = vSQL & ", CASE T1.�|�C���g�敪 WHEN '���p' THEN 0 ELSE 1 END AS POINT_SORT "
vSQL = vSQL & "FROM �|�C���g���� T1 WITH (NOLOCK) "
vSQL = vSQL & "LEFT JOIN �� T2 WITH (NOLOCK) "
vSQL = vSQL & "  ON T2.�󒍔ԍ� = T1.�󒍔ԍ� "
vSQL = vSQL & "  AND T2.�ڋq�ԍ� = T1.�ڋq�ԍ� "
vSQL = vSQL & "LEFT JOIN �󒍖��� T3 WITH (NOLOCK) "
vSQL = vSQL & " ON T3.�󒍔ԍ� = T1.�󒍔ԍ� AND T3.�󒍖��הԍ� = T1.�󒍖��הԍ� "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "  T1.�ڋq�ԍ�= " & wCustomerNo
vSQL = vSQL & "  AND T2.�폜�� IS NULL "
vSQL = vSQL & "  AND (T1.�|�C���g���t IS NOT NULL) "
vSQL = vSQL & "GROUP BY "
vSQL = vSQL & "  T1.�|�C���g�敪, T1.�|�C���g���t, T1.�󒍔ԍ� ,T1.�|�C���g����, T1.�g�p�󒍔ԍ�, T2.�o�׊�����, T2.�x�����@ "
vSQL = vSQL & "HAVING "
vSQL = vSQL & "  SUM(T1.�|�C���g) <> '0' "
vSQL = vSQL & "UNION "
vSQL = vSQL & "SELECT "
vSQL = vSQL & " CASE "
vSQL = vSQL & "   WHEN (T1.�|�C���g���t IS NOT NULL) AND (T1.�|�C���g�敪 = '�l��') "
vSQL = vSQL & "     THEN "
vSQL = vSQL & "       CASE "
vSQL = vSQL & "         WHEN T2.�x�����@ = '�����' THEN t2.�o�׊����� "
vSQL = vSQL & "         ELSE T1.�|�C���g���t "
vSQL = vSQL & "       END "
vSQL = vSQL & "   WHEN T1.�|�C���g���t IS NOT NULL AND T1.�|�C���g�敪 <> '�l��' "
vSQL = vSQL & "     THEN T1.�|�C���g���t "
vSQL = vSQL & "   ELSE "
'2020.11.16 takeuchi MOD start
'vSQL = vSQL & "     CASE "
'vSQL = vSQL & "       WHEN T2.�x�����@ = '�����' "
'vSQL = vSQL & "         THEN CONVERT(NVARCHAR, T2.�o�׊�����, 111) "
'vSQL = vSQL & "       ELSE "
'vSQL = vSQL & "         CONVERT(NVARCHAR, TA.�o�ד�, 111) "
'vSQL = vSQL & "     END "
vSQL = vSQL & "     GETDATE() "
'2020.11.16 takeuchi MOD end
vSQL = vSQL & " END AS �������� "
vSQL = vSQL & ", T1.�󒍔ԍ� AS �������ԍ� "
vSQL = vSQL & ", T1.�|�C���g�敪 AS �|�C���g���p�l�� "
'vSQL = vSQL & ", CASE WHEN T1.�|�C���g���t IS NULL THEN '������' ELSE '������' END AS �����敪 "

vSQL = vSQL & ",T1.�g�p�󒍔ԍ� "
vSQL = vSQL & ",CASE "
vSQL = vSQL & "   WHEN T1.�|�C���g���t IS NULL THEN "
vSQL = vSQL & "     CASE "
vSQL = vSQL & "       WHEN LEFT(T1.�g�p�󒍔ԍ�, 2) = 'RA' THEN �|�C���g�敪 "
vSQL = vSQL & "       ELSE '������' "
vSQL = vSQL & "     END "
vSQL = vSQL & "   ELSE '������' END AS �����敪 "

vSQL = vSQL & ", T1.�|�C���g���t AS �|�C���g�l���� "
vSQL = vSQL & ", T1.�|�C���g���� "
vSQL = vSQL & ", SUM(T1.�|�C���g) AS �|�C���g "
vSQL = vSQL & ", CASE T1.�|�C���g�敪 WHEN '���p' THEN 0 ELSE 1 END AS POINT_SORT "
vSQL = vSQL & "FROM �|�C���g���� T1 WITH (NOLOCK) "
vSQL = vSQL & "   , �� T2 WITH (NOLOCK) "
vSQL = vSQL & "   , �󒍖��� T3 WITH (NOLOCK) "
vSQL = vSQL & "     inner join (SELECT T4.�󒍔ԍ� AS �󒍔ԍ�, T4.�󒍖��הԍ� AS �󒍖��הԍ�, MAX(�o�ד�) AS �o�ד� "
vSQL = vSQL & "                   FROM �o�ז��� T4 WITH (NOLOCK) ,�o�� T5 WITH (NOLOCK) "
vSQL = vSQL & "                  WHERE T4.�o�הԍ� = T5.�o�הԍ� "
vSQL = vSQL & "                  GROUP BY T4.�󒍔ԍ� ,T4.�󒍖��הԍ�) TA "
vSQL = vSQL & "             on  T3.�󒍔ԍ� = TA.�󒍔ԍ� AND T3.�󒍖��הԍ� = TA.�󒍖��הԍ� "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "  T1.�ڋq�ԍ�= " & wCustomerNo
vSQL = vSQL & "  AND T2.�󒍔ԍ� = T1.�󒍔ԍ� AND T2.�ڋq�ԍ� = T1.�ڋq�ԍ� "
vSQL = vSQL & "  AND T3.�󒍔ԍ� = T1.�󒍔ԍ� AND T3.�󒍖��הԍ� = T1.�󒍖��הԍ� "
vSQL = vSQL & "  AND T2.�폜�� IS NULL "
vSQL = vSQL & "  AND ((T1.�|�C���g���t IS NULL AND T2.�x�����@ <> '�����' AND T3.�󒍐��� = T3.�o�׍��v���� AND T3.�󒍐��� > 0) "
vSQL = vSQL & "       OR "
vSQL = vSQL & "       (T1.�|�C���g���t IS NULL AND T2.�x�����@ = '�����' AND T2.�o�׊����� IS NOT NULL)) "
vSQL = vSQL & "GROUP BY "
vSQL = vSQL & "  T1.�|�C���g�敪, T1.�|�C���g���t, T1.�󒍔ԍ�, T1.�|�C���g����, T1.�g�p�󒍔ԍ�, T2.�o�׊�����, T2.�x�����@, TA.�o�ד� "
vSQL = vSQL & "HAVING "
vSQL = vSQL & "  SUM(T1.�|�C���g) <> '0' "
vSQL = vSQL & ") AS x "
vSQL = vSQL & " GROUP BY "
vSQL = vSQL & "  x.��������, x.�������ԍ�, x.�|�C���g���p�l��, x.�����敪, x.�|�C���g�l����, x.�|�C���g����, x.POINT_SORT "
vSQL = vSQL & "ORDER BY "
vSQL = vSQL & "  x.��������, x.�������ԍ�, x.POINT_SORT ASC"
'--------------------------------------------------------
' �V�K �����܂� 2020.06.25 GV
'--------------------------------------------------------





'@@@@Response.Write(vSQL)


Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic
If vRS.EOF = False Then

	' �S������JSON�f�[�^�ɃZ�b�g
	oJSON.data.Add "cnt" ,vRS.RecordCount

	If (wOrderNo = "") Then
		' ���X�g�ǉ�
		oJSON.data.Add "list" ,oJSON.Collection()
	End If

	'--- �w��y�[�W��\������ׂ̃��R�[�h�ʒu�t��(SearchList�̏����ɕ키)
	'vRS.PageSize = PAGE_SIZE
	'If wPage > ((vRS.RecordCount + (PAGE_SIZE - 1)) / PAGE_SIZE) Then		'MAX�y�[�W�𒴂���ꍇ�͍ŏI�y�[�W��
	'	wPage = Fix(vRS.RecordCount / PAGE_SIZE)
	'vRS.PageSize = wPageSize
'	If wPage > ((vRS.RecordCount + (wPageSize - 1)) / wPageSize) Then		'MAX�y�[�W�𒴂���ꍇ�͍ŏI�y�[�W��
'		wPage = Round((vRS.RecordCount / wPageSize) + 0.5)
'	End If
	' ���R�[�h������̃y�[�W��
	vAllPage = Round((vRS.RecordCount / wPageSize) + 0.5)

	' ���R�[�h�ʒu�̈ʒu�t��
	'vRS.AbsolutePage = wPage

	vOffset = vRS.RecordCount - (wPage * wPageSize)

'Response.Write "vRS.RecordCount=" & vRS.RecordCount & "<br>"
'Response.Write "wPageSize=" & wPageSize & "<br>"
'Response.Write "wPage=" & wPage & "<br>"
'Response.Write "vAllPage=" & vAllPage & "<br>"
'Response.Write "vOffset=" & vOffset & "<br>"
'Response.Write "a=" & fix(vRS.RecordCount / wPageSize) & "<br>"

	'�Ō�̃y�[�W�̏ꍇ
	If wPage = vAllPage Then
		vOffset = 0
		vAdjust = 2
		'���R�[�h�� - (�y�[�W�T�C�Y * fix(���R�[�h�� / �y�[�W�T�C�Y))
		maxLoop = vRS.RecordCount - (wPageSize * fix(vRS.RecordCount / wPageSize)) - 1
	Else
		vAdjust = 1
		maxLoop = wPageSize - 1
	End If
'Response.Write "vOffset=" & vOffset & "<br>"
'Response.Write "vAdjust=" & vAdjust & "<br>"
'Response.Write "maxLoop=" & maxLoop & "<br>"

	'For i = 0 To (vRS.PageSize - 1)
	For i = 0 To (vRS.RecordCount - 1)

		' --------------------------------
		' �|�C���g
		If (IsNull(vRS("�|�C���g"))) Then
			vPoint = 0
		Else
			vPoint = CStr(Trim(vRS("�|�C���g")))
		End If

		'�|�C���g�c��ݐ�
		'2016.07.14 GV mod start
		'vPointZan = vPointZan + CLng(vPoint)
		If (CStr(Trim(vRS("�����敪"))) <> "������") Then
				vPointZan = vPointZan + CLng(vPoint)
		End If
		'2016.07.14 GV mod end

		' --------------------------------

		'�󒍔ԍ��̎w�肪����ꍇ�A�����܂ł̃|�C���g�c���擾����
		If (wOrderNo <> "") Then

			If (wOrderNo = CStr(Trim(vRS("�������ԍ�")))) Then
				vOnoFlg = True
			End If

			' �w��̎󒍔ԍ��܂Ń��[�v���B���قȂ�󒍔ԍ��ɂȂ����ꍇ
			If (vOnoFlg = True) And (wOrderNo <> CStr(Trim(vRS("�������ԍ�")))) Then
				'�ݐσ|�C���g�c���P�O�ɂ��ǂ�
				vPointZan = vPointZan - CLng(vPoint)

				' ���X�g�ǉ�
				oJSON.data.Add "o_no" ,wOrderNo
				oJSON.data.Add "pt_zan" ,vPointZan

				
				'With oJSON.data("list")
				'	.Add j, oJSON.Collection()
				'	With .item(j)
				'		'.Add "o_dt" ,formatDateYYYYMMDD(vOrderDate)
				'		.Add "o_no" ,wOrderNo
				'		'.Add "kubun" ,vKubun
				'		'.Add "pt_dt" ,formatDateYYYYMMDD(vPointDate)
				'		'.Add "pt_expire" ,formatDateYYYYMMDD(vPointExpire)
				'		'.Add "pt" ,vPoint
				'		.Add "pt_zan" ,vPointZan
				'	End With
				'End With

				'���[�v�E�o
				Exit For
			End If
		Else
		'�K�v�ȃ��R�[�h�ʒu�̏ꍇ�ɁAJSON�f�[�^�𐶐�����
			'If (wPage <= vAllPage) And (i >= (vOffset)) And (i <= (vOffset + wPageSize - vAdjust)) Then
			If (wPage <= vAllPage) And (i >= (vOffset)) And (addCnt <= (maxLoop)) Then
				'��������(�|�C���g���t)
				If (IsNull(vRS("��������"))) Then
					vOrderDate = ""
				Else
					vOrderDate = CStr(Trim(vRS("��������")))
				End If

				'�������ԍ�(�󒍔ԍ�)
				If (IsNull(vRS("�������ԍ�"))) Then
					vOrderNo = ""
				Else
					vOrderNo = CStr(Trim(vRS("�������ԍ�")))
				End If

				'�|�C���g���p�l��(�|�C���g�敪)
				If (IsNull(vRS("�|�C���g���p�l��"))) Then
					vKubun = ""
				Else
					vKubun = CStr(Trim(vRS("�|�C���g���p�l��")))
				End If

				'2016.07.14 GV mod start
				If (CStr(Trim(vRS("�����敪"))) = "������") Then
					vKubun = "������"
				End If
				'2016.07.14 GV mod end


				'�|�C���g�l����(�|�C���g���t)
				If (IsNull(vRS("�|�C���g�l����"))) Then
					vPointDate = ""
				Else
					vPointDate = CStr(Trim(vRS("�|�C���g�l����")))
				End If

				'�|�C���g����
				If (IsNull(vRS("�|�C���g����"))) Then
					vPointExpire = ""
				Else
					vPointExpire = CStr(Trim(vRS("�|�C���g����")))
				End If


				' ���X�g�ǉ�
				With oJSON.data("list")
					.Add j, oJSON.Collection()
					With .item(j)
						.Add "o_dt" ,formatDateYYYYMMDD(vOrderDate)
						.Add "o_no" ,vOrderNo
						'.Add "order_type" ,vOrderType
						'.Add "payment_method" ,vPaymentMethod
						.Add "kubun" ,vKubun
						.Add "pt_dt" ,formatDateYYYYMMDD(vPointDate)
						.Add "pt_expire" ,formatDateYYYYMMDD(vPointExpire)
						.Add "pt" ,vPoint
						.Add "pt_zan" ,vPointZan
					End With
				End With

				addCnt = addCnt + 1
			End If

			' �C�e���[�^���C���N�������g
			j = j + 1
		End If

		vRS.MoveNext

		If vRS.EOF Then
			Exit For
		End If
	Next
End If

'���R�[�h�Z�b�g�����
vRS.Close

'���R�[�h�Z�b�g�̃N���A
Set vRS = Nothing

' -------------------------------------------------
' JSON�f�[�^�̕ԋp
' -------------------------------------------------
' �w�b�_�o��
Response.AddHeader "Content-Type", "application/json; charset=shift_jis"
Response.AddHeader "Cache-Control", "no-cache,must-revalidate"
Response.AddHeader "Pragma", "no-cache"

' JSON�f�[�^�̏o��
Response.Write oJSON.JSONoutput()

End Function


'========================================================================
'
'	Function	���t���̃t�H�[�}�b�g (YYYY�NMM��DD��)
'
'========================================================================
Function formatDateYYYYMMDD(pdatDate)

Dim vDate

If IsNull(pdatDate) = True Then
	' Null �͌v�Z�s�\
	Exit Function
End If

If IsDate(pdatDate) = False Then
	' ���t���łȂ���Όv�Z�s�\
	Exit Function
End If

vDate = DatePart("yyyy", pdatDate) & "/"

If DatePart("m", pdatDate) <= 9 Then
	vDate = vDate & "0" & DatePart("m", pdatDate)
Else
	vDate = vDate & DatePart("m", pdatDate)
End If

vDate = vDate & "/"

If DatePart("d", pdatDate) <= 9 Then
	vDate = vDate & "0" & DatePart("d", pdatDate)
Else
	vDate = vDate & DatePart("d", pdatDate)
End If

vDate = vDate & ""

formatDateYYYYMMDD = vDate

End Function

'========================================================================
%>
