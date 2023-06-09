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
'2021.08.04 GV �V�K�쐬�B(get_point_history_list_v3_json�𓥏P)(#2859)
'2021.08.27 GV ���ς����Ԃ̒����͏��O����悤���C�B(#2909)
'2021.09.03 GV �l���\��|�C���g�̃N�G���C���B(#2921)
'2022.01.07 GV ���l�����C�Ή��B(#3040)
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
Dim wStatus						'PHP���瑗�M���ꂽ�|�C���g�敪�p�����[�^
Dim wPointKubun					'�N�G���ɑg�ݍ��ރ|�C���g�敪

'=======================================================================
'	�󂯓n�������o�� & �����ݒ�
'=======================================================================
' Get�p�����[�^
wCustomerNo = ReplaceInput(Trim(Request("cno")))
wOrderNo = ReplaceInput(Trim(Request("ono")))
wPage = ReplaceInput(Trim(Request("page")))	' �y�[�W�ʒu
wPageSize = ReplaceInput(Trim(Request("page_size")))
wStatus = ReplaceInput(Trim(Request("status")))

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

'�X�e�[�^�X
If wStatus = "" Or IsNumeric(wStatus) = False Then
	wStatus = 0
Else
	wStatus = CLng(wStatus)
End If

' �|�C���g�敪
If wStatus = 1 Then
	wPointKubun = "�l���\��"
ElseIf wStatus = 2 Then
	wPointKubun = "�l��"
ElseIf wStatus = 3 Then
	wPointKubun = "���p"
ElseIf wStatus = 4 Then
	wPointKubun = "����"
ElseIf wStatus = 99 Then
	wPointKubun = "'�ԕi', '�ԋ�', '����', '�s��'"
Else
	wPointKubun = ""
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
Dim vKubun1
Dim vKubun2
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
Dim makerName
Dim itemName
Dim itemId
Dim shipCompDate
Dim kakutokuYoteiPoint
Dim standardDate
Dim webItem

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
makerName = ""
itemName = ""
itemId = ""
shipCompDate = ""
kakutokuYoteiPoint = 0
webItem = ""

'--- �Y���ڋq�̃|�C���g���ׂ̎��o��
vSQL = ""
' �S�̂̊l���\��|�C���g
vSQL = vSQL & "SELECT "
vSQL = vSQL & "  sum(T1.�|�C���g) AS �|�C���g "
vSQL = vSQL & "FROM "
vSQL = vSQL & "  �|�C���g���� T1 WITH (NOLOCK) "
vSQL = vSQL & "INNER JOIN �� T2 WITH (NOLOCK) "'2021.08.27 GV add
vSQL = vSQL & "  ON  T2.�󒍔ԍ� = T1.�󒍔ԍ� "'2021.08.27 GV add
vSQL = vSQL & "  AND T2.�ڋq�ԍ� = T1.�ڋq�ԍ� "'2021.08.27 GV add
vSQL = vSQL & "WHERE "
vSQL = vSQL & "  T1.�ڋq�ԍ� = " & wCustomerNo
vSQL = vSQL & " AND "
vSQL = vSQL & "  T1.�|�C���g�敪 = '�l��' "
vSQL = vSQL & "AND "
vSQL = vSQL & "  T1.�|�C���g���t IS NULL "
vSQL = vSQL & "AND " '2021.09.02 GV add
vSQL = vSQL & "  ((T1.�g�p�󒍔ԍ� NOT LIKE 'RA%') OR (T1.�g�p�󒍔ԍ� IS NULL)) " '2021.09.02 GV add
vSQL = vSQL & "AND "
vSQL = vSQL & "  T2.�󒍓� IS NOT NULL " '2021.08.27 GV add

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic
If vRS.EOF = False Then
		If (IsNull(vRS("�|�C���g"))) Then
			kakutokuYoteiPoint =  0
		Else
			kakutokuYoteiPoint = CStr(Trim(vRS("�|�C���g")))
		End If
End If

'���R�[�h�Z�b�g�����
vRS.Close

'���R�[�h�Z�b�g�̃N���A
Set vRS = Nothing

'JSON�f�[�^�ɒǉ�
oJSON.data.Add "obtain_yotei_pt" ,kakutokuYoteiPoint



'--------------------------------------------------------
' �������� 2021.07.09 GV
' ��ver. �� get_point_history_list_v3_json.asp ���Q�l
'--------------------------------------------------------
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "  x.��� "
vSQL = vSQL & " ,x.�\���p��� "
vSQL = vSQL & " ,x.�o�׊����� "
vSQL = vSQL & " ,x.�󒍔ԍ� "
vSQL = vSQL & " ,x.�󒍖��הԍ� "
vSQL = vSQL & " ,x.���[�J�[�� "
vSQL = vSQL & " ,x.���i�� "
vSQL = vSQL & " ,x.�F "
vSQL = vSQL & " ,x.�K�i "
vSQL = vSQL & " ,x.���iID "
vSQL = vSQL & " ,x.Web���i�t���O "
vSQL = vSQL & " ,x.�|�C���g���p�l�� "
vSQL = vSQL & " ,x.�|�C���g�敪 "
vSQL = vSQL & " ,x.�g�p�󒍔ԍ� "
vSQL = vSQL & " ,x.�����敪 "
vSQL = vSQL & " ,x.�|�C���g�l���� "
vSQL = vSQL & " ,x.�|�C���g���� "
vSQL = vSQL & " ,x.�|�C���g "
vSQL = vSQL & " ,x.POINT_SORT "
vSQL = vSQL & "FROM "
vSQL = vSQL & "(SELECT "
vSQL = vSQL & "  T2.���ϓ� AS ��� "
vSQL = vSQL & " ,CONVERT(NVARCHAR, T2.���ϓ�, 111) AS �\���p��� "
vSQL = vSQL & " ,CONVERT(NVARCHAR, T2.�o�׊�����, 111) as �o�׊����� "
vSQL = vSQL & " ,T1.�󒍔ԍ� AS �󒍔ԍ� "
vSQL = vSQL & " ,T3.�󒍖��הԍ� "
vSQL = vSQL & " ,mk.���[�J�[�� "
vSQL = vSQL & " ,T3.���i�� "
vSQL = vSQL & " ,T3.�F "
vSQL = vSQL & " ,T3.�K�i "
vSQL = vSQL & " ,z.���iID "
vSQL = vSQL & " ,i.Web���i�t���O "
vSQL = vSQL & " ,T1.�|�C���g�敪 AS �|�C���g���p�l�� "
vSQL = vSQL & " ,(CASE "
vSQL = vSQL & "     WHEN T1.�|�C���g���t IS NULL AND T1.�|�C���g�敪 = '�l��' "
vSQL = vSQL & "       THEN '�l���\��' "
vSQL = vSQL & "     WHEN  T1.�|�C���g���t IS NOT NULL AND T1.�|�C���g�敪 = '�l��' "
vSQL = vSQL & "       THEN '�l��' "
vSQL = vSQL & "     ELSE T1.�|�C���g�敪 "
vSQL = vSQL & "   END) AS �|�C���g�敪 "
vSQL = vSQL & " ,T1.�g�p�󒍔ԍ� "
vSQL = vSQL & " ,(CASE "
vSQL = vSQL & "     WHEN T1.�|�C���g���t IS NULL THEN "
vSQL = vSQL & "       CASE "
vSQL = vSQL & "         WHEN LEFT(T1.�g�p�󒍔ԍ�, 2) = 'RA' THEN '-' "
vSQL = vSQL & "         ELSE '������' "
vSQL = vSQL & "       END "
vSQL = vSQL & "     ELSE '������' "
vSQL = vSQL & "   END) AS �����敪 "
vSQL = vSQL & " ,T1.�|�C���g���t AS �|�C���g�l���� "
vSQL = vSQL & " ,T1.�|�C���g���� "
vSQL = vSQL & " ,SUM(T1.�|�C���g) AS �|�C���g "
vSQL = vSQL & " ,CASE T1.�|�C���g�敪 WHEN '���p' THEN 0 ELSE 1 END AS POINT_SORT "
vSQL = vSQL & "FROM "
vSQL = vSQL & "  �|�C���g���� T1 WITH (NOLOCK) "
vSQL = vSQL & "LEFT JOIN �� T2 WITH (NOLOCK) "
vSQL = vSQL & "  ON  T2.�󒍔ԍ� = T1.�󒍔ԍ� "
vSQL = vSQL & "  AND T2.�ڋq�ԍ� = T1.�ڋq�ԍ� "
vSQL = vSQL & "LEFT JOIN �󒍖��� T3 WITH (NOLOCK) "
vSQL = vSQL & " ON  T3.�󒍔ԍ�     = T1.�󒍔ԍ� "
vSQL = vSQL & " AND T3.�󒍖��הԍ� = T1.�󒍖��הԍ� "
vSQL = vSQL & "LEFT JOIN ���[�J�[ mk WITH (NOLOCK) "
vSQL = vSQL & " ON mk.���[�J�[�R�[�h = T3.���[�J�[�R�[�h "
vSQL = vSQL & "LEFT JOIN �F�K�i�ʍ݌� z WITH (NOLOCK) "
vSQL = vSQL & "  ON z.���[�J�[�R�[�h = T3.���[�J�[�R�[�h "
vSQL = vSQL & " AND z.���i�R�[�h     = T3.���i�R�[�h "
vSQL = vSQL & " AND z.�F             = T3.�F "
vSQL = vSQL & " AND z.�K�i           = T3.�K�i "
vSQL = vSQL & "LEFT JOIN ���i i WITH (NOLOCK) "
vSQL = vSQL & "  ON i.���[�J�[�R�[�h = z.���[�J�[�R�[�h "
vSQL = vSQL & " AND i.���i�R�[�h     = z.���i�R�[�h "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "  T1.�ڋq�ԍ� = " & wCustomerNo
vSQL = vSQL & "  AND T1.�󒍔ԍ� > 0 "
vSQL = vSQL & "  AND T2.�󒍓� IS NOT NULL " '2021.08.27 GV add

'2021.09.02 GV add start
vSQL = vSQL & "  AND "
vSQL = vSQL & "  (CASE "
vSQL = vSQL & "     WHEN "
vSQL = vSQL & "       LEFT(T1.�g�p�󒍔ԍ�, 2) = 'RA' "
vSQL = vSQL & "       THEN "
vSQL = vSQL & "         CASE "
vSQL = vSQL & "           WHEN T1.�|�C���g���t IS NOT NULL THEN 1 "
vSQL = vSQL & "           ELSE 0 "
vSQL = vSQL & "         END "
vSQL = vSQL & "     ELSE 1 "
vSQL = vSQL & "   END) = 1 "
'2021.09.02 GV add end

vSQL = vSQL & "GROUP BY "
vSQL = vSQL & " T1.�󒍔ԍ�,T1.�|�C���g�敪,T1.�|�C���g���t,T1.�g�p�󒍔ԍ�,T1.�|�C���g���� "
vSQL = vSQL & " ,T2.�󒍓�, T2.���ϓ�, T2.�o�׊�����, T3.�󒍖��הԍ� "
vSQL = vSQL & " ,T3.���i��, T3.�F, T3.�K�i,z.���iID, i.Web���i�t���O "
vSQL = vSQL & " ,mk.���[�J�[�� "
vSQL = vSQL & "HAVING "
vSQL = vSQL & "  SUM(T1.�|�C���g) <> '0' "
vSQL = vSQL & "UNION "
vSQL = vSQL & "SELECT "
vSQL = vSQL & "  T1.�|�C���g���t AS ��� "
vSQL = vSQL & " ,CONVERT(NVARCHAR, T1.�|�C���g���t, 111) as �\���p��� "
vSQL = vSQL & " ,CONVERT(NVARCHAR, T1.�|�C���g���t, 111) as �o�׊����� "
vSQL = vSQL & " ,T1.�󒍔ԍ� AS �󒍔ԍ� "
vSQL = vSQL & " ,T1.�󒍖��הԍ� "
vSQL = vSQL & " ,'' AS ���[�J�[�� "
'vSQL = vSQL & " ,T1.���l AS ���i�� " ' 2022.01.07 GV mod
vSQL = vSQL & " ,CASE T1.�|�C���g�敪 WHEN '����' THEN '' ELSE T1.���l END AS ���i�� " ' 2022.01.07 GV mod
vSQL = vSQL & " ,'' AS �F "
vSQL = vSQL & " ,'' AS �K�i "
vSQL = vSQL & " ,NULL AS ���iID "
vSQL = vSQL & " ,'' AS Web���i�t���O "
vSQL = vSQL & " ,T1.�|�C���g�敪 AS �|�C���g���p�l�� "
vSQL = vSQL & " ,(CASE "
vSQL = vSQL & "     WHEN T1.�|�C���g���t IS NULL AND T1.�|�C���g�敪 = '�l��' "
vSQL = vSQL & "       THEN '�l���\��' "
vSQL = vSQL & "     WHEN  T1.�|�C���g���t IS NOT NULL AND T1.�|�C���g�敪 = '�l��' "
vSQL = vSQL & "       THEN '�l��' "
vSQL = vSQL & "     ELSE T1.�|�C���g�敪 "
vSQL = vSQL & "   END) AS �|�C���g�敪 "
vSQL = vSQL & " ,T1.�g�p�󒍔ԍ� "
vSQL = vSQL & " ,(CASE "
vSQL = vSQL & "     WHEN T1.�|�C���g���t IS NULL THEN "
vSQL = vSQL & "       CASE "
vSQL = vSQL & "         WHEN LEFT(T1.�g�p�󒍔ԍ�, 2) = 'RA' THEN '-' "
vSQL = vSQL & "         ELSE '������' "
vSQL = vSQL & "       END "
vSQL = vSQL & "     ELSE '������' "
vSQL = vSQL & "   END) AS �����敪 "
vSQL = vSQL & " ,T1.�|�C���g���t AS �|�C���g�l���� "
vSQL = vSQL & " ,T1.�|�C���g���� "
vSQL = vSQL & " ,SUM(T1.�|�C���g) AS �|�C���g "
vSQL = vSQL & " ,CASE T1.�|�C���g�敪 WHEN '���p' THEN 0 ELSE 1 END AS POINT_SORT "
vSQL = vSQL & "FROM "
vSQL = vSQL & "  �|�C���g���� T1 WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "  T1.�ڋq�ԍ� =" & wCustomerNo
vSQL = vSQL & "  AND T1.�󒍔ԍ� = 0 "

'2021.09.02 GV add start
vSQL = vSQL & " AND "
vSQL = vSQL & " (CASE "
vSQL = vSQL & "     WHEN "
vSQL = vSQL & "       LEFT(T1.�g�p�󒍔ԍ�, 2) = 'RA' "
vSQL = vSQL & "       THEN "
vSQL = vSQL & "         CASE "
vSQL = vSQL & "           WHEN T1.�|�C���g���t IS NOT NULL THEN 1 "
vSQL = vSQL & "           ELSE 0 "
vSQL = vSQL & "         END "
vSQL = vSQL & "     ELSE 1 "
vSQL = vSQL & "   END) = 1 "
'2021.09.02 GV add end

vSQL = vSQL & "GROUP BY "
vSQL = vSQL & "  T1.�|�C���g���t,T1.�󒍔ԍ�,T1.�󒍖��הԍ�,T1.���l,T1.�|�C���g�敪,T1.�g�p�󒍔ԍ�,�|�C���g���� "
vSQL = vSQL & "HAVING "
vSQL = vSQL & "  SUM(T1.�|�C���g) <> '0' "
vSQL = vSQL & ") as x "
vSQL = vSQL & "WHERE "
vSQL = vSQL & " 1 = 1 "

If wPointKubun <> "" THen
	If wStatus = 99 Then
		vSQL = vSQL & " AND x.�|�C���g�敪 IN (" & wPointKubun & ") "
	Else
		vSQL = vSQL & " AND x.�|�C���g�敪 = '" & wPointKubun & "' "
	End If
End If


vSQL = vSQL & "ORDER BY "
vSQL = vSQL & "  x.���, x.�󒍔ԍ�, x.�󒍖��הԍ� DESC, x.POINT_SORT ASC "
'--------------------------------------------------------
' �����܂� 2021.07.09 GV
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
			If (wOrderNo = CStr(Trim(vRS("�󒍔ԍ�")))) Then
				vOnoFlg = True
			End If

			' �w��̎󒍔ԍ��܂Ń��[�v���B���قȂ�󒍔ԍ��ɂȂ����ꍇ
			If (vOnoFlg = True) And (wOrderNo <> CStr(Trim(vRS("�󒍔ԍ�")))) Then
				'�ݐσ|�C���g�c���P�O�ɂ��ǂ�
				vPointZan = vPointZan - CLng(vPoint)

				' ���X�g�ǉ�
				oJSON.data.Add "o_no" ,wOrderNo
				oJSON.data.Add "pt_zan" ,vPointZan

				'���[�v�E�o
				Exit For
			End If
		Else
		'�K�v�ȃ��R�[�h�ʒu�̏ꍇ�ɁAJSON�f�[�^�𐶐�����
			If (wPage <= vAllPage) And (i >= (vOffset)) And (addCnt <= (maxLoop)) Then
				'���
				standardDate = CStr(Trim(vRS("�\���p���")))
	
				'�o�׊�����
				If (IsNull(vRS("�o�׊�����"))) Then
					shipCompDate = ""
				Else
					shipCompDate = CStr(Trim(vRS("�o�׊�����")))
				End If

				'�󒍔ԍ�
				If (IsNull(vRS("�󒍔ԍ�"))) Then
					vOrderNo = ""
				Else
					vOrderNo = CStr(Trim(vRS("�󒍔ԍ�")))
				End If

				'���[�J�[��
				If (IsNull(vRS("���[�J�[��"))) Then
					makerName = ""
				Else
					makerName = CStr(Trim(vRS("���[�J�[��")))
				End If

				'���i��
				If (IsNull(vRS("���i��"))) Then
					itemName = ""
				Else
					itemName = CStr(Trim(vRS("���i��")))

					'�F
					If (IsNull(vRS("�F"))) Then
					ElseIF Trim(vRS("�F")) <> "" Then
						itemName = itemName & " / " & CStr(Trim(vRS("�F")))
					End If

					'�K�i
					If (IsNull(vRS("�K�i"))) Then
					ElseIF Trim(vRS("�K�i")) <> "" Then
						itemName = itemName & " / " & CStr(Trim(vRS("�K�i")))
					End If
				End If

				'���iID
				If (IsNull(vRS("���iID"))) Then
					itemId = ""
				Else
					itemId = CStr(Trim(vRS("���iID")))
				End If

				'Web���i�t���O
				If (IsNull(vRS("Web���i�t���O"))) Then
					webItem = ""
				Else
					webItem = CStr(Trim(vRS("Web���i�t���O")))
				End If


				'�|�C���g���p�l��(�|�C���g�敪)
				If (IsNull(vRS("�|�C���g���p�l��"))) Then
					vKubun1 = ""
				Else
					vKubun1 = CStr(Trim(vRS("�|�C���g���p�l��")))
				End If

				If (CStr(Trim(vRS("�����敪"))) = "������") Then
					vKubun1 = "������"
				End If

				'�|�C���g���p�l��(�|�C���g�敪)
				If (IsNull(vRS("�|�C���g�敪"))) Then
					vKubun2 = ""
				Else
					vKubun2 = CStr(Trim(vRS("�|�C���g�敪")))
				End If

				'�|�C���g�l����
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
						.Add "std_dt" ,standardDate
						.Add "ship_comp_dt" ,formatDateYYYYMMDD(shipCompDate)
						.Add "o_dt" ,formatDateYYYYMMDD(vOrderDate)
						.Add "o_no" ,vOrderNo
						'.Add "order_type" ,vOrderType
						'.Add "payment_method" ,vPaymentMethod
						.Add "kubun1" ,vKubun1
						.Add "kubun2" ,vKubun2
						.Add "m_name" ,makerName
						.Add "i_name" ,itemName
						.Add "i_id" ,itemId
						.Add "web_flag" ,webItem
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
