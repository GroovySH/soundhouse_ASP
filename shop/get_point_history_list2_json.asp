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
'2016.02.10 GV �V�K�쐬
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
Dim i
Dim j
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

Set oJSON = New aspJSON

' ������
i         = 0
j         = 0
vPointZan = 0
vAllPage  = 0
vOffset   = 0
vAdjust   = 1

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

'��������
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "  CONVERT(NVARCHAR, T1.�|�C���g���t, 111) AS ��������, "
vSQL = vSQL & "  T1.�󒍔ԍ� AS �������ԍ�, "
'vSQL = vSQL & "  MAX(T2.�󒍌`��) AS ���������@, "
'vSQL = vSQL & "  MAX(T2.�x�����@) AS ���x�����@, "
vSQL = vSQL & "  T1.�|�C���g�敪 AS �|�C���g���p�l��, "
vSQL = vSQL & "  T1.�|�C���g���t AS �|�C���g�l����, "
vSQL = vSQL & "  T1.�|�C���g����, "
vSQL = vSQL & "  SUM(T1.�|�C���g) AS �|�C���g, "
vSQL = vSQL & "  CASE T1.�|�C���g�敪 "
vSQL = vSQL & "    WHEN '���p' THEN 0 "
vSQL = vSQL & "    ELSE 1 "
vSQL = vSQL & "  END AS POINT_SORT "
vSQL = vSQL & "FROM "
vSQL = vSQL & "  �|�C���g���� T1 WITH (NOLOCK) "
'vSQL = vSQL & "  LEFT JOIN �� T2 WITH (NOLOCK) "
'vSQL = vSQL & "    ON (T2.�󒍔ԍ� = T1.�󒍔ԍ� "
'vSQL = vSQL & "    AND T2.�ڋq�ԍ� = T1.�ڋq�ԍ�) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "  T1.�ڋq�ԍ�= " & wCustomerNo
vSQL = vSQL & " AND T1.�|�C���g���t IS NOT NULL "
vSQL = vSQL & "GROUP BY "
vSQL = vSQL & "  T1.�|�C���g�敪, "
vSQL = vSQL & "  T1.�|�C���g���t, "
vSQL = vSQL & "  T1.�󒍔ԍ� "
vSQL = vSQL & "  ,T1.�|�C���g���� "
vSQL = vSQL & "HAVING "
vSQL = vSQL & "  SUM(T1.�|�C���g) <> '0' "
vSQL = vSQL & "ORDER BY "
vSQL = vSQL & "  ��������, "
vSQL = vSQL & "  �������ԍ�, "
vSQL = vSQL & "  POINT_SORT ASC"


'@@@@Response.Write(vSQL)

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic
If vRS.EOF = False Then

	' �S������JSON�f�[�^�ɃZ�b�g
	oJSON.data.Add "count" ,vRS.RecordCount

	' ���X�g�ǉ�
	oJSON.data.Add "list" ,oJSON.Collection()

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

	'�Ō�̃y�[�W�̏ꍇ
	If wPage = vAllPage Then
		vOffset = 0
		vAdjust = 2
	Else
		vAdjust = 1
	End If

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
		vPointZan = vPointZan + CLng(vPoint)

		' --------------------------------
		'�K�v�ȃ��R�[�h�ʒu�̏ꍇ�ɁAJSON�f�[�^�𐶐�����
		'If (wPage <= vAllPage) And (i >= (wPageSize * (wPage - 1))) And (i <= (wPageSize * wPage - 1)) Then
		If (wPage <= vAllPage) And (i >= (vOffset)) And (i <= (vOffset + wPageSize - vAdjust)) Then
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

			'���������@(MAX(T2.�󒍌`��))
			'If (IsNull(vRS("���������@"))) Then
			'	vOrderType = ""
			'Else
			'	vOrderType = CStr(Trim(vRS("���������@")))
			'End If

			'���x�����@(MAX(T2.�x�����@))
			'If (IsNull(vRS("���x�����@"))) Then
			'	vPaymentMethod = ""
			'Else
			'	vPaymentMethod = CStr(Trim(vRS("���x�����@")))
			'End If

			'�|�C���g���p�l��(�|�C���g�敪)
			If (IsNull(vRS("�|�C���g���p�l��"))) Then
				vKubun = ""
			Else
				vKubun = CStr(Trim(vRS("�|�C���g���p�l��")))
			End If

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
					.Add "order_date" ,formatDateYYYYMMDD(vOrderDate)
					.Add "order_no" ,vOrderNo
					'.Add "order_type" ,vOrderType
					'.Add "payment_method" ,vPaymentMethod
					.Add "kubun" ,vKubun
					.Add "point_date" ,formatDateYYYYMMDD(vPointDate)
					.Add "point_expire" ,formatDateYYYYMMDD(vPointExpire)
					.Add "point" ,vPoint
					.Add "point_zan" ,vPointZan
				End With
			End With
		End If

		' �C�e���[�^���C���N�������g
		j = j + 1

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
