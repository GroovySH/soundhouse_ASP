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
'	�w�������ꗗ�y�[�W�ɂ����闘�p�����|�C���g�����擾
'
'
'�ύX����
'2016/03/11 GV �V�K�쐬�B(Web�����ύX�L�����Z���@�\)
'2020.06.30 GV �~���������X�g�Ή��B(#2841)
'
'========================================================================
'On Error Resume Next

Dim Connection
Dim ConnectionEmax

Dim wErrMsg						' �G���[���b�Z�[�W (���̃y�[�W����n����郁�b�Z�[�W)
Dim wDispMsg					' �ʏ탁�b�Z�[�W(�G���[�ȊO) (���̃y�[�W����n����郁�b�Z�[�W)
Dim wErrDesc
Dim wMsg						' �G���[���b�Z�[�W (�{�y�[�W�ō쐬���郁�b�Z�[�W)
Dim wCustomerNo					' �ڋq�ԍ�
Dim wOrderNo					' �󒍔ԍ�
Dim oJSON						' JSON�I�u�W�F�N�g


'=======================================================================
'	�󂯓n�������o�� & �����ݒ�
'=======================================================================
' Get�p�����[�^
wCustomerNo = ReplaceInput(Trim(Request("cno")))
wOrderNo = ReplaceInput(Trim(Request("ono")))


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
Dim vRS
Dim vHTML
Dim vPointDate
Dim vPoint
Dim vPointZan
Dim vUseOrderNo
Dim vBeforeOrderNo
Dim vBeforeOrderDetailNo
Dim vAddFlag
Dim vTotalObtainPoint	' ���v�l���|�C���g

Set oJSON = New aspJSON

' �l�����X�g�ǉ�
oJSON.data.Add "obtain" ,oJSON.Collection()
' ���p���X�g�ǉ�
oJSON.data.Add "used" ,oJSON.Collection()

' �C�e���[�^������
i = 0

vBeforeOrderNo = null
vBeforeOrderDetailNo = null
vAddFlag = false

vTotalObtainPoint = 0

'-----------------------------------------------------------
' �l���|�C���g���̎擾
'-----------------------------------------------------------
'vSQL = ""
'vSQL = vSQL & "SELECT "
'vSQL = vSQL & "    a.�󒍔ԍ� "
'vSQL = vSQL & "  , a.�󒍖��הԍ�"
'vSQL = vSQL & "      , a.�󒍖��׎}��"
'vSQL = vSQL & "      , a.�|�C���g�敪"
'vSQL = vSQL & "      , a.�|�C���g���t"
'vSQL = vSQL & "      , a.�|�C���g"
'vSQL = vSQL & "      , a.�|�C���g�c"
'vSQL = vSQL & "      , a.�g�p�󒍔ԍ�"
'vSQL = vSQL & "      , a.����o�^��"
'vSQL = vSQL & " FROM "
'vSQL = vSQL & "    " & gLinkServer & "�|�C���g���ח��� a WITH (NOLOCK) "
'vSQL = vSQL & " WHERE "
'vSQL = vSQL & "  (CONVERT(VARCHAR(100), �󒍔ԍ�)+"
'vSQL = vSQL & "CONVERT(VARCHAR(100),�󒍖��הԍ�)+"
'vSQL = vSQL & "CONVERT(VARCHAR(100), �󒍖��׎}��)+"
'vSQL = vSQL & "CONVERT(varchar(100), ����o�^��,121))"
'vSQL = vSQL & "   IN ("
'vSQL = vSQL & "     SELECT"
'vSQL = vSQL & "       (CONVERT(VARCHAR(100), �󒍔ԍ�)+"
'vSQL = vSQL & "CONVERT(VARCHAR(100),�󒍖��הԍ�)+"
'vSQL = vSQL & "CONVERT(VARCHAR(100), �󒍖��׎}��)+"
'vSQL = vSQL & "CONVERT(varchar(100), MAX(����o�^��),121))"
'vSQL = vSQL & "     FROM"
'vSQL = vSQL & "       " & gLinkServer & "�|�C���g���ח��� WITH (NOLOCK) "
'vSQL = vSQL & "     WHERE"
'vSQL = vSQL & "           �ڋq�ԍ� =  " & wCustomerNo
'vSQL = vSQL & "       AND �X�V�敪 = 'Updated'"
'vSQL = vSQL & "       AND �|�C���g�敪 = '�l��'"
'vSQL = vSQL & "       AND �|�C���g�c is not null "
'vSQL = vSQL & "       AND �|�C���g���� is not null "
'vSQL = vSQL & "       AND �ŏI�X�V���� = '�o�׎w��'"
'vSQL = vSQL & "       AND �󒍔ԍ� IN (" & wOrderNo & ")"
'vSQL = vSQL & "     GROUP BY"
'vSQL = vSQL & "      �󒍔ԍ�, �󒍖��הԍ�, �󒍖��׎}��"
'vSQL = vSQL & "   )"
'vSQL = vSQL & " ORDER BY"
'vSQL = vSQL & "   �󒍔ԍ� ASC, �󒍖��הԍ� ASC, �󒍖��׎}�� ASC"
vSQL = createPointSql(wCustomerNo, wOrderNo, "�l��")

'@@@@Response.Write(vSQL&"<br>")

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

// ���R�[�h�����݂����ꍇ�AJSON�I�u�W�F�N�g���쐬
If vRS.EOF = False Then
	createJsonObject vRS, "obtain"
End If


'���R�[�h�Z�b�g�����
vRS.Close


'-----------------------------------------------------------
' ���p�|�C���g���̎擾
'-----------------------------------------------------------
' �C�e���[�^������
i = 0

vBeforeOrderNo = null

'--- �Y���ڋq�̃|�C���g���ׂ̎��o��
'vSQL = ""
'vSQL = vSQL & "SELECT "
'vSQL = vSQL & "    a.�󒍔ԍ� "
'vSQL = vSQL & "  , a.�󒍖��הԍ�"
'vSQL = vSQL & "      , a.�󒍖��׎}��"
'vSQL = vSQL & "      , a.�|�C���g�敪"
'vSQL = vSQL & "      , a.�|�C���g���t"
'vSQL = vSQL & "      , a.�|�C���g"
'vSQL = vSQL & "      , a.�|�C���g�c"
'vSQL = vSQL & "      , a.�g�p�󒍔ԍ�"
'vSQL = vSQL & "      , a.����o�^��"
'vSQL = vSQL & " FROM "
'vSQL = vSQL & "    " & gLinkServer & "�|�C���g���ח��� a WITH (NOLOCK) "
'vSQL = vSQL & " WHERE "
'vSQL = vSQL & "  (CONVERT(VARCHAR(100), �󒍔ԍ�)+"
'vSQL = vSQL & "CONVERT(VARCHAR(100),�󒍖��הԍ�)+"
'vSQL = vSQL & "CONVERT(VARCHAR(100), �󒍖��׎}��)+"
'vSQL = vSQL & "CONVERT(varchar(100), ����o�^��,121))"
'vSQL = vSQL & "   IN ("
'vSQL = vSQL & "     SELECT"
'vSQL = vSQL & "       (CONVERT(VARCHAR(100), �󒍔ԍ�)+"
'vSQL = vSQL & "CONVERT(VARCHAR(100),�󒍖��הԍ�)+"
'vSQL = vSQL & "CONVERT(VARCHAR(100), �󒍖��׎}��)+"
'vSQL = vSQL & "CONVERT(varchar(100), MAX(����o�^��),121))"
'vSQL = vSQL & "     FROM"
'vSQL = vSQL & "       " & gLinkServer & "�|�C���g���ח��� WITH (NOLOCK) "
'vSQL = vSQL & "     WHERE"
'vSQL = vSQL & "           �ڋq�ԍ� =  " & wCustomerNo
'vSQL = vSQL & "       AND �X�V�敪 = 'Inserted'"
'vSQL = vSQL & "       AND �|�C���g�敪 = '���p'"
'vSQL = vSQL & "       AND �|�C���g�c is null "
'vSQL = vSQL & "       AND �|�C���g���� is null "
'vSQL = vSQL & "       AND �ŏI�X�V���� = '��'"
'vSQL = vSQL & "       AND �󒍔ԍ� IN (" & wOrderNo & ")"
'vSQL = vSQL & "     GROUP BY"
'vSQL = vSQL & "      �󒍔ԍ�, �󒍖��הԍ�, �󒍖��׎}��"
'vSQL = vSQL & "   )"
'vSQL = vSQL & " ORDER BY"
'vSQL = vSQL & "   �󒍔ԍ� ASC, �󒍖��הԍ� ASC, �󒍖��׎}�� ASC"
vSQL = createPointSql(wCustomerNo, wOrderNo, "���p")

'@@@@Response.Write(vSQL)

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

// ���R�[�h�����݂����ꍇ�AJSON�I�u�W�F�N�g���쐬
If vRS.EOF = False Then
	createJsonObject vRS, "used"
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

vDate = DatePart("yyyy", pdatDate) & "�N"

If DatePart("m", pdatDate) <= 9 Then
	vDate = vDate & "0" & DatePart("m", pdatDate)
Else
	vDate = vDate & DatePart("m", pdatDate)
End If

vDate = vDate & "��"

If DatePart("d", pdatDate) <= 9 Then
	vDate = vDate & "0" & DatePart("d", pdatDate)
Else
	vDate = vDate & DatePart("d", pdatDate)
End If

vDate = vDate & "��"

formatDateYYYYMMDD = vDate

End Function

'========================================================================
'
'	Function	�|�C���g���̎擾SQL
'
'========================================================================
Function createPointSql(customerNo, orderNo, kubun)
	Dim vSQL
	vSQL = ""
	vSQL = vSQL & "SELECT "
	vSQL = vSQL & "    a.�󒍔ԍ� "
	vSQL = vSQL & "  , a.�󒍖��הԍ�"
	vSQL = vSQL & "      , a.�󒍖��׎}��"
	vSQL = vSQL & "      , a.�|�C���g�敪"
	vSQL = vSQL & "      , a.�|�C���g���t"
	vSQL = vSQL & "      , a.�|�C���g"
	vSQL = vSQL & "      , a.�|�C���g�c"
	vSQL = vSQL & "      , a.�g�p�󒍔ԍ�"
	vSQL = vSQL & "      , a.�|�C���g�ԍ�"
	vSQL = vSQL & " FROM "
	vSQL = vSQL & "    " & gLinkServer & "�|�C���g���� a WITH (NOLOCK) "
	vSQL = vSQL & "     WHERE"
	vSQL = vSQL & "           �ڋq�ԍ� =  " & customerNo
	vSQL = vSQL & "       AND �|�C���g�敪 = '" & kubun & "'"
	vSQL = vSQL & "       AND �󒍔ԍ� IN (" & orderNo & ")"
	vSQL = vSQL & " ORDER BY"
	vSQL = vSQL & "   �󒍔ԍ� ASC, �󒍖��הԍ� ASC, �󒍖��׎}�� ASC"

	createPointSql = vSQL
End Function

'========================================================================
'
'	Function	DB����擾�����f�[�^����I�u�W�F�N�g�𐶐�
' JSON�I�u�W�F�N�g�Ŕz���ǉ�����ɂ́A�����o�ϐ��̃L�[�𐔒l�ɂ��邪�A
' �󒍔ԍ��̌��ł͒ǉ��ł��Ȃ��B
'
'========================================================================
Function createJsonObject(vRS, kubun)
	Dim pointDate
	Dim point
	Dim pointZan
	Dim useOrderNo
	Dim addFlag
	Dim beforeOrderNo
	Dim beforeOrderDetailNo
	Dim totalPoint
	Dim beforePointNo '2021.06.30 GV add

	beforeOrderNo = null
	beforeOrderDetailNo = null
	addFlag = false
	totalPoint = 0
	beforePointNo = null ' 2021.06.30 GV add

	' ���R�[�h�Z�b�g�̍Ō�܂Ń��[�v
	Do Until vRS.EOF

		' �|�C���g���t
		If (IsNull(vRS("�|�C���g���t"))) Then
			pointDate = ""
		Else
			pointDate = CStr(Trim(vRS("�|�C���g���t")))
		End If

		' �|�C���g
		If (IsNull(vRS("�|�C���g"))) Then
			point = 0
		Else
			point = CStr(Trim(vRS("�|�C���g")))
		End If

		totalPoint = totalPoint + vRS("�|�C���g")

		' �|�C���g�c
		If (IsNull(vRS("�|�C���g�c"))) Then
			pointZan = 0
		Else
			pointZan = CStr(Trim(vRS("�|�C���g�c")))
		End If

		' �g�p�󒍔ԍ�
		If (IsNull(vRS("�g�p�󒍔ԍ�"))) Then
			useOrderNo = ""
		Else
			useOrderNo = CStr(Trim(vRS("�g�p�󒍔ԍ�")))
		End If

		' �󒍔ԍ����P�O�̃��[�v���ƈႤ�ꍇ
		If (IsNull(beforeOrderNo) = True) Then
			addFlag = True
		ElseIf (beforeOrderNo <> vRS("�󒍔ԍ�")) Then
			addFlag = True
		Else
			addFlag = false
		End If

		If (addFlag) Then
			beforeOrderNo = vRS("�󒍔ԍ�")

			beforeOrderDetailNo = null
			beforePointNo = null '2021.06.30 GV add

			With oJSON.data(kubun)
				.Add "o"&CStr(beforeOrderNo) ,oJSON.Collection()
			End With
		End If

		'�󒍖��הԍ���1�O�̃��[�v���ƈႤ�ꍇ
		If (IsNull(beforeOrderDetailNo) = True) Then
			addFlag = True
		ElseIf (beforeOrderDetailNo <> vRS("�󒍖��הԍ�")) Then
			addFlag = True
		Else
			addFlag = false
		End If

		'�|�C���g�ԍ���1�O�̃��[�v���ƈႤ�ꍇ
		If (IsNull(beforePointNo) = True) Then
			addFlag = True
		ElseIf (beforePointNo <> vRS("�|�C���g�ԍ�")) Then
			addFlag = True
		Else
			addFlag = false
		End If



		If (addFlag = True) Then
			beforeOrderDetailNo = vRS("�󒍖��הԍ�")
			beforePointNo = vRS("�|�C���g�ԍ�") '2021.06.30 GV add

			With oJSON.data(kubun)
				'With .item(beforeOrderNo)
				'With .item("o"&CStr(beforeOrderNo))  '2021.06.30 GV mod
				'	.Add "d"&beforeOrderDetailNo ,oJSON.Collection()  '2021.06.30 GV mod
				With .item("o"&CStr(beforeOrderNo))
					.Add "d"&beforePointNo ,oJSON.Collection()
				End With
			End With
		End If

		' �l�����X�g�ǉ�
		With oJSON.data(kubun).item("o"&CStr(beforeOrderNo))
			'With .item("d"&beforeOrderDetailNo) '2021.06.30 GV mod
			'	.Add "sub"&vRS("�󒍖��׎}��"), oJSON.Collection()  '2021.06.30 GV mod
			With .item("d"&beforePointNo)
				.Add "sub"&vRS("�󒍖��׎}��"), oJSON.Collection()
			End With
		End With

		With oJSON.data(kubun).item("o"&CStr(beforeOrderNo))
			'With .item("d"&beforeOrderDetailNo) '2021.06.30 GV mod
			With .item("d"&beforePointNo)
				With .item("sub"&CStr(vRS("�󒍖��׎}��")))
					.Add "o_no" ,CStr(Trim(vRS("�󒍔ԍ�")))
					.Add "od_no" ,CStr(Trim(vRS("�󒍖��הԍ�")))
					.Add "od_sub_no" ,CStr(Trim(vRS("�󒍖��׎}��")))
					.Add "kubun" ,CStr(Trim(vRS("�|�C���g�敪")))
					.Add "pt_dt" ,formatDateYYYYMMDD(pointDate)
					.Add "pt" ,point
					.Add "pt_zan" ,pointZan
					.Add "use_o_no" ,useOrderNo
				End With
			End With
		End With

		' ���R�[�h�Z�b�g�̃|�C���^�����̍s�ֈړ�
		vRS.MoveNext
	Loop

	If kubun = "obtain" Then
		oJSON.data.Add "total_obtain_pt" ,totalPoint
	ElseIf kubun = "used" Then
		oJSON.data.Add "total_used_pt" ,totalPoint
	End If

'createJsonObject = oJSON
End Function
'========================================================================
%>
