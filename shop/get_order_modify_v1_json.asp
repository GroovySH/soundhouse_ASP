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
'	�w������
'	�ύX�\��Ԃ̏ꍇ"Y"���A�s�̏ꍇ��"N"��ԋp����
'
' �ύX�\���
'   Web�����i�C���^�[�l�b�g�A�X�}�[�g�t�H���j�ł���B
'   �x�������@���u���[���v�ȊO�ł���B
'   �󒍃X�e�[�^�X���u�󒍁v(�o�׎w������)�łȂ��B
'   ���[�J�[�����i���܂܂�Ă��Ȃ��B
'   ���c���K���݌ɐ���=0�łȂ��B
'
'
'�ύX����
'2016/02/04 GV �V�K�쐬�B(�����ύX�L�����Z���@�\)
'2018.01.12 GV �����m�F�����؂ꌩ�ς��蒍���͕ύX�L�����Z���s�B
'
'========================================================================
'On Error Resume Next

Dim ConnectionEmax

Dim wErrMsg						' �G���[���b�Z�[�W (���̃y�[�W����n����郁�b�Z�[�W)
Dim wErrDesc
Dim wMsg						' �G���[���b�Z�[�W (�{�y�[�W�ō쐬���郁�b�Z�[�W)
Dim wCustomerNo					' �ڋq�ԍ�
Dim wOrderNo					' �󒍔ԍ�
Dim wDefer						' �ύX���[�h(Y/N)
Dim wFlg						' ���s�t���O
Dim oJSON						' JSON�I�u�W�F�N�g
Dim modifyFlag					' �ύX�\�t���O
Dim cancelFlag					' �L�����Z���\�t���O
Dim wNgReason					' �s���R
Dim wDepositFlag   				' ���������t���O
Dim wDepositAmount 				' �������v���z
Dim wWebModCancelFlg			' Web�����ύX�L�����Z�����t���O
Dim wCItem						' �L�����Z�����i
Dim cItems						' �z�񉻂����L�����Z�����i
Dim btnOn						' �{�^���\���t���O
Dim wOrderDate					' �󒍓�
Dim wHachuHikiateZero			' �����������ʃ[���t���O
Dim wTekiseiHachuSuuSei			' �K���݌�0���F�K�i�ʍ݌�.�������ʂ�����
Dim wDepositTerm				' �����m�F�����i���j

'=======================================================================
'	�󂯓n�������o�� & �����ݒ�
'=======================================================================
wFlg = True

' Get�p�����[�^
' �ڋq�ԍ�
wCustomerNo = ReplaceInput_NoCRLF(Trim(Request("cno")))
' ���l�̂݃`�F�b�N (ASP�͑S�p�ł������Ȃ�True��Ԃ�)
If (IsNumeric(wCustomerNo) = False) Or (cf_checkNumeric(wCustomerNo) = False) Then
	wFlg = False
End If

' �󒍔ԍ�
wOrderNo = ReplaceInput_NoCRLF(Trim(Request("ono")))
' ���l�̂݃`�F�b�N (ASP�͑S�p�ł������Ȃ�True��Ԃ�)
If (IsNumeric(wOrderNo) = False) Or (cf_checkNumeric(wOrderNo) = False) Then
	wFlg = False
End If

' �ۗ����[�h
wDefer = ReplaceInput_NoCRLF(Trim(Request("defer")))
wDefer = UCase(wDefer)
If (wDefer <> "Y") And (wDefer <> "N") And (wDefer <> "") Then
	wFlg = False
End If

' �L�����Z�����i
wCItem = ReplaceInput_NoCRLF(Trim(Request("c_item")))
If (wCItem <> "") Then
	cItems = Split(wCItem, "_")
End If


wNgReason = ""
wDepositFlag = ""
wDepositAmount = 0
wOrderDate = ""


'=======================================================================
'	Execute main
'=======================================================================
Call connect_db()

Call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   ' member�n�̑��̃y�[�W�����ɂȂ炤
If Err.Description <> "" Then
End If

Call close_db()

Call sendResponse()

'========================================================================
'
'	Function	Connect database
'
'========================================================================
Function connect_db()

Set ConnectionEmax = Server.CreateObject("ADODB.Connection")
ConnectionEmax.Open g_connectionEmax

End function

'========================================================================
'
'	Function	Close database
'
'========================================================================
Function close_db()

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
Dim vRS1          '�󒍃��R�[�h�Z�b�g
Dim vRS2          '�󒍖��׃��R�[�h�Z�b�g
Dim vRS3          '�X�V���R�[�h�Z�b�g
Dim okFlag
Dim wSQL
'Dim orderDate
Dim deleteDate
Dim promote
' 2018.01.12 GV add start
Dim wItemChar1
Dim wItemChar2
Dim wItemNum1
Dim wItemNum2
Dim wItemDate1
Dim wItemDate2
' 2018.01.12 GV end

' JSON�I�u�W�F�N�g����
Set oJSON = New aspJSON

okFlag = True
modifyFlag = "Y"  '�ύX�\�t���O
cancelFlag = "Y"  '�L�����Z���\�t���O
btnOn  = "Y"      '�{�^���\���t���O
wHachuHikiateZero = "" ' �����������ʃ[��
wTekiseiHachuSuuSei = "N" '�K���݌�0���F�K�i�ʍ݌�.�������ʂ�����

'�R���g���[���}�X�^���猩�ς���L���������擾 2018.01.12 GV add
call getEmaxCntlMst("��","�����m�F�҂�����","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)
If (IsNull(wItemNum1)) Then
	wDepositTerm = 10
Else
	wDepositTerm = wItemNum1
End If

' ���͒l������̏ꍇ
If (wFlg = True) Then
	'�󒍂̎��o��
	wSQL = ""
	wSQL = wSQL & "SELECT "
	wSQL = wSQL & "  �󒍌`�� "
	wSQL = wSQL & " ,�x�����@ "
	wSQL = wSQL & " ,�󒍓� "
	wSQL = wSQL & " ,�폜�� "
	wSQL = wSQL & " ,�폜�� "
	wSQL = wSQL & " ,Web�����ύX�L�����Z�����t���O "
	wSQL = wSQL & " ,�������v���z "
	wSQL = wSQL & " ,���������t���O "
	wSQL = wSQL & " ,�o�׊����� "
	wSQL = wSQL & " ,���̑����v���z "
	wSQL = wSQL & " ,ISNULL(�z����񖾍׎w��t���O, 'N') AS �z����񖾍׎w��t���O "
	wSQL = wSQL & " ,���ϓ� " '2018.01.12 GV add
'	wSQL = wSQL & "  FROM �� WITH(UPDLOCK) "
	wSQL = wSQL & "  FROM �� WITH(NOLOCK) "
	wSQL = wSQL & " WHERE �󒍔ԍ� = " & wOrderNo
	wSQL = wSQL & "  AND �ڋq�ԍ� = " & wCustomerNo
	wSQL = wSQL & "  AND �폜�� IS NULL "
	'@@@@Response.Write wSQL & "<br>"

	Set vRS1 = Server.CreateObject("ADODB.Recordset")
	vRS1.Open wSQL, ConnectionEmax, adOpenStatic, adLockOptimistic
	'vRS1.Open wSQL, ConnectionEmax, adOpenStatic, adLockPessimistic

	'���R�[�h�����݂��Ă���ꍇ
	If vRS1.EOF = False Then
		'Web�����ύX�L�����Z�����t���O
		If (IsNull(vRS1("Web�����ύX�L�����Z�����t���O"))) Then
			wWebModCancelFlg = "N"
		Else
			If (Trim(vRS1("Web�����ύX�L�����Z�����t���O")) <> "Y") Then
				wWebModCancelFlg = "N"
			Else
				wWebModCancelFlg = "Y"
			End If
		End If

		'Web�����ύX�L�����Z�����t���O��Y�̏ꍇ�A�����܂��̉\�������邽��
		'�ύX�L�����Z���s��
		If (wWebModCancelFlg = "Y") Then
			okFlag = False '�t���ONG
			wNgReason = "8"
			modifyFlag = "N"
			cancelFlag = "N"
			btnOn = "N"
			wFlg = okFlag

			'���R�[�h�Z�b�g�����
			vRS1.Close

			'���R�[�h�Z�b�g�̃N���A
			Set vRS1 = Nothing

			'�֐��E�o
			Exit Function
		End If

		'�o�׊������Ă���ꍇ�A�{�^����\��
		If (IsNull(vRS1("�o�׊�����")) = False) Then
			btnOn = "N"
			modifyFlag = "N"
			cancelFlag = "N"
		End If

		If (vRS1("�󒍌`��") = "�C���^�[�l�b�g") Or (vRS1("�󒍌`��") = "�X�}�[�g�t�H��") Then
			'�ύX�\�i�Ȃɂ����Ȃ��j
		Else
			okFlag = False '�t���ONG
			wNgReason = "1"
			modifyFlag = "N"
			cancelFlag = "N"
			btnOn = "N"
		End If

		If (okFlag = True) Then
			If (Mid(vRS1("�x�����@"), 1, 3) = "���[��") Then
				okFlag = False '�t���ONG
				wNgReason = "2"
				modifyFlag = "N"
				cancelFlag = "N"
				btnOn = "N" '2018.01.12 GV add
			End If
		End If

		If (okFlag = True) Then
			If (vRS1("���̑����v���z") <> 0) Then
				okFlag = False '�t���ONG
				wNgReason = "10"
				modifyFlag = "N"
				cancelFlag = "N"
				btnOn = "N" '2018.01.12 GV add
			End If
		End If

		If (okFlag = True) Then
			'Emax �œ͂���𕡐��ݒ肵�Ă���ꍇ
			If (vRS1("�z����񖾍׎w��t���O") = "Y") Then
				okFlag = False '�t���ONG
				wNgReason = "11"
				modifyFlag = "N"
				cancelFlag = "N"
				btnOn = "N" '2018.01.12 GV add
			End If
		End If

		'���������t���O
		If (IsNull(vRS1("���������t���O"))) Then
			wDepositFlag = ""
		Else
			wDepositFlag = CStr(Trim(vRS1("���������t���O")))
		End If

		' �������v���z
		If (IsNull(vRS1("�������v���z"))) Then
			wDepositAmount = 0
		Else
			wDepositAmount = CDbl(vRS1("�������v���z"))
		End If

		'�󒍓�����
		If (IsNull(vRS1("�󒍓�")) = True) Or (vRS1("�󒍓�") = "") Then
			wOrderDate = ""
		Else
			wOrderDate = vRS1("�󒍓�")
		End If

		'�폜������
		If (IsNull(vRS1("�폜��")) = True) Or (vRS1("�폜��") = "") Then
			deleteDate = ""
		Else
			deleteDate = vRS1("�폜��")
		End If

		'2018.01.12 GV add start
		'�폜����Ă��Ȃ��A���ς����ԁA�����������Ă��Ȃ�
		If (okFlag = True) Then
			If ((deleteDate = "") And (wOrderDate = "") And (wDepositFlag <> "Y")) Then
				'���ϓ���Null�łȂ��A�{���Ƃ̍���������m�F�����ȏ�
				If (IsNull(vRS1("���ϓ�")) = False) And (DateDiff("d", vRS1("���ϓ�"), Now()) >= CInt(wDepositTerm)) Then
					okFlag = False '�t���ONG
					wNgReason = "12"
					modifyFlag = "N"
					cancelFlag = "N"
					btnOn = "N"
				End If
			End If
		End If
		'2018.01.12 GV add end

		If (okFlag = True) Then
			'�󒍖��׃��R�[�h���擾
			wSQL = ""
			wSQL = wSQL & "SELECT "
			wSQL = wSQL & "  od.�󒍖��הԍ� "
			wSQL = wSQL & " ,od.���i�R�[�h "
			wSQL = wSQL & " ,od.���[�J�[�����t���O "
			wSQL = wSQL & " ,od.�󒍒P�� "
			wSQL = wSQL & " ,od.�o�׎w�����v���� "
			wSQL = wSQL & ", od.�󒍖��ה��l "
			'wSQL = wSQL & " ,z.�K���݌ɐ��� "
			wSQL = wSQL & ", ISNULL(od.�K���݌ɐ���, 0) AS �K���݌ɐ��� " '�������̓K���݌�
			wSQL = wSQL & " ,od.������������ "
			wSQL = wSQL & " ,i.�Z�b�g���i�t���O "
			wSQL = wSQL & " ,i.Web���i�t���O "
			wSQL = wSQL & " ,z.�������� as z��������"
			wSQL = wSQL & " FROM �󒍖��� od WITH (NOLOCK) "
			wSQL = wSQL & " INNER JOIN �F�K�i�ʍ݌� z WITH (NOLOCK) "
			wSQL = wSQL & "   ON z.���[�J�[�R�[�h = od.���[�J�[�R�[�h "
			wSQL = wSQL & "  AND z.���i�R�[�h = od.���i�R�[�h "
			wSQL = wSQL & "  AND z.�F = od.�F "
			wSQL = wSQL & "  AND z.�K�i = od.�K�i "

			wSQL = wSQL & " INNER JOIN ���i i WITH (NOLOCK) "
			wSQL = wSQL & "   ON i.���[�J�[�R�[�h = z.���[�J�[�R�[�h "
			wSQL = wSQL & "  AND i.���i�R�[�h = z.���i�R�[�h "

			wSQL = wSQL & " WHERE "
			wSQL = wSQL & "      od.�󒍔ԍ� = " & wOrderNo
			wSQL = wSQL & " AND od.�Z�b�g�i�e���הԍ� = 0 "

			'@@@@Response.Write wSQL & "<br>"

			Set vRS2 = Server.CreateObject("ADODB.Recordset")
			vRS2.Open wSQL, ConnectionEmax, adOpenStatic, adLockOptimistic
	
			'���R�[�h�����݂��Ă���ꍇ
			If vRS2.EOF = False Then

				'�����������ʃ[���t���O��������
				wHachuHikiateZero = "Y"

				'���[�v�J�n
				Do while vRS2.EOF = False

					'�����������ʂ��[���łȂ����̂��P�ł�����ꍇ
					If (vRS2("������������") <> 0) Then
						wHachuHikiateZero = "N"
					End If

					'�K���݌ɐ���0���A�F�K�i�ʍ݌ɂ̔������ʂ������̂��̂��P�ł�����ꍇ
					If (vRS2("�K���݌ɐ���") = 0) And (vRS2("z��������") > 0) Then
						wTekiseiHachuSuuSei = "Y"
					End If

					'�̑��i����
					promote = "N"
					If (CDbl(Trim(vRS2("�󒍒P��"))) = 0) Then
						'�󒍖��ה��l�Ɂu�̑��i�v�Ɗ܂܂��ꍇ�A
						If (InStr(Trim(vRS2("�󒍖��ה��l")), "�̑��i") > 0) Then
							promote = "Y"
						ElseIf (InStr(Trim(vRS2("���i�R�[�h")), "HOTMENU") > 0) Then
							promote = "Y"
						End If
					End If

					'�o�׎w����1�ł�����
					If (vRS2("�o�׎w�����v����") > 0) Then
						okFlag = False '�t���ONG
						wNgReason = "3"
						modifyFlag = "N"
						cancelFlag = "N"
						btnOn = "N"
					Else
						'�󒍂��폜����Ă���ꍇ�A�{�^���\���Ȃ�
						If (deleteDate <> "") Then
							btnOn = "N"
						End If
					End If

					'���[�J�[����
					If (vRS2("���[�J�[�����t���O") = "Y") Then
						okFlag = False '�t���ONG
						wNgReason = "4"
						modifyFlag = "N"
						cancelFlag = "N"
						btnOn = "N" '2018.01.12 GV add
					End If

					' �̑��i�łȂ�
					If promote = "N" Then
						'Web�Ɍf�ڂ��Ă��Ȃ�
						If Trim(vRS2("Web���i�t���O")) <> "Y" Then
							okFlag = False '�t���ONG
							wNgReason = "9"
							modifyFlag = "N"
							cancelFlag = "N"
							btnOn = "N" '2018.01.12 GV add
						End If
					End If

					If (okFlag = True) Then
						'��荞�܂ꂽ�����̏��
						If (wOrderDate = "") And (deleteDate = "") Then
							'modifyFlag = "Y"
						Else
							'�Z�b�g�i�̏ꍇ�͓K���݌ɐ��ʂ��݂Ȃ�
							If (vRS2("�Z�b�g���i�t���O") = "Y") Then
								If ((wOrderDate <> "") And (deleteDate = "")) Then
									'modifyFlag = "Y"
								Else
									okFlag = False '�t���ONG
									wNgReason = "5"
									cancelFlag = "N"
								End If
							Else
								'�̑��i
								If (promote = "Y") Then
									If ((wOrderDate <> "") And (deleteDate = "")) Then
										'modifyFlag = "Y"
									Else
										okFlag = False '�t���ONG
										wNgReason = "5"
										modifyFlag = "N"
										cancelFlag = "N"
										btnOn = "N" '2018.01.12 GV add
									End If
								Else
									' ���c���K���݌ɐ��ʂ�0�łȂ��ꍇ�AOK
									If (((wOrderDate <> "") And (deleteDate = "")) And (vRS2("�K���݌ɐ���") > 0)) Then
										'modifyFlag = "Y"
									Else 
										'�L�����Z�����i���w�肵�Ă���ꍇ
										If (wCItem <> "") Then
											'�K���݌�0�̂��̂��A�L�����Z�����悤�Ƃ��鏤�i�Ɋ܂܂�Ă���ꍇ
											If in_array(vRS2("�󒍖��הԍ�"), cItems) Then
												wNgReason = "5"
												cancelFlag = "N" '�L�����Z���͕s�����A�ύX�͎󂯕t����
											Else
											End If
										Else
											If (((wOrderDate <> "") And (deleteDate = "")) And (vRS2("�K���݌ɐ���") < 1)) Then
												wNgReason = "5"
												cancelFlag = "N" '�L�����Z���͕s�����A�ύX�͎󂯕t����
											Else
												okFlag = False '�t���ONG
												wNgReason = "5"
												modifyFlag = "N"
												cancelFlag = "N"
												btnOn = "N" '2018.01.12 GV add
											End If
										End If
									End If
								End If ' �̑��i
							End If '�K���݌�
						End If '���t
					End If

					'NG�̏ꍇ
					'If okFlag = False Then
					'	'���[�v�E�o
					'	Exit Do
					'End If

					'���̍s�ֈړ�
					vRS2.MoveNext
				Loop
			Else
				'�󒍖��׃��R�[�h���Ȃ��ꍇ�ANG
				okFlag = False
				wNgReason = "6"
				modifyFlag = "N"
				cancelFlag = "N"
				btnOn = "N"
			End If

			'�󒍖��׃��R�[�h�Z�b�g�����
			vRS2.Close

			'�󒍖��׃��R�[�h�Z�b�g�̃N���A
			Set vRS2 = Nothing
		End If '�󒍖��׎擾�@�����
	Else
		'�󒍃��R�[�h���Ȃ��ꍇ�ANG
		okFlag = False
		wNgReason = "7"
		modifyFlag = "N"
		cancelFlag = "N"
		btnOn = "N"
	End If

	'�K���݌ɐ���0���A�F�K�i�ʍ݌ɂ̔������ʂ������̂��̂��P�ł�����ꍇ
	'If wTekiseiHachuSuuSei = "Y" Then
	'	okFlag = False
	'	wNgReason = "12"
	'	modifyFlag = "N"
	'	cancelFlag = "N"
	'	'btnOn = "N"
	'End If

	'�ύX�܂��̓L�����Z�����\�̏ꍇ
	If (modifyFlag = "Y") Or (cancelFlag = "Y") Then
		'Web�����ύX�L�����Z�����t���O���X�V����ꍇ
		If (wDefer <> "") Then
			'---- �g�����U�N�V�����J�n
			ConnectionEmax.BeginTrans

			'�󒍂̎��o��
			wSQL = ""
			wSQL = wSQL & "SELECT "
			wSQL = wSQL & "  Web�����ύX�L�����Z�����t���O "
			'wSQL = wSQL & " FROM �� WITH(UPDLOCK) "
			wSQL = wSQL & " FROM �� "
			wSQL = wSQL & " WHERE �󒍔ԍ� = " & wOrderNo
			wSQL = wSQL & "  AND �ڋq�ԍ� = " & wCustomerNo
			wSQL = wSQL & "  AND �폜�� IS NULL "
			'@@@@Response.Write wSQL & "<br>"

			Set vRS3 = Server.CreateObject("ADODB.Recordset")
			vRS3.Open wSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

			'���R�[�h�����݂��Ă���ꍇ
			If vRS3.EOF = False Then
				vRS3("Web�����ύX�L�����Z�����t���O") = wDefer
				vRS3.update

				'�R�~�b�g
				ConnectionEmax.CommitTrans
			Else
				'���[���o�b�N
				ConnectionEmax.RollbackTrans
			End If

			'�X�V���R�[�h�Z�b�g�����
			vRS3.Close

			'�X�V���R�[�h�Z�b�g�̃N���A
			Set vRS3 = Nothing
		End If
	End If

	'�󒍃��R�[�h�Z�b�g�����
	vRS1.Close

	'�󒍃��R�[�h�Z�b�g�̃N���A
	Set vRS1 = Nothing
Else
	'���͒l��NG�̏ꍇ
	okFlag = False
	wNgReason = "99"
End If

wFlg = okFlag

End Function


'========================================================================
'
'	Function	JSON�ԋp
'
'========================================================================
Function sendResponse()

	' �S������JSON�f�[�^�ɃZ�b�g
	oJSON.data.Add "ono" ,wOrderNo
	oJSON.data.Add "cno" ,wCustomerNo
	oJSON.data.Add "o_dt" ,wOrderDate
	oJSON.data.Add "deposit" ,wDepositFlag
	oJSON.data.Add "deposit_am" ,wDepositAmount
	oJSON.data.Add "defer" ,wDefer
	oJSON.data.Add "modifying", wWebModCancelFlg
	oJSON.data.Add "btn_on", btnOn
	oJSON.data.Add "mod", modifyFlag
	oJSON.data.Add "cancel", cancelFlag
	oJSON.data.Add "reason" ,wNgReason
	oJSON.data.Add "h_hiki_zero" ,wHachuHikiateZero
	oJSON.data.Add "tekisei_h_sei", wTekiseiHachuSuuSei

	'�ύX�ƃL�����Z���̎�t�����𕪂������߁A�ȉ��͕s�v
	'If wFlg = True Then
	'	oJSON.data.Add "result" ,"Y"
	'Else
	'	oJSON.data.Add "result" ,"N"
	'End If

	' -------------------------------------------------
	' JSON�f�[�^�̕ԋp
	' -------------------------------------------------
	' �w�b�_�o��
	Response.AddHeader "Content-Type", "application/json; charset=shift_jis"
	Response.AddHeader "Cache-Control", "no-cache,must-revalidate"
	Response.AddHeader "Pragma", "no-cache"
	Response.AddHeader "X-Content-Type-Options", "nosniff"

	' JSON�f�[�^�̏o��
	Response.Write oJSON.JSONoutput()

End Function

'========================================================================
'
'	Function	�z�񑶍݃`�F�b�N
'
'========================================================================
Function in_array(needle, arr)
	in_array = False
	Dim element
	Dim i

	For i=0 To UBound(arr)
		element = CStr(needle)
		If Trim(arr(i)) = Trim(element) Then
			in_array = True
			Exit Function
		End If
	Next
End Function

'========================================================================
'
'	Function	Emax�̃R���g���[���}�X�^����f�[�^�擾
'
'========================================================================

Function getEmaxCntlMst(pSubSystemCd, pItemCd, pItemSubCd, pItemChar1, pItemChar2, pItemNum1, pItemNum2, pItemDate1, pItemDate2)

Dim RS_cntl
Dim v_sql

'---- �R���g���[���}�X�^���o��

v_sql = ""
v_sql = v_sql & "SELECT a.*"
v_sql = v_sql & "  FROM �R���g���[���}�X�^ a WITH (NOLOCK)"
v_sql = v_sql & " WHERE a.sub_system_cd = '" & pSubSystemCd & "'"
v_sql = v_sql & "   AND a.item_cd = '" & pItemCd & "'"
v_sql = v_sql & "   AND a.item_sub_cd = '" & pItemSubCd & "'"

'@@@@@@response.write(v_sql)

Set RS_cntl = Server.CreateObject("ADODB.Recordset")
RS_cntl.Open v_sql, ConnectionEmax, adOpenStatic

If RS_cntl.EOF <> True Then
	pItemChar1 = RS_cntl("item_char1")
	pItemChar2 = RS_cntl("item_char2")
	pItemNum1 = RS_cntl("item_num1")
	pItemNum2 = RS_cntl("item_num2")
	pItemDate1 = RS_cntl("item_date1")
	pItemDate2 = RS_cntl("item_date2")
End If

RS_cntl.Close

End Function
'========================================================================
%>
