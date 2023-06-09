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
<!--#include file="../common/HttpsSecurity.inc"-->

<%
'========================================================================
'
'	�I�[�_�[���o�^
'		POST���ꂽ�������󒍂֓o�^�B
'		���͂��ꂽ�f�[�^�[�̃`�F�b�N�B
'
'�ύX����
'2011/02/16 GV(dy) OrderinfoInsert.asp�����ɍ�蒼��
'2011/02/16 hn �z���\���Z�b�g�d�l�ύX
'2011/04/14 hn SessionID�֘A�ύX'
'2011/05/02 hn ���}�g�̏ꍇ�Ŏ��Ԏw�肠��́A�[���w��K�{�`�F�b�N�ǉ�
'2011/06/29 an #867 �^����Ђɐ��Z�ǉ��A�^����Ђ̌�����@�ύX�A�w��\�[���̃`�F�b�N���@�ύX
'2011/07/25 hn ���}�g�͎g�p���Ȃ��悤�ɕύX
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2011/08/11 an #1090 �ߑO�w�莞�A���Z�̌ߑO�s�n��̏ꍇ�͉^����Ђ�����ɕύX
'2011/09/12 an #1111/1130 �^����Ќ�������ɍ���Œ�t���O/���[�h�^�C���`�F�b�N�ǉ�
'2012/07/26 nt ���d�ʎ��A�^����А���@�\��ǉ�
'2012/08/25 nt ���������֎~�n��̐���@�\��ǉ�
'2013/02/18 GV #1525 ������̎w��[������
'2014/08/05 GV �[�i���\���ύX�Ή�
'
'========================================================================
On Error Resume Next
Response.Expires = -1			' Do not cache
Response.buffer = true

'---- Session���
Dim wUserID
Dim wUserName
Dim wMSG

'---- �󂯓n����������ϐ�
Dim cmd
Dim ship_address_no
Dim ship_invoice_fl
Dim customer_kn
Dim customer_email
Dim telephone
Dim KabusokuAm
Dim payment_method
Dim furikomi_nm
Dim ikkatsu_fl
Dim freight_forwarder
Dim delivery_fl
Dim delivery_mm
Dim delivery_dd
Dim delivery_tm
Dim eigyousho_dome_fl
Dim receipt_fl
Dim receipt_nm
Dim receipt_memo
Dim RebateFl
Dim i_tokuchuu_fl
Dim i_daibiki_fuka_fl

'---- �͐���
Dim wShipNm
Dim wShipZip
Dim wShipPrefecture
Dim wShipAddress
Dim wShipTel

Dim wDeliveryDt
Dim wCustZip
Dim wCustPrefecture
Dim wRitouFl							'�����t���O
Dim wFreightAm							'����
Dim wFreightForwarder					'�z�����
Dim wSoukoCnt							'�o�בq�ɐ�
Dim wKoguchi							'����

Dim wAfter13FL							'13���ȍ~
Dim wAfter14FL							'14���ȍ~
Dim wAfter15FL							'15���ȍ~
Dim wAfter16FL							'16���ȍ~
'Dim w94HokkaidouRitouFL				'�z���悪�@��B�E�l���E�k�C���E����   2011/06/29 an del
Dim w94ChugokuHokkaidouFL				'�z���悪�@��B�E�l���E�����E�k�C��   2011/06/29 an add
Dim wHolidayFL							'���j�Փ�
Dim wAvailableDate						'�w��\��
Dim wSatFl								'�y�j��
Dim wFriFL								'���j��
Dim wErrDesc   '2011/08/01 an add

Dim kErrFlg		'���d�ʃG���[�t���O 2012/07/26 nt add
Dim wSagawaLTFl	'���������֎~�t���O 2012/08/25 nt add

'---- DB
Dim Connection

'Const w9Shuu4KokuHokkaido = "������,���茧,���ꌧ,�啪��,�F�{��,�{�茧,��������,���ꌧ,���쌧,������,���Q��,���m��,�k�C��"    '2011/06/29 an del
Const w9Shuu4KokuChugokuHokkaido = "������,���茧,���ꌧ,�啪��,�F�{��,�{�茧,��������,���쌧,������,���Q��,���m��,���挧,���R��,������,�L����,�R����,�k�C��"   '2011/06/29 an add
Const cAddDaysToNyukaYoteibi = 2

'=======================================================================
'	�󂯓n�������o��
'=======================================================================
'---- Session�ϐ�
wUserID = Session("UserID")
wUserName = Session("userName")
wMsg = Session.contents("msg")

'---- �󂯓n�������o��
cmd = Left(ReplaceInput(Trim(Request("cmd"))), 10)
ship_address_no = ReplaceInput(Request("ship_address_no"))
ship_invoice_fl = Left(ReplaceInput(Trim(Request("ship_invoice_fl"))), 1)
customer_kn = Left(ReplaceInput(Trim(Request("customer_kn"))), 60)
customer_email = Left(ReplaceInput(Trim(Request("customer_email"))), 60)
telephone = Left(ReplaceInput(Trim(Request("telephone"))), 20)
KabusokuAm = ReplaceInput(Trim(Request("KabusokuAm")))
payment_method = Left(ReplaceInput(Trim(Request("payment_method"))), 10)
furikomi_nm = Left(ReplaceInput(Trim(Request("furikomi_nm"))), 30)
ikkatsu_fl = Left(ReplaceInput(Trim(Request("ikkatsu_fl"))), 1)
freight_forwarder = Left(ReplaceInput(Trim(Request("freight_forwarder"))), 8)
delivery_fl = ReplaceInput(Trim(Request("delivery_fl")))
delivery_mm = ReplaceInput(Trim(Request("delivery_mm")))
delivery_dd = ReplaceInput(Trim(Request("delivery_dd")))
delivery_tm = ReplaceInput(Trim(Request("delivery_tm")))
eigyousho_dome_fl = Left(ReplaceInput(Trim(Request("eigyousho_dome_fl"))), 1)
receipt_fl = Left(ReplaceInput(Trim(Request("receipt_fl"))), 1)
receipt_nm = Left(ReplaceInput(Trim(Request("receipt_nm"))), 30)
receipt_memo = Left(ReplaceInput(Trim(Request("receipt_memo"))), 25)
RebateFl = Left(ReplaceInput(Trim(Request("RebateFl"))), 1)
i_tokuchuu_fl = Left(ReplaceInput(Request("i_tokuchuu_fl")), 1)
i_daibiki_fuka_fl = Left(ReplaceInput(Request("i_daibiki_fuka_fl")), 1)

'2014/08/05 GV add start
'OrderInfoEnter.asp �Ŕ�\���������ꍇ
fwriteErrorLog("[BEFORE] payment_method='"&payment_method&"' // ship_address_no='"&ship_address_no&"' // ship_invoice_fl='"&ship_invoice_fl&"'")
If IsNULL(ship_invoice_fl) = true Then
fwriteErrorLog("ship_invoice_fl is null.")
end if

'If ship_invoice_fl = "" Then								'2014/08/21 comment out
If (ship_invoice_fl = "") Or (ship_invoice_fl = " ") Then	'2014/08/21 add
	ship_invoice_fl = "Y"
End If
fwriteErrorLog("[AFTER] ship_invoice_fl='"&ship_invoice_fl&"'")
'2014/08/05 GV add start

'---- �Z�b�V�����؂�`�F�b�N
If wUserID = ""Then
	Response.Redirect g_HTTP
End If

Session("msg") = ""
wMSG = ""

'=======================================================================
'	Execute main
'=======================================================================
Call connect_db()
Call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "OrderinfoInsert.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

Call close_db()

If Err.Description <> "" Then
	Response.Redirect g_HTTP & "shop/Error.asp"
End If

'---- �G���[�������Ƃ��͒������e�m�F�y�[�W�A�G���[������Β������e�w��y�[�W��
If wMSG = "" Then
	Select Case cmd
		Case "next"
			If payment_method = "���[��" Then
				Server.Transfer "OrderLoan.asp"
			Else
				Server.Transfer "OrderConfirm.asp"
			End If
		Case "address"
			Server.Transfer "OrderShipAddress.asp"
	End Select
Else
	Session("msg") = wMSG
	Server.Transfer "OrderInfoEnter.asp"
End If

'========================================================================
'
'	Function	Connect database
'
'========================================================================
Function connect_db()

'---- Connect database
Set Connection = Server.CreateObject("ADODB.Connection")
Connection.Open g_connection

End function

'========================================================================
'
'	Function	Close database
'
'========================================================================
Function close_db()

Connection.Close
Set Connection= Nothing    '2011/08/01 an add

End Function

'========================================================================
'
'	Function	Main
'
'========================================================================
Function main()

'---- ���󒍏��X�V
Call update_order_header()

If cmd = "next" Then
	'---- ���̓f�[�^�[�̃`�F�b�N
	Call validate_data()
End If

End Function

'========================================================================
'
'	Function	���󒍏��̍X�V
'
'========================================================================
Function update_order_header()

Dim RSv
Dim vSQL

'---- �͐���擾
If isNumeric(ship_address_no) = False Then
	ship_address_no = 1
End If

Call GetTodokesakiInfo(ship_address_no)

'---- �����t���O�̐ݒ�
Call setRitouFlag(wShipZip)

'---- ����֎~�t���O�̐ݒ�
Call setSagawaLTFlag(wShipZip)

'---- �����E�z����ЁE�o�בq�ɐ��E���� �v�Z
Call fCalcShipping(gSessionID, "�ꊇ", wFreightAm, wFreightForwarder, wSoukoCnt, wKoguchi)		'2011/04/14 hn mod

'---- ����Recordset���o��
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    *"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    ����"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "    SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic, adLockOptimistic

'---- �ڋq���
RSv("�ڋq�ԍ�") = wUserID
RSv("�ڋqE_mail") = customer_email
RSv("�ڋq�d�b�ԍ�") = telephone
RSv("���σt���O") = ""

'---- �x�����@
RSv("�x�����@") = payment_method

Select Case payment_method
	Case "��s�U��"
		If furikomi_nm = "" Then
			furikomi_nm = customer_kn
		End If
		RSv("�U�����`�l") = furikomi_nm
		RSv("���[����������t���O") = ""
		RSv("��]���[����") = 0
		RSv("���[������") = 0
		RSv("���[�����z") = 0
		RSv("���[�����") = ""
		RSv("�I�����C�����[���\���t���O") = ""

	Case "���[��"
		RSv("�U�����`�l") = ""

	Case Else
		RSv("�U�����`�l") = ""
		RSv("���[����������t���O") = ""
		RSv("��]���[����") = 0
		RSv("���[������") = 0
		RSv("���[�����z") = 0
		RSv("���[�����") = ""
		RSv("�I�����C�����[���\���t���O") = ""

End Select

'---- �͐���
RSv("�͐�Z���A��") = ship_address_no
RSv("�͐於�O") = wShipNm
RSv("�͐�X�֔ԍ�") = wShipZip
RSv("�͐�s���{��") = wShipPrefecture
RSv("�͐�Z��") = wShipAddress
RSv("�͐�d�b�ԍ�") = wShipTel

If ship_address_no = 1 Then
	RSv("�͐�敪") = "S"
	RSv("�͐�[�i�����t�t���O") = "Y"		'�͐悪�Z���Ɠ����̏ꍇ�͖�������Y
Else
	RSv("�͐�敪") = "D"
	RSv("�͐�[�i�����t�t���O") = ship_invoice_fl
End If

If delivery_mm <> "" And delivery_dd <> "" Then
	If cf_NumToChar(DatePart("m", Date()), 2) & cf_NumToChar(DatePart("d", Date()), 2) > (delivery_mm & delivery_dd) Then
		wDeliveryDt = Cstr(Clng(DatePart("yyyy", Date())) + 1) & "/" & delivery_mm & "/" & delivery_dd
	Else
		wDeliveryDt = DatePart("yyyy", Date()) & "/" & delivery_mm & "/" & delivery_dd
	End If
	If isDate(wDeliveryDt) = False Then
		wDeliveryDt = ""
	End If
End If

If wDeliveryDt <> "" Then
	RSv("�w��[��") = wDeliveryDt
Else
	RSv("�w��[��") = NULL
End If

'---- �z����� 
'---- �^����Ѓ`�F�b�N���ύX
call CheckFreightForwarder()   '2011/06/29 an add

'---- ����1���ŋ�A�֎~���i���܂܂�ĂȂ��d�ʏ��i�łȂ��ꍇ�́A�z����Ђ��u���}�g�^�A�v�ɋ����ύX   '2011/06/29 an del s
'If wRitouFl = "Y" And wKoguchi = 1 And checkKuuyukinshiShouhin() = 0 And checkJyuuryouShouhin() = 0 Then
'	freight_forwarder = "2"
'End If     '2011/06/29 an del e

RSv("�^����ЃR�[�h") = freight_forwarder

RSv("���Ԏw��") = delivery_tm

RSv("�c�Ə��~�߃t���O") = eigyousho_dome_fl

If RSv("�x�����@") = "�����" Then
	RSv("�ꊇ�o�׃t���O") = "Y"
Else
	RSv("�ꊇ�o�׃t���O") = ikkatsu_fl
End If

If wRitouFl = "Y" Then
	RSv("�ꊇ�o�׃t���O") = "Y"
End If

'---- �̎���
RSv("�̎������s�t���O") = receipt_fl
If receipt_fl = "Y" Then
	If receipt_nm <> "" Then
		RSv("�̎�������") = receipt_nm
	Else
		RSv("�̎�������") = wUserName
	End If
	If receipt_memo <> "" Then
		RSv("�̎����A������") = receipt_memo
	Else
		RSv("�̎����A������") = "�����@���Ƃ���"
	End If
Else
	RSv("�̎�������") = ""
	RSv("�̎����A������") = ""
End If

RSv("���x�[�g�g�p�t���O") = RebateFl
RSv("�ߕs�����E���z") = KabusokuAm

RSv("�ŏI�X�V��") = Now()

RSv.Update
RSv.Close

End Function

'========================================================================
'
'	Function	�͐�̌ڋq���̎擾
'
'========================================================================
Function GetTodokesakiInfo(vAddressNo)

Dim RSv
Dim vSQL

vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    a.�Z������"
vSQL = vSQL & "  , a.�ڋq�X�֔ԍ�"
vSQL = vSQL & "  , a.�ڋq�s���{��"
vSQL = vSQL & "  , a.�ڋq�Z��"
vSQL = vSQL & "  , b.�ڋq�d�b�ԍ�"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    Web�ڋq�Z�� a WITH (NOLOCK)"
vSQL = vSQL & "  , Web�ڋq�Z���d�b�ԍ� b WITH (NOLOCK)"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "        a.�ڋq�ԍ� = " & wUserID
vSQL = vSQL & "    AND a.�Z���A�� = " & vAddressNo
vSQL = vSQL & "    AND b.�d�b�敪 = '�d�b'"
vSQL = vSQL & "    AND b.�ڋq�ԍ� = a.�ڋq�ԍ�"
vSQL = vSQL & "    AND b.�Z���A�� = a.�Z���A��"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic, adLockOptimistic

wShipNm = RSv("�Z������")
wShipZip = RSv("�ڋq�X�֔ԍ�")
wShipPrefecture = RSv("�ڋq�s���{��")
wShipAddress = RSv("�ڋq�Z��")
wShipTel = RSv("�ڋq�d�b�ԍ�")

RSv.Close

End Function

'========================================================================
'
'	Function	���̓f�[�^�[�̃`�F�b�N
'
'========================================================================
Function validate_data()

Dim vAddress
Dim vDateMMDD
Dim vDateTimeMsg '2011/02/16 hn add

'---- �x�����@
If payment_method = "" Then
	wMSG = wMSG & "���x�������@��I�����Ă��������B<br>"
End If

'---- �U���l���`
If cf_checkKataKana(furikomi_nm) = False Then
	wMSG = wMSG & "�U���l���`�͑S�p�J�^�J�i�݂̂œ��͂��Ă��������B<br>"
End If

'---- ������s���i�`�F�b�N
If (i_tokuchuu_fl = "Y" Or i_daibiki_fuka_fl = "Y") And payment_method = "�����" Then
	wMSG = wMSG & "��^���i�܂��͓��ʎ�z�̏��i���܂܂�Ă��邽�߁A������ł̂������͎�t�ł��܂���B���̂��x�������@�ւ̕ύX�����肢���܂��B<br>"
End If

'---- ���͐�
If payment_method = "���[��" And ship_address_no <> 1 Then
	wMSG = wMSG & "���[���ł��x�����̏ꍇ�A���͂���͂��q�l�̏Z���݂̂ƂȂ�A�ʂ̏Z���ւ̔z���w��͂ł��܂���B<br>"
End If

'---- �z�B���w��
If delivery_fl = "Y" And delivery_mm = "" And delivery_dd = "" And delivery_tm = "" Then
	wMSG = wMSG & "�z�B���̓��t�܂��͎��Ԃ���͂��Ă��������B<br>"
End If

If (delivery_mm <> "" And delivery_dd = "") Or (delivery_mm = "" And delivery_dd <> "") Then
	wMSG = wMSG & "�z�B���̎w�肪����������܂���B<br>"
	wDeliveryDt = ""
End If

If delivery_mm <> "" And delivery_dd <> "" Then
	vDateMMDD = cf_NumToChar(DatePart("m", Date()), 2) & cf_NumToChar(DatePart("d", Date()), 2)
	If vDateMMDD > (delivery_mm & delivery_dd) Then
		wDeliveryDt = Cstr(Clng(DatePart("yyyy", Date())) + 1) & "/" & delivery_mm & "/" & delivery_dd
	Else
		wDeliveryDt = DatePart("yyyy", Date()) & "/" & delivery_mm & "/" & delivery_dd
	End If
	If isDate(wDeliveryDt) = False Then
		wMSG = wMSG & "�z�B���̎w�肪����������܂���B<br>"
		wDeliveryDt = ""
	End If
Else
	wDeliveryDt = ""
End If

'---- �z�B���w��(�����w�莞�̃`�F�b�N)
If wDeliveryDt <> "" Then

	'---- �z�B���w��\���`�F�b�N�i���c�i�œ��ח\�肪�Ȃ����A���ח\��+2�����O�͎w��NG�j
	If checkNyukaYoteibi() = False Then
		wMSG = wMSG & "�݌ɂ̂Ȃ����i���������Ɋ܂܂�Ă��邽�߁A�z�B��]���̎w��͂ł��܂���B<br>"
	Else
		'---- ���[���̏ꍇ�A�z�����w��s��
		If payment_method = "���[��" Then
			wMSG = wMSG & "���x�������@�����[���̏ꍇ�A�z�B��]���̎w��͂ł��܂���B<br>"

		Else
			wAfter13FL = False
			wAfter14FL = False
			wAfter15FL = False
			wAfter16FL = False
			'w94HokkaidouRitouFL = False    '2011/06/29 an del
			w94ChugokuHokkaidouFL = False   '2011/06/29 an add
			wHolidayFL = False
			wSatFL = False
			wFriFL = False

			'---- 13���ȍ~���ǂ����`�F�b�N
			If DatePart("h", Now()) >= 13 Then
				wAfter13FL = True
			End If

			'---- 14���ȍ~���ǂ����`�F�b�N
			If DatePart("h", Now()) >= 14 Then
				wAfter14FL = True
			End If

			'---- 15���ȍ~���ǂ����`�F�b�N
			If DatePart("h", Now()) >= 15 Then
				wAfter15FL = True
			End If

			'---- 16���ȍ~���ǂ����`�F�b�N
			If DatePart("h", Now()) >= 16 Then
				wAfter16FL = True
			End If

			'---- ���������j��Փ����ǂ����`�F�b�N
			If (DatePart("w", Date()) = vbSunday) Or (checkHoliday(Date()) = True) Then
				wHolidayFL = True
			End If

			'---- �������y�j�����ǂ����`�F�b�N
			If DatePart("w", Date()) = vbSaturday Then
				wSatFL = True
			End If

			'---- ���������j�����ǂ����`�F�b�N
			If DatePart("w", Date()) = vbFriday Then
				wFriFL = True
			End If

			'---- �z���悪�@��B�E�l���E�����E�k�C�����ǂ����`�F�b�N  
			'If wRitouFl = "Y" Or Instr(w9Shuu4KokuHokkaido, wShipPrefecture) > 0 Then
			if Instr(w9Shuu4KokuChugokuHokkaido, wShipPrefecture) > 0 Then
				'w94HokkaidouRitouFL = True     '2011/06/29 an del
				w94ChugokuHokkaidouFL = True    '2011/06/29 an add
			End If

			'---- �w��\�����Z�b�g
			wAvailableDate = setAvailableDate()

			'---- �z�����`�F�b�N
			If wDeliveryDt < wAvailableDate Then

				'2011/02/16 hn mod s
				'---- �x��
				If wHolidayFl = True Then
					vDateTimeMsg = "�x����"
				else

					'---- �R���r�j �y�j�� 13���ȍ~
					If payment_method = "�R���r�j�x��" AND wSatFl = True AND wAfter13Fl = true Then
						vDateTimeMsg = "�R���r�j�x���̓y�j��13���ȍ~��"
					end if

					'---- �R���r�j ���� 14���ȍ~
					If payment_method = "�R���r�j�x��" AND wSatFl = False AND wAfter14Fl = true Then
						vDateTimeMsg = "�R���r�j�x����14���ȍ~��"
					end if

					'---- ��s�U�� 14���ȍ~
					If payment_method = "��s�U��" AND wAfter14Fl = true Then
						vDateTimeMsg = "��s�U����14���ȍ~��"
					end if

					'---- ��s�U�� �y�j��
					If payment_method = "��s�U��" AND wSatFL = true Then
						vDateTimeMsg = "��s�U���̓y�j����"
					end if

					'---- ������E�J�[�h 16���ȍ~
					If (payment_method = "�����" OR payment_method = "�N���W�b�g�J�[�h") AND wAfter16Fl = true Then
						vDateTimeMsg = payment_method & "��16���ȍ~��"
					end if
				end if

				'---- ��B��l��������E�k�C��
				'If w94HokkaidouRitouFL = True Then                             '2011/06/29 an del
					'wMSG = wMSG & "���͂��悪��B�E�l���E�k�C���E�����ŁA"     '2011/06/29 an del
				If w94ChugokuHokkaidouFL = True Then                            '2011/06/29 an add
					wMSG = wMSG & "���͂��悪��B�E�l���E�����E�k�C���ŁA"      '2011/06/29 an add
				end if
				
				'---- ����                      '2011/06/29 an add s
				If wRitouFl = "Y" Then
					wMSG = wMSG & "���͂��悪����E�����ŁA"
				end if                          '2011/06/29 an add e

				wMSG = wMSG & vDateTimeMsg & "�z�����w��́A" & wAvailableDate & "�ȍ~���w�肵�Ă��������B<br>"
				'2011/02/16 hn mod e

			End If

			'---- 60���ȓ�
			'2013/02/18 GV #1525 MOD START
'			If (DateDiff("d", DateAdd("d", 60, Date()), wDeliveryDt) > 0) Then
'				wMSG = wMSG & "�z�����w���60���ȓ��̓��t���w�肵�Ă��������B<br>"
			If (payment_method = "�����" AND DateDiff("d", DateAdd("d", 10, Date()), wDeliveryDt) > 0) Then
				wMSG = wMSG & "�z�����w���10���ȓ��̓��t���w�肵�Ă��������B<br>"
			ElseIf (payment_method <> "�����" AND DateDiff("d", DateAdd("d", 60, Date()), wDeliveryDt) > 0) Then
				wMSG = wMSG & "�z�����w���60���ȓ��̓��t���w�肵�Ă��������B<br>"
			'2013/02/18 GV #1525 MOD END
			End If
		End If
	End If
End If

'---- �c�Ə��~��
If payment_method = "���[��" And eigyousho_dome_fl = "Y" Then
	wMSG = wMSG & "���[���ł��x�����̏ꍇ�A�c�Ə��~�ߎw��͂ł��܂���B<br>"
End If

'---- �d�ʏ��i����̂Ƃ��̓��}�g�̓G���[  2011/06/29 an del���ڋq�ɂ͑I���ł��Ȃ����ߍ폜
'If freight_forwarder = "2" Then
'	If checkJyuuryouShouhin() > 0 Then
'		wMSG = wMSG & "�������ɏd�ʏ��i���܂܂�Ă��܂��̂Ń��}�g�^�A�̎w��͂ł��܂���B<br>"
'	End If
'End If

'---- ���}�g������͕������̓G���[  2011/06/29 an del���ڋq�ɂ͑I���ł��Ȃ����ߍ폜
'If freight_forwarder = "2" And payment_method = "�����" Then
'	If wKoguchi > 1 Then
'		wMSG = wMSG & "�������͕������ƂȂ�܂��B���}�g�^�A�ł̑�����̎w��͂ł��܂���B�^����Ђ�ύX���邩�A���x�����@��ύX���Ă��������B<br>"
'	End If
'End If

'---- ���}�g+�z�����w��Ȃ�+���Ԏw��̓G���[	2011/05/02 hn add
if (freight_forwarder = "2") AND (delivery_tm <> "") AND (delivery_mm = "")then
	wMSG = wMSG & "�z�����Ԃ��w�肳���ꍇ�́A�z�������w�肵�Ă��������B<br>"
end if

'---- ����+����+���Ԏw��̓G���[
If wRitouFl = "Y" And freight_forwarder = "1" And delivery_tm <> "" Then
	wMSG = wMSG & "���q�l�̂��͂���ւ͎��Ԏw����s���܂���B<br>"
End If

'---- �J�[�h�ŉc�Ə��~�ߎw��̓G���[
If payment_method = "�N���W�b�g�J�[�h" And eigyousho_dome_fl = "Y" Then
	wMSG = wMSG & "�N���W�b�g�J�[�h�ł������̏ꍇ�́A�c�Ə��~�߂̎w��͂ł��܂���B<br>"
End If

'---- �����,���[��,�R���r�j/�X�֋ǎx���̏ꍇ�̎��؂͏o���Ȃ��`�F�b�N
If receipt_fl = "Y" And (payment_method = "�����" Or payment_method = "���[��" Or payment_method = "�R���r�j�x��") Then
	wMSG = wMSG & "�w��̂��x�������@�ł��w���̍ہC�̎����͔��s�ł��܂���B<br>"
End If

'---- �̎��؈���܂��͒A���������͂Ţ�K�v��`�F�b�N����ĂȂ����̓G���[
If receipt_fl <> "Y" And (receipt_nm <> "" Or receipt_memo <> "") Then
	wMSG = wMSG & "�̎������K�v�ȏꍇ�́u�K�v�v���`�F�b�N���Ă��������B�s�v�ȏꍇ�͈����܂��͒A���������N���A���Ă��������B<br>"
End If

'2012/07/26 nt del
'---- �����ŋ�A�֎~���i���܂܂�Ă���ꍇ�͍���̂�OK
'If wRitouFl = "Y" And freight_forwarder <> 1 And checkKuuyukinshiShouhin() > 0 Then
'	wMSG = wMSG & "��A�֎~���i���܂܂�Ă��܂��̂ŉ^����Ђ�����}�ւɕύX���Ă��������B<br>"
'End If

'2012/07/26 nt add start
'---- �u���Z�v+�u���d�ʕi����v+�u�����v+�u������v�̎��̓G���[
if kErrFlg = false and wRitouFl = "Y"  and payment_method = "�����" then
	if wMSG <> "" then wMSG = wMSG & "<br>" end if
	wMSG = wMSG & "���q�l�̂������́A�d�ʁA�������͑傫�����K��l�𒴂��鏤�i���܂ވׁA��������ȊO�̂��x�����@��I�����Ă��������B<br>"
end if

'---- �u���Z�v+�u���d�ʕi����v+�u�����w��F�L��v�̎��̓G���[
if kErrFlg = false and delivery_fl = "Y" then
	if wMSG <> "" then wMSG = wMSG & "<br>" end if
	wMSG = wMSG & "���q�l�̂������́A�d�ʁA�������͑傫�����K��l�𒴂��鏤�i���܂ވׁA�z�B�����Ȃ��ł̂��͂��ƂȂ�܂��B<br>"
end if
'2012/07/26 nt add end

'2012/08/25 nt add start
if kErrFlg = false and wSagawaLTFl = "Y"  and payment_method = "�����" then
	wMSG = wMSG & "���q�l�̂������́A������������󂯂ł��Ȃ��n��ׁ̈A��������ȊO�̂��x�����@��I�����Ă��������B<br>"
end if
'2012/08/25 nt add end

If wMSG <> "" Then
	wMSG = "<b>�ȉ��̓��̓G���[��������Ă��������B</b><br><br>" & wMSG
End If

End Function

'========================================================================
'
'	Function	�����t���O�̐ݒ�
'
'		parm:		�z����X�֔ԍ�
'		return:	����E�����Ȃ�@wRitouFl = Y
'				�����ȊO�@�@�@�@wRitouFl = N
'
'========================================================================
Function setRitouFlag(p_zip)

Dim vZip
Dim RSv
Dim vSQL

vZip = Replace(p_zip, "-", "")

If vZip = "" Then
	wRitouFl  = "N"
	Exit Function
End If

'---- �����`�F�b�N
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    �X�֔ԍ�"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    ���� WITH (NOLOCK)"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "    �X�֔ԍ� = '" & vZip & "'"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic

If RSv.EOF = True Then
	wRitouFl = "N"
	'---- ����͗�������               '2011/06/29 an add s
	if wShipPrefecture = "���ꌧ" then
		wRitouFl = "Y"                 '2011/06/29 an add e
	end if
Else
	wRitouFl = "Y"
End If

RSv.Close

End Function

'2012/08/25 nt add function
'========================================================================
'
'	Function	���������֎~�t���O�̐ݒ�
'
'		parm:		�z����X�֔ԍ�
'		return:	���������֎~�n��Ȃ�@wSagawaLTFl = Y
'				�֎~�n��ȊO�@�@�@�@�@�@wSagawaLTFl = N
'
'========================================================================
Function setSagawaLTFlag(p_zip)

Dim vZip
Dim RSv
Dim vSQL

vZip = Replace(p_zip, "-", "")

If vZip = "" Then
	wSagawaLTFl  = "N"
	Exit Function
End If

'---- ���쐧���`�F�b�N
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    �X�֔ԍ�"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    ���쐧�� WITH (NOLOCK)"
vSQL = vSQL & " WHERE �X�֔ԍ� = '" & vZip & "'"
vSQL = vSQL & "   AND ����s�t���O='Y'"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic

If RSv.EOF = True Then
	wSagawaLTFl = "N"
Else
	wSagawaLTFl = "Y"
End If

RSv.Close

End Function

'========================================================================
'
'	Function	�Փ��`�F�b�N
'
'		input : �`�F�b�N�����(YYYY/MM/DD)
'		return:	�Փ��Ȃ�@True
'				�Փ��ȊO�@False
'
'========================================================================
Function checkHoliday(p_date)

Dim RSv
Dim vSQL

'---- �Փ��`�F�b�N
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    �N����"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    �J�����_�[ WITH (NOLOCK)"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "        �N���� = '" & p_date & "'"
vSQL = vSQL & "    AND �x���t���O = 'Y'"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic

If RSv.EOF = True Then
	checkHoliday  = False
Else
	checkHoliday = True
End If

RSv.Close

End Function

'========================================================================
'
'	Function	�z���\���Z�b�g    2010/12/17 an mod  2011/02/16 hn mod
'
'		return:	�z���\�� (YYYY/MM/DD)
'
'   �x�����@���R���r�j
'     �y�j���ȊO
'     14���ȑO�F��t��������
'     14���ȍ~�F��t��������

'     �y�j��
'     13���ȑO�F��t��������
'     13���ȍ~�F��t��������
'
'   �x�����@����s�U���i�j���֌W�Ȃ��j
'     14���ȑO�F��t��������
'     14���ȍ~�F��t��������
'
'   �x�����@��������i�j���֌W�Ȃ��j
'     16���ȑO�F��t��������
'     16���ȍ~�F��t��������
'
'   �x�����@�����[��
'     �[���w��s��
'
'==============================
'   ��t�����A���j�A�j���A�i��s�U���̏ꍇ�͓y�j�����܂ށj�̏ꍇ�͎��̉c�Ɠ�����t���Ƃ���B
'
'   �[���w��\���́A
'   ��t��+1���@�i���L�ȊO�j
'   ��t��+2���@�i��B�A�l���A�����A�k�C���j
'   ��t��+5���@�i����A�����j
'
'========================================================================
'
Function setAvailableDate()

Dim vOrderDate

vOrderDate = cf_FormatDate(Date(), "YYYY/MM/DD")

'---- �R���r�j�x���F�y�j���ȊO�@14���ȍ~�͎�t���͗���
if payment_method = "�R���r�j�x��" AND wSatFl = false AND wAfter14Fl = true then
	vOrderDate = cf_FormatDate(DateAdd("d", 1, Date()), "YYYY/MM/DD")
end if

'---- �R���r�j�x���F�y�j���@13���ȍ~�͎�t���͗���
if payment_method = "�R���r�j�x��" AND wSatFl = true AND wAfter13Fl = true then
	vOrderDate = cf_FormatDate(DateAdd("d", 1, Date()), "YYYY/MM/DD")
end if

'---- ��s�U���F14���ȍ~�͎�t���͗���
if payment_method = "��s�U��" AND wAfter14Fl = true then
	vOrderDate = cf_FormatDate(DateAdd("d", 1, Date()), "YYYY/MM/DD")
end if

'---- ������E�J�[�h�F16���ȍ~�͎�t���͗���
If (payment_method = "�����" OR payment_method = "�N���W�b�g�J�[�h") AND wAfter16Fl = true Then
	vOrderDate = cf_FormatDate(DateAdd("d", 1, Date()), "YYYY/MM/DD")
end if

'---- ��t�����A���j�A�j���A�i��s�U���̏ꍇ�͓y�j�����܂ށj�̏ꍇ�͎��̉c�Ɠ�����t���Ƃ���B
Do
	if  (DatePart("w", vOrderDate) = vbSunday) OR (checkHoliday(vOrderDate) = True) then
		vOrderDate = cf_FormatDate(DateAdd("d", 1, vOrderDate), "YYYY/MM/DD")
	else 
		if  payment_method = "��s�U��" AND (DatePart("w", vOrderDate) = vbSaturday) then
			vOrderDate = cf_FormatDate(DateAdd("d", 1, vOrderDate), "YYYY/MM/DD")
		else 
			exit do
		end if
	end if
Loop

'---- �w��\��
'if w94HokkaidouRitouFl = true then                '2011/06/29 an del
'---- �����F��t��+5��                             '2011/06/29 an mod s
if wRitouFl = "Y" then
	setAvailableDate = cf_FormatDate(DateAdd("d", 5, vOrderDate), "YYYY/MM/DD")
else
	'---- ��B��l��������E�k�C���F��t��+2��
	if w94ChugokuHokkaidouFL = true then
		setAvailableDate = cf_FormatDate(DateAdd("d", 2, vOrderDate), "YYYY/MM/DD")
	'---- ���̑��F��t��+1��                       '2011/06/29 an mod e
	else
		setAvailableDate = cf_FormatDate(DateAdd("d", 1, vOrderDate), "YYYY/MM/DD")
	end if
end if

End function

'========================================================================
'
'	Function	�d�ʏ��i�`�F�b�N
'
'		return:	�d�ʏ��i����
'
'========================================================================
Function checkJyuuryouShouhin()

Dim RSv
Dim vSQL

'---- �d�ʏ��i�������o��
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    COUNT(*) AS �d�ʏ��i����"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    ���󒍖��� a WITH (NOLOCK)"
vSQL = vSQL & "  , Web���i b WITH (NOLOCK)"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "        b.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
vSQL = vSQL & "    AND b.���i�R�[�h = a.���i�R�[�h"
vSQL = vSQL & "    AND a.SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod
vSQL = vSQL & "    AND b.�����敪 = '�d�ʏ��i'"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic

checkJyuuryouShouhin = RSv("�d�ʏ��i����")

RSv.Close

End Function

'2012/07/26 nt comment mod
'========================================================================
'
'	Function	��A�֎~���i�`�F�b�N�ː��d�ʕi�`�F�b�N
'               12/07/26�A��A�֎~�t���O�͖��g�p�̂��߁A�{�t���O�͐��d�ʕi
'               ���ʂ̃t���O�ւƓ]�p�B
'
'		return:	��A�֎~���i�����ː��d�ʕi����
'
'========================================================================
Function checkKuuyukinshiShouhin()

Dim RSv
Dim vSQL

'---- ��A�֎~���i�������o��
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    COUNT(*) AS ��A�֎~���i����"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    ���󒍖��� a WITH (NOLOCK)"
vSQL = vSQL & "  , Web���i b WITH (NOLOCK)"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "        b.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
vSQL = vSQL & "    AND b.���i�R�[�h = a.���i�R�[�h"
vSQL = vSQL & "    AND a.SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod
vSQL = vSQL & "    AND b.��A�֎~�t���O = 'Y'"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic

checkKuuyukinshiShouhin =  RSv("��A�֎~���i����")

RSv.Close

End Function

'========================================================================
'
'	Function	���c�̗L���Ɠ��ח\������m�F���A�z�B���w��\���`�F�b�N
'
'		return:	 �z�B���w��\�Ȃ� True
'                �z�B���w��s�Ȃ� False
'
'========================================================================
Function checkNyukaYoteibi()

Dim RSv
Dim vSQL
Dim vHikiateKanouQt
Dim vSetCount
Dim vMaxNyukaYoteibi

checkNyukaYoteibi = True

vHikiateKanouQt = ""
vSetCount = ""
vMaxNyukaYoteibi =""

'---- ���󒍖��׎��o��
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "     a.���[�J�[�R�[�h"
vSQL = vSQL & "   , a.���i�R�[�h"
vSQL = vSQL & "   , a.�F"
vSQL = vSQL & "   , a.�K�i"
vSQL = vSQL & "   , a.�󒍐���"
vSQL = vSQL & "   , b.�����\����"
vSQL = vSQL & "   , b.�����\���ח\���"
vSQL = vSQL & "   , c.�Z�b�g���i�t���O"
vSQL = vSQL & " FROM"
vSQL = vSQL & "     ���󒍖��� a WITH (NOLOCK)"
vSQL = vSQL & "   , Web�F�K�i�ʍ݌� b WITH (NOLOCK)"
vSQL = vSQL & "   , Web���i c WITH (NOLOCK)"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "         b.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
vSQL = vSQL & "     AND b.���i�R�[�h = a.���i�R�[�h"
vSQL = vSQL & "     AND b.�F = a.�F"
vSQL = vSQL & "     AND b.�K�i = a.�K�i"
vSQL = vSQL & "     AND c.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
vSQL = vSQL & "     AND c.���i�R�[�h = a.���i�R�[�h"
vSQL = vSQL & "     AND b.�I���� IS NULL"
vSQL = vSQL & "     AND a.SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic

If RSv.EOF = False Then

	'---- �S���󒍖��׏��i���`�F�b�N
	Do Until RSv.EOF = True

		'�Z�b�g�i�̏ꍇ�̓Z�b�g�i�S�̂̍݌ɐ��ʁA���ח\������`�F�b�N����
		If RSv("�Z�b�g���i�t���O") = "Y" Then

			'---- �Z�b�g�i�݌ɐ��ʁAMAX�����\�\��擾
			Call GetSetCount(RSv("���[�J�[�R�[�h"), RSv("���i�R�[�h"), RSv("�F"), RSv("�K�i"), vSetCount, vMaxNyukaYoteibi)

			If vMaxNyukaYoteibi < Date() Then			'���ח\��̂Ȃ��Z�b�g�i�̏ꍇ��2000/01/01�������Ă���
				vMaxNyukaYoteibi = ""
			Else
				vMaxNyukaYoteibi = vMaxNyukaYoteibi		'�Z�b�g�i�S�̂�MAX���ח\����ŏ㏑��
			End If

			vHikiateKanouQt = vSetCount					'�Z�b�g�i�S�̂�MIN�݌ɐ��ʂŏ㏑��

		Else
			vHikiateKanouQt = RSv("�����\����")
			vMaxNyukaYoteibi = RSv("�����\���ח\���")
		End If

		If RSv("�󒍐���")  > vHikiateKanouQt Then		'���c�i�̏ꍇ�͓��ח\������`�F�b�N
			'---- ���ח\��̂Ȃ����i��1�ł�����ꍇ�͔z�B���w��s��
			If IsNULL(vMaxNyukaYoteibi) = True Or vMaxNyukaYoteibi = "" Then
				checkNyukaYoteibi = False
				Exit Do
			Else
				'---- �w��\���i���ח\���+2�ȍ~�j
				wAvailableDate = cf_FormatDate(DateAdd("d", cAddDaysToNyukaYoteibi, vMaxNyukaYoteibi), "YYYY/MM/DD")

				'---- ���ׂ��Ԃɍ���Ȃ����i��1�ł�����ꍇ�͔z�B���w��s��
				If wDeliveryDt < wAvailableDate Then
					checkNyukaYoteibi = False
					Exit Do
				End If
			End If
		End If

		RSv.MoveNext
	Loop
End If

RSv.Close

End Function

'========================================================================
'
'	Function	�^����Ѓ`�F�b�N���ύX    2011/06/29 an add
'
'========================================================================

Function CheckFreightForwarder()

Dim RSv
Dim vSQL

Dim vItemChar1   '2011/08/11 an add s
Dim vItemChar2
Dim vItemNum1
Dim vItemNum2
Dim vItemDate1
Dim vItemDate2   '2011/08/11 an add e

'---- ����
If wRitouFl = "Y" then
'2011/07/25 hn del s
'	'---- ����1���ŋ�A�֎~���i���܂܂�ĂȂ��d�ʏ��i�łȂ��ꍇ�́A�z����Ђ��u���}�g�^�A�v�ɋ����ύX
'	if wKoguchi = 1 And checkKuuyukinshiShouhin() = 0 And checkJyuuryouShouhin() = 0 Then
'		freight_forwarder = "2"
'	'---- ��L�ȊO�̗����͍���
'	else
'2011/07/25 hn del e
		freight_forwarder = "1"
'	end if	'2011/07/25 hn del 

'---- �����ȊO
else
	
	'---- �K��^����Ђ��Z�b�g
	vSQL = ""
	vSQL = vSQL & "SELECT �^����ЃR�[�h"
	vSQL = vSQL & "  FROM ���ʋK��^����� WITH (NOLOCK)"
	vSQL = vSQL & " WHERE �� = '" & wShipPrefecture & "'"
	
	Set RSv = Server.CreateObject("ADODB.Recordset")
	RSv.Open vSQL, Connection, adOpenStatic
	
	if RSv.EOF = false then
		freight_forwarder = RSv("�^����ЃR�[�h")
	end if
	
	RSv.Close

	'---- ����Œ�Ȃ獲��ɕύX   2011/09/11 an add s
	vSQL = ""
	vSQL = vSQL & "SELECT ����Œ�t���O"
	vSQL = vSQL & "  FROM Web�ڋq WITH (NOLOCK)"
	vSQL = vSQL & " WHERE �ڋq�ԍ� = " & wUserID
	
	Set RSv = Server.CreateObject("ADODB.Recordset")
	RSv.Open vSQL, Connection, adOpenStatic
	
	if RSv.EOF = false then
		if RSv("����Œ�t���O") = "Y" then
			freight_forwarder = "1"
		end if
	end if
	
	RSv.Close                      '2011/09/11 an add e
	
end if

'---- ���Z�Ŕz���s�̏ꍇ�͉^����Ђ�ύX
'if freight_forwarder = "5" AND wDeliveryDt <> "" then   '2011/08/11 an del
if freight_forwarder = "5" then                          '2011/08/11 an mod s

	'---- �z�B�w�������̏ꍇ�A���Z�z�B�s���łȂ����`�F�b�N
	if wDeliveryDt <> "" then
		'---- ���j�Ȃ獲��
		if DatePart("w", wDeliveryDt) = vbSunday then
			freight_forwarder = "1"
		'---- ���Z�z���s���Ȃ獲��
		else
			vSQL = ""
			vSQL = vSQL & "SELECT ���Z�z�B�s���t���O"
			vSQL = vSQL & "  FROM �J�����_�[ WITH (NOLOCK)"
			vSQL = vSQL & " WHERE �N���� = '" & wDeliveryDt & "'"

			Set RSv = Server.CreateObject("ADODB.Recordset")
			RSv.Open vSQL, Connection, adOpenStatic
			
			if RSv.EOF = false then
				if RSv("���Z�z�B�s���t���O") = "Y" then
					freight_forwarder = "1"
				end if
			end if
			
			RSv.Close
		end if
	end if

	'---- ���Z�d���R�[�h�}�X�^�`�F�b�N    2011/09/12 an mod s
	vSQL = ""
	vSQL = vSQL & "SELECT �z�B�ߑO�ߌ�"
	vSQL = vSQL & "     , ���[�h�^�C��"
	vSQL = vSQL & "  FROM ���Z�d���R�[�h�}�X�^ WITH (NOLOCK)"
	vSQL = vSQL & " WHERE �X�֔ԍ� = '" & Replace(wShipZip,"-","") & "'"

	Set RSv = Server.CreateObject("ADODB.Recordset")
	RSv.Open vSQL, Connection, adOpenStatic
	
	if RSv.EOF = false then
		
		'---- ���Ԏw��_���Z 01�̕�����擾�i�ߑO�w��j
		call getCntlMst("��","���Ԏw��_���Z","01", vItemChar1, vItemChar2, vItemNum1, vItemNum2, vItemDate1, vItemDate2)
	
		'---- ���Z�ߑO�s�n��Ȃ獲��ɕύX
		if delivery_tm = vItemChar1 then
			if RSv("�z�B�ߑO�ߌ�") = "P" then
				freight_forwarder = "1"
			end if
		end if
		
		'---- ���[�h�^�C�� >= 2�Ȃ獲��ɕύX
		if RSv("���[�h�^�C��") >= 2 then
			freight_forwarder = "1"
		end if
	end if
	
	RSv.Close      '2011/08/11 an mod e  2011/09/12 an mod e
	
end if

'2012/07/26 nt add start
'---- ���d�ʕi�A�^����Ђ̑I�萧��
'----�u���d�ʕi�v���܂߂΁A�����t���Łu���Z�v�֕ύX
kErrFlg = true
if checkKuuyukinshiShouhin() <> 0 then
	freight_forwarder = "5"

	'---- �u�����v�ŁA���u������v�w��̏ꍇ�̓G���[�Ώ�
	if wRitouFl = "Y"  and payment_method = "�����" then
		kErrFlg = false

	'---- �u�����w��v������ꍇ�̓G���[�Ώ�
	elseif delivery_fl = "Y" then
		kErrFlg = false

	end if
end if
'2012/07/26 nt add end

'2012/08/25 nt add start
'---- �u���������֎~�n��v�ŁA���u������v�w��̏ꍇ�̓G���[�Ώ�
if wSagawaLTFl = "Y" and payment_method = "�����" then
	kErrFlg = false
end if
'2012/08/25 nt add end


End Function

%>
