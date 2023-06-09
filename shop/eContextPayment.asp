<%@ LANGUAGE="VBScript" %>
<%
'�l�b�g�n�E�X�˂��ƃn�E�X�l�b�g�͂���
'�T�E���h�n�E�X
 Option Explicit
%>
<!--#include file="../common/ADOVBS.inc"-->
<!--#include file="../common/system_common_econ.inc"-->
<!--#include file="../common/shop_common_functions.inc"-->
<!--#include file="../common/bfunctions1.asp"-->
<%
'========================================================================
'
'	eontext�������o�^
'		eContext����Ăяo�����
'
'�X�V����
'2009/04/30 �G���[����error.asp�ֈړ�
'2011/02/12 ss �����ς݂̏ꍇ�C�[�R���� -2 �̃G���[�߂�l��Ԃ��Ă��邪�A
'              ����l 1 ��Ԃ�DB�X�V�͂��Ȃ��悤�ύX
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2012/01/30 GV �������(��M�f�[�^)���O�̏o�͈ʒu �ύX
'
'========================================================================

On Error Resume Next

Dim OrderID		'�T�C�g�����ԍ�
Dim ShopID		'�T�C�g�V���b�vID
Dim ID				'�f�[�^ID�@�����ʒm�F0
Dim PayDate		'������ YY/MM/DD HH:MM:SS
Dim PayBy			'�������@�敪 0:���� 1:�N���W�b�g�J�[�h
Dim CvsCode		'�R���r�j��ƃR�[�h
Dim KssspCode	'�R���r�j�X�܃R�[�h
Dim InputID		'�ڋq���d�b�ԍ����ɓ��͂������e
Dim OrdAmount	'�������v

Dim Connection
Dim RS

Dim wSQL
Dim wHTML
Dim wMSG
Dim wErrDesc   '2011/08/01 an add

Dim FS
Dim FS_Data
Dim DataFileName

'=======================================================================

'---- execute main process
call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "eContextPayment.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
'	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc)
end if                                           '2011/08/01 an add e

call close_db()

if Err.Description <> "" then
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

'---- eContext�փ��b�Z�[�W���M
if wMSG = "" then
	wHTML = "1 eContextPayment.asp ����"
else
	if wMSG = "�����ς�" then			'2011/02/12 ss add
		wHTML = "1 eContextPayment.asp ����"	'2011/02/12 ss add
	else						'2011/02/12 ss add
		wHTML = "-2 eContextPayment.asp " & wMSG
	end if						'2011/02/12 ss add
end if

Response.write(wHTML)

'=======================================================================
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

'---- ��M�f�[�^�[�̎��o��
OrderID = ReplaceInput(Trim(Request("OrderID")))
ShopID = ReplaceInput(Trim(Request("ShopID")))
ID = ReplaceInput(Trim(Request("ID")))
PayDate = ReplaceInput(Trim(Request("PayDate")))
PayBy = ReplaceInput(Trim(Request("PayBy")))
CvsCode = ReplaceInput(Trim(Request("CvsCode")))
KssspCode = ReplaceInput(Trim(Request("KssspCode")))
InputID = ReplaceInput(Trim(Request("InputID")))
OrdAmount = ReplaceInput(Trim(Request("OrdAmount")))

'---- ���̓f�[�^�[�̃`�F�b�N
call validate_data()

'---- eContext�������o�^
if wMSG = "" then
	call InserteContextPayment()
end if

'---- ��M�f�[�^�i�[
Set FS = CreateObject("Scripting.FileSystemObject")
' 2012/01/30 GV Mod Start
'DataFileName = "eContextData/�����ʒm" & Year(Date()) & Right("0" & Month(Date()), 2) & Right("0" & Day(Date()), 2) & ".txt"
'DataFileName = Server.MapPath(DataFileName)		'Map log file
DataFileName = "�����ʒm" & Year(Date()) & Right("0" & Month(Date()), 2) & Right("0" & Day(Date()), 2) & ".txt"
DataFileName = g_LogRoot & g_eContextDataLog & DataFileName
' 2012/01/30 GV Mod End
Set FS_Data = FS.OpenTextFile(DataFileName, 8, true)			'File open - Append Mode

FS_Data.WriteLine(cf_FormatTime(Now(), "HH:MM:SS") & " OrderID=" & OrderID & ",ShopID=" & ShopID & ",ID=" & ID & ",PayDate=" & PayDate & ",PayBy=" & PayBy & ",CvsCode=" & CvsCode & ",KssspCode=" & KssspCode & ",InpitID=" & InputID & ",OrdAmount=" & OrdAmount & ",MSG=" & wMSG)

FS_Data.Close

End Function

'========================================================================
'
'	Function	���̓f�[�^�[�̃`�F�b�N
'
'========================================================================
'
Function validate_data()

'---- �T�C�g�V���b�vID
if ShopID <> g_eContext_ID then
	wMSG = wMSG & "�T�C�g�R�[�h�s�� ShopID=" & ShopID & " "
end if

'---- �f�[�^ID�@�����ʒm�F0
if ID <> "0" then
	wMSG = wMSG & "�f�[�^ID�s�� ID=" & ID & " "
end if

'---- OrderID
if OrderID = "" Or IsNumeric(OrderID) = false then
	wMSG = wMSG & "OrderID�Ȃ� "
end if

'---- PayDate
if IsDate(PayDate) = false then
	wMSG = wMSG & "PayDate�s�� "
end if

'---- OrdAmount
if IsNumeric(OrdAmount) = false then
	wMSG = wMSG & "OrdAmount�s�� "
end if

if wMSG <> "" then
	exit function
end if

'---- �󒍔ԍ�
wSQL = ""
wSQL = wSQL & "SELECT �󒍍��v���z"
wSQL = wSQL & "  FROM Web�� WITH (NOLOCK)"
wSQL = wSQL & " WHERE �󒍔ԍ� = " & OrderID

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RS.EOF = true then
	wMSG = wMSG & "�Y���󒍏��Ȃ� "
else
	'---- �󒍍��v
	if isNumeric(OrdAmount) = false then
			wMSG = wMSG & "�󒍍��v�s�� OrdAmount=" & OrdAmount & " "
	else
		if Clng(OrdAmount) <> RS("�󒍍��v���z") then
			wMSG = wMSG & "�󒍍��v�s��v OrdAmount=" & OrdAmount & " �󒍋��z=" & RS("�󒍍��v���z") & " "
		end if
	end if
end if

RS.Close

'2011/02/12 ss add ��
if wMSG <> "" then
	exit function
end if
'2011/02/12 ss add ��

'---- �����ς݃`�F�b�N
wSQL = ""
wSQL = wSQL & "SELECT ������"
wSQL = wSQL & "  FROM eContext����"
wSQL = wSQL & " WHERE �󒍔ԍ� = " & OrderID

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RS.EOF = false then
'		wMSG = wMSG & "�����ς� ������=" & cf_FormatDate(RS("������"), "YYYY/MM/DD") & " " & cf_FormatTime(RS("������"), "HH:MM:SS") & " "	'2011/02/12 ss del
		wMSG = wMSG & "�����ς�"	'2011/02/12 ss add
end if

RS.Close

End Function

'========================================================================
'
'	Function	eContext�������o�^
'
'========================================================================
'
Function InserteContextPayment()

'----
wSQL = ""
wSQL = wSQL & "SELECT *"
wSQL = wSQL & "  FROM eContext����"
wSQL = wSQL & " WHERE 1 = 2"

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic, adLockOptimistic

RS.AddNew

RS("�󒍔ԍ�") = OrderID
RS("������") = PayDate
RS("eContext�����敪") = Left(PayBy, 1)
RS("��ƃR�[�h") = Left(CvsCode, 10)
RS("�R���r�j�X�܃R�[�h") = Left(KssspCode, 20)
RS("�ڋq���͓d�b�ԍ�") = Left(InputID, 20)
RS("�������z") = OrdAmount
RS("������M��") = Now()

RS.Update
RS.Close

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

%>
