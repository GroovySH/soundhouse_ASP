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
'	�F�B�ɂ����߂鑗�M
'
'�X�V����
'2007/08/23 ���i�A�N�Z�X�����o�^�i�F�B�ɂ����߂�j
'2007/09/10 ���i�A�N�Z�X�����o�^��N���ʂɕύX
'2009/04/30 �G���[����error.asp�ֈړ�
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'
'========================================================================

On Error Resume Next

Dim userEmail
Dim UserName

Dim Item
Dim ToAddr
Dim FromName
Dim Message1
Dim Message

Dim wItem
Dim MakerCd
Dim ProductCd

Dim wItemChar1
Dim wItemChar2
Dim wItemNum1
Dim wItemNum2
Dim wItemDate1
Dim wItemDate2

Dim wBody

Dim Connection
Dim RS

Dim wSQL
Dim wHTML
Dim wMSG
Dim wErrDesc   '2011/08/01 an add

'========================================================================

Response.buffer = true

'---- UserID ���o��

'---- �Ăяo��������̃f�[�^���o��
Item = ReplaceInput(Request("Item"))
ToAddr = ReplaceInput_NoCRLF(Request("ToAddr"))  '2011/08/01 an mod
FromName = ReplaceInput(Request("FromName"))
Message = ReplaceInput(Request("Message"))
Message1 = ReplaceInput(Request("Message1"))

wItem = Split(Item, "^")
MakerCd = Trim(wItem(0))
ProductCd = Trim(wItem(1))

'---- Execute main
call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "TellaFriendSend.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
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
'========================================================================
'
Function main()

Dim i
Dim vYYYYMM
Dim OBJ_NewMail
Dim RSv

'---- �������߃��[���쐬�i�ڋq��)
call getCntlMst("Web","Email","��ʃg���[��", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)

wBody = Message1 & vbNewLine & Message & vbNewLine & vbNewLine & wItemChar1

Set OBJ_NewMail = Server.CreateObject("CDO.Message") 

OBJ_NewMail.from = "shop@soundhouse.co.jp"
OBJ_NewMail.to = ToAddr

OBJ_NewMail.subject = FromName & "�@�l����A�������߃��[�����͂��Ă��܂�"
OBJ_NewMail.TextBody = wBody
OBJ_NewMail.BodyPart.Charset = "iso-2022-jp"

OBJ_NewMail.Send

Set OBJ_NewMail = Nothing

'---- ���i�A�N�Z�X�����o�^�i�F�B�ɂ����߂�j
vYYYYMM = Year(Now()) & Right("0" & Month(Now()),2)
wSQL = ""
wSQL = wSQL & "SELECT *"
wSQL = wSQL & "  FROM ���i�A�N�Z�X����"
wSQL = wSQL & " WHERE ���[�J�[�R�[�h = '" & MakerCd & "'"
wSQL = wSQL & "   AND ���i�R�[�h = '" & ProductCd & "'"
wSQL = wSQL & "   AND �N�� = '" & vYYYYMM & "'"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RSv.EOF = true then
	RSv.AddNew

	RSv("���[�J�[�R�[�h") = MakerCd
	RSv("���i�R�[�h") = ProductCd
	RSv("�N��") = vYYYYMM
	RSv("�F�B�ɂ����ߌ���") = 1
else
	RSv("�F�B�ɂ����ߌ���") = RSv("�F�B�ɂ����ߌ���") + 1
end if

RSv.Update
RSv.close

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
<html lang="ja">
<head>
<meta charset="Shift_JIS">
<title>���F�B�ɂ����߂�b�T�E���h�n�E�X</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/ask.css" type="text/css">
</head>

<body>

<!--#include file="../Navi/NaviTop.inc"-->

<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="��������{���ł�"></a></span>
  
  <!-- �R���e���cstart -->
  <div id="globalContents">
    <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
      <p class="home"><a href="<%=g_HTTP%>"><img src="<%=g_RelLink%>images/icon_home.gif" alt="HOME"></a></p>
      <ul id="path">
        <li class="now">���F�B�ɂ����߂�</li>
      </ul>
    </div></div></div>

    <h1 class="title">���F�B�ɂ����߂�</h1>
    <p>�ȉ��̓��e�Ń��[���𑗐M���܂����B<br>���肪�Ƃ��������܂����B</p>
    
    <table class="form">
      <tr>
        <th>����</th>
        <td><%=ToAddr%></td>
      </tr>
      <tr>
        <th>����</th>
        <td><%=FromName%> �l����A�������߃��[�����͂��Ă��܂�</td>
      </tr>
      <tr>
        <th>���b�Z�[�W</th>
        <td><p><%=Replace(wBody, vbNewLine, "<br>")%></p></td>
      </tr>
    </table>
    <p><a href="ProductDetail.asp?Item=<%=Item%>">���i�y�[�W�֖߂�</a></p>

</div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<div class="tooltip"><p>ASK</p></div>
<!--#include file="../Navi/NaviScript.inc"-->
<script type="text/javascript" src="jslib/ask.js"></script>
</body>
</html>