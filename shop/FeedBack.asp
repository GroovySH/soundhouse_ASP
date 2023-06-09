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
'	�o�׌�A���P�[�g
'
'�X�V����
'2007/04/23 ���͗������`�F�b�N�ǉ�
'2010/03/17 hn ���O�C���`�F�b�N��ǉ�
'2010/07/16 st �A���P�[�g����������500�����ȓ��ɑΉ�
'2010/12/07 an ����3,4�̏��ԓ���ւ����t�H�[�����͕ύX���Ȃ�
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2011/08/31 an �g�т���A�N�Z�X����mobi��Redirect����
'2012/06/08 GV #1367 �A���P�[�g�ł̍w���҃`�F�b�N
'2012/07/25 ok �w���҃`�F�b�N��EmaxDB�̎󒍃e�[�u������m�F����悤�ύX
'========================================================================

On Error Resume Next

Dim userID		'2010/03/17 hn add
Dim OrderNo

Dim q1
Dim q1Name
Dim q1Department
Dim q1Comment
Dim q2
Dim q2Comment
Dim q3         '����4�ɊY��
Dim q3Comment  '����4�ɊY��
Dim q4         '����3�ɊY��
Dim q4Other    '����3�ɊY��
Dim q5
Dim q5Comment
Dim q6
Dim q6Other
Dim q7
Dim q7Comment

Dim wSQL
Dim wMsg
Dim wErrDesc   '2011/08/01 an add

'2010/07/16 st add s
Dim Connection
Dim RS
Dim wCanWriteFl
'2010/07/16 st add e
Dim ConnectionEmax		'2012/07/25 ok Add
'========================================================================

Response.buffer = true

wMsg = ""

'---- Session�ɃA���P�[�g���M�t���O������΂�����g�p		2010/07/16 st add
if Session("CanWriteFl") <> "" then
	wCanWriteFl = Session("CanWriteFl")
else
	wCanWriteFl = "N"
end if

'---- UserID ���o��	'2010/03/17 hn add
userID = Session("userID")

'---- �p�����[�^��荞��
OrderNo = ReplaceInput(Request("OrderNo"))

q1 = ReplaceInput(Left(Request("q1"), 10))
q1Name = ReplaceInput(Left(Request("q1Name"), 10))
q1Department = ReplaceInput(Left(Request("q1Department"), 10))
q1Comment = ReplaceInput(Left(Request("q1Comment"), 500))
q2 = ReplaceInput(Left(Request("q2"), 10))
q2Comment = ReplaceInput(Left(Request("q2Comment"), 500))
q3 = ReplaceInput(Left(Request("q3"), 10))                  '����4�ɊY��
q3Comment = ReplaceInput(Left(Request("q3Comment"), 500))   '����4�ɊY��
q4 = ReplaceInput(Left(Request("q4"), 150))                 '����3�ɊY��
q4Other = ReplaceInput(Left(Request("q4Other"), 50))        '����3�ɊY��
q5 = ReplaceInput(Left(Request("q5"), 10))
q5Comment = ReplaceInput(Left(Request("q5Comment"), 500))
q6 = ReplaceInput(Left(Request("q6"), 50))
q6Other = ReplaceInput(Left(Request("q6Other"), 500))
q7 = ReplaceInput(Left(Request("q7"), 10))
q7Comment = ReplaceInput(Left(Request("q7Comment"), 500))

'Response.Write("UserId:" & userID & "<br>")
'Response.Write("OrderNo:" & Request("OrderNo") & "<br>")

'---- �g�т���A�N�Z�X���ꂽ�ꍇ��mobi��Redirect    2011/08/31 an add s
if gPhoneType = "NMB" then
	Response.Redirect g_HTTPmobi & "shop/FeedBack.asp?OrderNo=" & OrderNo
elseif gPhoneType = "SP" then
	Response.Redirect g_HTTPsp & "shop/FeedBack.asp?OrderNo=" & OrderNo
end if                                             '2011/08/31 an add e

'---- Session�ɃI�[�_�[�ԍ�������΂�����g�p		2010/03/17 hn add
if OrderNo = "" then
	 OrderNo = Session("OrderNo")
else
	Session("OrderNo") = OrderNo
end if

'---- ���O�C�������Ă��邩�ǂ����̃`�F�b�N	2010/03/17 hn add
if userID = "" then
	wMsg = "���O�C�������Ă��������B" 
	Session("msg") = wMSG
	Response.Redirect g_HTTPS & "shop/Login.asp?called_from=feedback"
end if

'---- ���ږ����͂��������I�[�o�[���̏�������			2010/07/16 st add s
if Session("msg") <> "" then
	wMsg = "<p class='error'>" & Session("msg") & "</p>"
	Session("msg") = ""
else
	call connect_db()
	call main()
	
		'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
	if Err.Description <> "" then
		wErrDesc = "FeedBack.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
		call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
	end if                                               '2011/08/01 an add e

	call close_db()
	
	if Err.Description <> "" then                     '2011/08/01 an add s
		Response.Redirect g_HTTP & "shop/Error.asp"
	end if                                            '2011/08/01 an add e

end if

'2010/07/16 st add e

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

'2012/07/25 ok Add
Set ConnectionEmax = Server.CreateObject("ADODB.Connection")
ConnectionEmax.Open g_connectionEmax

End function
'========================================================================
'
'	Function	main proc
'
'========================================================================
'
Function main()

if OrderNo = "" then
	wMsg = "<p class='error'>�󒍔ԍ�������܂���̂ł��̃A���P�[�g�͑��M�ł��܂���B</p>"
else
    '//Add GV #1367 Start
    If BuyerCheck = False Then
        wMsg = "<p class='error'>���w���җl�̃��[�U�[ID�ƈقȂ邽�߂��̃A���P�[�g�͑��M�ł��܂���B</p>"
        Exit Function
    End If
    '//Add GV #1367 End

	'---- �A���P�[�g�o�^�L���`�F�b�N
	wSQL = ""
	wSQL = wSQL & "SELECT *"
	wSQL = wSQL & "  FROM �o�׌�A���P�[�g WITH (NOLOCK)"
	wSQL = wSQL & " WHERE �󒍔ԍ� = " & OrderNo 
		  
	Set RS = Server.CreateObject("ADODB.Recordset")
	RS.Open wSQL, Connection, adOpenStatic

	if RS.EOF = false then
		wMsg = "<p class='error'>" & "�󒍔ԍ�[" &  OrderNo & "] �Ɋւ���A���P�[�g�́A���ɓo�^����Ă��܂��B" & "</p>"
	else
		wCanWriteFL = "Y"
	end if
	RS.close
end if

End function

'Add GV #1367 START
'========================================================================
'
'	Function	BuyerCheck
'
'========================================================================
Function BuyerCheck()

    wSQL = ""
	wSQL = wSQL & "SELECT *"
'2012/07/25 ok Add Start
	wSQL = wSQL & "  FROM �� a WITH (NOLOCK) "
	wSQL = wSQL & "WHERE "
    wSQL = wSQL & "     a.�ڋq�ԍ� = " & userID & " AND " 
    wSQL = wSQL & "     a.�󒍔ԍ� = " & OrderNo

	Set RS = Server.CreateObject("ADODB.Recordset")
	RS.Open wSQL, ConnectionEmax, adOpenStatic, adLockOptimistic
'2012/07/25 ok Add End

'2012/07/25 ok Del Start
'	wSQL = wSQL & "  FROM Web�ڋq a WITH (NOLOCK)"
'    wSQL = wSQL & " INNER JOIN Web�� b ON a.�ڋq�ԍ� = b.�ڋq�ԍ�"
'	wSQL = wSQL & " WHERE " 
'    wSQL = wSQL & "     a.�ڋq�ԍ� = " & userID & " AND " 
'    wSQL = wSQL & "     b.�󒍔ԍ� = " & OrderNo

'    Set RS = Server.CreateObject("ADODB.Recordset")
'	RS.Open wSQL, Connection, adOpenStatic
'2012/07/25 ok Del End

'    Response.Write RS.RecordCount
    If RS.RecordCount > 0 Then
        BuyerCheck = True
    Else
        BuyerCheck = False
    End If
    
    RS.Close

End Function
'Add GV #1367 END

'========================================================================
'
'	Function	Close database
'
'========================================================================
'
Function close_db()

Connection.close
Set Connection= Nothing    '2011/08/01 an add

'2012/07/25 ok Add
ConnectionEmax.Close
Set ConnectionEmax = Nothing

End function

'========================================================================
%>

<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="Shift_JIS">
<title>���w���җl�����A���P�[�g�b�T�E���h�n�E�X</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/feedback.css" type="text/css">

<script type="text/javascript">
//
// ====== 	Function:	FeedBack_onSubmit
//
function FeedBack_onSubmit(pForm){
	if (pForm.q1Comment.value.length > 500){
		alert("����1�ɓ��͂��ꂽ��������500�����𒴂��Ă��܂���@500�����ȓ��ł��肢���܂��B");
		return false;
	}	
	if (pForm.q2Comment.value.length > 500){
		alert("����2�ɓ��͂��ꂽ��������500�����𒴂��Ă��܂���@500�����ȓ��ł��肢���܂��B");
		return false;
	}
	if (pForm.q3Comment.value.length > 500){
		alert("����4�ɓ��͂��ꂽ��������500�����𒴂��Ă��܂���@500�����ȓ��ł��肢���܂��B");
		return false;
	}
	if (pForm.q5Comment.value.length > 500){
		alert("����5�ɓ��͂��ꂽ��������500�����𒴂��Ă��܂���@500�����ȓ��ł��肢���܂��B");
		return false;
	}
	if (pForm.q6Other.value.length > 500){
		alert("����6�ɓ��͂��ꂽ��������500�����𒴂��Ă��܂���@500�����ȓ��ł��肢���܂��B");
		return false;
	}
	if (pForm.q7Comment.value.length > 500){
		alert("����7�ɓ��͂��ꂽ��������500�����𒴂��Ă��܂���@500�����ȓ��ł��肢���܂��B");
		return false;
	}
	return true;
}

</script>

</head>
<body>

<!--#include file="../Navi/NaviTop.inc"-->

<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="��������{���ł�"></a></span>
  
  <!-- �R���e���cstart -->
  <div id="globalContents" class="feedback">
    <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
      <p class="home"><a href="<%=g_HTTP%>"><img src="<%=g_RelLink%>images/icon_home.gif" alt="HOME"></a></p>
      <ul id="path">
        <li class="now">���w���җl�����A���P�[�g</li>
      </ul>
    </div></div></div>

    <h1 class="title">���w���җl�����A���P�[�g</h1>
    <p>���̓x�̓T�E���h�n�E�X�������p���������A���ɂ��肪�Ƃ��������܂����B<br>���i�̂������A�z�B�A���͂��������i�̏�ԂȂǏ\�����������������܂����ł��傤���B</p>
    <p>���ǂ��T�E���h�n�E�X����ԑ�؂ɂ������Ă���܂��u�S�̂��������T�[�r�X�v�����b�g�[�ɁA���q�l�ɏ\�����������Ă���������T�[�r�X��񋟂���悤�w�͂������Ă���܂��B���C�t���̓_���������܂�����A�ǂ�ȏ����Ȏ��ł��������Ȃ����m�点���������B�Ɩ��p�I�[�f�B�I�A�y��A�Ɩ��̑����f�p�[�g�Ƃ��č���Ƃ����q�l�̂��v�]�ɉ\�Ȍ��肨�������Ă��������ƍl���Ă��܂��̂ŁA�����͂����肢�������܂��B</p>
    
    <%=wMsg%>
    
    <form name="fFeedBack" action="FeedBackStore.asp" method="post" onSubmit="return FeedBack_onSubmit(this)">
    
    <p>�󒍔ԍ��F<%=OrderNo%><input type="hidden" name="OrderNo" value="<%=OrderNo%>"></p>
    
    <ol class="form">
    	<li>
        	<h2>1. �X�^�b�t�̉���</h2>
            <p>�Ⴆ�΁A�������₨�₢���킹�̍ۂ̑Ή��͂������ł����ł��傤���H</p>
            <ul>
            	<li><input type="radio" id="q1_5" name="q1" value="��ϖ���" <% if q1 = "��ϖ���" then %> checked <% end if %>><label for="q1_5">��ϖ���</label></li>
                <li><input type="radio" id="q1_4" name="q1" value="����" <% if q1 = "����" then %> checked <% end if %>><label for="q1_4">����</label></li>
                <li><input type="radio" id="q1_3" name="q1" value="����" <% if q1 = "����" then %> checked <% end if %>><label for="q1_3">����</label></li>
                <li><input type="radio" id="q1_2" name="q1" value="�s��" <% if q1 = "�s��" then %> checked <% end if %>><label for="q1_2">�s��</label></li>
                <li><input type="radio" id="q1_1" name="q1" value="��ϕs��" <% if q1 = "��ϕs��" then %> checked <% end if %>><label for="q1_1">��ϕs��</label></li>
            </ul>
            <p>���q�l�ւ̉��Β��A���ɗD�G�Ǝv���܂����]�ƈ�������܂����炲�L�����������B</p>
            <ul>
            	<li>���O�F<input name="q1Name" type="text" size="20" value="<%=q1Name%>"></li>
                <li>�����F<input name="q1Department" type="text" size="20" value="<%=q1Department%>"></li>
            </ul>
            <p>���ӌ����ǂ���(500�����܂�)</p>
            <textarea name="q1Comment" rows="3"><%=q1Comment%></textarea>
        </li>
        <li>
        	<h2>2. �z�[���y�[�W�̓��e</h2>
            <p>�Ⴆ�΁A���i�̌����̂��₷����A���̏[���x�A�^��_�������ɉ����ł����ł��傤���H</p>
            <ul>
            	<li><input type="radio" id="q2_5" name="q2" value="��ώg���₷��" <% if q2 = "��ώg���₷��" then %> checked <% end if %>><label for="q2_5">��ώg���₷��</label></li>
                <li><input type="radio" id="q2_4" name="q2" value="�g���₷��" <% if q2 = "�g���₷��" then %> checked <% end if %>><label for="q2_4">�g���₷��</label></li>
                <li><input type="radio" id="q2_3" name="q2" value="����" <% if q2 = "����" then %> checked <% end if %>><label for="q2_3">����</label></li>
                <li><input type="radio" id="q2_2" name="q2" value="�g���ɂ���" <% if q2 = "�g���ɂ���" then %> checked <% end if %>><label for="q2_2">�g���ɂ���</label></li>
                <li><input type="radio" id="q2_1" name="q2" value="��ώg���ɂ���" <% if q2 = "��ώg���ɂ���" then %> checked <% end if %>><label for="q2_1">��ώg���ɂ���</label></li>
            </ul>
            <p>���ӌ����ǂ���(500�����܂�)</p>
            <textarea name="q2Comment" rows="3"><%=q2Comment%></textarea>
        </li>
        <li>
        	<h2>3. �w�ǂ��Ă���G��(�����񓚉�)</h2>
            <p>���L�̂����A�悭�ǂނ��̂������ł����\�ł��̂Ń`�F�b�N���Ă��������B<br>�Y���G�����Ȃ��ꍇ�A���̑��̗��ɂ��L�����������B</p>
            <ul>
            	<li><input type="checkbox" id="q4_1" name="q4" value="�T�E���h�����R�[�f�B���O�E�}�K�W��" <% if InStr(q4,"�T�E���h&amp;���R�[�f�B���O�E�}�K�W��") <> "0" then %> checked <% end if %>><label for="q4_1">�T�E���h&amp;���R�[�f�B���O�E�}�K�W��</label></li>
                <li><input type="checkbox" id="q4_2" name="q4" value="�M�^�[�E�}�K�W��" <% if InStr(q4,"�M�^�[�E�}�K�W��") <> "0" then %> checked <% end if %>><label for="q4_2">�M�^�[�E�}�K�W��</label></li>
                <li><input type="checkbox" id="q4_3" name="q4" value="�x�[�X�E�}�K�W��" <% if InStr(q4,"�x�[�X�E�}�K�W��") <> "0" then %> checked <% end if %>><label for="q4_3">�x�[�X�E�}�K�W��</label></li>
                <li><input type="checkbox" id="q4_4" name="q4" value="���Y�����h�����E�}�K�W��" <% if InStr(q4,"���Y�����h�����E�}�K�W��") <> "0" then %> checked <% end if %>><label for="q4_4">���Y�����h�����E�}�K�W��</label></li>
                <li><input type="checkbox" id="q4_5" name="q4" value="�Q�b�J��" <% if InStr(q4,"�Q�b�J��") <> "0" then %> checked <% end if %>><label for="q4_5">�Q�b�J��</label></li>
                <li><input type="checkbox" id="q4_6" name="q4" value="�A�R�[�X�e�B�b�N�E�M�^�[�E�}�K�W��" <% if InStr(q4,"�A�R�[�X�e�B�b�N�E�M�^�[�E�}�K�W��") <> "0" then %> checked <% end if %>><label for="q4_6">�A�R�[�X�e�B�b�N�E�M�^�[�E�}�K�W��</label></li>
                <li><input type="checkbox" id="q4_7" name="q4" value="GROOVE" <% if InStr(q4,"GROOVE") <> "0" then %> checked <% end if %>><label for="q4_7">GROOVE</label></li>
                <li><input type="checkbox" id="q4_8" name="q4" value="�����O�M�^�[" <% if InStr(q4,"�����O�M�^�[") <> "0" then %> checked <% end if %>><label for="q4_8">�����O�M�^�[</label></li>
                <li><input type="checkbox" id="q4_9" name="q4" value="DTM�}�K�W��" <% if InStr(q4,"DTM�}�K�W��") <> "0" then %> checked <% end if %>><label for="q4_9">DTM�}�K�W��</label></li>
                <li><input type="checkbox" id="q4_10" name="q4" value="�r�f�I�T����" <% if InStr(q4,"�r�f�I�T����") <> "0" then %> checked <% end if %>><label for="q4_10">�r�f�I�T����</label></li>
                <li><input type="checkbox" id="q4_11" name="q4" value="�r�f�I��" <% if InStr(q4,"�r�f�I��") <> "0" then %> checked <% end if %>><label for="q4_11">�r�f�I��</label></li>
                <li><input type="checkbox" id="q4_12" name="q4" value="�J���I�P�t�@��" <% if InStr(q4,"�J���I�P�t�@��") <> "0" then %> checked <% end if %>><label for="q4_12">�J���I�P�t�@��</label></li>
                <li><input type="checkbox" id="q4_13" name="q4" value="�̂̎蒟" <% if InStr(q4,"�̂̎蒟") <> "0" then %> checked <% end if %>><label for="q4_13">�̂̎蒟</label></li>
             </ul>
             <p>���̑�<input name="q4Other" type="text" size="50" value="<%=q4Other%>"></p>
        </li>
        <li>
        	<h2>4. �G���L���̓��e</h2>
            <p>�T�E���h�n�E�X�̎G���L���������ɂȂ������q���܂ɂ����₢�����܂��B<br>�Ⴆ�΁A���i�ɂ��ċ�������������A�w���̍ۂɎQ�l�ɂȂ���e�ł����ł��傤���H</p>
            <ul>
            	<li><input type="radio" id="q3_5" name="q3" value="��ϖ���"  <% if q3 = "��ϖ���" then %> checked <% end if %>><label for="q3_5">��ϖ���</label></li>
                <li><input type="radio" id="q3_4" name="q3" value="����" <% if q3 = "����" then %> checked <% end if %>><label for="q3_4">����</label></li>
                <li><input type="radio" id="q3_3" name="q3" value="����" <% if q3 = "����" then %> checked <% end if %>><label for="q3_3">����</label></li>
                <li><input type="radio" id="q3_2" name="q3" value="�s��" <% if q3 = "�s��" then %> checked <% end if %>><label for="q3_2">�s��</label></li>
                <li><input type="radio" id="q3_1" name="q3" value="��ϕs��" <% if q3 = "��ϕs��" then %> checked <% end if %>><label for="q3_1">��ϕs��</label></li>
            </ul>
            <p>���ӌ����ǂ���(500�����܂�)</p>
            <textarea name="q3Comment" rows="3"><%=q3Comment%></textarea>
        </li>
        <li>
        	<h2>5. �J�^���O(�z�b�g���j���[�E�z�b�g�X�^�b�t)�̓��e</h2>
            <p>�Ⴆ�΁A���i�X�y�b�N�̎ʐ^�A�X�y�b�N�̌��₷����A���e�̏[���x�͂������ł��傤���H</p>
            <ul>
            	<li><input type="radio" id="q5_5" name="q5" value="��ϖ���" <% if q5 = "��ϖ���" then %> checked <% end if %>><label for="q5_5">��ϖ���</label></li>
                <li><input type="radio" id="q5_4" name="q5" value="����" <% if q5 = "����" then %> checked <% end if %>><label for="q5_4">����</label></li>
                <li><input type="radio" id="q5_3" name="q5" value="����" <% if q5 = "����" then %> checked <% end if %>><label for="q5_3">����</label></li>
                <li><input type="radio" id="q5_2" name="q5" value="�s��" <% if q5 = "�s��" then %> checked <% end if %>><label for="q5_2">�s��</label></li>
                <li><input type="radio" id="q5_1" name="q5" value="��ϕs��" <% if q5 = "��ϕs��" then %> checked <% end if %>><label for="q5_1">��ϕs��</label></li>
            </ul>
            <p>���ӌ����ǂ���(500�����܂�)</p>
            <textarea name="q5Comment" rows="3"><%=q5Comment%></textarea>
        </li>
        <li>
        	<h2>6. �T�E���h�n�E�X�����I�т������������R</h2>
            <ul>
            	<li><input type="radio" id="q6_5" name="q6" value="�O�ɗ��p�����Ƃ��̈�ۂ��ǂ�����" <% if q6 = "�O�ɗ��p�����Ƃ��̈�ۂ��ǂ�����" then %> checked <% end if %>><label for="q6_5">�O�ɗ��p�����Ƃ��̈�ۂ��ǂ�����</label></li>
                <li><input type="radio" id="q6_4" name="q6" value="�l�ɂ����߂���" <% if q6 = "�l�ɂ����߂���" then %> checked <% end if %>><label for="q6_4">�l�ɂ����߂���</label></li>
                <li><input type="radio" id="q6_3" name="q6" value="�G���̍L��������" <% if q6 = "�G���̍L��������" then %> checked <% end if %>><label for="q6_3">�G���̍L��������</label></li>
                <li><input type="radio" id="q6_2" name="q6" value="�C���^�[�l�b�g��������" <% if q6 = "�C���^�[�l�b�g��������" then %> checked <% end if %>><label for="q6_2">�C���^�[�l�b�g��������</label></li>
            </ul>
            <p>���ӌ����ǂ���(500�����܂�)</p>
            <textarea name="q6Other" rows="3"><%=q6Other%></textarea>
        </li>
        <li>
        	<h2>7. ���m�荇���̕��ɂ����p���������߂��������܂����H</h2>
            <ul>
            	<li><input type="radio" id="q7_5" name="q7" value="�͂�" <% if q7 = "�͂�" then %> checked <% end if %>><label for="q7_5">�͂�</label></li>
                <li><input type="radio" id="q7_4" name="q7" value="������" <% if q7 = "������" then %> checked <% end if %>><label for="q7_4">������</label></li>
            </ul>
            <p>���ӌ����ǂ���(500�����܂�)</p>
            <textarea name="q7Comment" rows="3"><%=q7Comment%></textarea>
        </li>
    </ol>
    
    <input type="hidden" name="q8" value="">
	<input type="hidden" name="q9" value="">
    
    <% if wCanWriteFl = "Y" then %>
    <p>��낵����Α��M�{�^���������Ă��������B</p>
    <p class="btnBox"><input type="submit" value="���M" class="opover"></p>
	<% end if %>

	</form>
    
</div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>
