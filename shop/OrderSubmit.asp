<%@ LANGUAGE="VBScript" %>
<%
'ネットハウスねっとハウスネットはうす
'サウンドハウス
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
'	オーダー登録・送信処理
'
'------------------------------------------------------------------------
'	
'		届先が新規に指定されたときは顧客情報に追加
'		離島の場合　送料=送料*離島掛率
'		仮受注情報をWeb受注情報へコピーし、仮受注情報を削除。
'		カードオーダーの場合は与信確認後上記を行う。
'		オーダー受付メールの送信。（顧客 & ショップ)
'
'------------------------------------------------------------------------
'更新履歴
'2004/12/20 注文メール文章変更
'2004/12/20 カード有効期限に/がないデータの対策
'2004/12/21 Thanksページへの戻りをHTTPSに変更
'2004/12/27 注文確認時(OrderConfirm)別Windowから商品追加されて送信ボタンを押された時の対処
'2005/04/05 カード情報を受注データから取り出すように変更
'2005/05/13 OBJ_NewMail.BodyPart.Charset = "iso-2022-jp"をセット
'2005/06/20 オリコローン追加
'2005/08/24 在庫情報表示を共通関数使用に変更し、受注数量<引当可能数量の場合は顧客にも在庫数を表示
'2005/08/31 自動返信メールから単価を抜く
'2005/09/28 メールサブジェクトに支払方法を追加
'2005/09/29 メールサブジェクトにコンタクト管理用情報を追加
'2005/10/07 SH宛てメール送信元をSHとする
'2006/06/22 Thanks.asp呼び出しパラメータにURLEncode追加
'2006/07/05 BlueGateオーソリから呼び出し時は、受注番号が渡される。
'2006/10/24 eContextからの呼び出し時は、振り込み票URL、決済選択用URLが渡される｡(コンビニ支払)
'2006/11/27 受注金額にコンビニ支払い手数料込みに変更
'2006/11/30 コンビニ支払いメール文章変更
'2007/01/11 コンビニ支払い時、eContext支払方法URL・eContext振込票URLをWeb受注に登録
'2007/01/12 「コンビニ支払」を「コンビニ/郵便局支払」に表示変更
'2007/01/30 メールヘッダ文章変更
'2007/02/28 受注形態を追加
'2007/04/20 色規格別在庫の引当可能数量を更新、商品の完売日をセット
'2007/05/09 代引き、クレジットのときは、金額をメールに表示
'2008/04/14 リベート機能追加、カード情報取り出し部分削除（顧客過不足金更新時マイナスはエラー：不正顧客対策）
'2008/04/21 無条件に届け先をセット
'2008/05/07 注文者情報をWeb受注にセット
'2008/05/14 HTTPSチェック対応
'2008/05/23 入力データチェック強化（LEFT, Numeric, EOF他)
'2008/09/16 コンビニ支払表示をコンビニ支払", "ネットバンキング・ゆうちょ・コンビニ払いに変更
'2008/12/12 Email、顧客には在庫数量非通知に。Shopはセット品以外在庫数量通知 + 文言修正
'2009/04/30 エラー時にerror.aspへ移動
'2009/06/17 廃番で色規格があるとき、最初の色の引当可能在庫=0で完売日がセットされる問題を修正
'2009/12/07 an「在庫稀少」を「在庫僅少」に変更
'2009/12/17 hn レコメンド用変更（商品購買ログ出力）→コメントアウト
'2010/03/04 an レコメンド用変更（商品購買ログ出力）有効化
'2010/05/07 an 納期予定が「XX/XX頃予定」の場合は、"程かかります"をつけないように修正
'2010/08/11 an 自動応答メール送信時にShopBCCにBCCするように修正
'2010/12/20 hn 見積もりの時も個数限定単価の場合は個数限定受注済数量を更新する
'2011/01/28 GV(ay) 届先登録処理の削除
'2011/04/14 hn SessionID関連変更
'2011/06/01 if-web 自動送信メールの運送会社表示部分を削除
'2011/06/29 an #867 佐川、ヤマトの場合は時間指定を各社に応じて読み替え
'2011/08/01 an #1087 Error.aspログ出力対応
'2011/09/09 an #1123 自動返信メール修正→受注数より在庫数が少ないときは「一部在庫がありません」と表示
'2012/01/10 an レコメンド商品購買ログ出力停止
'2012/01/23 hn 受注明細にデータがない時はエラーとする
'2012/08/15 nt セット品時の配信メール内容不正（在庫状況）を修正
'2012/09/25 nt 領収証宛先・但し書きを変更
'2013/07/30 GV #1618 アフィリエイト重複送信対応
'
'========================================================================

On Error Resume Next
Response.Expires = -1			' Do not cache
Response.buffer = true

Dim userID
Dim userName
Dim msg

Dim customer_email
Dim customer_no

Dim OrderNo

Dim eConF
Dim eConK

Dim w_order_no
Dim w_body_hd
Dim w_body_dt1
Dim w_body_dt2
Dim w_body_tl

Dim w_comp_ryakushou
Dim w_order_estimate
Dim w_todokesaki_no
Dim w_payment_method
Dim w_loan_company
Dim w_product_am
Dim w_holiday_fl
Dim wKabusokuAM

Dim wSalesTaxRate
Dim wPrice
Dim wProdTotalAm

Dim wItemChar1
Dim wItemChar2
Dim wItemNum1
Dim wItemNum2
Dim wItemDate1
Dim wItemDate2

Dim Connection
Dim RS_order_header
Dim RS_order_detail
Dim RS_web_order_header
Dim RS_web_order_detail
Dim RS_web_customer
Dim RS_cntl
Dim RS_customer
Dim RS_company
Dim RS_prod
Dim RS_set
Dim RS_calender

Dim wSQL
Dim w_html
Dim w_msg
Dim wErrDesc   '2011/08/01 an add

'=======================================================================

'---- UserID 取り出し
userID = Session("userID")
userName = Session("userName")

'---- セッション切れチェック
if userID = ""then
	Response.Redirect g_HTTP
end if

Session("msg") = ""
w_msg = ""

OrderNo = ReplaceInput(Request("OrderNo"))
eConf = ReplaceInput(Trim(Request("eConf")))
eConK = ReplaceInput(Trim(Request("eConK")))

'---- Execute main
call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	Connection.RollbackTrans
	wErrDesc = "OrderSubmit.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

call close_db()

if Err.Description <> "" then	
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

'---- エラーが無いときはありがとうページ、エラーがあれば注文内容入力ページへ
if w_msg = "" then
	Session(g_cookie_name) = ""		'注文商品Cookieをクリア
	Session("OrderAtOnce") = "1"	'2013/07/30 GV #1618 add
	Response.Redirect "Thanks.asp?order_no=" & w_order_no & "&product_am=" & w_product_am & "&order_estimate=" & Server.URLEncode(w_order_estimate) & "&payment_method=" & Server.URLEncode(w_payment_method) & "&loan_company=" & Server.URLEncode(w_loan_company)
else
	Session("msg") = "<font color='#ff0000'>" & w_msg & "</font>"
	Response.Redirect "OrderInfoEnter.asp"
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

'---- トランザクション開始
Connection.BeginTrans

'---- 消費税率取出し
call getCntlMst("共通","消費税率","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)			'消費税率
wSalesTaxRate = Clng(wItemNum1)

'---- 会社情報の取り出し
call get_company()

'---- 受注情報の登録
call insert_web_order_header()
if w_msg <> "" then
	Connection.RollbackTrans
	exit function			'if error exit
end if

call insert_web_order_detail()
if w_msg <> "" then
	Connection.RollbackTrans
	exit function			'if error exit
end if

'---- 仮受注情報削除
call delete_web_order()

'---- トランザクション終了
Connection.CommitTrans

'---- メール送信
call send_order_mail()

End Function

'========================================================================
'
'	Function	Web受注の登録
'
'========================================================================
'
Function insert_web_order_header()
Dim i

Dim vItemChar1
Dim vItemChar2
Dim vItemNum1
Dim vItemNum2
Dim vItemDate1
Dim vItemDate2

'---- 仮受注の取り出し
wSQL = ""
wSQL = wSQL & "SELECT a.*"
wSQL = wSQL & "  FROM 仮受注 a"
wSQL = wSQL & " WHERE SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod
	  
Set RS_order_header = Server.CreateObject("ADODB.Recordset")
RS_order_header.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RS_order_header.EOF = true then
	w_msg = "NoData"
	exit function
end if

'2011/01/28 GV Mod Start
''---- 顧客届先情報登録
'if (Trim(RS_order_header("届先住所連番")) = 0) then
'	w_todokesaki_no = insert_todokesaki()
'else
'	w_todokesaki_no = Trim(RS_order_header("届先住所連番"))
'end if

w_todokesaki_no = Trim(RS_order_header("届先住所連番"))
'2011/01/28 GV Mod End

'---- insert 受注
wSQL = ""
wSQL = wSQL & "SELECT *"
wSQL = wSQL & "  FROM Web受注"
wSQL = wSQL & " WHERE 1 = 2"
 
Set RS_web_order_header = Server.CreateObject("ADODB.Recordset")
RS_web_order_header.Open wSQL, Connection, adOpenStatic, adLockOptimistic

'---- 受注/見積　セーブ
if RS_order_header("見積フラグ") = "Y" then
	w_order_estimate = "お見積"
else
	w_order_estimate = "ご注文"
end if

'---- 受注番号取り出し(BlueGate呼び出し時は受注番号は採番されている)
if OrderNo = "" then
	w_order_no = CLng(get_cntl_no("共通","番号","Web受注"))
else
	if isNumeric(OrderNo) = false then
		w_msg = w_msg & "受注番号エラー"
		exit function
	end if
	w_order_no = CLng(OrderNo)
end if

'---- 受注作成
RS_web_order_header.AddNew

RS_web_order_header("受注番号") = w_order_no

For i=0 to RS_web_order_header.Fields.Count - 1
	if RS_order_header(i).Name <> "SessionID" then
		if isnull(RS_order_header(RS_order_header(i).Name)) = false then
			RS_web_order_header(RS_order_header(i).Name) = RS_order_header(RS_order_header(i).Name)
		end if
	end if
Next

'---- 佐川、ヤマトで時間指定ありの場合は時間指定を各運送会社に対応した時間に読み替え   2011/06/29 an add s
if Trim(RS_order_header("時間指定")) <> "" then
	if  RS_order_header("運送会社コード") = "1" OR RS_order_header("運送会社コード") = "2" then

		call getCntlMst("受注","時間指定読み替え",RS_order_header("時間指定"), vItemChar1, vItemChar2, vItemNum1, vItemNum2, vItemDate1, vItemDate2)

		'---- 佐川
		if RS_order_header("運送会社コード") = "1" then
			RS_web_order_header("時間指定") = vItemChar1
		end if
		'---- ヤマト
		if RS_order_header("運送会社コード") = "2" then
			RS_web_order_header("時間指定") = vItemChar2
		end if
	end if
end if                                                                                '2011/06/29 an add e
 
'---- 届先情報のセット
RS_web_order_header("届先住所連番") = w_todokesaki_no
CAll SetTodokesaki(w_todokesaki_no)

'---- 注文者情報のセット
CAll SetChuumonsha()

'---- コンビニ支払URL情報登録　'2007/01/11
if RS_web_order_header("支払方法") = "コンビニ支払" then
	RS_web_order_header("eContext支払方法URL") = eConK
	RS_web_order_header("eContext振込票URL") = eConF
end if

RS_web_order_header("受注形態") = "インターネット"

RS_web_order_header("最終更新日") = now()
RS_web_order_header("入力日") = now()

RS_web_order_header.update

'---- 顧客振込名義人登録
if (Trim(RS_web_order_header("振込名義人")) <> "") then
	call update_furikomimeiginin()
end if

'---- 支払方法、他 セーブ
w_payment_method = RS_web_order_header("支払方法")
w_product_am = RS_web_order_header("商品合計金額")
if w_payment_method = "ローン" then
	w_loan_company = RS_web_order_header("ローン会社")
end if

wKabusokuAm = RS_web_order_header("過不足相殺金額")

if Trim(RS_web_order_header("リベート使用フラグ")) = "Y" then
		call updateKabusokuAm()
end if

'---- リベート使用で0円になったとき
if RS_web_order_header("受注合計金額") = 0 then
	RS_web_order_header("支払方法") = "現金"
	RS_web_order_header.update
end if

'---- メールヘッダ、トレーラの編集
call edit_mail_ht()

RS_web_order_header.close

End function

'========================================================================
'
'	Function	Web受注明細の登録
'
'========================================================================
'
Function insert_web_order_detail()
Dim i
Dim vTotalAm

'---- 仮受注明細の取り出し
wSQL = ""
wSQL = wSQL & "SELECT a.*"
wSQL = wSQL & "     , b.セット商品フラグ"
wSQL = wSQL & "     , b.メーカー直送取寄区分"
wSQL = wSQL & "     , b.希少数量"
wSQL = wSQL & "     , b.取扱中止日"
wSQL = wSQL & "     , b.廃番日"
wSQL = wSQL & "     , b.商品名"
wSQL = wSQL & "     , c.引当可能数量"
wSQL = wSQL & "     , c.引当可能入荷予定日"
wSQL = wSQL & "  FROM 仮受注明細 a WITH (NOLOCK)"
wSQL = wSQL & "     , Web商品 b WITH (NOLOCK)"
wSQL = wSQL & "     , Web色規格別在庫 c WITH (NOLOCK)"
wSQL = wSQL & " WHERE b.メーカーコード = a.メーカーコード"
wSQL = wSQL & "   AND b.商品コード = a.商品コード"
wSQL = wSQL & "   AND c.メーカーコード = a.メーカーコード"
wSQL = wSQL & "   AND c.商品コード = a.商品コード"
wSQL = wSQL & "   AND c.色 = a.色"
wSQL = wSQL & "   AND c.規格 = a.規格"
wSQL = wSQL & "   AND SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod
wSQL = wSQL & " ORDER BY 受注明細番号"
	  
Set RS_order_detail = Server.CreateObject("ADODB.Recordset")
RS_order_detail.Open wSQL, Connection, adOpenStatic

'---- 受注明細がなければエラー	2012/01/23 hn add
if RS_order_detail.EOF = true then
	w_msg = w_msg & "ご注文がありません。"
	exit function
end if

'---- insert 受注明細
wSQL = ""
wSQL = wSQL & "SELECT *"
wSQL = wSQL & "  FROM Web受注明細"
wSQL = wSQL & " WHERE 1 = 2"
 
Set RS_web_order_detail = Server.CreateObject("ADODB.Recordset")
RS_web_order_detail.Open wSQL, Connection, adOpenStatic, adLockOptimistic

vTotalAm = 0
'---- 受注明細作成
Do while RS_order_detail.EOF = false

	RS_web_order_detail.AddNew

	RS_web_order_detail("受注番号") = w_order_no

	For i=0 to RS_order_detail.Fields.Count - 1 - 8		'-1 -仮受注明細以外の項目数(上記SQL)
		if RS_order_detail(i).Name <> "SessionID" then
			RS_web_order_detail(RS_order_detail(i).Name) = RS_order_detail(RS_order_detail(i).Name)
		end if
	Next

	RS_web_order_detail.update

	'---- 色規格別在庫　引当可能数量　更新	'2007/04/20
	call updateInventory()
	if w_msg <> "" then
		exit function
	end if

	'---- B品、廃番品の完売かどうかのチェック	'2009/06/17
	if RS_order_detail("B品フラグ") = "Y"  OR isNull(RS_order_detail("廃番日")) = false then
		call updateKanbaibi()		'完売日をセット
	end if

	'---- 個数限定単価商品なら商品マスタ更新
'	if (RS_order_detail("個数限定単価フラグ") = "Y") AND (w_order_estimate = "ご注文") AND ((w_payment_method = "クレジットカード") OR (w_payment_method = "代引き")) then  '2010/12/20 hn del
	if RS_order_detail("個数限定単価フラグ") = "Y" then	'2010/12/20 hn add
		call updateProduct()
	end if

	'---- メール明細行編集
	call edit_mail_dt()

	'---- 受注合計金額計算
	vTotalAm = vTotalAm + RS_order_detail("受注金額")

	'----- レコメンド商品購買ログ登録   2009/12/17 add hn 2010/03/04 an 有効化 2012/01/10 an 停止
	'call AddRecommendPurchaseLog(RS_order_detail("メーカーコード"), RS_order_detail("商品コード"))

	RS_order_detail.MoveNext
Loop

'---- 受注合計金額チェック
if vTotalAm <> w_product_am then
	w_msg = w_msg & "｢商品合計金額不一致｣ 注文内容を再度ご確認願います。"
end if

RS_web_order_detail.close
RS_order_detail.close

End function

'2011/01/28 GV Del Start
''========================================================================
''
''	Function	届先情報の登録
''
''========================================================================
''
'Function insert_todokesaki()
'Dim i
'Dim v_Max_no
'
''---- 同一住所があるかどうかチェック
'wSQL = ""
'wSQL = wSQL & "SELECT 住所連番"
'wSQL = wSQL & "  FROM Web顧客住所 WITH (NOLOCK)"
'wSQL = wSQL & " WHERE 顧客番号 = " & userID
'wSQL = wSQL & "   AND 住所名称 = '" & Replace(RS_order_header("届先名前"),"'","''") & "'"
'wSQL = wSQL & "   AND 顧客郵便番号 = '" & RS_order_header("届先郵便番号") & "'"
'wSQL = wSQL & "   AND 顧客都道府県 = '" & RS_order_header("届先都道府県") & "'"
'wSQL = wSQL & "   AND 顧客住所 = '" & Replace(RS_order_header("届先住所"),"'","''") & "'"
'
'Set RS_customer = Server.CreateObject("ADODB.Recordset")
'RS_customer.Open wSQL, Connection, adOpenStatic, adLockOptimistic
'
'if RS_customer.EOF = false then				'同一住所あり
'	insert_todokesaki = RS_customer("住所連番")
'	Exit Function
'end if
'
''---- MAX住所連番の取り出し
'wSQL = ""
'wSQL = wSQL & "SELECT MAX(住所連番) AS MAX住所連番"
'wSQL = wSQL & "  FROM Web顧客住所 WITH (NOLOCK)"
'wSQL = wSQL & " WHERE 顧客番号 = " & userID
'	  
'Set RS_customer = Server.CreateObject("ADODB.Recordset")
'RS_customer.Open wSQL, Connection, adOpenStatic, adLockOptimistic
'
'v_max_no = RS_customer("MAX住所連番") + 1
'
''---- insert 顧客住所
'wSQL = ""
'wSQL = wSQL & "SELECT *"
'wSQL = wSQL & "  FROM Web顧客住所"
'wSQL = wSQL & " WHERE 1 = 2"
' 
'Set RS_customer = Server.CreateObject("ADODB.Recordset")
'RS_customer.Open wSQL, Connection, adOpenStatic, adLockOptimistic
'
'RS_customer.AddNew
'
'RS_customer("顧客番号") = UserID
'RS_customer("住所連番") = v_Max_no
'RS_customer("住所区分") = "届先"
'RS_customer("住所名称") = RS_order_header("届先名前")
'RS_customer("顧客郵便番号") = RS_order_header("届先郵便番号")
'RS_customer("顧客都道府県") = RS_order_header("届先都道府県")
'RS_customer("顧客住所") = RS_order_header("届先住所")
'RS_customer("勤務先フラグ") = "N"
'RS_customer("納品書送付可フラグ") = RS_order_header("届先納品書送付可フラグ")
'RS_customer("規定届先フラグ") = "N"
'RS_customer("最終更新日") = Now()
'RS_customer("最終更新者コード") = "Internet"
'
'RS_customer.update
'
''---- insert 顧客住所電話番号
'wSQL = ""
'wSQL = wSQL & "SELECT *"
'wSQL = wSQL & "  FROM Web顧客住所電話番号"
'wSQL = wSQL & " WHERE 1 = 2"
' 
'Set RS_customer = Server.CreateObject("ADODB.Recordset")
'RS_customer.Open wSQL, Connection, adOpenStatic, adLockOptimistic
'
'RS_customer.AddNew
'
'RS_customer("顧客番号") = UserID
'RS_customer("住所連番") = v_Max_no
'RS_customer("電話連番") = 1
'RS_customer("電話区分") = "電話"
'RS_customer("顧客電話番号") = RS_order_header("届先電話番号")
'RS_customer("検索用顧客電話番号") = cf_numeric_only(RS_order_header("届先電話番号"))
'RS_customer("最終更新日") = Now()
'RS_customer("最終更新者コード") = "Internet"
'
'RS_customer.update
'RS_customer.close
'
'insert_todokesaki = v_Max_no
'
'End function
'2011/01/28 GV Del End

'========================================================================
'
'	Function	届先情報のセット(Web受注へ)
'
'========================================================================
'
Function SetTodokesaki(pNo)
Dim RSv

'---- 
wSQL = ""
wSQL = wSQL & "SELECT a.*"
wSQL = wSQL & "     , b.顧客電話番号"
wSQL = wSQL & "  FROM Web顧客住所 a WITH (NOLOCK)"
wSQL = wSQL & "     , Web顧客住所電話番号 b WITH (NOLOCK)"
wSQL = wSQL & " WHERE b.顧客番号 = a.顧客番号"
wSQL = wSQL & "   AND b.住所連番 = a.住所連番"
wSQL = wSQL & "   AND b.電話連番 = 1"
wSQL = wSQL & "   AND a.顧客番号 = " & userID
wSQL = wSQL & "   AND a.住所連番 = " & pNo

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RSv.EOF = false then	
	RS_web_order_header("届先名前") = RSv("住所名称")
	RS_web_order_header("届先郵便番号") = RSv("顧客郵便番号")
	RS_web_order_header("届先都道府県") = RSv("顧客都道府県")
	RS_web_order_header("届先住所") = RSv("顧客住所") 
	RS_web_order_header("届先電話番号") = RSv("顧客電話番号")
else
	RS_web_order_header("届先名前") = ""
	RS_web_order_header("届先郵便番号") = ""
	RS_web_order_header("届先都道府県") = ""
	RS_web_order_header("届先住所") = ""
	RS_web_order_header("届先電話番号") = ""
end if

RSv.close

End function

'========================================================================
'
'	Function	注文者情報のセット(Web受注へ)
'
'========================================================================
'
Function SetChuumonsha()
Dim RSv

'---- 
wSQL = ""
wSQL = wSQL & "SELECT a.顧客名"
wSQL = wSQL & "     , b.顧客郵便番号"
wSQL = wSQL & "     , b.顧客都道府県"
wSQL = wSQL & "     , b.顧客住所"
wSQL = wSQL & "     , c.顧客電話番号"
wSQL = wSQL & "  FROM Web顧客 a WITH (NOLOCK)"
wSQL = wSQL & "     , Web顧客住所 b WITH (NOLOCK)"
wSQL = wSQL & "     , Web顧客住所電話番号 c WITH (NOLOCK)"
wSQL = wSQL & " WHERE b.顧客番号 = a.顧客番号"
wSQL = wSQL & "   AND c.顧客番号 = b.顧客番号"
wSQL = wSQL & "   AND c.住所連番 = b.住所連番"
wSQL = wSQL & "   AND a.顧客番号 = " & userID
wSQL = wSQL & "   AND b.住所連番 = 1"
wSQL = wSQL & "   AND c.電話連番 = 1"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RSv.EOF = false then	
	RS_web_order_header("注文者名前") = RSv("顧客名")
	RS_web_order_header("注文者郵便番号") = RSv("顧客郵便番号")
	RS_web_order_header("注文者都道府県") = RSv("顧客都道府県")
	RS_web_order_header("注文者住所") = RSv("顧客住所") 
	RS_web_order_header("注文者電話番号") = RSv("顧客電話番号")
else
	RS_web_order_header("注文者名前") = ""
	RS_web_order_header("注文者郵便番号") = ""
	RS_web_order_header("注文者都道府県") = ""
	RS_web_order_header("注文者住所") = ""
	RS_web_order_header("注文者電話番号") = ""
end if

RSv.close

End function

'========================================================================
'
'	Function	振込名義人の更新
'
'========================================================================
'
Function update_furikomimeiginin()
Dim i

'---- 振込名義人の更新
wSQL = ""
wSQL = wSQL & "SELECT 振込名義人"
wSQL = wSQL & "       , 最終更新日"
wSQL = wSQL & "       , 最終更新者コード"
wSQL = wSQL & "    FROM Web顧客"
wSQL = wSQL & "   WHERE 顧客番号 = " & RS_order_header("顧客番号")

Set RS_customer = Server.CreateObject("ADODB.Recordset")
RS_customer.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RS_customer.EOF = false then	
	RS_customer("振込名義人") = RS_order_header("振込名義人")
	RS_customer("最終更新日") = Now()
	RS_customer("最終更新者コード") = "Internet"
end if

RS_customer.update
RS_customer.close

End function

'========================================================================
'
'	Function	入金過不足金額の更新
'
'========================================================================
'
Function updateKabusokuAm()
Dim i
Dim vCustKabusoku

'---- 入金過不足金額の更新
wSQL = ""
wSQL = wSQL & "SELECT 入金過不足金額"
wSQL = wSQL & "       , 最終更新日"
wSQL = wSQL & "       , 最終更新者コード"
wSQL = wSQL & "       , 最終更新処理名"
wSQL = wSQL & "    FROM Web顧客"
wSQL = wSQL & "   WHERE 顧客番号 = " & RS_order_header("顧客番号")

Set RS_customer = Server.CreateObject("ADODB.Recordset")
RS_customer.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RS_customer.EOF = false then	
	vCustKabusoku = RS_customer("入金過不足金額") - wKabusokuAm
	if vCustKabusoku < 0 then
		w_msg = w_msg & "リベート金額が不足しています。 注文内容を再度ご確認願います。<br>"
	else
		RS_customer("入金過不足金額") = RS_customer("入金過不足金額") - wKabusokuAm
		RS_customer("最終更新日") = Now()
		RS_customer("最終更新者コード") = "Internet"
		RS_customer("最終更新処理名") = "OrderSubmit.asp"
	end if
else
	w_msg = w_msg & "顧客情報がありません。<br>"
end if

RS_customer.update
RS_customer.close

End function

'========================================================================
'
'	Function	仮受注の削除
'
'========================================================================
'
Function delete_web_order()

'---- 仮受注の削除
RS_order_header.delete
RS_order_header.close

End function

'========================================================================
'
'	Function	レコメンド商品購買ログ	2009/12/17
'
'========================================================================
'
Function AddRecommendPurchaseLog(pMakerCd, pProductCd)

Dim RSv

'---- 
wSQL = ""
wSQL = wSQL & "SELECT *"
wSQL = wSQL & "  FROM レコメンド商品購買ログ"
wSQL = wSQL & " WHERE 1 = 2"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

'---- レコメンド商品購買ログ登録
RSv.AddNew

RSv("レコメンドユーザーID") = gSessionID				'2011/04/14 hn mod
RSv("メーカーコード") = pMakerCd
RSv("商品コード") = pProductCd
RSv("ユーザーエージェント") = Request.ServerVariables("HTTP_USER_AGENT")
RSv("アクセス日") = Now()

RSv.Update
RSv.close

End function

'========================================================================
'
'	Function	コントロールマスタから番号採番
'
'		parm: sub_sustem_cd, item_cd, item_sub_cd
'		return:	番号
'
'========================================================================
'
Function get_cntl_no(p_sub_system_cd, p_item_cd, p_item_sub_cd)

'---- コントロールマスタ取り出し
wSQL = ""
wSQL = wSQL & "SELECT item_num1"
wSQL = wSQL & "  FROM コントロールマスタ"
wSQL = wSQL & " WHERE sub_system_cd = '" & p_sub_system_cd & "'"
wSQL = wSQL & "   AND item_cd = '" & p_item_cd & "'"
wSQL = wSQL & "   AND item_sub_cd = '" & p_item_sub_cd & "'"
	  
'@@@@@@response.write(wSQL)

Set RS_cntl = Server.CreateObject("ADODB.Recordset")
RS_cntl.Open wSQL, Connection, adOpenStatic, adLockOptimistic

RS_cntl("item_num1") = Clng(RS_cntl("item_num1")) + 1
get_cntl_no = RS_cntl("item_num1")

RS_cntl.update
RS_cntl.close

End function

'========================================================================
'
'	Function	色規格別在庫の引当可能数量を更新	2007/04/20
'
'========================================================================
'
Function updateInventory()

Dim RSv

wSQL = ""
wSQL = wSQL & "SELECT 引当可能数量"
wSQL = wSQL & "     , B品引当可能数量"
wSQL = wSQL & "  FROM Web色規格別在庫"
wSQL = wSQL & " WHERE メーカーコード = '" & RS_order_detail("メーカーコード")  & "'"
wSQL = wSQL & "   AND 商品コード = '" & RS_order_detail("商品コード")  & "'"
wSQL = wSQL & "   AND 色 = '" & RS_order_detail("色")  & "'"
wSQL = wSQL & "   AND 規格 = '" & RS_order_detail("規格")  & "'"
	  
'@@@@@@response.write(wSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RS_order_detail("B品フラグ") <> "Y" then
	if isNull(RS_order_detail("廃番日")) = false AND RSv("引当可能数量") < RS_order_detail("受注数量") then
		w_msg = w_msg & RS_order_detail("商品名") & "は、在庫が" & RSv("引当可能数量") & "個しかありません。　数量を変更してご注文ください。<br>"
		exit function
	else
		RSv("引当可能数量") = RSv("引当可能数量") - RS_order_detail("受注数量")
	end if

else
	if RSv("B品引当可能数量") >= RS_order_detail("受注数量") then
		RSv("B品引当可能数量") = RSv("B品引当可能数量") - RS_order_detail("受注数量")
	else
		w_msg = w_msg & RS_order_detail("商品名") & "は、在庫が" & RSv("B品引当可能数量") & "個しかありません。　数量を変更してご注文ください。<br>"
		exit function
	end if
end if

RSv.update
RSv.close

End function

'========================================================================
'
'	Function	商品の完売日をセット	2007/04/20
'
'========================================================================
'
Function updateKanbaibi()

Dim RSv

wSQL = ""
wSQL = wSQL & "SELECT SUM(引当可能数量) AS 引当可能数量"
wSQL = wSQL & "     , SUM(B品引当可能数量) AS B品引当可能数量"
wSQL = wSQL & "  FROM Web色規格別在庫"
wSQL = wSQL & " WHERE メーカーコード = '" & RS_order_detail("メーカーコード")  & "'"
wSQL = wSQL & "   AND 商品コード = '" & RS_order_detail("商品コード")  & "'"
	  
'@@@@@@response.write(wSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if (RS_order_detail("B品フラグ") = "Y" and RSv("B品引当可能数量") <= 0) OR (isNull(RS_order_detail("廃番日")) = false AND RSv("引当可能数量") <= 0) then

	RSv.close
	
	wSQL = ""
	wSQL = wSQL & "SELECT 完売日"
	wSQL = wSQL & "  FROM Web商品"
	wSQL = wSQL & " WHERE メーカーコード = '" & RS_order_detail("メーカーコード")  & "'"
	wSQL = wSQL & "   AND 商品コード = '" & RS_order_detail("商品コード")  & "'"
		  
	'@@@@@@response.write(wSQL)

	Set RSv = Server.CreateObject("ADODB.Recordset")
	RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

	RSv("完売日") = Now()

	RSv.update

end if

RSv.close

End function

'========================================================================
'
'	Function	商品マスタの個数限定受注済数量を更新
'
'========================================================================
'
Function updateProduct()

wSQL = ""
wSQL = wSQL & "SELECT 個数限定受注済数量"
wSQL = wSQL & "     , 個数限定数量"
wSQL = wSQL & "     , 特価商品フラグ"
wSQL = wSQL & "  FROM Web商品"
wSQL = wSQL & " WHERE メーカーコード = '" & RS_order_detail("メーカーコード")  & "'"
wSQL = wSQL & "   AND 商品コード = '" & RS_order_detail("商品コード")  & "'"
	  
'@@@@@@response.write(wSQL)

Set RS_prod = Server.CreateObject("ADODB.Recordset")
RS_prod.Open wSQL, Connection, adOpenStatic, adLockOptimistic

RS_prod("個数限定受注済数量") = RS_prod("個数限定受注済数量") + RS_order_detail("受注数量")
if RS_prod("個数限定受注済数量") >= RS_prod("個数限定数量") then
	RS_prod("特価商品フラグ") = ""
end if

RS_prod.update
RS_prod.close

End function

'========================================================================
'
'	Function	メールヘッダ、トレーラの編集
'
'========================================================================
'
Function edit_mail_ht()
Dim i
Dim v_temp
Dim vPaymentMethod

'---- 顧客データ取り出し
call get_customer()

'---- メールヘッダ
w_body_hd = "●　受付日時：" & FormatDateTime(RS_web_order_header("入力日"), 0) & "　" & w_order_estimate & "　" & w_order_no

w_body_hd = w_body_hd & vbNewLine & vbNewLine
w_body_hd = w_body_hd & "－－－－－　お客様　－－－－－" & vbNewLine
w_body_hd = w_body_hd & "名前　　　：　" & RS_customer("顧客名") & vbNewLine
w_body_hd = w_body_hd & "ふりがな ：　" & RS_customer("顧客フリガナ") & vbNewLine
w_body_hd = w_body_hd & "住所　　　：　〒" & RS_customer("顧客郵便番号") & "　" & RS_customer("顧客都道府県") & RS_customer("顧客住所") & vbNewLine
w_body_hd = w_body_hd & "電話番号：　" & RS_customer("顧客電話番号") & vbNewLine
w_body_hd = w_body_hd & "ＦＡＸ　　　：　" & RS_customer("FAX") & vbNewLine
w_body_hd = w_body_hd & "Ｅメール　：　" & RS_web_order_header("顧客E_mail") & vbNewLine
w_body_hd = w_body_hd & "顧客番号：　" & RS_web_order_header("顧客番号") & vbNewLine & vbNewLine
		
w_body_hd = w_body_hd & "－－－－－　お届け先　－－－－" & vbNewLine
if RS_web_order_header("届先区分") = "D" then

	if Trim(RS_web_order_header("届先住所連番")) <> "0" then
		call get_todokesaki(Trim(RS_web_order_header("届先住所連番")))

		w_body_hd = w_body_hd & "名前　　　：　" & RS_customer("住所名称") & vbNewLine
		w_body_hd = w_body_hd & "住所　　　：　〒" & RS_customer("顧客郵便番号") & "　" & RS_customer("顧客都道府県") & RS_customer("顧客住所") & vbNewLine
		w_body_hd = w_body_hd & "電話番号：　" & RS_customer("顧客電話番号") & vbNewLine
	else
		w_body_hd = w_body_hd & "名前　　　：　" & RS_web_order_header("届先名前") & vbNewLine
		w_body_hd = w_body_hd & "住所　　　：　〒" & RS_web_order_header("届先郵便番号") & "　" & RS_web_order_header("届先都道府県") & RS_web_order_header("届先住所") & vbNewLine
		w_body_hd = w_body_hd & "電話番号：　" & RS_web_order_header("届先電話番号") & vbNewLine
	end if

	if RS_web_order_header("届先納品書送付可フラグ") = "Y" then
		w_body_hd = w_body_hd & "納品書 　：　送付して良い" & vbNewLine & vbNewLine
	else
		w_body_hd = w_body_hd & "納品書 　：　送付しない" & vbNewLine & vbNewLine
	end if
else
	w_body_hd = w_body_hd & "＊＊　同上" & vbNewLine & vbNewLine
end if

w_body_hd = w_body_hd & "－－－－－　配送指定　－－－－" & vbNewLine

'2011/06/01 if-web del start
'if RS_web_order_header("運送会社コード") = "1" then
'	w_body_hd = w_body_hd & "運送会社　　　：　佐川急便" & vbNewLine
'end if
'if RS_web_order_header("運送会社コード") = "2" then
'	w_body_hd = w_body_hd & "運送会社　　　：　ヤマト運輸" & vbNewLine
'end if
'if RS_web_order_header("運送会社コード") = "3" then
'	w_body_hd = w_body_hd & "運送会社　　　：　福山通運" & vbNewLine
'end if
'2011/06/01 if-web del end

if Trim(RS_web_order_header("指定納期")) <> "" then
	w_body_hd = w_body_hd & "配送日指定 　：　" & RS_web_order_header("指定納期") & vbNewLine
end if

if Trim(RS_web_order_header("時間指定")) <> "" then
	w_body_hd = w_body_hd & "配送時間指定：　" & RS_web_order_header("時間指定") & vbNewLine
end if

if RS_web_order_header("営業所止めフラグ") = "Y" then
	w_body_hd = w_body_hd & "営業所止め" & vbNewLine
end if

if RS_web_order_header("一括出荷フラグ") = "Y" then
	w_body_hd = w_body_hd & "商品が全て揃ってから出荷を行う" & vbNewLine & vbNewLine
end if
if RS_web_order_header("一括出荷フラグ") = "N" then
	w_body_hd = w_body_hd & "在庫のある商品から出荷を行う" & vbNewLine & vbNewLine
end if

w_body_hd = w_body_hd & "備考　：" & vbNewLine & RS_web_order_header("見積備考") & vbNewLine

w_body_hd = w_body_hd & vbNewLine

'---- トレーラ
'---- 金額関連 2007/05/09
if w_payment_method = "クレジットカード" OR w_payment_method = "代引き" then
	w_body_tl = w_body_tl & "商品合計（税込み）：　商品合計金額未確定" & vbNewLine
	wPrice = Fix(RS_web_order_header("送料") * (100 + wSalesTaxRate) / 100)
	w_body_tl = w_body_tl & "送料（税込み）：　" & FormatCurrency(wPrice,0) & vbNewLine

		if w_payment_method = "代引き" then
			wPrice = Fix(RS_web_order_header("代引手数料") * (100 + wSalesTaxRate) / 100)
			w_body_tl = w_body_tl & "代引手数料（税込み）：　" & FormatCurrency(wPrice,0) & vbNewLine
		end if
	
	if RS_web_order_header("リベート使用フラグ") = "Y" then
		w_body_tl = w_body_tl & "クレジット/過不足金：　" & FormatCurrency(RS_web_order_header("過不足相殺金額") * (-1) ,0) & vbNewLine
	end if

	w_body_tl = w_body_tl & "合計金額（税込み）：　" & FormatCurrency(RS_web_order_header("受注合計金額"),0) & vbNewLine & vbNewLine

end if

'---- 支払方法
vPaymentMethod = w_payment_method

if RS_web_order_header("受注合計金額") = 0 then
	w_body_tl = w_body_tl	& "支払方法　：　お支払い不要" & vbNewLine 
else
	w_body_tl = w_body_tl	& "支払方法　：　" & Replace(vPaymentMethod, "コンビニ支払", "ネットバンキング・ゆうちょ・コンビニ払い") & vbNewLine 
end if

if RS_web_order_header("支払方法") = "銀行振込" then
	if RS_web_order_header("振込名義人") <> "" then
		w_body_tl = w_body_tl & "振込名義人：　" & RS_web_order_header("振込名義人") & vbNewLine
	end if
end if

if RS_web_order_header("支払方法") = "代引き" then
	w_body_tl = w_body_tl & vbNewLine
end if

if RS_web_order_header("支払方法") = "ローン" then
	if RS_web_order_header("ローン頭金ありフラグ") = "Y" then
		w_body_tl = w_body_tl & "頭金あり　　　：" & FormatCurrency(RS_web_order_header("ローン頭金"),0) & vbNewLine
	else
		w_body_tl = w_body_tl & "頭金なし" & vbNewLine
	end if

	if Trim(RS_web_order_header("オンラインローン申込フラグ")) <> "Y" then
		if Trim(RS_web_order_header("希望ローン回数")) <> "" then
			w_body_tl = w_body_tl & "希望ローン回数：　" & RS_web_order_header("希望ローン回数") & vbNewLine
		end if
		if Trim(RS_web_order_header("ローン金額")) <> "" then
			v_temp = RS_web_order_header("ローン金額")
			w_body_tl = w_body_tl & "ローン金額　：　" & FormatCurrency(Ccur(v_temp)) & vbNewLine
		end if
	else
		w_body_tl = w_body_tl & "オンラインローン申込" & vbNewLine
		if RS_web_order_header("ローン会社") = "セントラル" then
			w_body_tl = w_body_tl & "（セディナ利用）" & vbNewLine
		end if
		if RS_web_order_header("ローン会社") = "オリコ" then
			w_body_tl = w_body_tl & "（オリコ利用）" & vbNewLine
		end if
	end if
	w_body_tl = w_body_tl & vbNewLine
end if

'---- リベートメッセージ
if RS_web_order_header("リベート使用フラグ") = "Y" then
	w_body_tl = w_body_tl & vbNewLine & "クレジット/過不足金は、このご注文・見積りのみに充当されます。" & vbNewLine & "キャンセルしてご利用にならない場合は弊社営業宛までご連絡ください。" & vbNewLine & vbNewLine
end if

'---- 領収書
if RS_web_order_header("領収書発行フラグ") = "Y" then
	w_body_tl = w_body_tl & "領収書必要"
	if RS_web_order_header("領収書宛先") <> "" then
		'2012/09/25 nt mod
		w_body_tl = w_body_tl & "　　領収書宛名：" & RS_web_order_header("領収書宛先") & " 様"
		w_body_tl = w_body_tl & "　　但し書き：" & RS_web_order_header("領収書但し書き")
		'w_body_tl = w_body_tl & "　　領収証宛先：" & RS_web_order_header("領収書宛先") & " 様"
		'w_body_tl = w_body_tl & "　　領収証但し書き：" & RS_web_order_header("領収書但し書き")
	end if
end if

customer_email = RS_web_order_header("顧客E_mail")	'顧客メールアドレスセーブ
customer_no = RS_web_order_header("顧客番号")	'顧客番号セーブ

RS_customer.close

End function

'========================================================================
'
'	Function	メール明細行編集
'
'========================================================================
'
Function edit_mail_dt()

Dim v_body_dt
Dim v_inv1
Dim v_inv2
Dim v_product_nm
Dim vInventoryCd
Dim vProdTermFl

v_product_nm = RS_web_order_detail("商品名")
if Trim(RS_web_order_detail("色")) <> "" then
	v_product_nm = v_product_nm & "/" & RS_web_order_detail("色")
end if
if Trim(RS_web_order_detail("規格")) <> "" then
	v_product_nm = v_product_nm & "/" & RS_web_order_detail("規格")
end if

if RS_web_order_detail("商品名") <> RS_web_order_detail("商品コード") then
		v_product_nm = v_product_nm & " (" & RS_web_order_detail("商品コード") & ")"
end if

if RS_web_order_detail("B品フラグ") = "Y" then
		v_product_nm = v_product_nm & " (B品）"
end if

v_body_dt = ""
v_body_dt = v_body_dt & "メーカー	：　" & RS_web_order_detail("メーカー名") & vbNewLine
v_body_dt = v_body_dt & "商品名 　：　" & v_product_nm & vbNewLine

wPrice = calcPrice(RS_web_order_detail("受注単価"), wSalesTaxRate)
wProdTotalAm = wProdTotalAm + (wPrice * RS_web_order_detail("受注数量"))

w_html = w_html & "    <td align='right' width='100'>"

'---- 単価、数量、金額 2007/05/09
if w_payment_method = "クレジットカード" OR w_payment_method = "代引き" then
	v_body_dt = v_body_dt & "単価(税込)：　" & FormatCurrency(wPrice,0) & vbNewLine
	v_body_dt = v_body_dt & "数量 　　：　" & RS_web_order_detail("受注数量") & vbNewLine
	v_body_dt = v_body_dt & "金額(税込)：　" & FormatCurrency(wPrice * RS_web_order_detail("受注数量"),0) & vbNewLine
else
	v_body_dt = v_body_dt & "数量 　　：　" & RS_web_order_detail("受注数量") & vbNewLine
end if

'---- 廃番チェック
if  (isNull(RS_order_detail("取扱中止日")) = true AND isNull(RS_order_detail("廃番日")) = true) _
 OR (isNull(RS_order_detail("廃番日")) = false AND RS_order_detail("引当可能数量") > 0) then
	vProdTermFl = "N"
else
	vProdTermFl = "Y"
end if

'---- 在庫情報セット（顧客用）
vInventoryCd = GetInventoryStatus(RS_web_order_detail("メーカーコード"),RS_web_order_detail("商品コード"),RS_web_order_detail("色"),RS_web_order_detail("規格"),RS_order_detail("引当可能数量"),RS_order_detail("希少数量"),RS_order_detail("セット商品フラグ"),RS_order_detail("メーカー直送取寄区分"),RS_order_detail("引当可能入荷予定日"),vProdTermFl)
'---- 在庫が無ければ、「在庫：無し」と追記。在庫ある場合は記載していない。
if vInventoryCd <> "在庫あり" AND vInventoryCd <> "在庫僅少" then
	v_body_dt = v_body_dt & "在庫 　　：　無し" & vbNewLine
end if
'---- 納期予定を表記
v_body_dt = v_body_dt & "納期予定 ：　"
	
'---- w_body_dt1（社内）, w_body_dt2（顧客）
if vInventoryCd <> "問合せ" AND vInventoryCd <> "特別注文" AND vInventoryCd <> "お取寄せ" AND vInventoryCd <> "取扱中止" then
	'---- "約Xヶ月"のようにおおよその納期が記載される場合は文言を修正。
	if vInventoryCd <> "在庫あり" AND vInventoryCd <> "在庫僅少" AND Right(Trim(vInventoryCd), 2) <> "予定" then  '2010/05/07 an changed
		w_body_dt1 = w_body_dt1 & v_body_dt & vInventoryCd & "程かかります。" & vbNewLine & vbNewLine       '2011/09/09 an mod s
		w_body_dt2 = w_body_dt2 & v_body_dt & vInventoryCd & "程かかります。" & vbNewLine & vbNewLine
	'---- "在庫あり"か"在庫僅少"か"X/X頃予定"
	else
		'---- "在庫あり"でも受注数より在庫数が少ないときは「一部在庫がありません」と表示
		'if vInventoryCd = "在庫あり" AND ( RS_web_order_detail("受注数量") > RS_order_detail("引当可能数量")) then	'2012/08/15 nt mod
		'---- セット品の場合、引当可能数量が正確でないため、メッセージ判別は除外（vInventoryCdを正とする）
		if vInventoryCd = "在庫あり" AND ( RS_web_order_detail("受注数量") > RS_order_detail("引当可能数量") AND (RS_order_detail("セット商品フラグ") <> "Y")) then
			w_body_dt1 = w_body_dt1 & v_body_dt & "一部在庫がありません"
			w_body_dt2 = w_body_dt2 & v_body_dt & "一部在庫がありません" & vbNewLine & vbNewLine
		else
			'---- 上記以外は従来通り表記
			w_body_dt1 = w_body_dt1 & v_body_dt & vInventoryCd
			w_body_dt2 = w_body_dt2 & v_body_dt & vInventoryCd & vbNewLine & vbNewLine
		end if
		
		'---- 社内は引当可能数量を記載
		if RS_order_detail("セット商品フラグ") <> "Y" then
			w_body_dt1 = w_body_dt1 & "(" & RS_order_detail("引当可能数量") & "個)" & vbNewLine & vbNewLine
		else
			w_body_dt1 = w_body_dt1 & vbNewLine & vbNewLine
		end if  
	end if
else
	'---- "問合せ","特別注文","お取寄せ","取扱中止"なら従来通り表記
	w_body_dt1 = w_body_dt1 & v_body_dt & vInventoryCd & vbNewLine & vbNewLine
	w_body_dt2 = w_body_dt2 & v_body_dt & vInventoryCd & vbNewLine & vbNewLine                             '2011/09/09 an mod e
end if

'if vInventoryCd <> "問合せ" AND vInventoryCd <> "特別注文" AND vInventoryCd <> "お取寄せ" AND vInventoryCd <> "取扱中止" then  '2011/09/09 an del s
	''---- "約Xヶ月"のようにおおよその納期が記載される場合は文言を修正。
	'if vInventoryCd <> "在庫あり" AND vInventoryCd <> "在庫僅少" AND Right(Trim(vInventoryCd), 2) <> "予定" then   '2010/05/07 an changed
		'w_body_dt1 = w_body_dt1 & v_body_dt & vInventoryCd & "程かかります。" & vbNewLine & vbNewLine
	'else
		''---- "在庫あり"か"在庫希少"か"X/X頃予定"なら従来通り表記
		'if RS_order_detail("セット商品フラグ") <> "Y" then
			'w_body_dt1 = w_body_dt1 & v_body_dt & vInventoryCd & "(" & RS_order_detail("引当可能数量") & "個)" & vbNewLine & vbNewLine
		'else
			'w_body_dt1 = w_body_dt1 & v_body_dt & vInventoryCd & vbNewLine & vbNewLine
		'end if
	'end if	
'else
	''---- "問合せ","特別注文","お取寄せ","取扱中止"なら従来通り表記
	'w_body_dt1 = w_body_dt1 & v_body_dt & vInventoryCd & vbNewLine & vbNewLine
'end if            '2011/09/09 an del e

End function

'========================================================================
'
'	Function	メール送信
'
'========================================================================
'
Function send_order_mail()

Dim v_from_mail
Dim v_BCC_mail    '2010/08/11 an add
Dim OBJ_NewMail
Dim v_subject

'Set OBJ_NewMail = CreateObject("CDONTS.NewMail")
Set OBJ_NewMail = Server.CreateObject("CDO.Message") 

'---- コントロールマスタからメールアドレス取り出し
if w_order_estimate = "お見積" then
	call getCntlMst("共通","送信先Email","Web見積通知", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)
else
	if w_payment_method = "クレジットカード" then
		call getCntlMst("共通","送信先Email","Webカード受注通知", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)
	else
		call getCntlMst("共通","送信先Email","Web受注通知", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)
	end if
end if

v_from_mail = wItemChar1

'---- 顧客への自動応答メールにBCCするメールアドレス取りだし
call getCntlMst("共通","送信先Email","ShopBCC", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)  '2010/08/11 an add
v_BCC_mail = wItemChar1  '2010/08/11 an add

w_body_dt1 = "－－－－－　商品明細　－－－－" & vbNewLine & w_body_dt1
w_body_dt2 = "－－－－－　商品明細　－－－－" & vbNewLine & w_body_dt2

'---- トレーラ編集時に未確定だった商品合計金額を置き換え
w_body_tl = Replace(w_body_tl, "商品合計金額未確定", FormatCurrency(wProdTotalAm,0))

'---- メール送信　ショップ
OBJ_NewMail.from = v_from_mail
OBJ_NewMail.to = v_from_mail
OBJ_NewMail.subject = w_order_estimate & w_order_no & "/" & w_payment_method & " ["  & customer_no & "/Web-Emax/Web受注-" & w_payment_method & "]"
OBJ_NewMail.TextBody = w_body_hd & w_body_dt1 & w_body_tl
OBJ_NewMail.MimeFormatted = False

OBJ_NewMail.Send

Set OBJ_NewMail = Nothing

'---- メール送信　顧客
'---- コントロールマスタからヘッダ、トレーラ付加情報取出し
call CheckHoliday()		'	休日用追加ヘッダが必要かどうか確認

call getCntlMst("Web","Email","ヘッダ", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)

if w_Holiday_fl = "Y" then
	w_body_hd = wItemChar1 & vbNewLine & wItemChar2 & vbNewLine & vbNewLine & w_body_hd
else
	w_body_hd = wItemChar1 & vbNewLine & vbNewLine & w_body_hd
end if

call getCntlMst("Web","Email","トレーラ", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)
w_body_tl = w_body_tl & vbNewLine & vbNewLine & wItemChar1

if w_order_estimate = "お見積" then
	v_subject = "サウンドハウス　お見積受付確認メール（自動配信）" & w_order_no
else
	v_subject = "サウンドハウス　ご注文受付確認メール（自動配信）" & w_order_no
end if

Set OBJ_NewMail = Server.CreateObject("CDO.Message") 

OBJ_NewMail.from = v_from_mail
OBJ_NewMail.to = customer_email
OBJ_NewMail.bcc = v_BCC_mail   '2010/08/11 an add
OBJ_NewMail.subject = v_subject
OBJ_NewMail.TextBody = w_body_hd & w_body_dt2 & w_body_tl
OBJ_NewMail.BodyPart.Charset = "iso-2022-jp"

OBJ_NewMail.Send

Set OBJ_NewMail = Nothing

End function

'========================================================================
'
'	Function	顧客情報の取り出し
'
'========================================================================
'
Function get_customer()

'---- 顧客情報取り出し
wSQL = ""
wSQL = wSQL & "SELECT a.顧客名"
wSQL = wSQL & "       , a.顧客フリガナ"
wSQL = wSQL & "       , a.顧客E_mail1"
wSQL = wSQL & "       , b.顧客郵便番号"
wSQL = wSQL & "       , b.顧客都道府県"
wSQL = wSQL & "       , b.顧客住所"
wSQL = wSQL & "       , c.顧客電話番号"
wSQL = wSQL & "       , d.顧客電話番号 AS FAX"
wSQL = wSQL & "  FROM Web顧客 a WITH (NOLOCK)"
wSQL = wSQL & "     , Web顧客住所 b WITH (NOLOCK) LEFT JOIN Web顧客住所電話番号 d WITH (NOLOCK)"
wSQL = wSQL & "                                          ON d.顧客番号 = b.顧客番号"
wSQL = wSQL & "                                         AND d.住所連番 = b.住所連番"
wSQL = wSQL & "                                         AND d.電話区分 = 'FAX'" 
wSQL = wSQL & "     , Web顧客住所電話番号 c WITH (NOLOCK)"
wSQL = wSQL & " WHERE a.顧客番号 = " & userID
wSQL = wSQL & "   AND b.顧客番号 = a.顧客番号"
wSQL = wSQL & "   AND b.住所連番 = 1"
wSQL = wSQL & "   AND c.顧客番号 = a.顧客番号"
wSQL = wSQL & "   AND c.住所連番 = 1"
wSQL = wSQL & "   AND c.電話連番 = 1"
	  
Set RS_customer = Server.CreateObject("ADODB.Recordset")
RS_customer.Open wSQL, Connection, adOpenStatic, adLockOptimistic

End function

'========================================================================
'
'	Function	顧客届先情報の取り出し
'
'========================================================================
'
Function get_todokesaki(p_ship_address_no)

'---- 顧客届先情報取り出し
wSQL = ""
wSQL = wSQL & "SELECT b.住所連番"
wSQL = wSQL & "       , b.住所名称"
wSQL = wSQL & "       , b.顧客郵便番号"
wSQL = wSQL & "       , b.顧客都道府県"
wSQL = wSQL & "       , b.顧客住所"
wSQL = wSQL & "       , c.顧客電話番号"
wSQL = wSQL & "  FROM Web顧客住所 b WITH (NOLOCK)"
wSQL = wSQL & "     , Web顧客住所電話番号 c WITH (NOLOCK)"
wSQL = wSQL & " WHERE b.顧客番号 = " & userID
wSQL = wSQL & "   AND b.住所連番 = " & Clng(p_ship_address_no)
wSQL = wSQL & "   AND c.顧客番号 = b.顧客番号"
wSQL = wSQL & "   AND c.住所連番 = b.住所連番"
wSQL = wSQL & "   AND c.電話区分 = '電話'"
	  
'@@@@@@response.write(wSQL)

Set RS_customer = Server.CreateObject("ADODB.Recordset")
RS_customer.Open wSQL, Connection, adOpenStatic, adLockOptimistic

End function

'========================================================================
'
'	Function	本支店情報取り出し
'
'		return:	w_comp_ryakushou
'
'========================================================================
'
Function get_company()

'---- 本支店取り出し
wSQL = ""
wSQL = wSQL & "SELECT *"
wSQL = wSQL & "  FROM 本支店 WITH (NOLOCK)"
wSQL = wSQL & " WHERE 本支店コード = '1'"
	  
Set RS_company = Server.CreateObject("ADODB.Recordset")
RS_company.Open wSQL, Connection, adOpenStatic, adLockOptimistic

w_comp_ryakushou = RS_company("本支店略称")

RS_company.close

End function

'========================================================================
'
'	Function	休日用ヘッダが必要かどうかチェック
'
'		return:	w_holiday_fl
'
'========================================================================
'
Function CheckHoliday()

Dim v_time

w_holiday_fl = ""

'---- カレンダー情報取り出し(翌日)
wSQL = ""
wSQL = wSQL & "SELECT 休日フラグ"
wSQL = wSQL & "  FROM カレンダー WITH (NOLOCK)"
wSQL = wSQL & " WHERE 年月日 = '" & cf_FormatDate(DateAdd("d", 1, Date()), "YYYY/MM/DD") & "'"
	  
Set RS_calender = Server.CreateObject("ADODB.Recordset")
RS_calender.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RS_calender.EOF = false OR DatePart("w", DateAdd("d", 1, Date())) = vbSunday then		'翌日が休み
	'---- カレンダー情報取り出し(当日)
	wSQL = ""
	wSQL = wSQL  & "SELECT 休日フラグ"
	wSQL = wSQL  & "  FROM カレンダー WITH (NOLOCK)"
	wSQL = wSQL  & " WHERE 年月日 = '" & cf_FormatDate(Date(), "YYYY/MM/DD") & "'"
		  
	Set RS_calender = Server.CreateObject("ADODB.Recordset")
	RS_calender.Open wSQL, Connection, adOpenStatic, adLockOptimistic

	if RS_calender.EOF = false OR DatePart("w", Date()) = vbSunday then		'当日も休み
			w_holiday_fl = "Y"
	else
		if DatePart("w", Date()) = vbSaturday then	'土曜日
			if cf_FormatTime(Now(), "HH:MM") > "17:00" then
				w_holiday_fl = "Y"
			end if
		else				'平日
			if cf_FormatTime(Now(), "HH:MM") > "19:00" then
				w_holiday_fl = "Y"
			end if
		end if
	end if
end if

RS_calender.close

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
