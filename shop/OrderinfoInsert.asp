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
'	オーダー情報登録
'		POSTされた情報を仮受注へ登録。
'		入力されたデーターのチェック。
'
'変更履歴
'2011/02/16 GV(dy) OrderinfoInsert.aspを元に作り直し
'2011/02/16 hn 配送可能日セット仕様変更
'2011/04/14 hn SessionID関連変更'
'2011/05/02 hn ヤマトの場合で時間指定ありは、納期指定必須チェック追加
'2011/06/29 an #867 運送会社に西濃追加、運送会社の決定方法変更、指定可能納期のチェック方法変更
'2011/07/25 hn ヤマトは使用しないように変更
'2011/08/01 an #1087 Error.aspログ出力対応
'2011/08/11 an #1090 午前指定時、西濃の午前不可地域の場合は運送会社を佐川に変更
'2011/09/12 an #1111/1130 運送会社決定条件に佐川固定フラグ/リードタイムチェック追加
'2012/07/26 nt 嵩重量時、運送会社制御機能を追加
'2012/08/25 nt 佐川代引き禁止地域の制御機能を追加
'2013/02/18 GV #1525 代引きの指定納期制限
'2014/08/05 GV 納品書表示変更対応
'
'========================================================================
On Error Resume Next
Response.Expires = -1			' Do not cache
Response.buffer = true

'---- Session情報
Dim wUserID
Dim wUserName
Dim wMSG

'---- 受け渡し情報を受取る変数
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

'---- 届先情報
Dim wShipNm
Dim wShipZip
Dim wShipPrefecture
Dim wShipAddress
Dim wShipTel

Dim wDeliveryDt
Dim wCustZip
Dim wCustPrefecture
Dim wRitouFl							'離島フラグ
Dim wFreightAm							'送料
Dim wFreightForwarder					'配送会社
Dim wSoukoCnt							'出荷倉庫数
Dim wKoguchi							'個口数

Dim wAfter13FL							'13時以降
Dim wAfter14FL							'14時以降
Dim wAfter15FL							'15時以降
Dim wAfter16FL							'16時以降
'Dim w94HokkaidouRitouFL				'配送先が　九州・四国・北海道・離島   2011/06/29 an del
Dim w94ChugokuHokkaidouFL				'配送先が　九州・四国・中国・北海道   2011/06/29 an add
Dim wHolidayFL							'日曜祭日
Dim wAvailableDate						'指定可能日
Dim wSatFl								'土曜日
Dim wFriFL								'金曜日
Dim wErrDesc   '2011/08/01 an add

Dim kErrFlg		'嵩重量エラーフラグ 2012/07/26 nt add
Dim wSagawaLTFl	'佐川代引き禁止フラグ 2012/08/25 nt add

'---- DB
Dim Connection

'Const w9Shuu4KokuHokkaido = "福岡県,長崎県,佐賀県,大分県,熊本県,宮崎県,鹿児島県,沖縄県,香川県,徳島県,愛媛県,高知県,北海道"    '2011/06/29 an del
Const w9Shuu4KokuChugokuHokkaido = "福岡県,長崎県,佐賀県,大分県,熊本県,宮崎県,鹿児島県,香川県,徳島県,愛媛県,高知県,鳥取県,岡山県,島根県,広島県,山口県,北海道"   '2011/06/29 an add
Const cAddDaysToNyukaYoteibi = 2

'=======================================================================
'	受け渡し情報取り出し
'=======================================================================
'---- Session変数
wUserID = Session("UserID")
wUserName = Session("userName")
wMsg = Session.contents("msg")

'---- 受け渡し情報取り出し
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
'OrderInfoEnter.asp で非表示だった場合
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

'---- セッション切れチェック
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

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "OrderinfoInsert.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

Call close_db()

If Err.Description <> "" Then
	Response.Redirect g_HTTP & "shop/Error.asp"
End If

'---- エラーが無いときは注文内容確認ページ、エラーがあれば注文内容指定ページへ
If wMSG = "" Then
	Select Case cmd
		Case "next"
			If payment_method = "ローン" Then
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

'---- 仮受注情報更新
Call update_order_header()

If cmd = "next" Then
	'---- 入力データーのチェック
	Call validate_data()
End If

End Function

'========================================================================
'
'	Function	仮受注情報の更新
'
'========================================================================
Function update_order_header()

Dim RSv
Dim vSQL

'---- 届先情報取得
If isNumeric(ship_address_no) = False Then
	ship_address_no = 1
End If

Call GetTodokesakiInfo(ship_address_no)

'---- 離島フラグの設定
Call setRitouFlag(wShipZip)

'---- 佐川禁止フラグの設定
Call setSagawaLTFlag(wShipZip)

'---- 送料・配送会社・出荷倉庫数・個口数 計算
Call fCalcShipping(gSessionID, "一括", wFreightAm, wFreightForwarder, wSoukoCnt, wKoguchi)		'2011/04/14 hn mod

'---- 仮受注Recordset取り出し
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    *"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    仮受注"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "    SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic, adLockOptimistic

'---- 顧客情報
RSv("顧客番号") = wUserID
RSv("顧客E_mail") = customer_email
RSv("顧客電話番号") = telephone
RSv("見積フラグ") = ""

'---- 支払方法
RSv("支払方法") = payment_method

Select Case payment_method
	Case "銀行振込"
		If furikomi_nm = "" Then
			furikomi_nm = customer_kn
		End If
		RSv("振込名義人") = furikomi_nm
		RSv("ローン頭金ありフラグ") = ""
		RSv("希望ローン回数") = 0
		RSv("ローン頭金") = 0
		RSv("ローン金額") = 0
		RSv("ローン会社") = ""
		RSv("オンラインローン申込フラグ") = ""

	Case "ローン"
		RSv("振込名義人") = ""

	Case Else
		RSv("振込名義人") = ""
		RSv("ローン頭金ありフラグ") = ""
		RSv("希望ローン回数") = 0
		RSv("ローン頭金") = 0
		RSv("ローン金額") = 0
		RSv("ローン会社") = ""
		RSv("オンラインローン申込フラグ") = ""

End Select

'---- 届先情報
RSv("届先住所連番") = ship_address_no
RSv("届先名前") = wShipNm
RSv("届先郵便番号") = wShipZip
RSv("届先都道府県") = wShipPrefecture
RSv("届先住所") = wShipAddress
RSv("届先電話番号") = wShipTel

If ship_address_no = 1 Then
	RSv("届先区分") = "S"
	RSv("届先納品書送付可フラグ") = "Y"		'届先が住所と同じの場合は無条件にY
Else
	RSv("届先区分") = "D"
	RSv("届先納品書送付可フラグ") = ship_invoice_fl
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
	RSv("指定納期") = wDeliveryDt
Else
	RSv("指定納期") = NULL
End If

'---- 配送情報 
'---- 運送会社チェック＆変更
call CheckFreightForwarder()   '2011/06/29 an add

'---- 離島1個口で空輸禁止商品が含まれてなく重量商品でない場合は、配送会社を「ヤマト運輸」に強制変更   '2011/06/29 an del s
'If wRitouFl = "Y" And wKoguchi = 1 And checkKuuyukinshiShouhin() = 0 And checkJyuuryouShouhin() = 0 Then
'	freight_forwarder = "2"
'End If     '2011/06/29 an del e

RSv("運送会社コード") = freight_forwarder

RSv("時間指定") = delivery_tm

RSv("営業所止めフラグ") = eigyousho_dome_fl

If RSv("支払方法") = "代引き" Then
	RSv("一括出荷フラグ") = "Y"
Else
	RSv("一括出荷フラグ") = ikkatsu_fl
End If

If wRitouFl = "Y" Then
	RSv("一括出荷フラグ") = "Y"
End If

'---- 領収書
RSv("領収書発行フラグ") = receipt_fl
If receipt_fl = "Y" Then
	If receipt_nm <> "" Then
		RSv("領収書宛先") = receipt_nm
	Else
		RSv("領収書宛先") = wUserName
	End If
	If receipt_memo <> "" Then
		RSv("領収書但し書き") = receipt_memo
	Else
		RSv("領収書但し書き") = "音響機器代として"
	End If
Else
	RSv("領収書宛先") = ""
	RSv("領収書但し書き") = ""
End If

RSv("リベート使用フラグ") = RebateFl
RSv("過不足相殺金額") = KabusokuAm

RSv("最終更新日") = Now()

RSv.Update
RSv.Close

End Function

'========================================================================
'
'	Function	届先の顧客情報の取得
'
'========================================================================
Function GetTodokesakiInfo(vAddressNo)

Dim RSv
Dim vSQL

vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    a.住所名称"
vSQL = vSQL & "  , a.顧客郵便番号"
vSQL = vSQL & "  , a.顧客都道府県"
vSQL = vSQL & "  , a.顧客住所"
vSQL = vSQL & "  , b.顧客電話番号"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    Web顧客住所 a WITH (NOLOCK)"
vSQL = vSQL & "  , Web顧客住所電話番号 b WITH (NOLOCK)"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "        a.顧客番号 = " & wUserID
vSQL = vSQL & "    AND a.住所連番 = " & vAddressNo
vSQL = vSQL & "    AND b.電話区分 = '電話'"
vSQL = vSQL & "    AND b.顧客番号 = a.顧客番号"
vSQL = vSQL & "    AND b.住所連番 = a.住所連番"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic, adLockOptimistic

wShipNm = RSv("住所名称")
wShipZip = RSv("顧客郵便番号")
wShipPrefecture = RSv("顧客都道府県")
wShipAddress = RSv("顧客住所")
wShipTel = RSv("顧客電話番号")

RSv.Close

End Function

'========================================================================
'
'	Function	入力データーのチェック
'
'========================================================================
Function validate_data()

Dim vAddress
Dim vDateMMDD
Dim vDateTimeMsg '2011/02/16 hn add

'---- 支払方法
If payment_method = "" Then
	wMSG = wMSG & "お支払い方法を選択してください。<br>"
End If

'---- 振込人名義
If cf_checkKataKana(furikomi_nm) = False Then
	wMSG = wMSG & "振込人名義は全角カタカナのみで入力してください。<br>"
End If

'---- 代引き不可商品チェック
If (i_tokuchuu_fl = "Y" Or i_daibiki_fuka_fl = "Y") And payment_method = "代引き" Then
	wMSG = wMSG & "大型商品または特別手配の商品が含まれているため、代引きでのご注文は受付できません。他のお支払い方法への変更をお願いします。<br>"
End If

'---- お届先
If payment_method = "ローン" And ship_address_no <> 1 Then
	wMSG = wMSG & "ローンでお支払いの場合、お届け先はお客様の住所のみとなり、別の住所への配送指定はできません。<br>"
End If

'---- 配達日指定
If delivery_fl = "Y" And delivery_mm = "" And delivery_dd = "" And delivery_tm = "" Then
	wMSG = wMSG & "配達日の日付または時間を入力してください。<br>"
End If

If (delivery_mm <> "" And delivery_dd = "") Or (delivery_mm = "" And delivery_dd <> "") Then
	wMSG = wMSG & "配達日の指定が正しくありません。<br>"
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
		wMSG = wMSG & "配達日の指定が正しくありません。<br>"
		wDeliveryDt = ""
	End If
Else
	wDeliveryDt = ""
End If

'---- 配達日指定(翌日指定時のチェック)
If wDeliveryDt <> "" Then

	'---- 配達日指定可能かチェック（注残品で入荷予定がないか、入荷予定+2日より前は指定NG）
	If checkNyukaYoteibi() = False Then
		wMSG = wMSG & "在庫のない商品がご注文に含まれているため、配達希望日の指定はできません。<br>"
	Else
		'---- ローンの場合、配送日指定不可
		If payment_method = "ローン" Then
			wMSG = wMSG & "お支払い方法がローンの場合、配達希望日の指定はできません。<br>"

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

			'---- 13時以降かどうかチェック
			If DatePart("h", Now()) >= 13 Then
				wAfter13FL = True
			End If

			'---- 14時以降かどうかチェック
			If DatePart("h", Now()) >= 14 Then
				wAfter14FL = True
			End If

			'---- 15時以降かどうかチェック
			If DatePart("h", Now()) >= 15 Then
				wAfter15FL = True
			End If

			'---- 16時以降かどうかチェック
			If DatePart("h", Now()) >= 16 Then
				wAfter16FL = True
			End If

			'---- 今日が日曜･祭日かどうかチェック
			If (DatePart("w", Date()) = vbSunday) Or (checkHoliday(Date()) = True) Then
				wHolidayFL = True
			End If

			'---- 今日が土曜日かどうかチェック
			If DatePart("w", Date()) = vbSaturday Then
				wSatFL = True
			End If

			'---- 今日が金曜日かどうかチェック
			If DatePart("w", Date()) = vbFriday Then
				wFriFL = True
			End If

			'---- 配送先が　九州・四国・中国・北海道かどうかチェック  
			'If wRitouFl = "Y" Or Instr(w9Shuu4KokuHokkaido, wShipPrefecture) > 0 Then
			if Instr(w9Shuu4KokuChugokuHokkaido, wShipPrefecture) > 0 Then
				'w94HokkaidouRitouFL = True     '2011/06/29 an del
				w94ChugokuHokkaidouFL = True    '2011/06/29 an add
			End If

			'---- 指定可能日をセット
			wAvailableDate = setAvailableDate()

			'---- 配送日チェック
			If wDeliveryDt < wAvailableDate Then

				'2011/02/16 hn mod s
				'---- 休日
				If wHolidayFl = True Then
					vDateTimeMsg = "休日の"
				else

					'---- コンビニ 土曜日 13時以降
					If payment_method = "コンビニ支払" AND wSatFl = True AND wAfter13Fl = true Then
						vDateTimeMsg = "コンビニ支払の土曜日13時以降の"
					end if

					'---- コンビニ 平日 14時以降
					If payment_method = "コンビニ支払" AND wSatFl = False AND wAfter14Fl = true Then
						vDateTimeMsg = "コンビニ支払の14時以降の"
					end if

					'---- 銀行振込 14時以降
					If payment_method = "銀行振込" AND wAfter14Fl = true Then
						vDateTimeMsg = "銀行振込の14時以降の"
					end if

					'---- 銀行振込 土曜日
					If payment_method = "銀行振込" AND wSatFL = true Then
						vDateTimeMsg = "銀行振込の土曜日の"
					end if

					'---- 代引き・カード 16時以降
					If (payment_method = "代引き" OR payment_method = "クレジットカード") AND wAfter16Fl = true Then
						vDateTimeMsg = payment_method & "の16時以降の"
					end if
				end if

				'---- 九州･四国･中国・北海道
				'If w94HokkaidouRitouFL = True Then                             '2011/06/29 an del
					'wMSG = wMSG & "お届け先が九州・四国・北海道・離島で、"     '2011/06/29 an del
				If w94ChugokuHokkaidouFL = True Then                            '2011/06/29 an add
					wMSG = wMSG & "お届け先が九州・四国・中国・北海道で、"      '2011/06/29 an add
				end if
				
				'---- 離島                      '2011/06/29 an add s
				If wRitouFl = "Y" Then
					wMSG = wMSG & "お届け先が沖縄・離島で、"
				end if                          '2011/06/29 an add e

				wMSG = wMSG & vDateTimeMsg & "配送日指定は、" & wAvailableDate & "以降を指定してください。<br>"
				'2011/02/16 hn mod e

			End If

			'---- 60日以内
			'2013/02/18 GV #1525 MOD START
'			If (DateDiff("d", DateAdd("d", 60, Date()), wDeliveryDt) > 0) Then
'				wMSG = wMSG & "配送日指定は60日以内の日付を指定してください。<br>"
			If (payment_method = "代引き" AND DateDiff("d", DateAdd("d", 10, Date()), wDeliveryDt) > 0) Then
				wMSG = wMSG & "配送日指定は10日以内の日付を指定してください。<br>"
			ElseIf (payment_method <> "代引き" AND DateDiff("d", DateAdd("d", 60, Date()), wDeliveryDt) > 0) Then
				wMSG = wMSG & "配送日指定は60日以内の日付を指定してください。<br>"
			'2013/02/18 GV #1525 MOD END
			End If
		End If
	End If
End If

'---- 営業所止め
If payment_method = "ローン" And eigyousho_dome_fl = "Y" Then
	wMSG = wMSG & "ローンでお支払いの場合、営業所止め指定はできません。<br>"
End If

'---- 重量商品ありのときはヤマトはエラー  2011/06/29 an del→顧客には選択できないため削除
'If freight_forwarder = "2" Then
'	If checkJyuuryouShouhin() > 0 Then
'		wMSG = wMSG & "ご注文に重量商品が含まれていますのでヤマト運輸の指定はできません。<br>"
'	End If
'End If

'---- ヤマト代引きは複数個口はエラー  2011/06/29 an del→顧客には選択できないため削除
'If freight_forwarder = "2" And payment_method = "代引き" Then
'	If wKoguchi > 1 Then
'		wMSG = wMSG & "ご注文は複数個口となります。ヤマト運輸での代引きの指定はできません。運送会社を変更するか、お支払方法を変更してください。<br>"
'	End If
'End If

'---- ヤマト+配送日指定なし+時間指定はエラー	2011/05/02 hn add
if (freight_forwarder = "2") AND (delivery_tm <> "") AND (delivery_mm = "")then
	wMSG = wMSG & "配送時間を指定される場合は、配送日も指定してください。<br>"
end if

'---- 離島+佐川+時間指定はエラー
If wRitouFl = "Y" And freight_forwarder = "1" And delivery_tm <> "" Then
	wMSG = wMSG & "お客様のお届け先へは時間指定を行えません。<br>"
End If

'---- カードで営業所止め指定はエラー
If payment_method = "クレジットカード" And eigyousho_dome_fl = "Y" Then
	wMSG = wMSG & "クレジットカードでご注文の場合は、営業所止めの指定はできません。<br>"
End If

'---- 代引き,ローン,コンビニ/郵便局支払の場合領収証は出せないチェック
If receipt_fl = "Y" And (payment_method = "代引き" Or payment_method = "ローン" Or payment_method = "コンビニ支払") Then
	wMSG = wMSG & "指定のお支払い方法でご購入の際，領収書は発行できません。<br>"
End If

'---- 領収証宛先または但し書き入力で｢必要｣チェックされてない時はエラー
If receipt_fl <> "Y" And (receipt_nm <> "" Or receipt_memo <> "") Then
	wMSG = wMSG & "領収書が必要な場合は「必要」をチェックしてください。不要な場合は宛名または但し書きをクリアしてください。<br>"
End If

'2012/07/26 nt del
'---- 離島で空輸禁止商品が含まれている場合は佐川のみOK
'If wRitouFl = "Y" And freight_forwarder <> 1 And checkKuuyukinshiShouhin() > 0 Then
'	wMSG = wMSG & "空輸禁止商品が含まれていますので運送会社を佐川急便に変更してください。<br>"
'End If

'2012/07/26 nt add start
'---- 「西濃」+「嵩重量品あり」+「離島」+「代引き」の時はエラー
if kErrFlg = false and wRitouFl = "Y"  and payment_method = "代引き" then
	if wMSG <> "" then wMSG = wMSG & "<br>" end if
	wMSG = wMSG & "お客様のご注文は、重量、もしくは大きさが規定値を超える商品を含む為、代金引換以外のお支払方法を選択してください。<br>"
end if

'---- 「西濃」+「嵩重量品あり」+「日時指定：有り」の時はエラー
if kErrFlg = false and delivery_fl = "Y" then
	if wMSG <> "" then wMSG = wMSG & "<br>" end if
	wMSG = wMSG & "お客様のご注文は、重量、もしくは大きさが規定値を超える商品を含む為、配達日時なしでのお届けとなります。<br>"
end if
'2012/07/26 nt add end

'2012/08/25 nt add start
if kErrFlg = false and wSagawaLTFl = "Y"  and payment_method = "代引き" then
	wMSG = wMSG & "お客様のご注文は、代金引換をお受けできない地域の為、代金引換以外のお支払方法を選択してください。<br>"
end if
'2012/08/25 nt add end

If wMSG <> "" Then
	wMSG = "<b>以下の入力エラーを訂正してください。</b><br><br>" & wMSG
End If

End Function

'========================================================================
'
'	Function	離島フラグの設定
'
'		parm:		配送先郵便番号
'		return:	沖縄・離島なら　wRitouFl = Y
'				離島以外　　　　wRitouFl = N
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

'---- 離島チェック
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    郵便番号"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    離島 WITH (NOLOCK)"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "    郵便番号 = '" & vZip & "'"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic

If RSv.EOF = True Then
	wRitouFl = "N"
	'---- 沖縄は離島扱い               '2011/06/29 an add s
	if wShipPrefecture = "沖縄県" then
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
'	Function	佐川代引き禁止フラグの設定
'
'		parm:		配送先郵便番号
'		return:	佐川代引き禁止地域なら　wSagawaLTFl = Y
'				禁止地域以外　　　　　　wSagawaLTFl = N
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

'---- 佐川制限チェック
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    郵便番号"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    佐川制限 WITH (NOLOCK)"
vSQL = vSQL & " WHERE 郵便番号 = '" & vZip & "'"
vSQL = vSQL & "   AND 代引不可フラグ='Y'"

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
'	Function	祭日チェック
'
'		input : チェックする日(YYYY/MM/DD)
'		return:	祭日なら　True
'				祭日以外　False
'
'========================================================================
Function checkHoliday(p_date)

Dim RSv
Dim vSQL

'---- 祭日チェック
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    年月日"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    カレンダー WITH (NOLOCK)"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "        年月日 = '" & p_date & "'"
vSQL = vSQL & "    AND 休日フラグ = 'Y'"

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
'	Function	配送可能日セット    2010/12/17 an mod  2011/02/16 hn mod
'
'		return:	配送可能日 (YYYY/MM/DD)
'
'   支払方法＝コンビニ
'     土曜日以外
'     14時以前：受付日＝当日
'     14時以降：受付日＝翌日

'     土曜日
'     13時以前：受付日＝当日
'     13時以降：受付日＝翌日
'
'   支払方法＝銀行振込（曜日関係なし）
'     14時以前：受付日＝当日
'     14時以降：受付日＝翌日
'
'   支払方法＝代引き（曜日関係なし）
'     16時以前：受付日＝当日
'     16時以降：受付日＝翌日
'
'   支払方法＝ローン
'     納期指定不可
'
'==============================
'   受付日が、日曜、祝日、（銀行振込の場合は土曜日も含む）の場合は次の営業日を受付日とする。
'
'   納期指定可能日は、
'   受付日+1日　（下記以外）
'   受付日+2日　（九州、四国、中国、北海道）
'   受付日+5日　（沖縄、離島）
'
'========================================================================
'
Function setAvailableDate()

Dim vOrderDate

vOrderDate = cf_FormatDate(Date(), "YYYY/MM/DD")

'---- コンビニ支払：土曜日以外　14時以降は受付日は翌日
if payment_method = "コンビニ支払" AND wSatFl = false AND wAfter14Fl = true then
	vOrderDate = cf_FormatDate(DateAdd("d", 1, Date()), "YYYY/MM/DD")
end if

'---- コンビニ支払：土曜日　13時以降は受付日は翌日
if payment_method = "コンビニ支払" AND wSatFl = true AND wAfter13Fl = true then
	vOrderDate = cf_FormatDate(DateAdd("d", 1, Date()), "YYYY/MM/DD")
end if

'---- 銀行振込：14時以降は受付日は翌日
if payment_method = "銀行振込" AND wAfter14Fl = true then
	vOrderDate = cf_FormatDate(DateAdd("d", 1, Date()), "YYYY/MM/DD")
end if

'---- 代引き・カード：16時以降は受付日は翌日
If (payment_method = "代引き" OR payment_method = "クレジットカード") AND wAfter16Fl = true Then
	vOrderDate = cf_FormatDate(DateAdd("d", 1, Date()), "YYYY/MM/DD")
end if

'---- 受付日が、日曜、祝日、（銀行振込の場合は土曜日も含む）の場合は次の営業日を受付日とする。
Do
	if  (DatePart("w", vOrderDate) = vbSunday) OR (checkHoliday(vOrderDate) = True) then
		vOrderDate = cf_FormatDate(DateAdd("d", 1, vOrderDate), "YYYY/MM/DD")
	else 
		if  payment_method = "銀行振込" AND (DatePart("w", vOrderDate) = vbSaturday) then
			vOrderDate = cf_FormatDate(DateAdd("d", 1, vOrderDate), "YYYY/MM/DD")
		else 
			exit do
		end if
	end if
Loop

'---- 指定可能日
'if w94HokkaidouRitouFl = true then                '2011/06/29 an del
'---- 離島：受付日+5日                             '2011/06/29 an mod s
if wRitouFl = "Y" then
	setAvailableDate = cf_FormatDate(DateAdd("d", 5, vOrderDate), "YYYY/MM/DD")
else
	'---- 九州･四国･中国・北海道：受付日+2日
	if w94ChugokuHokkaidouFL = true then
		setAvailableDate = cf_FormatDate(DateAdd("d", 2, vOrderDate), "YYYY/MM/DD")
	'---- その他：受付日+1日                       '2011/06/29 an mod e
	else
		setAvailableDate = cf_FormatDate(DateAdd("d", 1, vOrderDate), "YYYY/MM/DD")
	end if
end if

End function

'========================================================================
'
'	Function	重量商品チェック
'
'		return:	重量商品件数
'
'========================================================================
Function checkJyuuryouShouhin()

Dim RSv
Dim vSQL

'---- 重量商品件数取り出し
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    COUNT(*) AS 重量商品件数"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    仮受注明細 a WITH (NOLOCK)"
vSQL = vSQL & "  , Web商品 b WITH (NOLOCK)"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "        b.メーカーコード = a.メーカーコード"
vSQL = vSQL & "    AND b.商品コード = a.商品コード"
vSQL = vSQL & "    AND a.SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod
vSQL = vSQL & "    AND b.送料区分 = '重量商品'"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic

checkJyuuryouShouhin = RSv("重量商品件数")

RSv.Close

End Function

'2012/07/26 nt comment mod
'========================================================================
'
'	Function	空輸禁止商品チェック⇒嵩重量品チェック
'               12/07/26、空輸禁止フラグは未使用のため、本フラグは嵩重量品
'               判別のフラグへと転用。
'
'		return:	空輸禁止商品件数⇒嵩重量品件数
'
'========================================================================
Function checkKuuyukinshiShouhin()

Dim RSv
Dim vSQL

'---- 空輸禁止商品件数取り出し
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    COUNT(*) AS 空輸禁止商品件数"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    仮受注明細 a WITH (NOLOCK)"
vSQL = vSQL & "  , Web商品 b WITH (NOLOCK)"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "        b.メーカーコード = a.メーカーコード"
vSQL = vSQL & "    AND b.商品コード = a.商品コード"
vSQL = vSQL & "    AND a.SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod
vSQL = vSQL & "    AND b.空輸禁止フラグ = 'Y'"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic

checkKuuyukinshiShouhin =  RSv("空輸禁止商品件数")

RSv.Close

End Function

'========================================================================
'
'	Function	注残の有無と入荷予定日を確認し、配達日指定可能かチェック
'
'		return:	 配達日指定可能なら True
'                配達日指定不可なら False
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

'---- 仮受注明細取り出し
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "     a.メーカーコード"
vSQL = vSQL & "   , a.商品コード"
vSQL = vSQL & "   , a.色"
vSQL = vSQL & "   , a.規格"
vSQL = vSQL & "   , a.受注数量"
vSQL = vSQL & "   , b.引当可能数量"
vSQL = vSQL & "   , b.引当可能入荷予定日"
vSQL = vSQL & "   , c.セット商品フラグ"
vSQL = vSQL & " FROM"
vSQL = vSQL & "     仮受注明細 a WITH (NOLOCK)"
vSQL = vSQL & "   , Web色規格別在庫 b WITH (NOLOCK)"
vSQL = vSQL & "   , Web商品 c WITH (NOLOCK)"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "         b.メーカーコード = a.メーカーコード"
vSQL = vSQL & "     AND b.商品コード = a.商品コード"
vSQL = vSQL & "     AND b.色 = a.色"
vSQL = vSQL & "     AND b.規格 = a.規格"
vSQL = vSQL & "     AND c.メーカーコード = a.メーカーコード"
vSQL = vSQL & "     AND c.商品コード = a.商品コード"
vSQL = vSQL & "     AND b.終了日 IS NULL"
vSQL = vSQL & "     AND a.SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic

If RSv.EOF = False Then

	'---- 全仮受注明細商品をチェック
	Do Until RSv.EOF = True

		'セット品の場合はセット品全体の在庫数量、入荷予定日をチェックする
		If RSv("セット商品フラグ") = "Y" Then

			'---- セット品在庫数量、MAX引当可能予定取得
			Call GetSetCount(RSv("メーカーコード"), RSv("商品コード"), RSv("色"), RSv("規格"), vSetCount, vMaxNyukaYoteibi)

			If vMaxNyukaYoteibi < Date() Then			'入荷予定のないセット品の場合は2000/01/01が入っている
				vMaxNyukaYoteibi = ""
			Else
				vMaxNyukaYoteibi = vMaxNyukaYoteibi		'セット品全体のMAX入荷予定日で上書き
			End If

			vHikiateKanouQt = vSetCount					'セット品全体のMIN在庫数量で上書き

		Else
			vHikiateKanouQt = RSv("引当可能数量")
			vMaxNyukaYoteibi = RSv("引当可能入荷予定日")
		End If

		If RSv("受注数量")  > vHikiateKanouQt Then		'注残品の場合は入荷予定日をチェック
			'---- 入荷予定のない商品が1つでもある場合は配達日指定不可
			If IsNULL(vMaxNyukaYoteibi) = True Or vMaxNyukaYoteibi = "" Then
				checkNyukaYoteibi = False
				Exit Do
			Else
				'---- 指定可能日（入荷予定日+2以降）
				wAvailableDate = cf_FormatDate(DateAdd("d", cAddDaysToNyukaYoteibi, vMaxNyukaYoteibi), "YYYY/MM/DD")

				'---- 入荷が間に合わない商品が1つでもある場合は配達日指定不可
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
'	Function	運送会社チェック＆変更    2011/06/29 an add
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

'---- 離島
If wRitouFl = "Y" then
'2011/07/25 hn del s
'	'---- 離島1個口で空輸禁止商品が含まれてなく重量商品でない場合は、配送会社を「ヤマト運輸」に強制変更
'	if wKoguchi = 1 And checkKuuyukinshiShouhin() = 0 And checkJyuuryouShouhin() = 0 Then
'		freight_forwarder = "2"
'	'---- 上記以外の離島は佐川
'	else
'2011/07/25 hn del e
		freight_forwarder = "1"
'	end if	'2011/07/25 hn del 

'---- 離島以外
else
	
	'---- 規定運送会社をセット
	vSQL = ""
	vSQL = vSQL & "SELECT 運送会社コード"
	vSQL = vSQL & "  FROM 県別規定運送会社 WITH (NOLOCK)"
	vSQL = vSQL & " WHERE 県 = '" & wShipPrefecture & "'"
	
	Set RSv = Server.CreateObject("ADODB.Recordset")
	RSv.Open vSQL, Connection, adOpenStatic
	
	if RSv.EOF = false then
		freight_forwarder = RSv("運送会社コード")
	end if
	
	RSv.Close

	'---- 佐川固定なら佐川に変更   2011/09/11 an add s
	vSQL = ""
	vSQL = vSQL & "SELECT 佐川固定フラグ"
	vSQL = vSQL & "  FROM Web顧客 WITH (NOLOCK)"
	vSQL = vSQL & " WHERE 顧客番号 = " & wUserID
	
	Set RSv = Server.CreateObject("ADODB.Recordset")
	RSv.Open vSQL, Connection, adOpenStatic
	
	if RSv.EOF = false then
		if RSv("佐川固定フラグ") = "Y" then
			freight_forwarder = "1"
		end if
	end if
	
	RSv.Close                      '2011/09/11 an add e
	
end if

'---- 西濃で配送不可の場合は運送会社を変更
'if freight_forwarder = "5" AND wDeliveryDt <> "" then   '2011/08/11 an del
if freight_forwarder = "5" then                          '2011/08/11 an mod s

	'---- 配達指定日ありの場合、西濃配達不可日でないかチェック
	if wDeliveryDt <> "" then
		'---- 日曜なら佐川
		if DatePart("w", wDeliveryDt) = vbSunday then
			freight_forwarder = "1"
		'---- 西濃配送不可日なら佐川
		else
			vSQL = ""
			vSQL = vSQL & "SELECT 西濃配達不可日フラグ"
			vSQL = vSQL & "  FROM カレンダー WITH (NOLOCK)"
			vSQL = vSQL & " WHERE 年月日 = '" & wDeliveryDt & "'"

			Set RSv = Server.CreateObject("ADODB.Recordset")
			RSv.Open vSQL, Connection, adOpenStatic
			
			if RSv.EOF = false then
				if RSv("西濃配達不可日フラグ") = "Y" then
					freight_forwarder = "1"
				end if
			end if
			
			RSv.Close
		end if
	end if

	'---- 西濃仕分コードマスタチェック    2011/09/12 an mod s
	vSQL = ""
	vSQL = vSQL & "SELECT 配達午前午後"
	vSQL = vSQL & "     , リードタイム"
	vSQL = vSQL & "  FROM 西濃仕分コードマスタ WITH (NOLOCK)"
	vSQL = vSQL & " WHERE 郵便番号 = '" & Replace(wShipZip,"-","") & "'"

	Set RSv = Server.CreateObject("ADODB.Recordset")
	RSv.Open vSQL, Connection, adOpenStatic
	
	if RSv.EOF = false then
		
		'---- 時間指定_西濃 01の文字列取得（午前指定）
		call getCntlMst("受注","時間指定_西濃","01", vItemChar1, vItemChar2, vItemNum1, vItemNum2, vItemDate1, vItemDate2)
	
		'---- 西濃午前不可地域なら佐川に変更
		if delivery_tm = vItemChar1 then
			if RSv("配達午前午後") = "P" then
				freight_forwarder = "1"
			end if
		end if
		
		'---- リードタイム >= 2なら佐川に変更
		if RSv("リードタイム") >= 2 then
			freight_forwarder = "1"
		end if
	end if
	
	RSv.Close      '2011/08/11 an mod e  2011/09/12 an mod e
	
end if

'2012/07/26 nt add start
'---- 嵩重量品、運送会社の選定制御
'----「嵩重量品」を含めば、条件付きで「西濃」へ変更
kErrFlg = true
if checkKuuyukinshiShouhin() <> 0 then
	freight_forwarder = "5"

	'---- 「離島」で、かつ「代引き」指定の場合はエラー対象
	if wRitouFl = "Y"  and payment_method = "代引き" then
		kErrFlg = false

	'---- 「日時指定」がある場合はエラー対象
	elseif delivery_fl = "Y" then
		kErrFlg = false

	end if
end if
'2012/07/26 nt add end

'2012/08/25 nt add start
'---- 「佐川代引き禁止地域」で、かつ「代引き」指定の場合はエラー対象
if wSagawaLTFl = "Y" and payment_method = "代引き" then
	kErrFlg = false
end if
'2012/08/25 nt add end


End Function

%>
