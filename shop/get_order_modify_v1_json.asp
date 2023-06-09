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
<!--#include file="../3rdParty/aspJSON1.17.asp"-->
<%
'========================================================================
'
'	購入履歴
'	変更可能状態の場合"Y"を、不可の場合は"N"を返却する
'
' 変更可能状態
'   Web注文（インターネット、スマートフォン）である。
'   支払い方法が「ローン」以外である。
'   受注ステータスが「受注」(出荷指示あり)でない。
'   メーカー直送品が含まれていない。
'   注残かつ適正在庫数量=0でない。
'
'
'変更履歴
'2016/02/04 GV 新規作成。(注文変更キャンセル機能)
'2018.01.12 GV 入金確認期限切れ見積もり注文は変更キャンセル不可。
'
'========================================================================
'On Error Resume Next

Dim ConnectionEmax

Dim wErrMsg						' エラーメッセージ (他のページから渡されるメッセージ)
Dim wErrDesc
Dim wMsg						' エラーメッセージ (本ページで作成するメッセージ)
Dim wCustomerNo					' 顧客番号
Dim wOrderNo					' 受注番号
Dim wDefer						' 変更モード(Y/N)
Dim wFlg						' 実行フラグ
Dim oJSON						' JSONオブジェクト
Dim modifyFlag					' 変更可能フラグ
Dim cancelFlag					' キャンセル可能フラグ
Dim wNgReason					' 不可理由
Dim wDepositFlag   				' 入金完了フラグ
Dim wDepositAmount 				' 入金合計金額
Dim wWebModCancelFlg			' Web注文変更キャンセル中フラグ
Dim wCItem						' キャンセル商品
Dim cItems						' 配列化したキャンセル商品
Dim btnOn						' ボタン表示フラグ
Dim wOrderDate					' 受注日
Dim wHachuHikiateZero			' 発注引当数量ゼロフラグ
Dim wTekiseiHachuSuuSei			' 適正在庫0かつ色規格別在庫.発注数量が正数
Dim wDepositTerm				' 入金確認期限（日）

'=======================================================================
'	受け渡し情報取り出し & 初期設定
'=======================================================================
wFlg = True

' Getパラメータ
' 顧客番号
wCustomerNo = ReplaceInput_NoCRLF(Trim(Request("cno")))
' 数値のみチェック (ASPは全角でも数字ならTrueを返す)
If (IsNumeric(wCustomerNo) = False) Or (cf_checkNumeric(wCustomerNo) = False) Then
	wFlg = False
End If

' 受注番号
wOrderNo = ReplaceInput_NoCRLF(Trim(Request("ono")))
' 数値のみチェック (ASPは全角でも数字ならTrueを返す)
If (IsNumeric(wOrderNo) = False) Or (cf_checkNumeric(wOrderNo) = False) Then
	wFlg = False
End If

' 保留モード
wDefer = ReplaceInput_NoCRLF(Trim(Request("defer")))
wDefer = UCase(wDefer)
If (wDefer <> "Y") And (wDefer <> "N") And (wDefer <> "") Then
	wFlg = False
End If

' キャンセル商品
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

'---- エラーメッセージをセッションデータに登録   ' member系の他のページ処理にならう
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
Dim vRS1          '受注レコードセット
Dim vRS2          '受注明細レコードセット
Dim vRS3          '更新レコードセット
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

' JSONオブジェクト生成
Set oJSON = New aspJSON

okFlag = True
modifyFlag = "Y"  '変更可能フラグ
cancelFlag = "Y"  'キャンセル可能フラグ
btnOn  = "Y"      'ボタン表示フラグ
wHachuHikiateZero = "" ' 発注引当数量ゼロ
wTekiseiHachuSuuSei = "N" '適正在庫0かつ色規格別在庫.発注数量が正数

'コントロールマスタから見積もり有効期限を取得 2018.01.12 GV add
call getEmaxCntlMst("受注","入金確認待ち期限","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)
If (IsNull(wItemNum1)) Then
	wDepositTerm = 10
Else
	wDepositTerm = wItemNum1
End If

' 入力値が正常の場合
If (wFlg = True) Then
	'受注の取り出し
	wSQL = ""
	wSQL = wSQL & "SELECT "
	wSQL = wSQL & "  受注形態 "
	wSQL = wSQL & " ,支払方法 "
	wSQL = wSQL & " ,受注日 "
	wSQL = wSQL & " ,削除日 "
	wSQL = wSQL & " ,削除日 "
	wSQL = wSQL & " ,Web注文変更キャンセル中フラグ "
	wSQL = wSQL & " ,入金合計金額 "
	wSQL = wSQL & " ,入金完了フラグ "
	wSQL = wSQL & " ,出荷完了日 "
	wSQL = wSQL & " ,その他合計金額 "
	wSQL = wSQL & " ,ISNULL(配送情報明細指定フラグ, 'N') AS 配送情報明細指定フラグ "
	wSQL = wSQL & " ,見積日 " '2018.01.12 GV add
'	wSQL = wSQL & "  FROM 受注 WITH(UPDLOCK) "
	wSQL = wSQL & "  FROM 受注 WITH(NOLOCK) "
	wSQL = wSQL & " WHERE 受注番号 = " & wOrderNo
	wSQL = wSQL & "  AND 顧客番号 = " & wCustomerNo
	wSQL = wSQL & "  AND 削除日 IS NULL "
	'@@@@Response.Write wSQL & "<br>"

	Set vRS1 = Server.CreateObject("ADODB.Recordset")
	vRS1.Open wSQL, ConnectionEmax, adOpenStatic, adLockOptimistic
	'vRS1.Open wSQL, ConnectionEmax, adOpenStatic, adLockPessimistic

	'レコードが存在している場合
	If vRS1.EOF = False Then
		'Web注文変更キャンセル中フラグ
		If (IsNull(vRS1("Web注文変更キャンセル中フラグ"))) Then
			wWebModCancelFlg = "N"
		Else
			If (Trim(vRS1("Web注文変更キャンセル中フラグ")) <> "Y") Then
				wWebModCancelFlg = "N"
			Else
				wWebModCancelFlg = "Y"
			End If
		End If

		'Web注文変更キャンセル中フラグがYの場合、同期まちの可能性があるため
		'変更キャンセル不可
		If (wWebModCancelFlg = "Y") Then
			okFlag = False 'フラグNG
			wNgReason = "8"
			modifyFlag = "N"
			cancelFlag = "N"
			btnOn = "N"
			wFlg = okFlag

			'レコードセットを閉じる
			vRS1.Close

			'レコードセットのクリア
			Set vRS1 = Nothing

			'関数脱出
			Exit Function
		End If

		'出荷完了している場合、ボタン非表示
		If (IsNull(vRS1("出荷完了日")) = False) Then
			btnOn = "N"
			modifyFlag = "N"
			cancelFlag = "N"
		End If

		If (vRS1("受注形態") = "インターネット") Or (vRS1("受注形態") = "スマートフォン") Then
			'変更可能（なにもしない）
		Else
			okFlag = False 'フラグNG
			wNgReason = "1"
			modifyFlag = "N"
			cancelFlag = "N"
			btnOn = "N"
		End If

		If (okFlag = True) Then
			If (Mid(vRS1("支払方法"), 1, 3) = "ローン") Then
				okFlag = False 'フラグNG
				wNgReason = "2"
				modifyFlag = "N"
				cancelFlag = "N"
				btnOn = "N" '2018.01.12 GV add
			End If
		End If

		If (okFlag = True) Then
			If (vRS1("その他合計金額") <> 0) Then
				okFlag = False 'フラグNG
				wNgReason = "10"
				modifyFlag = "N"
				cancelFlag = "N"
				btnOn = "N" '2018.01.12 GV add
			End If
		End If

		If (okFlag = True) Then
			'Emax で届け先を複数設定している場合
			If (vRS1("配送情報明細指定フラグ") = "Y") Then
				okFlag = False 'フラグNG
				wNgReason = "11"
				modifyFlag = "N"
				cancelFlag = "N"
				btnOn = "N" '2018.01.12 GV add
			End If
		End If

		'入金完了フラグ
		If (IsNull(vRS1("入金完了フラグ"))) Then
			wDepositFlag = ""
		Else
			wDepositFlag = CStr(Trim(vRS1("入金完了フラグ")))
		End If

		' 入金合計金額
		If (IsNull(vRS1("入金合計金額"))) Then
			wDepositAmount = 0
		Else
			wDepositAmount = CDbl(vRS1("入金合計金額"))
		End If

		'受注日判定
		If (IsNull(vRS1("受注日")) = True) Or (vRS1("受注日") = "") Then
			wOrderDate = ""
		Else
			wOrderDate = vRS1("受注日")
		End If

		'削除日判定
		If (IsNull(vRS1("削除日")) = True) Or (vRS1("削除日") = "") Then
			deleteDate = ""
		Else
			deleteDate = vRS1("削除日")
		End If

		'2018.01.12 GV add start
		'削除されていない、見積もり状態、入金完了していない
		If (okFlag = True) Then
			If ((deleteDate = "") And (wOrderDate = "") And (wDepositFlag <> "Y")) Then
				'見積日がNullでない、本日との差から入金確認期限以上
				If (IsNull(vRS1("見積日")) = False) And (DateDiff("d", vRS1("見積日"), Now()) >= CInt(wDepositTerm)) Then
					okFlag = False 'フラグNG
					wNgReason = "12"
					modifyFlag = "N"
					cancelFlag = "N"
					btnOn = "N"
				End If
			End If
		End If
		'2018.01.12 GV add end

		If (okFlag = True) Then
			'受注明細レコードを取得
			wSQL = ""
			wSQL = wSQL & "SELECT "
			wSQL = wSQL & "  od.受注明細番号 "
			wSQL = wSQL & " ,od.商品コード "
			wSQL = wSQL & " ,od.メーカー直送フラグ "
			wSQL = wSQL & " ,od.受注単価 "
			wSQL = wSQL & " ,od.出荷指示合計数量 "
			wSQL = wSQL & ", od.受注明細備考 "
			'wSQL = wSQL & " ,z.適正在庫数量 "
			wSQL = wSQL & ", ISNULL(od.適正在庫数量, 0) AS 適正在庫数量 " '注文時の適正在庫
			wSQL = wSQL & " ,od.発注引当数量 "
			wSQL = wSQL & " ,i.セット商品フラグ "
			wSQL = wSQL & " ,i.Web商品フラグ "
			wSQL = wSQL & " ,z.発注数量 as z発注数量"
			wSQL = wSQL & " FROM 受注明細 od WITH (NOLOCK) "
			wSQL = wSQL & " INNER JOIN 色規格別在庫 z WITH (NOLOCK) "
			wSQL = wSQL & "   ON z.メーカーコード = od.メーカーコード "
			wSQL = wSQL & "  AND z.商品コード = od.商品コード "
			wSQL = wSQL & "  AND z.色 = od.色 "
			wSQL = wSQL & "  AND z.規格 = od.規格 "

			wSQL = wSQL & " INNER JOIN 商品 i WITH (NOLOCK) "
			wSQL = wSQL & "   ON i.メーカーコード = z.メーカーコード "
			wSQL = wSQL & "  AND i.商品コード = z.商品コード "

			wSQL = wSQL & " WHERE "
			wSQL = wSQL & "      od.受注番号 = " & wOrderNo
			wSQL = wSQL & " AND od.セット品親明細番号 = 0 "

			'@@@@Response.Write wSQL & "<br>"

			Set vRS2 = Server.CreateObject("ADODB.Recordset")
			vRS2.Open wSQL, ConnectionEmax, adOpenStatic, adLockOptimistic
	
			'レコードが存在している場合
			If vRS2.EOF = False Then

				'発注引当数量ゼロフラグを初期化
				wHachuHikiateZero = "Y"

				'ループ開始
				Do while vRS2.EOF = False

					'発注引当数量がゼロでないものが１つでもある場合
					If (vRS2("発注引当数量") <> 0) Then
						wHachuHikiateZero = "N"
					End If

					'適正在庫数が0かつ、色規格別在庫の発注数量が正数のものが１つでもある場合
					If (vRS2("適正在庫数量") = 0) And (vRS2("z発注数量") > 0) Then
						wTekiseiHachuSuuSei = "Y"
					End If

					'販促品判定
					promote = "N"
					If (CDbl(Trim(vRS2("受注単価"))) = 0) Then
						'受注明細備考に「販促品」と含まれる場合、
						If (InStr(Trim(vRS2("受注明細備考")), "販促品") > 0) Then
							promote = "Y"
						ElseIf (InStr(Trim(vRS2("商品コード")), "HOTMENU") > 0) Then
							promote = "Y"
						End If
					End If

					'出荷指示が1つでもある
					If (vRS2("出荷指示合計数量") > 0) Then
						okFlag = False 'フラグNG
						wNgReason = "3"
						modifyFlag = "N"
						cancelFlag = "N"
						btnOn = "N"
					Else
						'受注が削除されている場合、ボタン表示なし
						If (deleteDate <> "") Then
							btnOn = "N"
						End If
					End If

					'メーカー直送
					If (vRS2("メーカー直送フラグ") = "Y") Then
						okFlag = False 'フラグNG
						wNgReason = "4"
						modifyFlag = "N"
						cancelFlag = "N"
						btnOn = "N" '2018.01.12 GV add
					End If

					' 販促品でない
					If promote = "N" Then
						'Webに掲載していない
						If Trim(vRS2("Web商品フラグ")) <> "Y" Then
							okFlag = False 'フラグNG
							wNgReason = "9"
							modifyFlag = "N"
							cancelFlag = "N"
							btnOn = "N" '2018.01.12 GV add
						End If
					End If

					If (okFlag = True) Then
						'取り込まれただけの状態
						If (wOrderDate = "") And (deleteDate = "") Then
							'modifyFlag = "Y"
						Else
							'セット品の場合は適正在庫数量をみない
							If (vRS2("セット商品フラグ") = "Y") Then
								If ((wOrderDate <> "") And (deleteDate = "")) Then
									'modifyFlag = "Y"
								Else
									okFlag = False 'フラグNG
									wNgReason = "5"
									cancelFlag = "N"
								End If
							Else
								'販促品
								If (promote = "Y") Then
									If ((wOrderDate <> "") And (deleteDate = "")) Then
										'modifyFlag = "Y"
									Else
										okFlag = False 'フラグNG
										wNgReason = "5"
										modifyFlag = "N"
										cancelFlag = "N"
										btnOn = "N" '2018.01.12 GV add
									End If
								Else
									' 注残かつ適正在庫数量が0でない場合、OK
									If (((wOrderDate <> "") And (deleteDate = "")) And (vRS2("適正在庫数量") > 0)) Then
										'modifyFlag = "Y"
									Else 
										'キャンセル商品を指定している場合
										If (wCItem <> "") Then
											'適正在庫0のものが、キャンセルしようとする商品に含まれている場合
											If in_array(vRS2("受注明細番号"), cItems) Then
												wNgReason = "5"
												cancelFlag = "N" 'キャンセルは不可だが、変更は受け付ける
											Else
											End If
										Else
											If (((wOrderDate <> "") And (deleteDate = "")) And (vRS2("適正在庫数量") < 1)) Then
												wNgReason = "5"
												cancelFlag = "N" 'キャンセルは不可だが、変更は受け付ける
											Else
												okFlag = False 'フラグNG
												wNgReason = "5"
												modifyFlag = "N"
												cancelFlag = "N"
												btnOn = "N" '2018.01.12 GV add
											End If
										End If
									End If
								End If ' 販促品
							End If '適正在庫
						End If '日付
					End If

					'NGの場合
					'If okFlag = False Then
					'	'ループ脱出
					'	Exit Do
					'End If

					'次の行へ移動
					vRS2.MoveNext
				Loop
			Else
				'受注明細レコードがない場合、NG
				okFlag = False
				wNgReason = "6"
				modifyFlag = "N"
				cancelFlag = "N"
				btnOn = "N"
			End If

			'受注明細レコードセットを閉じる
			vRS2.Close

			'受注明細レコードセットのクリア
			Set vRS2 = Nothing
		End If '受注明細取得　おわり
	Else
		'受注レコードがない場合、NG
		okFlag = False
		wNgReason = "7"
		modifyFlag = "N"
		cancelFlag = "N"
		btnOn = "N"
	End If

	'適正在庫数が0かつ、色規格別在庫の発注数量が正数のものが１つでもある場合
	'If wTekiseiHachuSuuSei = "Y" Then
	'	okFlag = False
	'	wNgReason = "12"
	'	modifyFlag = "N"
	'	cancelFlag = "N"
	'	'btnOn = "N"
	'End If

	'変更またはキャンセルが可能の場合
	If (modifyFlag = "Y") Or (cancelFlag = "Y") Then
		'Web注文変更キャンセル中フラグを更新する場合
		If (wDefer <> "") Then
			'---- トランザクション開始
			ConnectionEmax.BeginTrans

			'受注の取り出し
			wSQL = ""
			wSQL = wSQL & "SELECT "
			wSQL = wSQL & "  Web注文変更キャンセル中フラグ "
			'wSQL = wSQL & " FROM 受注 WITH(UPDLOCK) "
			wSQL = wSQL & " FROM 受注 "
			wSQL = wSQL & " WHERE 受注番号 = " & wOrderNo
			wSQL = wSQL & "  AND 顧客番号 = " & wCustomerNo
			wSQL = wSQL & "  AND 削除日 IS NULL "
			'@@@@Response.Write wSQL & "<br>"

			Set vRS3 = Server.CreateObject("ADODB.Recordset")
			vRS3.Open wSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

			'レコードが存在している場合
			If vRS3.EOF = False Then
				vRS3("Web注文変更キャンセル中フラグ") = wDefer
				vRS3.update

				'コミット
				ConnectionEmax.CommitTrans
			Else
				'ロールバック
				ConnectionEmax.RollbackTrans
			End If

			'更新レコードセットを閉じる
			vRS3.Close

			'更新レコードセットのクリア
			Set vRS3 = Nothing
		End If
	End If

	'受注レコードセットを閉じる
	vRS1.Close

	'受注レコードセットのクリア
	Set vRS1 = Nothing
Else
	'入力値がNGの場合
	okFlag = False
	wNgReason = "99"
End If

wFlg = okFlag

End Function


'========================================================================
'
'	Function	JSON返却
'
'========================================================================
Function sendResponse()

	' 全件数をJSONデータにセット
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

	'変更とキャンセルの受付条件を分けたため、以下は不要
	'If wFlg = True Then
	'	oJSON.data.Add "result" ,"Y"
	'Else
	'	oJSON.data.Add "result" ,"N"
	'End If

	' -------------------------------------------------
	' JSONデータの返却
	' -------------------------------------------------
	' ヘッダ出力
	Response.AddHeader "Content-Type", "application/json; charset=shift_jis"
	Response.AddHeader "Cache-Control", "no-cache,must-revalidate"
	Response.AddHeader "Pragma", "no-cache"
	Response.AddHeader "X-Content-Type-Options", "nosniff"

	' JSONデータの出力
	Response.Write oJSON.JSONoutput()

End Function

'========================================================================
'
'	Function	配列存在チェック
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
'	Function	Emaxのコントロールマスタからデータ取得
'
'========================================================================

Function getEmaxCntlMst(pSubSystemCd, pItemCd, pItemSubCd, pItemChar1, pItemChar2, pItemNum1, pItemNum2, pItemDate1, pItemDate2)

Dim RS_cntl
Dim v_sql

'---- コントロールマスタ取り出し

v_sql = ""
v_sql = v_sql & "SELECT a.*"
v_sql = v_sql & "  FROM コントロールマスタ a WITH (NOLOCK)"
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
