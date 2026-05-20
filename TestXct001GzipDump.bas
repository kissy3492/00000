' ============================================================
' XCT001 レスポンス保存版 ミニテスト
' ============================================================
' 目的:
'   gzip 自動解凍された JSON の中身を Raw_API シートに保存し、
'   「1㎡」「変動率」などの具体的なキー名を Ctrl+F で確認できるようにする
'
' 使い方:
'   1. v18.7.1 を開いた状態で
'   2. VBE で「標準モジュール」を新規追加
'   3. このコードを貼り付け
'   4. Sub TestXct001GzipDump の中の YOUR_API_KEY_HERE を APIキーに置換
'   5. F5 で実行
'   6. 完了後、Raw_API シートを開いて Ctrl+F で「1㎡」「変動率」を検索
' ============================================================

Public Sub TestXct001GzipDump()
    Const URL As String = "https://www.reinfolib.mlit.go.jp/ex-api/external/XCT001?year=2025&area=13&division=00"
    Const API_KEY As String = "YOUR_API_KEY_HERE"  ' ← ここを実際の APIキーに置換
    
    If API_KEY = "YOUR_API_KEY_HERE" Then
        MsgBox "APIキーを設定してから実行してください。", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' --- HTTP 取得 ---
    Dim http As Object
    On Error Resume Next
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    If Err.Number <> 0 Then
        MsgBox "MSXML2.XMLHTTP.6.0 が使えません: " & Err.Description, vbCritical
        Application.ScreenUpdating = True
        Exit Sub
    End If
    On Error GoTo 0
    
    http.Open "GET", URL, False
    http.setRequestHeader "Ocp-Apim-Subscription-Key", API_KEY
    http.setRequestHeader "Accept", "application/json"
    
    On Error Resume Next
    http.Send
    If Err.Number <> 0 Then
        MsgBox "通信エラー: " & Err.Description, vbCritical
        Application.ScreenUpdating = True
        Exit Sub
    End If
    On Error GoTo 0
    
    Dim status As Long: status = CLng(http.Status)
    Dim respText As String: respText = CStr(http.responseText)
    Set http = Nothing
    
    ' --- 検索用キーワードを事前にチェック ---
    Dim foundM2_U33A1 As Boolean      ' "1㎡" (U+33A1)
    Dim foundM2_ascii As Boolean      ' "1m2"
    Dim foundM2_caret As Boolean      ' "1m^2"
    Dim foundM2_kanji As Boolean      ' "1平方メートル"
    
    foundM2_U33A1 = (InStr(1, respText, "1" & ChrW$(&H33A1) & "当たりの価格") > 0)
    foundM2_ascii = (InStr(1, respText, "1m2当たりの価格") > 0)
    foundM2_caret = (InStr(1, respText, "1m^2当たりの価格") > 0)
    foundM2_kanji = (InStr(1, respText, "1平方メートル当たりの価格") > 0)
    
    Dim foundChangeRate As Boolean
    foundChangeRate = (InStr(1, respText, """変動率""") > 0)
    
    Dim foundLat As Boolean
    foundLat = (InStr(1, respText, """緯度""") > 0)
    
    ' --- 結果をシートに書き出す ---
    Dim wsName As String: wsName = "Xct001_検証ダンプ"
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(wsName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = wsName
    End If
    ws.Cells.Clear
    
    ws.Range("A1").Value = "==== XCT001 検証ダンプ ===="
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "HTTP ステータス"
    ws.Range("B2").Value = status
    ws.Range("A3").Value = "レスポンス長"
    ws.Range("B3").Value = Len(respText)
    
    ws.Range("A5").Value = "==== キー名検索結果 ===="
    ws.Range("A5").Font.Bold = True
    ws.Range("A6").Value = "1㎡ (U+33A1)"
    ws.Range("B6").Value = IIf(foundM2_U33A1, "✓ あり", "なし")
    ws.Range("A7").Value = "1m2 (ASCII)"
    ws.Range("B7").Value = IIf(foundM2_ascii, "✓ あり", "なし")
    ws.Range("A8").Value = "1m^2 (キャレット)"
    ws.Range("B8").Value = IIf(foundM2_caret, "✓ あり", "なし")
    ws.Range("A9").Value = "1平方メートル"
    ws.Range("B9").Value = IIf(foundM2_kanji, "✓ あり", "なし")
    ws.Range("A10").Value = "変動率"
    ws.Range("B10").Value = IIf(foundChangeRate, "✓ あり", "なし")
    ws.Range("A11").Value = "緯度"
    ws.Range("B11").Value = IIf(foundLat, "✓ あり", "なし")
    
    ws.Range("A13").Value = "==== レスポンス本文(先頭 5000 文字) ===="
    ws.Range("A13").Font.Bold = True
    ws.Range("A14").Value = Left$(respText, 5000)
    ws.Range("A14").WrapText = True
    
    ws.Range("A16").Value = "==== 変動率を含む 1 レコード（最初の出現箇所周辺 800 文字）===="
    ws.Range("A16").Font.Bold = True
    Dim posChange As Long: posChange = InStr(1, respText, """変動率""")
    If posChange > 0 Then
        Dim startPos As Long: startPos = posChange - 400
        If startPos < 1 Then startPos = 1
        ws.Range("A17").Value = Mid$(respText, startPos, 800)
        ws.Range("A17").WrapText = True
    Else
        ws.Range("A17").Value = "(変動率キーが見つかりませんでした)"
    End If
    
    ws.Columns("A:B").AutoFit
    ws.Activate
    ws.Range("A1").Select
    
    Application.ScreenUpdating = True
    
    Dim msg As String
    msg = "完了しました。" & vbCrLf & vbCrLf
    msg = msg & "シート「Xct001_検証ダンプ」を確認してください。" & vbCrLf & vbCrLf
    msg = msg & "キー名検索結果:" & vbCrLf
    msg = msg & "  1㎡ (U+33A1): " & IIf(foundM2_U33A1, "✓", "✗") & vbCrLf
    msg = msg & "  1m2: " & IIf(foundM2_ascii, "✓", "✗") & vbCrLf
    msg = msg & "  1m^2: " & IIf(foundM2_caret, "✓", "✗") & vbCrLf
    msg = msg & "  1平方メートル: " & IIf(foundM2_kanji, "✓", "✗") & vbCrLf
    msg = msg & "  変動率: " & IIf(foundChangeRate, "✓", "✗") & vbCrLf
    msg = msg & "  緯度: " & IIf(foundLat, "✓", "✗")
    
    MsgBox msg, vbInformation, "XCT001 検証完了"
End Sub
