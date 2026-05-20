' ============================================================
' v18.7.2 改修方針 事前検証用ミニテスト
' ============================================================
' 目的:
'   MSXML2.XMLHTTP.6.0 が XCT001 の gzip 自動解凍に有効か確認する
'
' 使い方:
'   1. v18.7.1 を開いた状態で
'   2. VBE で「標準モジュール」を新規追加
'   3. このコードを貼り付け
'   4. Sub TestXct001Gzip の中の YOUR_API_KEY_HERE を APIキーに置換
'   5. F5 で実行 → 結果が MsgBox で出る
'
' 期待される結果:
'   - 「成功: XX件」または「成功: 0件 (HTTP 404)」
'   - レスポンス先頭が読める JSON （[ や { から始まる）
'
' 失敗時:
'   - 文字化けが続く → MSXML2.XMLHTTP.6.0 が使えない環境
'   - エラー発生 → 別の解決策が必要
' ============================================================

Public Sub TestXct001Gzip()
    Const URL As String = "https://www.reinfolib.mlit.go.jp/ex-api/external/XCT001?year=2025&area=13&division=00"
    Const API_KEY As String = "YOUR_API_KEY_HERE"  ' ← ここを実際の APIキーに置換
    
    If API_KEY = "YOUR_API_KEY_HERE" Then
        MsgBox "APIキーを設定してから実行してください。", vbExclamation
        Exit Sub
    End If
    
    Dim http As Object
    On Error Resume Next
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    If Err.Number <> 0 Then
        MsgBox "MSXML2.XMLHTTP.6.0 が使えません: " & Err.Description, vbCritical
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
        Exit Sub
    End If
    On Error GoTo 0
    
    Dim status As Long: status = CLng(http.Status)
    Dim respText As String: respText = CStr(http.responseText)
    
    Dim head100 As String
    head100 = Left$(respText, 200)
    
    Dim hasJsonStart As Boolean
    hasJsonStart = (Left$(LTrim$(respText), 1) = "[" Or Left$(LTrim$(respText), 1) = "{")
    
    Dim recordKeyFound As Boolean
    recordKeyFound = (InStr(1, respText, """価格時点""") > 0 Or InStr(1, respText, """緯度""") > 0)
    
    Dim msg As String
    msg = "==== XCT001 gzip 自動解凍テスト ====" & vbCrLf & vbCrLf
    msg = msg & "HTTP ステータス: " & status & vbCrLf
    msg = msg & "レスポンス長: " & Len(respText) & " 文字" & vbCrLf
    msg = msg & "JSON 開始: " & IIf(hasJsonStart, "✓ あり", "✗ なし") & vbCrLf
    msg = msg & "「価格時点」または「緯度」キー: " & IIf(recordKeyFound, "✓ あり", "✗ なし") & vbCrLf & vbCrLf
    msg = msg & "==== レスポンス先頭 200 文字 ====" & vbCrLf
    msg = msg & head100 & vbCrLf & vbCrLf
    
    If status = 200 And hasJsonStart And recordKeyFound Then
        msg = msg & "判定: ✓ 成功 - MSXML2.XMLHTTP.6.0 で gzip 自動解凍されている"
    ElseIf status = 200 And Not hasJsonStart Then
        msg = msg & "判定: ✗ 失敗 - HTTP 200 だが解凍されていない（文字化け継続）"
    ElseIf status = 404 Then
        msg = msg & "判定: △ データなし (HTTP 404) - 別パラメータで再試行を"
    Else
        msg = msg & "判定: ? 想定外 - HTTP " & status
    End If
    
    MsgBox msg, vbInformation, "XCT001 gzip 検証結果"
    
    Set http = Nothing
End Sub
