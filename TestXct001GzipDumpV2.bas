' ============================================================
' XCT001 レスポンス保存版 ミニテスト v2（エラー対策版）
' ============================================================
' 目的:
'   gzip 自動解凍された JSON の中身を MsgBox とイミディエイトウィンドウに
'   出して、「1㎡」「変動率」などの具体的なキー名を確認する
'
' 変更点:
'   v1 でシート追加時のエラー1004 を回避するため、シート操作を最小化。
'   結果は MsgBox とデバッグ Print のみ。
'
' 使い方:
'   1. v18.7.1 を開いた状態で
'   2. VBE で「標準モジュール」を新規追加
'   3. このコードを貼り付け
'   4. Sub TestXct001GzipDumpV2 の中の YOUR_API_KEY_HERE を APIキーに置換
'   5. F5 で実行
'   6. MsgBox に検索結果が表示される
'   7. もっと詳しく見たい場合は VBE の「表示」→「イミディエイトウィンドウ」を開く
' ============================================================

Public Sub TestXct001GzipDumpV2()
    Const URL As String = "https://www.reinfolib.mlit.go.jp/ex-api/external/XCT001?year=2025&area=13&division=00"
    Const API_KEY As String = "YOUR_API_KEY_HERE"  ' ← ここを実際の APIキーに置換
    
    If API_KEY = "YOUR_API_KEY_HERE" Then
        MsgBox "APIキーを設定してから実行してください。", vbExclamation
        Exit Sub
    End If
    
    ' --- HTTP 取得 ---
    Dim http As Object
    On Error Resume Next
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    If Err.Number <> 0 Then
        MsgBox "MSXML2.XMLHTTP.6.0 が使えません: " & Err.Description, vbCritical
        Exit Sub
    End If
    Err.Clear
    On Error GoTo 0
    
    http.Open "GET", URL, False
    http.setRequestHeader "Ocp-Apim-Subscription-Key", API_KEY
    http.setRequestHeader "Accept", "application/json"
    
    On Error Resume Next
    http.Send
    If Err.Number <> 0 Then
        MsgBox "通信エラー: " & Err.Description, vbCritical
        Set http = Nothing
        Exit Sub
    End If
    Err.Clear
    On Error GoTo 0
    
    Dim status As Long: status = CLng(http.Status)
    Dim respText As String: respText = CStr(http.responseText)
    Set http = Nothing
    
    ' --- 検索 ---
    Dim foundM2_U33A1 As Boolean
    Dim foundM2_ascii As Boolean
    Dim foundM2_caret As Boolean
    Dim foundM2_kanji As Boolean
    Dim foundChangeRate As Boolean
    Dim foundLat As Boolean
    
    foundM2_U33A1 = (InStr(1, respText, "1" & ChrW$(&H33A1) & "当たりの価格") > 0)
    foundM2_ascii = (InStr(1, respText, "1m2当たりの価格") > 0)
    foundM2_caret = (InStr(1, respText, "1m^2当たりの価格") > 0)
    foundM2_kanji = (InStr(1, respText, "1平方メートル当たりの価格") > 0)
    foundChangeRate = (InStr(1, respText, """変動率""") > 0)
    foundLat = (InStr(1, respText, """緯度""") > 0)
    
    ' --- イミディエイトウィンドウへ出力（後でデバッグウィンドウで見られる）---
    Debug.Print "==== XCT001 検証ダンプ v2 ===="
    Debug.Print "HTTP ステータス: " & status
    Debug.Print "レスポンス長: " & Len(respText) & " 文字"
    Debug.Print "1㎡ (U+33A1): " & foundM2_U33A1
    Debug.Print "1m2 (ASCII): " & foundM2_ascii
    Debug.Print "1m^2 (キャレット): " & foundM2_caret
    Debug.Print "1平方メートル: " & foundM2_kanji
    Debug.Print "変動率: " & foundChangeRate
    Debug.Print "緯度: " & foundLat
    Debug.Print "---- レスポンス先頭 1500 文字 ----"
    Debug.Print Left$(respText, 1500)
    Debug.Print "---- 変動率含む周辺 600 文字 ----"
    Dim posChange As Long: posChange = InStr(1, respText, """変動率""")
    If posChange > 0 Then
        Dim startPos As Long: startPos = posChange - 200
        If startPos < 1 Then startPos = 1
        Debug.Print Mid$(respText, startPos, 600)
    Else
        Debug.Print "(変動率キーが見つかりませんでした)"
    End If
    Debug.Print "==== 終了 ===="
    
    ' --- MsgBox で結果サマリ ---
    Dim msg As String
    msg = "==== XCT001 検証結果 v2 ====" & vbCrLf & vbCrLf
    msg = msg & "HTTP: " & status & vbCrLf
    msg = msg & "レスポンス長: " & Len(respText) & " 文字" & vbCrLf & vbCrLf
    msg = msg & "==== ㎡記号 4 パターン ====" & vbCrLf
    msg = msg & "  1㎡ (U+33A1): " & IIf(foundM2_U33A1, "あり", "なし") & vbCrLf
    msg = msg & "  1m2: " & IIf(foundM2_ascii, "あり", "なし") & vbCrLf
    msg = msg & "  1m^2: " & IIf(foundM2_caret, "あり", "なし") & vbCrLf
    msg = msg & "  1平方メートル: " & IIf(foundM2_kanji, "あり", "なし") & vbCrLf & vbCrLf
    msg = msg & "==== その他 ====" & vbCrLf
    msg = msg & "  変動率: " & IIf(foundChangeRate, "あり", "なし") & vbCrLf
    msg = msg & "  緯度: " & IIf(foundLat, "あり", "なし") & vbCrLf & vbCrLf
    msg = msg & "詳細はイミディエイトウィンドウへ出力済み" & vbCrLf
    msg = msg & "(VBE で Ctrl+G または 表示→イミディエイトウィンドウ)"
    
    MsgBox msg, vbInformation, "XCT001 検証結果 v2"
End Sub
