# Oculus v0.9 — ローカルAI（PaddleOCR）導入手順

Oculus本体（oculus.html）にはAI連携コードが組み込み済みです。
以下のファイルを `lib/` に置くだけで、摘要・その他列の認識が自動的にローカルAIへ切り替わります。
**未設置でも本体は従来どおり動作します（自動フォールバック）。外部通信は一切発生しません。**

-----

## 1. 配置するファイル（合計 約20MB）

```
oculus.html
lib/
 ├ opencv.js / pdf.min.js / tesseract.min.js / ...   ← 既存
 ├ ort.min.js                  ← 追加① ONNX Runtime Web 本体
 ├ ort-wasm-simd.wasm          ← 追加② 同 WASMバイナリ（SIMD版）
 ├ ort-wasm.wasm               ← 追加③ 同 WASMバイナリ（非SIMD・古いPC用フォールバック）
 └ ppocr/
    ├ rec_japan.onnx           ← 追加④ PaddleOCR日本語認識モデル（ONNX変換済み）
    └ japan_dict.txt           ← 追加⑤ 文字辞書（モデルと必ずペア）
```

## 2. 入手方法

### ①〜③ ONNX Runtime Web

npm レジストリから zip を直接取得できます（要・自宅等のネット環境）。

- <https://registry.npmjs.org/onnxruntime-web/-/onnxruntime-web-1.18.0.tgz>
- 展開して `package/dist/` 内の `ort.min.js`、`ort-wasm-simd.wasm`、`ort-wasm.wasm` の3点を `lib/` 直下へ。

※ バージョンは 1.17〜1.18 系を推奨。3ファイルは**必ず同一バージョン**で揃えること。
※ Oculusは `numThreads=1` で初期化するため、`*-threaded.wasm` 系は不要（簡易HTTPサーバのCOOP/COEP制約を回避する設計）。

### ④ 認識モデル（rec_japan.onnx）

PaddleOCR公式の日本語認識モデルを paddle2onnx で変換します。Python環境のあるPCで:

```bash
pip install paddlepaddle paddle2onnx

# 公式推論モデルの取得・展開
curl -O https://paddleocr.bj.bcebos.com/PP-OCRv3/multilingual/japan_PP-OCRv3_rec_infer.tar
tar xf japan_PP-OCRv3_rec_infer.tar

# ONNXへ変換
paddle2onnx --model_dir japan_PP-OCRv3_rec_infer \
  --model_filename inference.pdmodel \
  --params_filename inference.pdiparams \
  --save_file rec_japan.onnx \
  --opset_version 11
```

できた `rec_japan.onnx`（約10MB）を `lib/ppocr/` へ。

### ⑤ 文字辞書（japan_dict.txt）

PaddleOCRリポジトリの同梱辞書をそのまま使用:

- <https://raw.githubusercontent.com/PaddlePaddle/PaddleOCR/main/ppocr/utils/dict/japan_dict.txt>

`lib/ppocr/` へ。**モデルと辞書は同じ世代のペアであること**（v3モデルにはv3時点の辞書）。

## 3. 動作確認

1. いつも通りローカルサーバ経由で oculus.html を開く
1. 詳細設定の「摘要・その他列はローカルAI(PaddleOCR)で読む」がON（既定ON）であることを確認
1. OCR実行 → ログに `PaddleOCR(ローカルAI) 準備完了: 辞書XXXX字` と出れば有効
1. 結果表で、AIが読んだ摘要セルに緑の「AI」バッジが付く
1. 外部通信ゼロの確認: F12 → Network タブで、リクエスト先がすべて `127.0.0.1`（自分自身）であることを確認

## 4. 動作仕様（本体組み込み済み）

- 対象: 「摘要」「その他」列のみ。数字列（出金/入金/残高）は従来どおりTesseract英語モデル＋残高検算（数字はこちらの方が高速・十分高精度）
- AIには**2値化前の原画像**を渡す（AIモデルは自然画像で学習されているため前処理済み画像より精度が出る）
- AI結果が低信頼の場合はTesseractでも読み、良い方を自動採用
- AI → 摘要辞書補正 → の順で適用（併用で効果が重なる）
- 失敗・未設置時は自動的にTesseractへフォールバック（処理は止まらない）

## 5. トラブルシューティング

|症状                   |原因と対処                                                                         |
|---------------------|------------------------------------------------------------------------------|
|ログに「PaddleOCR未設置のため…」|ファイル配置・ファイル名を確認（`lib/ppocr/rec_japan.onnx`、`japan_dict.txt`、`lib/ort.min.js`） |
|摘要が文字化け・デタラメな漢字      |モデルと辞書の世代不一致。出力クラス数 = 辞書行数 + 2（blank＋空白）になる組合せで揃え直す                           |
|`WebAssembly` 関連エラー  |`ort-wasm-simd.wasm` / `ort-wasm.wasm` が `lib/` 直下にあるか、ort.min.js とバージョン一致かを確認|
|極端に遅い                |古いPCでSIMD非対応の可能性（非SIMD版で動作中）。摘要列だけなので実用上は許容範囲のはず                              |
|AIの方が精度が悪いセルがある      |低信頼時はTesseractと比較選択済み。さらに摘要辞書に正しい語を登録すれば辞書補正が最終防衛線になる                         |

## 6. セキュリティ上の整理（説明用）

- 追加するのはオープンソースの実行ライブラリ（ONNX Runtime / Apache-2.0）と学習済みモデルファイル（PaddleOCR / Apache-2.0）のみ
- 推論はブラウザ内WASMで完結し、画像・テキストとも外部送信なし
- モデルファイルは静的データであり、コード実行能力を持たない