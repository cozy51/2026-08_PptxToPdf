# Office to PDF

Microsoft PowerPointまたはWordを自動操作し、複数のPPTX・DOC・DOCXファイルをまとめてPDFへ変換するシンプルなデスクトップアプリです。

## 動作条件

- **Windows専用**です。
- PPTXの変換にはMicrosoft **PowerPoint**、DOC・DOCXの変換にはMicrosoft **Word**がインストール済みである必要があります。
- PowerPointおよびWordのCOM APIを使用します。
- **LibreOfficeは不要**です（使用しません）。
- Python 3.10以降を推奨します。

## インストール方法

コマンドプロンプトまたはPowerShellで、このフォルダへ移動して実行します。

```powershell
python -m venv .venv
.venv\Scripts\activate
python -m pip install -r requirements.txt
```

`tkinter`は通常のWindows版Pythonに含まれています。Pythonは[python.org](https://www.python.org/downloads/windows/)のWindows用インストーラーから導入できます。

## 起動方法

仮想環境を有効にしてから実行します。

```powershell
.venv\Scripts\activate
python main.py
```

## 使い方

1. 「Officeファイルを選択」をクリックしてファイルを選ぶか、エクスプローラーからPPTX・DOC・DOCXファイルを選択一覧へドラッグ＆ドロップします。複数ファイルを一度にドロップできます。
2. 「PDFに変換」をクリックします。
3. 同名のPDFが既にある場合は、ファイルごとに上書きするか選択します。
4. 変換結果を画面下部のログで確認します。

PDFは元のOfficeファイルと同じフォルダに同じベース名で保存されます（例: `sample.pptx` → `sample.pdf`、`document.docx` → `document.pdf`）。元のファイルは読み取り専用で開くため、変更や上書きは行いません。

変換中は本アプリ専用のPowerPointまたはWordインスタンスをバックグラウンドで起動します。変換完了後は各ファイルを閉じ、そのインスタンスだけを終了します。ユーザーが別に起動しているPowerPointやWordは終了しません。

## エラーが発生した場合

ログの `失敗` の後にエラーが発生した処理段階が表示されます。対応するPowerPointまたはWordで対象ファイルを手動で開けること、ファイルが同期済みであること、および出力先フォルダへ書き込めることを確認してください。

Officeまたはpywin32の環境によって `ExportAsFixedFormat` を使用できない場合は、自動的に `SaveAs`（PDF形式）で再試行します。ログに再試行メッセージが表示された後で `成功` と表示されれば、PDFへの変換は正常に完了しています。
