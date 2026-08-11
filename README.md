# PPTX to PDF

Microsoft PowerPointを自動操作し、複数のPPTXファイルをまとめてPDFへ変換するシンプルなデスクトップアプリです。

## 動作条件

- **Windows専用**です。
- Microsoft Office / **PowerPointがインストール済み**である必要があります。
- PowerPointのCOM APIを使用します。
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

1. 「PPTXファイルを選択」をクリックし、変換するファイルを1つ以上選択します。
2. 「PDFに変換」をクリックします。
3. 同名のPDFが既にある場合は、ファイルごとに上書きするか選択します。
4. 変換結果を画面下部のログで確認します。

PDFは元のPPTXと同じフォルダに同じベース名で保存されます（例: `sample.pptx` → `sample.pdf`）。元のPPTXは読み取り専用で開くため、変更や上書きは行いません。

変換中は本アプリ専用のPowerPointインスタンスをバックグラウンドで起動します。変換完了後は各プレゼンテーションを閉じ、そのインスタンスだけを終了します。ユーザーが別に起動しているPowerPointは終了しません。

## エラーが発生した場合

ログの `失敗` の後に、エラーが発生した処理段階（例: `PowerPointで開く`、`PDFとして書き出す`）が表示されます。PowerPointで対象ファイルを手動で開けること、ファイルが同期済みであること、および出力先フォルダへ書き込めることを確認してください。

PowerPointまたはpywin32の環境によって `ExportAsFixedFormat` を使用できない場合は、自動的にPowerPointの `SaveAs`（PDF形式）で再試行します。ログに再試行メッセージが表示された後で `成功` と表示されれば、PDFへの変換は正常に完了しています。
