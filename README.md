# Office to PDF

Microsoft PowerPoint・Word・Excelを自動操作し、複数のOfficeファイルをまとめてPDFへ変換するシンプルなデスクトップアプリです。

## 動作条件

- **Windows専用**です。
- PPTXにはMicrosoft **PowerPoint**、DOC・DOCXにはMicrosoft **Word**、XLS・XLSX・XLSM・XLSBにはMicrosoft **Excel**がインストール済みである必要があります。
- PowerPoint、WordおよびExcelのCOM APIを使用します。
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

1. 「Officeファイルを選択」をクリックしてファイルを選ぶか、エクスプローラーから対応ファイルを選択一覧へドラッグ＆ドロップします。PPTX、DOC、DOCX、XLS、XLSX、XLSM、XLSBに対応し、複数ファイルを一度に追加できます。
2. 不要なファイルは一覧で選択し、「選択項目を削除」をクリックするか `Delete` キーを押すと取り除けます。`Ctrl` または `Shift` キーで複数選択できます。
3. Excelの印刷設定を自動調整する場合だけ、「Excelの印刷範囲・用紙方向・倍率をPDF化前に最適化する」にチェックを入れます。初期状態はオフです。
4. 「PDFに変換」をクリックします。
5. 同名のPDFが既にある場合は、ファイルごとに上書きするか選択します。
6. 変換に成功したファイルは「処理が完了したファイル」へ移動します。再変換する場合は対象を選択し、「選択項目を未処理リストへ戻す」をクリックするか `Delete` キーを押します。
7. 変換結果を画面下部のログで確認します。失敗したファイルは未処理リストに残ります。

PDFは元のOfficeファイルと同じフォルダに同じベース名で保存されます（例: `sample.pptx` → `sample.pdf`、`document.docx` → `document.pdf`、`book.xlsx` → `book.pdf`）。元のファイルは読み取り専用で開くため、変更や上書きは行いません。

Excelの印刷設定最適化を選択した場合のみ、PDF化の直前に、表示中の各シートについて既存の手動改ページを解除してから使用セル範囲を印刷範囲に設定します。A4用紙、列数に応じた縦・横方向、横中央配置に加え、使用範囲の実寸と余白から横幅が1ページに収まる数値倍率を自動計算します。完全な空シートと非表示シートは対象外です。この印刷設定は元のExcelファイルには保存されません。

Excelの印刷設定には、PCの既定プリンターやプリンタードライバーに依存して利用できない項目があります。個別の設定に失敗した場合はログへ警告を表示し、その項目だけを省略してPDF変換を続行します。

変換中は本アプリ専用のPowerPoint、WordまたはExcelインスタンスをバックグラウンドで起動します。変換完了後は各ファイルを閉じ、そのインスタンスだけを終了します。ユーザーが別に起動しているOfficeアプリは終了しません。

## エラーが発生した場合

ログの `失敗` の後にエラーが発生した処理段階が表示されます。対応するOfficeアプリで対象ファイルを手動で開けること、ファイルが同期済みであること、および出力先フォルダへ書き込めることを確認してください。

PowerPointまたはWordで `ExportAsFixedFormat` を使用できない場合は、自動的に `SaveAs`（PDF形式）で再試行します。ログに再試行メッセージが表示された後で `成功` と表示されれば、PDFへの変換は正常に完了しています。
