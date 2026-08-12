# Office to PDF

Microsoft PowerPoint・Word・Excelを自動操作し、複数のOfficeファイルをまとめてPDFへ変換するシンプルなデスクトップアプリです。

## 動作条件

- **Windows専用**です。
- PPTXにはMicrosoft **PowerPoint**、DOC・DOCXにはMicrosoft **Word**、XLS・XLSX・XLSM・XLSBにはMicrosoft **Excel**がインストール済みである必要があります。
- PowerPoint、WordおよびExcelのCOM APIを使用します。
- **LibreOfficeは不要**です（使用しません）。
- ソースから起動・ビルドするPCではPython 3.10以降を推奨します。配布先PCにPythonは不要です。

## Windows用exeをビルドする

ビルドは**Windows 10/11上**で行ってください。ビルドPCに必要なものは次のとおりです。

- Python 3.10以降（python.org版を推奨）
- 依存ライブラリを取得するためのインターネット接続

コマンドプロンプトまたはPowerShellでリポジトリ直下へ移動し、ビルド専用の仮想環境を作ってPyInstallerを実行します。

```powershell
python -m venv .venv-build
.venv-build\Scripts\activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
pyinstaller --noconfirm --clean OfficePdfConverter.spec
```

成功すると、配布用ファイルが次の場所に生成されます。

```text
dist\OfficePdfConverter.exe
```

`--clean` は前回のビルドキャッシュを破棄します。作り直す場合は `dist` と `build` を削除してから実行してください。手順の詳細は [インストールマニュアル](docs/install_manual.md) にも記載しています。

### GitHub Actionsでビルドする（Windows PCが無い場合）

手元にWindows環境が無くてもexeを入手できるよう、GitHubのWindowsランナーでビルドするワークフロー [`.github/workflows/build-exe.yml`](.github/workflows/build-exe.yml) を用意しています。

1. GitHubのリポジトリで「Actions」タブを開きます。
2. 左側の一覧から「Build Windows exe」を選びます。
3. 「Run workflow」をクリックし、ビルドしたいブランチを選んで実行します。
4. 数分で完了します。実行結果のページ下部「Artifacts」から `OfficePdfConverter` をダウンロードすると、ZIPの中にexeとインストールマニュアルのPDFが入っています。

`v1.0` のような `v` で始まるタグをプッシュした場合は、同じ成果物がGitHubのリリースへ自動で添付されます。

PyInstallerはクロスコンパイルに対応していないため、Windows用exeはWindows上でしかビルドできません。このワークフローも `windows-latest` のランナーで実行しています。

### spec ファイルについて

`OfficePdfConverter.spec` は1ファイル形式（one-file）、ウィンドウ形式（コンソール非表示）でビルドします。`pywin32`、`tkinterdnd2`、tkdndのDLL/Tclデータ、アプリアイコンなど、実行に必要なPython側のライブラリとリソースはexeへ格納されます。

ビルドで生成される `.venv-build`、`build`、`dist` の3フォルダーは `.gitignore` で除外済みです。これらにはexeやicoといったバイナリが含まれ、バイナリファイルを扱えない環境では差分表示がエラーになるため、リポジトリへコミットしないでください。

### 配布方法

`dist\OfficePdfConverter.exe` **1ファイルだけ**を配布先のWindows 10/11 PCへコピーしてください。インストーラーやPython環境は不要で、exeをダブルクリックするとコンソール画面を出さずにアプリのGUIだけが起動します。初回起動時は、ウイルス対策ソフトによる確認のため少し時間がかかる場合があります。

Microsoft Office本体はexeへ同梱されません。変換に使用する形式に応じて、配布先PCにデスクトップ版のMicrosoft PowerPoint、Word、Excelがインストールされ、起動・ライセンス認証済みである必要があります。本アプリは配布先PCにあるOfficeをCOM APIで自動操作します。

配布前には、OfficeがインストールされたWindows 10/11 PCでexeを起動し、各対象形式を実際にPDF変換できることを確認してください。組織外へ配布する場合は、必要に応じてexeへコード署名を行うと、Windowsの発行元確認を受けやすくなります。

### インストールマニュアル（PDF）

導入手順をまとめたマニュアルを同梱しています。本文は [`docs/install_manual.md`](docs/install_manual.md) にテキストで管理し、配布用のPDFは次のコマンドで生成します。

```powershell
python -m pip install reportlab
python tools\make_manual.py
```

`docs\インストールマニュアル.pdf` が作られます。PDFはバイナリのためリポジトリには含めていません。マニュアルは「exeを使う人向けの導入手順」と「exeをビルドする人向けの環境構築手順」の2部構成で、SmartScreenの警告への対処やよくあるトラブルも記載しています。

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

黄色の一覧がこれから変換するファイル、緑色の一覧が変換完了ファイルです。

PDFは元のOfficeファイルと同じフォルダに同じベース名で保存されます（例: `sample.pptx` → `sample.pdf`、`document.docx` → `document.pdf`、`book.xlsx` → `book.pdf`）。元のファイルは読み取り専用で開くため、変更や上書きは行いません。

Excelの印刷設定最適化を選択した場合のみ、PDF化の直前に、表示中の各シートについて既存の手動改ページを解除してから使用セル範囲を印刷範囲に設定します。右端セルの文字がセル幅からはみ出す場合は、全角・半角文字数とフォントサイズから表示幅を概算し、文字が切れないよう印刷範囲を右へ拡張します。元から横向きのシートを縦向きへ変更することはありません。縦向きのシートは、使用範囲の実幅がA4縦の印刷可能幅を超える場合だけ横向きへ変更します。そのうえで横中央配置と、使用範囲の実寸・余白から算出した数値倍率を適用します。完全な空シートと非表示シートは対象外です。この印刷設定は元のExcelファイルには保存されません。

Excelの印刷設定には、PCの既定プリンターやプリンタードライバーに依存して利用できない項目があります。個別の設定に失敗した場合はログへ警告を表示し、その項目だけを省略してPDF変換を続行します。

変換中は本アプリ専用のPowerPoint、WordまたはExcelインスタンスをバックグラウンドで起動します。変換完了後は各ファイルを閉じ、そのインスタンスだけを終了します。ユーザーが別に起動しているOfficeアプリは終了しません。

## アプリアイコン

ウィンドウとタスクバーには専用アイコンが表示されます。メッセージダイアログなど、後から開くウィンドウにも同じアイコンが適用されます。

アイコンは `assets` フォルダーにBase64テキスト（`app_icon_16.png.b64` など4サイズ）として同梱しています。バイナリファイルを扱えない環境でもリポジトリをそのまま利用できるようにするためで、起動時に復号して読み込みます。読み込みに失敗した場合はアイコン設定だけを省略し、アプリは通常どおり起動します。

好みのアイコンへ差し替える場合は、PNG画像を `assets\app_icon.png` という名前で置いてください。このファイルがあるときは同梱アイコンより優先されます。exeで実行している場合は、exeと同じフォルダーに `assets` フォルダーを作ってその中へ置きます。

同梱アイコンの絵柄そのものを描き変える場合は、`tools/make_icon.py` の図形定義を編集して次を実行します。標準ライブラリだけで動作します。

```powershell
python tools\make_icon.py
```

## Officeがインストールされていない場合

本アプリはOffice本体を同梱せず、PCにインストール済みのPowerPoint・Word・ExcelをCOM APIで操作してPDFを作成します。そのため、必要なOfficeが無い場合は次のように案内します。

- 起動直後に、使用できるOfficeと見つからないOfficeをログへ表示します。
- どのOfficeも見つからない場合は、起動時に案内ダイアログを表示します。
- 「PDFに変換」を押したとき、選択中のファイルに必要なOfficeが無ければ、変換を始める前に対象のOffice名と確認事項をダイアログで表示して中止します。
- Officeの登録は残っているのに起動できない場合は、COMの「無効なクラス文字列」といったメッセージではなく、どのOfficeが使えないのかをログとダイアログで表示します。

インストール済みかどうかは、レジストリのProgID（`PowerPoint.Application` など）の登録有無で判定します。Officeを起動せずに確認できるため、変換前のチェックが速く終わります。

## エラーが発生した場合

ログの `失敗` の後にエラーが発生した処理段階が表示されます。対応するOfficeアプリで対象ファイルを手動で開けること、ファイルが同期済みであること、および出力先フォルダへ書き込めることを確認してください。

PowerPointまたはWordで `ExportAsFixedFormat` を使用できない場合は、自動的に `SaveAs`（PDF形式）で再試行します。ログに再試行メッセージが表示された後で `成功` と表示されれば、PDFへの変換は正常に完了しています。
