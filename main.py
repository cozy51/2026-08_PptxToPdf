from __future__ import annotations

import gc
import queue
import threading
import tkinter as tk
import unicodedata
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import pythoncom
import win32com.client
from tkinterdnd2 import DND_FILES, TkinterDnD


# COMへ渡す値はPythonのboolではなく、OfficeのMsoTriStateと同じ整数値を
# 明示的に使用する。環境によっては動的COMディスパッチがboolやキーワード
# 引数を正しく変換できず、"The Python instance can not be converted to a COM
# object" になるため。
MSO_TRUE = -1
MSO_FALSE = 0
PDF_FORMAT = 2  # PowerPoint: ppFixedFormatTypePDF
SAVE_AS_PDF = 32  # PowerPoint: ppSaveAsPDF
WORD_PDF_FORMAT = 17  # Word: wdExportFormatPDF / wdFormatPDF
EXCEL_PDF_FORMAT = 0  # Excel: xlTypePDF
EXCEL_PORTRAIT = 1  # Excel: xlPortrait
EXCEL_LANDSCAPE = 2  # Excel: xlLandscape
EXCEL_A4 = 9  # Excel: xlPaperA4
EXCEL_VISIBLE = -1  # Excel: xlSheetVisible
LANDSCAPE_COLUMN_THRESHOLD = 8
EXCEL_A4_PORTRAIT_WIDTH_POINTS = 595.28
EXCEL_A4_LANDSCAPE_WIDTH_POINTS = 841.89
EXCEL_DEFAULT_MARGIN_POINTS = 36.0
EXCEL_MIN_ZOOM = 10
EXCEL_MAX_OVERFLOW_COLUMNS = 50
EXCEL_MAX_OVERFLOW_ROWS = 10000
SUPPORTED_EXTENSIONS = {
    ".pptx",
    ".doc",
    ".docx",
    ".xls",
    ".xlsx",
    ".xlsm",
    ".xlsb",
}


class PptxToPdfApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Office to PDF")
        self.root.geometry("760x680")
        self.root.minsize(600, 560)

        self.selected_files: list[Path] = []
        self.completed_files: list[Path] = []
        self.events: queue.Queue[tuple[str, str]] = queue.Queue()
        self.converting = False
        self.optimize_excel_var = tk.BooleanVar(value=False)

        self._build_ui()
        self._configure_drag_and_drop()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self) -> None:
        frame = ttk.Frame(self.root, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)
        frame.rowconfigure(2, weight=1)
        frame.rowconfigure(6, weight=1)

        controls = ttk.Frame(frame)
        controls.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        self.select_button = ttk.Button(
            controls, text="Officeファイルを選択", command=self._select_files
        )
        self.select_button.pack(side=tk.LEFT)
        ttk.Label(
            controls, text="  またはPowerPoint・Word・Excelファイルを一覧へドロップ"
        ).pack(side=tk.LEFT)

        files_frame = ttk.LabelFrame(frame, text="選択したファイル", padding=6)
        files_frame.grid(row=1, column=0, sticky="nsew")
        files_frame.columnconfigure(0, weight=1)
        files_frame.rowconfigure(0, weight=1)

        self.file_list = tk.Listbox(files_frame, selectmode=tk.EXTENDED)
        self.file_list.grid(row=0, column=0, sticky="nsew")
        files_scroll = ttk.Scrollbar(
            files_frame, orient=tk.VERTICAL, command=self.file_list.yview
        )
        files_scroll.grid(row=0, column=1, sticky="ns")
        self.file_list.configure(yscrollcommand=files_scroll.set)
        self.file_list.bind("<Delete>", self._remove_selected_files)

        self.remove_button = ttk.Button(
            files_frame, text="選択項目を削除", command=self._remove_selected_files
        )
        self.remove_button.grid(row=1, column=0, sticky=tk.W, pady=(6, 0))

        completed_frame = ttk.LabelFrame(frame, text="処理が完了したファイル", padding=6)
        completed_frame.grid(row=2, column=0, sticky="nsew", pady=(8, 0))
        completed_frame.columnconfigure(0, weight=1)
        completed_frame.rowconfigure(0, weight=1)

        self.completed_list = tk.Listbox(completed_frame, selectmode=tk.EXTENDED)
        self.completed_list.grid(row=0, column=0, sticky="nsew")
        completed_scroll = ttk.Scrollbar(
            completed_frame, orient=tk.VERTICAL, command=self.completed_list.yview
        )
        completed_scroll.grid(row=0, column=1, sticky="ns")
        self.completed_list.configure(yscrollcommand=completed_scroll.set)
        self.completed_list.bind("<Delete>", self._restore_completed_files)

        self.restore_button = ttk.Button(
            completed_frame,
            text="選択項目を未処理リストへ戻す",
            command=self._restore_completed_files,
        )
        self.restore_button.grid(row=1, column=0, sticky=tk.W, pady=(6, 0))

        self.convert_button = ttk.Button(
            frame, text="PDFに変換", command=self._start_conversion
        )
        self.convert_button.grid(row=3, column=0, sticky=tk.W, pady=(10, 4))

        self.optimize_excel_check = ttk.Checkbutton(
            frame,
            text="Excelの印刷範囲・用紙方向・倍率をPDF化前に最適化する",
            variable=self.optimize_excel_var,
        )
        self.optimize_excel_check.grid(row=4, column=0, sticky=tk.W, pady=(0, 10))

        ttk.Label(frame, text="処理結果ログ").grid(row=5, column=0, sticky=tk.W)
        log_frame = ttk.Frame(frame)
        log_frame.grid(row=6, column=0, sticky="nsew", pady=(4, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log = tk.Text(log_frame, height=10, state=tk.DISABLED, wrap=tk.WORD)
        self.log.grid(row=0, column=0, sticky="nsew")
        log_scroll = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log.yview)
        log_scroll.grid(row=0, column=1, sticky="ns")
        self.log.configure(yscrollcommand=log_scroll.set)

    def _configure_drag_and_drop(self) -> None:
        self.file_list.drop_target_register(DND_FILES)
        self.file_list.dnd_bind("<<Drop>>", self._handle_drop)

    def _handle_drop(self, event: tk.Event) -> str:
        if self.converting:
            self._write_log("変換中はファイルを追加できません。")
            return "break"

        # Tclのsplitlistを使用すると、空白や日本語を含み波括弧で囲まれた
        # Windowsパスも、文字列を手作業で分割せず安全に取得できる。
        dropped_names = self.root.tk.splitlist(event.data)
        office_files = [
            Path(name).resolve()
            for name in dropped_names
            if Path(name).is_file()
            and Path(name).suffix.lower() in SUPPORTED_EXTENSIONS
        ]

        if not office_files:
            self._write_log("ドロップされた項目に対応するOfficeファイルがありません。")
            return "break"

        self._add_files(office_files)
        self._write_log(f"ドラッグ＆ドロップで{len(office_files)}件を追加しました。")
        return "break"

    def _add_files(self, paths: list[Path], *, replace: bool = False) -> None:
        if replace:
            self.selected_files = []

        known_paths = {
            str(path).casefold()
            for path in self.selected_files + self.completed_files
        }
        for path in paths:
            path_key = str(path).casefold()
            if path_key not in known_paths:
                self.selected_files.append(path)
                known_paths.add(path_key)

        self.file_list.delete(0, tk.END)
        for path in self.selected_files:
            self.file_list.insert(tk.END, str(path))

    def _remove_selected_files(self, _event: tk.Event | None = None) -> str:
        if self.converting:
            self._write_log("変換中は選択項目を削除できません。")
            return "break"

        selected_indices = list(self.file_list.curselection())
        if not selected_indices:
            self._write_log("削除するファイルを一覧から選択してください。")
            return "break"

        for index in reversed(selected_indices):
            del self.selected_files[index]
            self.file_list.delete(index)

        self._write_log(f"選択リストから{len(selected_indices)}件を削除しました。")
        return "break"

    def _restore_completed_files(self, _event: tk.Event | None = None) -> str:
        if self.converting:
            self._write_log("変換中は完了したファイルを戻せません。")
            return "break"

        selected_indices = list(self.completed_list.curselection())
        if not selected_indices:
            self._write_log("未処理リストへ戻すファイルを選択してください。")
            return "break"

        restored = [self.completed_files[index] for index in selected_indices]
        for index in reversed(selected_indices):
            del self.completed_files[index]
        self.selected_files.extend(restored)
        self._refresh_file_lists()
        self._write_log(f"未処理リストへ{len(restored)}件を戻しました。")
        return "break"

    def _mark_completed(self, source_path: Path) -> None:
        source_key = str(source_path).casefold()
        self.selected_files = [
            path for path in self.selected_files if str(path).casefold() != source_key
        ]
        if all(str(path).casefold() != source_key for path in self.completed_files):
            self.completed_files.append(source_path)
        self._refresh_file_lists()

    def _refresh_file_lists(self) -> None:
        self.file_list.delete(0, tk.END)
        for path in self.selected_files:
            self.file_list.insert(tk.END, str(path))
        self.completed_list.delete(0, tk.END)
        for path in self.completed_files:
            self.completed_list.insert(tk.END, str(path))

    def _select_files(self) -> None:
        names = filedialog.askopenfilenames(
            title="Officeファイルを選択",
            filetypes=[
                (
                    "対応するOfficeファイル",
                    "*.pptx *.doc *.docx *.xls *.xlsx *.xlsm *.xlsb",
                ),
                ("PowerPoint プレゼンテーション", "*.pptx"),
                ("Word 文書", "*.doc *.docx"),
                ("Excel ブック", "*.xls *.xlsx *.xlsm *.xlsb"),
            ],
        )
        if not names:
            return

        self._add_files([Path(name).resolve() for name in names], replace=True)
        self._write_log(f"{len(self.selected_files)} 件のファイルを選択しました。")

    def _start_conversion(self) -> None:
        if not self.selected_files:
            messagebox.showwarning(
                "ファイル未選択", "PowerPoint・Word・Excelファイルを選択してください。"
            )
            return

        files_to_convert: list[Path] = []
        for pptx_path in self.selected_files:
            pdf_path = pptx_path.with_suffix(".pdf")
            if pdf_path.exists():
                answer = messagebox.askyesnocancel(
                    "上書き確認",
                    f"次のPDFは既に存在します。上書きしますか？\n\n{pdf_path}",
                )
                if answer is None:
                    self._write_log("変換をキャンセルしました。")
                    return
                if not answer:
                    self._write_log(f"スキップ: {pptx_path.name}")
                    continue
            files_to_convert.append(pptx_path)

        if not files_to_convert:
            self._write_log("変換対象のファイルがありません。")
            return

        self.converting = True
        self.select_button.configure(state=tk.DISABLED)
        self.remove_button.configure(state=tk.DISABLED)
        self.restore_button.configure(state=tk.DISABLED)
        self.optimize_excel_check.configure(state=tk.DISABLED)
        self.convert_button.configure(state=tk.DISABLED)
        self._write_log("変換を開始します。")

        worker = threading.Thread(
            target=self._convert_files,
            args=(files_to_convert, self.optimize_excel_var.get()),
            daemon=False,
        )
        worker.start()
        self.root.after(100, self._process_events)

    def _convert_files(self, paths: list[Path], optimize_excel: bool) -> None:
        powerpoint = None
        word = None
        excel = None
        pythoncom.CoInitialize()
        try:
            for source_path in paths:
                current_step = "ファイルの確認"
                try:
                    if not source_path.is_file():
                        raise FileNotFoundError("ファイルが見つかりません")

                    extension = source_path.suffix.lower()
                    if extension == ".pptx":
                        current_step = "PowerPointを起動"
                        if powerpoint is None:
                            powerpoint = win32com.client.DispatchEx(
                                "PowerPoint.Application"
                            )
                        self._convert_powerpoint_file(powerpoint, source_path)
                    elif extension in {".doc", ".docx"}:
                        current_step = "Wordを起動"
                        if word is None:
                            word = win32com.client.DispatchEx("Word.Application")
                            word.DisplayAlerts = MSO_FALSE
                        self._convert_word_file(word, source_path)
                    elif extension in {".xls", ".xlsx", ".xlsm", ".xlsb"}:
                        current_step = "Excelを起動"
                        if excel is None:
                            excel = win32com.client.DispatchEx("Excel.Application")
                            excel.DisplayAlerts = MSO_FALSE
                        self._convert_excel_file(excel, source_path, optimize_excel)
                    else:
                        raise ValueError("対応していない拡張子です")
                except Exception as exc:
                    self.events.put(
                        (
                            "log",
                            f"失敗 [{current_step}]: {source_path.name} ({exc})",
                        )
                    )
        finally:
            if excel is not None:
                try:
                    excel.Quit()
                except Exception as exc:
                    self.events.put(("log", f"Excelの終了時にエラーが発生しました: {exc}"))
                finally:
                    excel = None
            if word is not None:
                try:
                    word.Quit()
                except Exception as exc:
                    self.events.put(("log", f"Wordの終了時にエラーが発生しました: {exc}"))
                finally:
                    word = None
            if powerpoint is not None:
                try:
                    powerpoint.Quit()
                except Exception as exc:
                    self.events.put(("log", f"PowerPointの終了時にエラーが発生しました: {exc}"))
                finally:
                    powerpoint = None
            gc.collect()
            pythoncom.CoUninitialize()
            self.events.put(("done", "すべての処理が終了しました。"))

    def _convert_powerpoint_file(self, powerpoint: object, source_path: Path) -> None:
        presentation = None
        pdf_path = source_path.with_suffix(".pdf")
        try:
            # COMメソッドにはキーワード引数やPythonのboolを渡さない。
            presentation = powerpoint.Presentations.Open(
                str(source_path), MSO_TRUE, MSO_FALSE, MSO_FALSE
            )
            try:
                presentation.ExportAsFixedFormat(str(pdf_path), PDF_FORMAT)
            except Exception as export_error:
                self._log_export_retry(export_error)
                try:
                    presentation.SaveAs(str(pdf_path), SAVE_AS_PDF)
                except Exception as save_as_error:
                    self._raise_export_error(export_error, save_as_error)
            self.events.put(("success", str(source_path)))
        finally:
            if presentation is not None:
                try:
                    presentation.Close()
                except Exception as exc:
                    self.events.put(
                        ("log", f"警告: {source_path.name} を閉じられませんでした ({exc})")
                    )

    def _convert_word_file(self, word: object, source_path: Path) -> None:
        document = None
        pdf_path = source_path.with_suffix(".pdf")
        try:
            # ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False
            document = word.Documents.Open(
                str(source_path), MSO_FALSE, MSO_TRUE, MSO_FALSE
            )
            try:
                document.ExportAsFixedFormat(str(pdf_path), WORD_PDF_FORMAT)
            except Exception as export_error:
                self._log_export_retry(export_error)
                try:
                    document.SaveAs2(str(pdf_path), WORD_PDF_FORMAT)
                except Exception as save_as_error:
                    self._raise_export_error(export_error, save_as_error)
            self.events.put(("success", str(source_path)))
        finally:
            if document is not None:
                try:
                    document.Close(MSO_FALSE)
                except Exception as exc:
                    self.events.put(
                        ("log", f"警告: {source_path.name} を閉じられませんでした ({exc})")
                    )

    def _convert_excel_file(
        self, excel: object, source_path: Path, optimize_print_settings: bool
    ) -> None:
        workbook = None
        pdf_path = source_path.with_suffix(".pdf")
        try:
            # UpdateLinks=False, ReadOnly=True。元のブックは変更しない。
            workbook = excel.Workbooks.Open(str(source_path), MSO_FALSE, MSO_TRUE)
            if optimize_print_settings:
                self._optimize_excel_print_settings(workbook)
            else:
                self.events.put(("log", "Excelの印刷設定の最適化を省略しました。"))
            # ExcelはPowerPoint/Wordと引数の順序が異なり、形式が先になる。
            workbook.ExportAsFixedFormat(EXCEL_PDF_FORMAT, str(pdf_path))
            self.events.put(("success", str(source_path)))
        finally:
            if workbook is not None:
                try:
                    workbook.Close(MSO_FALSE)
                except Exception as exc:
                    self.events.put(
                        ("log", f"警告: {source_path.name} を閉じられませんでした ({exc})")
                    )

    def _optimize_excel_print_settings(self, workbook: object) -> None:
        optimized_sheets = 0
        for worksheet in workbook.Worksheets:
            try:
                if worksheet.Visible != EXCEL_VISIBLE:
                    continue

                used_range = worksheet.UsedRange
                # 値がない完全な空シートはPDFの印刷対象に追加しない。
                if used_range.Count == 1 and used_range.Value2 is None:
                    continue

                page_setup = worksheet.PageSetup
                try:
                    # 既存の手動改ページが残っているとFitToPagesWideより優先され、
                    # 横方向が複数ページのままになるため、先に解除する。
                    worksheet.ResetAllPageBreaks()
                except Exception as exc:
                    self.events.put(
                        (
                            "log",
                            f"警告: シート「{worksheet.Name}」の既存改ページを解除"
                            f"できませんでした。設定可能な範囲で続けます ({exc})",
                        )
                    )
                orientation = (
                    EXCEL_LANDSCAPE
                    if used_range.Columns.Count > LANDSCAPE_COLUMN_THRESHOLD
                    else EXCEL_PORTRAIT
                )
                print_area, print_width = self._calculate_excel_print_area(
                    worksheet, used_range
                )
                zoom = self._calculate_excel_zoom(
                    print_width, page_setup, orientation, worksheet.Name
                )
                settings = (
                    ("PrintArea", print_area, "印刷範囲"),
                    ("PaperSize", EXCEL_A4, "用紙サイズ(A4)"),
                    ("Orientation", orientation, "印刷方向"),
                    # Excel環境によってZoom=FalseとFitToPages*が1004エラーに
                    # なるため、使用範囲の実寸から数値の倍率を計算する。
                    ("Zoom", zoom, f"拡大縮小率({zoom}%)"),
                    ("CenterHorizontally", True, "横中央配置"),
                )
                applied_settings = sum(
                    self._try_set_excel_page_property(
                        page_setup, property_name, value, worksheet.Name, label
                    )
                    for property_name, value, label in settings
                )
                if applied_settings:
                    optimized_sheets += 1
            except Exception as exc:
                # PageSetupの一部はプリンタードライバーやExcelのバージョンに
                # 依存する。最適化できなくてもPDF変換そのものは続行する。
                self.events.put(
                    (
                        "log",
                        f"警告: シート「{worksheet.Name}」の印刷設定を取得できないため、"
                        f"設定可能な範囲で変換を続けます ({exc})",
                    )
                )

        self.events.put(
            ("log", f"Excelの印刷設定を最適化しました（{optimized_sheets}シート）。")
        )

    def _try_set_excel_page_property(
        self,
        page_setup: object,
        property_name: str,
        value: object,
        sheet_name: str,
        label: str,
    ) -> int:
        try:
            setattr(page_setup, property_name, value)
            return 1
        except Exception as exc:
            self.events.put(
                (
                    "log",
                    f"警告: シート「{sheet_name}」の{label}を設定できませんでした。"
                    f"この設定を省略して変換を続けます ({exc})",
                )
            )
            return 0

    def _calculate_excel_zoom(
        self,
        content_width: float,
        page_setup: object,
        orientation: int,
        sheet_name: str,
    ) -> int:
        page_width = (
            EXCEL_A4_LANDSCAPE_WIDTH_POINTS
            if orientation == EXCEL_LANDSCAPE
            else EXCEL_A4_PORTRAIT_WIDTH_POINTS
        )
        try:
            left_margin = float(page_setup.LeftMargin)
            right_margin = float(page_setup.RightMargin)
        except Exception:
            left_margin = EXCEL_DEFAULT_MARGIN_POINTS
            right_margin = EXCEL_DEFAULT_MARGIN_POINTS

        printable_width = max(page_width - left_margin - right_margin, 1.0)
        # 丸め誤差やプリンタードライバーの印刷不能領域を考慮して5%余裕を持つ。
        required_zoom = int(printable_width / max(float(content_width), 1.0) * 95)
        if required_zoom < EXCEL_MIN_ZOOM:
            self.events.put(
                (
                    "log",
                    f"警告: シート「{sheet_name}」は横幅が広いため、Excelの最小倍率"
                    f"{EXCEL_MIN_ZOOM}%で出力します。横方向が複数ページになる場合があります。",
                )
            )
        return max(EXCEL_MIN_ZOOM, min(required_zoom, 100))

    def _calculate_excel_print_area(
        self, worksheet: object, used_range: object
    ) -> tuple[str, float]:
        first_row = int(used_range.Row)
        first_column = int(used_range.Column)
        row_count = int(used_range.Rows.Count)
        column_count = int(used_range.Columns.Count)
        last_row = first_row + row_count - 1
        last_column = first_column + column_count - 1
        required_extra_width = 0.0

        rows_to_check = min(row_count, EXCEL_MAX_OVERFLOW_ROWS)
        for row in range(first_row, first_row + rows_to_check):
            cell = worksheet.Cells(row, last_column)
            value = cell.Value2
            if not isinstance(value, str) or not value or bool(cell.WrapText):
                continue

            cell_width = float(
                cell.MergeArea.Width if bool(cell.MergeCells) else cell.Width
            )
            font_size = float(cell.Font.Size or 11)
            rendered_width = self._estimate_excel_text_width(value, font_size)
            required_extra_width = max(
                required_extra_width, rendered_width - cell_width
            )

        extended_column = last_column
        remaining_width = max(required_extra_width, 0.0)
        added_width = 0.0
        while remaining_width > 0 and (
            extended_column - last_column < EXCEL_MAX_OVERFLOW_COLUMNS
            and extended_column < 16384
        ):
            extended_column += 1
            column_width = float(worksheet.Columns(extended_column).Width)
            added_width += column_width
            remaining_width -= column_width

        if extended_column > last_column:
            self.events.put(
                (
                    "log",
                    f"シート「{worksheet.Name}」はセルからはみ出す文字を考慮し、"
                    f"印刷範囲を右へ{extended_column - last_column}列拡張しました。",
                )
            )
        if row_count > rows_to_check:
            self.events.put(
                (
                    "log",
                    f"警告: シート「{worksheet.Name}」は行数が多いため、"
                    f"先頭{rows_to_check}行で文字のはみ出しを判定しました。",
                )
            )

        top_left = worksheet.Cells(first_row, first_column)
        bottom_right = worksheet.Cells(last_row, extended_column)
        return (
            worksheet.Range(top_left, bottom_right).Address,
            float(used_range.Width) + added_width,
        )

    @staticmethod
    def _estimate_excel_text_width(value: str, font_size: float) -> float:
        # 全角文字を1文字、半角文字を約0.55文字として概算する。改行がある
        # 場合は最も長い行だけを使用し、セル内の左右余白分を加える。
        line_units = []
        for line in value.expandtabs(4).splitlines() or [""]:
            units = sum(
                1.0
                if unicodedata.east_asian_width(character) in {"W", "F", "A"}
                else 0.55
                for character in line
            )
            line_units.append(units)
        return max(line_units, default=0.0) * font_size + 4.0

    def _log_export_retry(self, export_error: Exception) -> None:
        self.events.put(
            (
                "log",
                "ExportAsFixedFormatを使用できなかったため、"
                f"SaveAsで再試行します ({export_error})",
            )
        )

    @staticmethod
    def _raise_export_error(
        export_error: Exception, save_as_error: Exception
    ) -> None:
        raise RuntimeError(
            "ExportAsFixedFormatとSaveAsの両方に失敗しました。"
            f" ExportAsFixedFormat: {export_error}; SaveAs: {save_as_error}"
        ) from save_as_error

    def _process_events(self) -> None:
        try:
            while True:
                event, message = self.events.get_nowait()
                if event == "success":
                    source_path = Path(message)
                    self._mark_completed(source_path)
                    self._write_log(
                        f"成功: {source_path.name} → {source_path.with_suffix('.pdf').name}"
                    )
                else:
                    self._write_log(message)
                if event == "done":
                    self.converting = False
                    self.select_button.configure(state=tk.NORMAL)
                    self.remove_button.configure(state=tk.NORMAL)
                    self.restore_button.configure(state=tk.NORMAL)
                    self.optimize_excel_check.configure(state=tk.NORMAL)
                    self.convert_button.configure(state=tk.NORMAL)
        except queue.Empty:
            pass

        if self.converting:
            self.root.after(100, self._process_events)

    def _write_log(self, message: str) -> None:
        self.log.configure(state=tk.NORMAL)
        self.log.insert(tk.END, message + "\n")
        self.log.see(tk.END)
        self.log.configure(state=tk.DISABLED)

    def _on_close(self) -> None:
        if self.converting:
            messagebox.showinfo("変換中", "変換が終了してからウィンドウを閉じてください。")
            return
        self.root.destroy()


def main() -> None:
    root = TkinterDnD.Tk()
    PptxToPdfApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
