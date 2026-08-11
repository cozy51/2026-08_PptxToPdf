from __future__ import annotations

import gc
import queue
import threading
import tkinter as tk
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
SUPPORTED_EXTENSIONS = {".pptx", ".doc", ".docx"}


class PptxToPdfApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Office to PDF")
        self.root.geometry("760x520")
        self.root.minsize(600, 420)

        self.selected_files: list[Path] = []
        self.events: queue.Queue[tuple[str, str]] = queue.Queue()
        self.converting = False

        self._build_ui()
        self._configure_drag_and_drop()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self) -> None:
        frame = ttk.Frame(self.root, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)
        frame.rowconfigure(4, weight=1)

        controls = ttk.Frame(frame)
        controls.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        self.select_button = ttk.Button(
            controls, text="Officeファイルを選択", command=self._select_files
        )
        self.select_button.pack(side=tk.LEFT)
        ttk.Label(
            controls, text="  またはPPTX・Wordファイルを下の一覧へドラッグ＆ドロップ"
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

        self.convert_button = ttk.Button(
            frame, text="PDFに変換", command=self._start_conversion
        )
        self.convert_button.grid(row=2, column=0, sticky=tk.W, pady=10)

        ttk.Label(frame, text="処理結果ログ").grid(row=3, column=0, sticky=tk.W)
        log_frame = ttk.Frame(frame)
        log_frame.grid(row=4, column=0, sticky="nsew", pady=(4, 0))
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

        known_paths = {str(path).casefold() for path in self.selected_files}
        for path in paths:
            path_key = str(path).casefold()
            if path_key not in known_paths:
                self.selected_files.append(path)
                known_paths.add(path_key)

        self.file_list.delete(0, tk.END)
        for path in self.selected_files:
            self.file_list.insert(tk.END, str(path))

    def _select_files(self) -> None:
        names = filedialog.askopenfilenames(
            title="Officeファイルを選択",
            filetypes=[
                ("対応するOfficeファイル", "*.pptx *.doc *.docx"),
                ("PowerPoint プレゼンテーション", "*.pptx"),
                ("Word 文書", "*.doc *.docx"),
            ],
        )
        if not names:
            return

        self._add_files([Path(name).resolve() for name in names], replace=True)
        self._write_log(f"{len(self.selected_files)} 件のファイルを選択しました。")

    def _start_conversion(self) -> None:
        if not self.selected_files:
            messagebox.showwarning(
                "ファイル未選択", "PPTXまたはWordファイルを選択してください。"
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
        self.convert_button.configure(state=tk.DISABLED)
        self._write_log("変換を開始します。")

        worker = threading.Thread(
            target=self._convert_files, args=(files_to_convert,), daemon=False
        )
        worker.start()
        self.root.after(100, self._process_events)

    def _convert_files(self, paths: list[Path]) -> None:
        powerpoint = None
        word = None
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
            self.events.put(("log", f"成功: {source_path.name} → {pdf_path.name}"))
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
            self.events.put(("log", f"成功: {source_path.name} → {pdf_path.name}"))
        finally:
            if document is not None:
                try:
                    document.Close(MSO_FALSE)
                except Exception as exc:
                    self.events.put(
                        ("log", f"警告: {source_path.name} を閉じられませんでした ({exc})")
                    )

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
                self._write_log(message)
                if event == "done":
                    self.converting = False
                    self.select_button.configure(state=tk.NORMAL)
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
