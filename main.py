from __future__ import annotations

import gc
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import pythoncom
import win32com.client


# COMへ渡す値はPythonのboolではなく、OfficeのMsoTriStateと同じ整数値を
# 明示的に使用する。環境によっては動的COMディスパッチがboolやキーワード
# 引数を正しく変換できず、"The Python instance can not be converted to a COM
# object" になるため。
MSO_TRUE = -1
MSO_FALSE = 0
PDF_FORMAT = 2  # PowerPoint: ppFixedFormatTypePDF


class PptxToPdfApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("PPTX to PDF")
        self.root.geometry("760x520")
        self.root.minsize(600, 420)

        self.selected_files: list[Path] = []
        self.events: queue.Queue[tuple[str, str]] = queue.Queue()
        self.converting = False

        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self) -> None:
        frame = ttk.Frame(self.root, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)
        frame.rowconfigure(4, weight=1)

        self.select_button = ttk.Button(
            frame, text="PPTXファイルを選択", command=self._select_files
        )
        self.select_button.grid(row=0, column=0, sticky=tk.W, pady=(0, 8))

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

    def _select_files(self) -> None:
        names = filedialog.askopenfilenames(
            title="PPTXファイルを選択",
            filetypes=[("PowerPoint プレゼンテーション", "*.pptx")],
        )
        if not names:
            return

        self.selected_files = [Path(name).resolve() for name in names]
        self.file_list.delete(0, tk.END)
        for path in self.selected_files:
            self.file_list.insert(tk.END, str(path))
        self._write_log(f"{len(self.selected_files)} 件のファイルを選択しました。")

    def _start_conversion(self) -> None:
        if not self.selected_files:
            messagebox.showwarning("ファイル未選択", "PPTXファイルを選択してください。")
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
        pythoncom.CoInitialize()
        try:
            powerpoint = win32com.client.DispatchEx("PowerPoint.Application")

            for pptx_path in paths:
                presentation = None
                current_step = "ファイルの確認"
                try:
                    if not pptx_path.is_file():
                        raise FileNotFoundError("ファイルが見つかりません")

                    pdf_path = pptx_path.with_suffix(".pdf")
                    current_step = "PowerPointで開く"
                    # pywin32の動的ディスパッチでの互換性を高めるため、COM
                    # メソッドにはキーワード引数やPythonのboolを渡さない。
                    presentation = powerpoint.Presentations.Open(
                        str(pptx_path), MSO_TRUE, MSO_FALSE, MSO_FALSE
                    )
                    current_step = "PDFとして書き出す"
                    presentation.ExportAsFixedFormat(str(pdf_path), PDF_FORMAT)
                    self.events.put(("log", f"成功: {pptx_path.name} → {pdf_path.name}"))
                except Exception as exc:
                    self.events.put(
                        (
                            "log",
                            f"失敗 [{current_step}]: {pptx_path.name} ({exc})",
                        )
                    )
                finally:
                    if presentation is not None:
                        try:
                            presentation.Close()
                        except Exception as exc:
                            self.events.put(
                                ("log", f"警告: {pptx_path.name} を閉じられませんでした ({exc})")
                            )
                        finally:
                            presentation = None
        except Exception as exc:
            self.events.put(("log", f"PowerPointを起動できませんでした: {exc}"))
        finally:
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
    root = tk.Tk()
    PptxToPdfApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
