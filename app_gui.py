# app_gui_moderno.py
import os
import threading
from pathlib import Path

import pandas as pd
import tkinter as tk
import tkinter.scrolledtext as st
import tkinter.font as tkfont
from tkinter import filedialog, messagebox

import ttkbootstrap as tb
from ttkbootstrap.constants import *
from ttkbootstrap.toast import ToastNotification

from extract_core import run_batch

DEFAULT_INPUT = Path("pdfs_entrada")
DEFAULT_OUTPUT = Path("dados_extraidos.xlsx")
APP_TITLE = "Extrator de Cursos em PDFs – TCDF"


class App(tb.Window):
    def __init__(self):
        super().__init__(title=APP_TITLE, themename="darkly")
        self.geometry("1100x680")
        self.minsize(980, 600)

        # ----- Vars -----
        self.input_dir = tk.StringVar(value=str(DEFAULT_INPUT))
        self.output_xlsx = tk.StringVar(value=str(DEFAULT_OUTPUT))

        self.pages = tk.StringVar(value="1")  # "1" ou "1,2"
        self.y_min = tk.DoubleVar(value=285.0)
        self.y_max = tk.DoubleVar(value=465.0)
        self.x_pres_ini = tk.DoubleVar(value=455.0)
        self.x_pres_fim = tk.DoubleVar(value=475.0)
        self.x_misto_ini = tk.DoubleVar(value=515.0)
        self.x_misto_fim = tk.DoubleVar(value=535.0)
        self.x_dist_ini = tk.DoubleVar(value=575.0)
        self.x_dist_fim = tk.DoubleVar(value=595.0)
        self.y_tol = tk.IntVar(value=12)
        self.export_dbg = tk.BooleanVar(value=False)

        self._hover_rowid = None
        self._build_ui()

    # ---------------- UI ----------------
    def _build_ui(self):
        # Topbar
        top = tb.Frame(self, padding=(16, 12, 16, 0))
        top.pack(fill=X)
        tb.Label(top, text=APP_TITLE, font="-size 16 -weight bold").pack(side=LEFT)
        tb.Button(top, text="Tema", bootstyle=SECONDARY, command=self._toggle_theme)\
            .pack(side=RIGHT, padx=6)
        tb.Button(top, text="Sobre", bootstyle=SECONDARY, command=self._sobre)\
            .pack(side=RIGHT)

        body = tb.Frame(self, padding=16)
        body.pack(fill=BOTH, expand=YES)

        left = tb.Frame(body)
        left.pack(side=LEFT, fill=Y)

        right = tb.Frame(body)
        right.pack(side=RIGHT, fill=BOTH, expand=YES, padx=(12, 0))

        # ----- Card: Arquivos -----
        card_files = tb.Labelframe(left, text="Arquivos", padding=12)
        card_files.pack(fill=X, pady=(0, 12))

        self._row_entry_browse(card_files, "Pasta de PDFs:", self.input_dir, self.pick_input)
        self._row_entry_browse(card_files, "Arquivo XLSX de saída:", self.output_xlsx,
                               self.pick_output, "Salvar…")

        # ----- Card: Parâmetros -----
        card_params = tb.Labelframe(left, text="Parâmetros", padding=12)
        card_params.pack(fill=X)

        self._grid_labeled(card_params, "Páginas (ex: 1 ou 1,2):", self.pages, 0, 0, width=14)
        self._grid_range(card_params, "Faixa Y dos cursos (mín/máx):", self.y_min, self.y_max, 1)
        self._grid_range(card_params, "Checkbox Presencial (X0..X1):", self.x_pres_ini, self.x_pres_fim, 2)
        self._grid_range(card_params, "Checkbox Misto (X0..X1):", self.x_misto_ini, self.x_misto_fim, 3)
        self._grid_range(card_params, "Checkbox À distância (X0..X1):", self.x_dist_ini, self.x_dist_fim, 4)

        self._grid_labeled(card_params, "Tolerância Y (px):", self.y_tol, 5, 0, width=10)
        tb.Checkbutton(card_params, text="Exportar PNGs de depuração",
                       variable=self.export_dbg, bootstyle="round-toggle")\
          .grid(row=6, column=0, columnspan=2, sticky=W, pady=(6, 0))

        # Botão principal
        btn_frame = tb.Frame(left)
        btn_frame.pack(fill=X, pady=12)
        self.btn_run = tb.Button(btn_frame, text="Executar extração",
                                 bootstyle=SUCCESS, command=self.run_extract_thread)
        self.btn_run.pack(fill=X)

        # ----- Card: Resultado -----
        card_res = tb.Labelframe(right, text="Resultados", padding=8)
        card_res.pack(fill=BOTH, expand=YES)

        # >>> CORRIGIDO: vírgula após "lotacao" e incluí "modalidade" <<<
        cols = ("arquivo", "requerente", "cargo", "lotacao", "curso_titulo", "curso_horas", "modalidade")
        self.tree = tb.Treeview(card_res, columns=cols, show="headings", height=16, bootstyle=INFO)

        # estilo geral
        self.style.configure("Treeview", rowheight=28, font=("-size", 10))
        self.style.configure("Treeview.Heading", font=("-size", 10, "bold"))

        # cabeçalhos e alinhamentos
        for c in cols:
            self.tree.heading(c, text=c.upper(), command=lambda col=c: self._sort_by(col, False))
            if c == "curso_horas":
                self.tree.column(c, width=90, anchor=E, stretch=False)  # números à direita
            elif c in ("curso_titulo", "cargo"):
                self.tree.column(c, width=260, anchor=W, stretch=True)  # campos longos
            else:
                self.tree.column(c, width=160, anchor=W, stretch=True)

        self.tree.pack(fill=BOTH, expand=YES)

        # zebra + hover (com fallbacks de cor)
        palette = self.style.colors
        secondary_bg = getattr(palette, "secondarybg", getattr(palette, "secondary", palette.inputbg))
        base_bg = getattr(palette, "bg", palette.inputbg)

        self.tree.tag_configure("odd", background=secondary_bg)
        self.tree.tag_configure("even", background=base_bg)
        text_on_primary = getattr(palette, "selectfg", getattr(palette, "fg", "#ffffff"))
        self.tree.tag_configure("hover", background=palette.primary, foreground=text_on_primary)

        self.tree.bind("<Motion>", self._on_tree_hover)

        # Progress + Log
        bottom = tb.Frame(card_res, padding=(0, 8, 0, 0))
        bottom.pack(fill=X)
        self.progress = tb.Progressbar(bottom, mode="indeterminate")
        self.progress.pack(fill=X, pady=(0, 8))

        self.log = st.ScrolledText(bottom, height=6)
        self.log.pack(fill=X)

        # Status bar
        self.status = tb.Label(self, text="Pronto", anchor=W, bootstyle=SECONDARY)
        self.status.pack(side=BOTTOM, fill=X)

    # ------------- helpers UI -------------
    def _row_entry_browse(self, parent, label, var, command, btn_text="Escolher…"):
        row = tb.Frame(parent)
        row.pack(fill=X, pady=4)
        tb.Label(row, text=label, width=22, anchor=W).pack(side=LEFT)
        tb.Entry(row, textvariable=var).pack(side=LEFT, fill=X, expand=YES, padx=(0, 6))
        tb.Button(row, text=btn_text, bootstyle=PRIMARY, command=command).pack(side=LEFT)

    def _grid_labeled(self, parent, label, var, r, c, width=16):
        tb.Label(parent, text=label).grid(row=r, column=c, sticky=W, pady=4)
        tb.Entry(parent, textvariable=var, width=width).grid(row=r, column=c + 1, sticky=W, padx=(8, 0))

    def _grid_range(self, parent, label, var_min, var_max, r):
        tb.Label(parent, text=label).grid(row=r, column=0, sticky=W, pady=4)
        row = tb.Frame(parent)
        row.grid(row=r, column=1, sticky=W)
        tb.Entry(row, textvariable=var_min, width=10).pack(side=LEFT)
        tb.Label(row, text=" a ").pack(side=LEFT, padx=4)
        tb.Entry(row, textvariable=var_max, width=10).pack(side=LEFT)

    def _toggle_theme(self):
        cur = self.style.theme.name
        self.style.theme_use("flatly" if cur != "flatly" else "darkly")

    def _toast(self, title, msg):
        ToastNotification(title=title, message=msg, duration=4000,
                          position=(self.winfo_x() + 40, self.winfo_y() + 60)).show_toast()

    def _sobre(self):
        messagebox.showinfo("Sobre",
                            "Extrator de Cursos em PDFs – TCDF\nUI moderna com ttkbootstrap.\n© você :)")

    # ---------- seleção de arquivos ----------
    def pick_input(self):
        p = filedialog.askdirectory(initialdir=self.input_dir.get() or os.getcwd())
        if p:
            self.input_dir.set(p)

    def pick_output(self):
        p = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                         filetypes=[("Excel", "*.xlsx")],
                                         initialfile=DEFAULT_OUTPUT.name)
        if p:
            self.output_xlsx.set(p)

    # ---------- Log ----------
    def append_log(self, text: str):
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.update_idletasks()

    # ---------- Tree helpers ----------
    def _clear_tree(self):
        for i in self.tree.get_children():
            self.tree.delete(i)

    def _get_tree_font(self):
        style_font = self.style.lookup("Treeview", "font")
        try:
            return tkfont.nametofont(style_font) if style_font else tkfont.nametofont("TkDefaultFont")
        except tk.TclError:
            return tkfont.nametofont("TkDefaultFont")

    def _autofit_tree(self, min_w=None, max_w=None, pad=16):
        """Ajusta largura por conteúdo (cabeçalho + linhas)."""
        f = self._get_tree_font()
        for col in self.tree["columns"]:
            header_text = self.tree.heading(col)["text"]
            max_px = f.measure(header_text)
            for iid in self.tree.get_children(""):
                val = self.tree.set(iid, col)
                if val is None:
                    continue
                px = f.measure(str(val))
                if px > max_px:
                    max_px = px
            w = max_px + pad
            if min_w and col in min_w:
                w = max(w, min_w[col])
            if max_w and col in max_w:
                w = min(w, max_w[col])
            stretch = (col != "curso_horas")  # horas não estica
            self.tree.column(col, width=int(w), stretch=stretch)

    def _sort_by(self, col, descending):
        """Ordena a treeview pela coluna; alterna asc/desc."""
        data = []
        for iid in self.tree.get_children(""):
            val = self.tree.set(iid, col)
            if col == "curso_horas":
                try:
                    sort_key = int(val)
                except (TypeError, ValueError):
                    sort_key = float("inf")
            else:
                sort_key = (val or "").lower()
            data.append((sort_key, iid))
        data.sort(reverse=descending)
        for idx, (_, iid) in enumerate(data):
            self.tree.move(iid, "", idx)
            self.tree.item(iid, tags=("odd" if idx % 2 else "even",))
        self.tree.heading(col, command=lambda c=col: self._sort_by(c, not descending))

    def _on_tree_hover(self, event):
        """Realça a linha sob o mouse (efeito hover)."""
        rowid = self.tree.identify_row(event.y)
        # remove highlight anterior
        if getattr(self, "_hover_rowid", None) and self._hover_rowid != rowid:
            tags = list(self.tree.item(self._hover_rowid, "tags"))
            if "hover" in tags:
                tags.remove("hover")
                self.tree.item(self._hover_rowid, tags=tuple(tags))
        # aplica highlight atual
        if rowid:
            tags = list(self.tree.item(rowid, "tags"))
            if "hover" not in tags:
                tags.append("hover")
                self.tree.item(rowid, tags=tuple(tags))
            self._hover_rowid = rowid

    # ---------- Overlay ----------
    def _show_overlay(self, text="Processando…"):
        self.update_idletasks()
        self.overlay = tb.Toplevel(self)
        self.overlay.overrideredirect(True)
        self.overlay.attributes("-topmost", True)
        self.overlay.geometry(
            f"{self.winfo_width()}x{self.winfo_height()}+{self.winfo_rootx()}+{self.winfo_rooty()}"
        )
        try:
            self.overlay.attributes("-alpha", 0.88)
        except Exception:
            pass

        frm = tb.Frame(self.overlay, padding=24)
        frm.place(relx=0.5, rely=0.5, anchor="center")
        tb.Label(frm, text=text, font="-size 12 -weight bold").pack(pady=(0, 8))
        self.ov_prog = tb.Progressbar(frm, mode="indeterminate")
        self.ov_prog.pack(fill=X)
        self.ov_prog.start(12)
        self.overlay.grab_set()

    def _hide_overlay(self):
        if hasattr(self, "ov_prog"):
            self.ov_prog.stop()
        if hasattr(self, "overlay"):
            try:
                self.overlay.grab_release()
                self.overlay.destroy()
            except Exception:
                pass

    # ---------- Execução ----------
    def run_extract_thread(self):
        t = threading.Thread(target=self.run_extract, daemon=True)
        t.start()

    def run_extract(self):
        try:
            # UI feedback
            self.btn_run.configure(state=DISABLED)
            self.progress.start(12)
            self._show_overlay("Extraindo cursos dos PDFs…")
            self.status.configure(text="Processando PDFs…")
            self.log.delete("1.0", tk.END)
            self._clear_tree()

            # coleta de parâmetros
            input_dir = Path(self.input_dir.get())
            output_xlsx = Path(self.output_xlsx.get())

            pages = []
            for p in self.pages.get().split(","):
                p = p.strip()
                if p:
                    pages.append(int(p))

            y_range = (float(self.y_min.get()), float(self.y_max.get()))
            x_cols = {
                "presencial": (float(self.x_pres_ini.get()), float(self.x_pres_fim.get())),
                "misto": (float(self.x_misto_ini.get()), float(self.x_misto_fim.get())),
                "à distância": (float(self.x_dist_ini.get()), float(self.x_dist_fim.get())),
            }
            y_tol = int(self.y_tol.get())
            export_dbg = bool(self.export_dbg.get())
            annotations_dir = Path("debug_checagem") if export_dbg else None

            # chamada principal
            self.append_log(f"🔎 Lendo PDFs em: {input_dir}")
            df: pd.DataFrame = run_batch(
                input_dir=input_dir,
                output_xlsx=output_xlsx,
                course_pages=pages,
                course_y_range=y_range,
                checkbox_columns=x_cols,
                y_tolerance=y_tol,
                export_annotations=export_dbg,
                annotations_dir=annotations_dir,
            )

            # render preview + zebra (tratando NaN -> "")
            tag = "even"
            for _, r in df.iterrows():
                values = []
                for c in self.tree["columns"]:
                    v = r.get(c)
                    v = "" if (v is None or (isinstance(v, float) and pd.isna(v))) else v
                    values.append(v)
                self.tree.insert("", "end", values=values, tags=(tag,))
                tag = "odd" if tag == "even" else "even"

            # auto-ajuste de largura
            self._autofit_tree(
                min_w={
                    "arquivo": 140,
                    "requerente": 180,
                    "cargo": 200,
                    "lotacao": 160,
                    "curso_titulo": 240,
                    "curso_horas": 70,
                    "modalidade": 120,
                },
                max_w={"curso_titulo": 520, "cargo": 360},
            )

            self.append_log(f"✅ Planilha salva em: {output_xlsx.resolve()}")
            self.status.configure(text=f"Planilha salva em: {output_xlsx}")
            self._toast("Extração concluída", f"Planilha salva em:\n{output_xlsx}")

        except Exception as e:
            messagebox.showerror("Erro", str(e))
            self.append_log(f"❌ Erro: {e}")
            self.status.configure(text="Erro na extração")
        finally:
            self.progress.stop()
            self._hide_overlay()
            self.btn_run.configure(state=NORMAL)


if __name__ == "__main__":
    App().mainloop()
