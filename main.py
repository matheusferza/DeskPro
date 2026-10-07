from __future__ import annotations

import json
import os
import re
import sys
import threading
import tkinter as tk
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import Any, Dict, List

import ttkbootstrap as ttk

from core import (
    export_excel_file,
    get_data_path,
    import_excel_file,
    load_records,
    parse_date,
    prepare_record,
    save_records,
)

BG = "#f4f7f6"
TEXT = "#16211f"
TEAL = "#0d7d72"
TEAL_HOVER = "#0a5f57"
TEAL_DEEP = "#0b5c54"
GOLD = "#b8862c"
GOLD_DEEP = "#8a6817"
BORDER = "#dfe9e5"
PANEL = "#ffffff"
HEADER_BG = "#0d7d72"
HEADER_TEXT = "#f4fdfb"
SUCCESS_BG = "#e3f3e9"
SUCCESS_TEXT = "#1f7a4d"
ALERT_BG = "#fbe7e4"
ALERT_TEXT = "#a83224"
INFO_BG = "#e7edf7"
INFO_TEXT = "#33508f"
CARD_ACCENT = "#ecf7f5"
CARD_ALERT = "#fff0ed"
CARD_SUCCESS = "#e8f5ed"
def get_icon_path() -> Path:
    candidates = []
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        candidates.append(exe_dir / "icon.ico")
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            candidates.append(Path(meipass) / "icon.ico")
    else:
        app_dir = Path(__file__).resolve().parent
        candidates.append(app_dir / "icon.ico")

    for candidate in candidates:
        if candidate.exists():
            return candidate
    return candidates[0] if candidates else Path("icon.ico")


ICON_PATH = get_icon_path()


class LaboratorioApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("🦷 Laboratório — Controle de Trabalhos")
        self.root.geometry("1280x760")
        self.root.minsize(820, 500)
        self.root.configure(bg=BG)

        self.data_path = get_data_path()
        self.records: List[Dict[str, Any]] = load_records(self.data_path)
        self.selected_record = None

        self.build_style()
        self.build_ui()
        self.refresh_table()
        self.refresh_summary()

    def build_style(self) -> None:
        style = ttk.Style(theme="flatly")
        self.style = style
        style.configure("App.TFrame", background=BG)
        style.configure("Header.TFrame", background=HEADER_BG)
        style.configure("Card.TFrame", background=PANEL, relief="flat")
        style.configure("Primary.TButton", background=TEAL, foreground="white", font=("Segoe UI", 10, "bold"))
        style.map("Primary.TButton", background=[("active", TEAL_HOVER), ("pressed", TEAL_HOVER)])
        style.configure("Gold.TButton", background=GOLD, foreground="white", font=("Segoe UI", 10, "bold"))
        style.map("Gold.TButton", background=[("active", GOLD_DEEP), ("pressed", GOLD_DEEP)])
        style.configure("Secondary.TButton", background="#ffffff", foreground=TEXT)
        style.map("Secondary.TButton", background=[("active", "#eef4f1"), ("pressed", "#edf3f0")])
        style.configure("TEntry", fieldbackground="#ffffff", foreground=TEXT)
        style.configure("TCombobox", fieldbackground="#ffffff", foreground=TEXT)
        style.configure("TCheckbutton", background=BG, foreground=TEXT)
        style.configure("Treeview", background="#ffffff", fieldbackground="#ffffff", foreground=TEXT)
        style.configure("Treeview", rowheight=32, font=("Segoe UI", 10))
        style.configure("Treeview.Heading", background=TEAL_DEEP, foreground="white", font=("Segoe UI", 9, "bold"))
        style.map("Treeview.Heading", background=[("active", TEAL)])
        style.map("Treeview", background=[("selected", TEAL)], foreground=[("selected", "white")])
        style.configure("Compact.TLabel", background=BG, foreground=TEXT)
        style.configure("HeaderTitle.TLabel", background=HEADER_BG, foreground=HEADER_TEXT, font=("Segoe UI", 22, "bold"))
        style.configure("HeaderSubtitle.TLabel", background=HEADER_BG, foreground="#d9f5f1", font=("Segoe UI", 10))
        style.configure("CardTitle.TLabel", background=PANEL, foreground=TEXT, font=("Segoe UI", 12, "bold"))
        style.configure("Muted.TLabel", background=BG, foreground="#586562", font=("Segoe UI", 10))
        style.configure("Success.TLabel", background=SUCCESS_BG, foreground=SUCCESS_TEXT, font=("Segoe UI", 10, "bold"), padding=(8, 4))
        style.configure("Danger.TLabel", background=ALERT_BG, foreground=ALERT_TEXT, font=("Segoe UI", 10, "bold"), padding=(8, 4))
        style.configure("Info.TLabel", background=INFO_BG, foreground=INFO_TEXT, font=("Segoe UI", 10, "bold"), padding=(8, 4))
        style.configure("Card.TCheckbutton", background=PANEL, foreground=TEXT)

    def build_ui(self) -> None:
        main = ttk.Frame(self.root, padding=(18, 12, 18, 12), style="App.TFrame")
        main.pack(fill="both", expand=True)

        header = ttk.Frame(main, style="Header.TFrame")
        header.pack(fill="x", pady=(0, 12), ipady=12, ipadx=16)

        title = ttk.Label(header, text="🦷 Laboratório — Controle de Trabalhos", style="HeaderTitle.TLabel")
        title.pack(anchor="w")
        subtitle = ttk.Label(header, text="Precisão em cada etapa, do gesso à entrega.", style="HeaderSubtitle.TLabel")
        subtitle.pack(anchor="w", pady=(2, 0))

        summary_bar = ttk.Frame(main, style="App.TFrame")
        summary_bar.pack(fill="x", pady=(0, 14))

        self.summary_total = self.create_summary_card(summary_bar, "Total", "0", CARD_ACCENT, TEAL_DEEP, "SummaryTotal.TFrame")
        self.summary_delay = self.create_summary_card(summary_bar, "Atrasados", "0", CARD_ALERT, ALERT_TEXT, "SummaryAlert.TFrame")
        self.summary_done = self.create_summary_card(summary_bar, "Entregues", "0", CARD_SUCCESS, SUCCESS_TEXT, "SummarySuccess.TFrame")

        content = ttk.Frame(main, style="App.TFrame")
        content.pack(fill="both", expand=True)

        self.form_panel = ttk.Frame(content, style="App.TFrame", width=420)
        self.form_panel.pack(fill="y", side="left", padx=(0, 14))
        self.form_panel.pack_propagate(False)

        self.form_canvas = tk.Canvas(self.form_panel, background=BG, borderwidth=0, highlightthickness=0)
        self.form_scrollbar = ttk.Scrollbar(self.form_panel, orient="vertical", command=self.form_canvas.yview)
        self.form_canvas.configure(yscrollcommand=self.form_scrollbar.set)
        self.form_canvas.pack(side="left", fill="both", expand=True)
        self.form_scrollbar.pack(side="right", fill="y")

        form_card = ttk.Frame(self.form_canvas, style="Card.TFrame", padding=14)
        self.form_canvas_window = self.form_canvas.create_window((0, 0), window=form_card, anchor="nw")
        form_card.bind("<Configure>", self._update_form_scrollregion)
        self.form_canvas.bind("<Configure>", self._resize_form_content)
        self.root.bind_all("<MouseWheel>", self._on_form_mousewheel)

        form_title = ttk.Label(form_card, text="Cadastrar trabalho", style="CardTitle.TLabel")
        form_title.pack(anchor="w", pady=(0, 10))

        self.form_vars = {
            "COD": tk.StringVar(),
            "DATA ENTRADA": tk.StringVar(),
            "CLIENTES": tk.StringVar(),
            "PACIENTE": tk.StringVar(),
            "SUP/INF": tk.StringVar(value="Superior"),
            "SAIDA": tk.StringVar(),
            "OBSERVAÇÕES": tk.StringVar(),
            "ENTREGUE": tk.BooleanVar(value=False),
        }

        self.fields = []
        field_rows = (
            (("Código", "COD"), ("Data de entrada", "DATA ENTRADA"), ("Data prevista de saída", "SAIDA")),
            (("Cliente / clínica", "CLIENTES"), ("Paciente", "PACIENTE"), ("Sup/Inf", "SUP/INF")),
        )
        for row_fields in field_rows:
            row = ttk.Frame(form_card, style="Card.TFrame")
            row.pack(fill="x", pady=(0, 8))
            for column, (label, key) in enumerate(row_fields):
                row.grid_columnconfigure(column, weight=1, uniform="form-fields")
                field = ttk.Frame(row, style="Card.TFrame")
                field.grid(row=0, column=column, sticky="ew", padx=(0 if column == 0 else 6, 0))
                ttk.Label(
                    field,
                    text=label,
                    font=("Segoe UI", 9, "bold"),
                    background=PANEL,
                    foreground=TEXT,
                    wraplength=115,
                    justify="left",
                ).pack(anchor="w", fill="x")
                if key == "SUP/INF":
                    value_widget = ttk.Combobox(
                        field,
                        textvariable=self.form_vars[key],
                        values=["Superior", "Inferior", "Ambas"],
                        state="readonly",
                    )
                    self.combo_supinf = value_widget
                else:
                    value_widget = ttk.Entry(field, textvariable=self.form_vars[key])
                value_widget.pack(fill="x", pady=(3, 0))
                self.fields.append((key, value_widget))

        obs_row = ttk.Frame(form_card, style="Card.TFrame")
        obs_row.pack(fill="x", pady=(0, 8))
        ttk.Label(obs_row, text="Observações", font=("Segoe UI", 9, "bold"), background=PANEL, foreground=TEXT).pack(anchor="w")
        obs_entry = ttk.Entry(obs_row, textvariable=self.form_vars["OBSERVAÇÕES"])
        obs_entry.pack(fill="x", pady=(3, 0))

        delivery_row = ttk.Frame(form_card, style="Card.TFrame")
        delivery_row.pack(fill="x", pady=(0, 8))
        delivered_check = ttk.Checkbutton(delivery_row, text="Já foi entregue", variable=self.form_vars["ENTREGUE"], style="Card.TCheckbutton")
        delivered_check.pack(side="left", anchor="w")

        self.status_var = tk.StringVar(value="No prazo")
        self.status_label = ttk.Label(delivery_row, textvariable=self.status_var, style="Success.TLabel")
        self.status_label.pack(side="right", anchor="e")

        actions = ttk.Frame(form_card, style="Card.TFrame")
        actions.pack(fill="x")

        self.save_button = ttk.Button(actions, text="Salvar trabalho", command=self.save_work, style="Primary.TButton")
        self.save_button.pack(side="left", padx=(0, 8))
        self.delete_button = ttk.Button(actions, text="Excluir", command=self.delete_selected, style="Secondary.TButton")
        self.delete_button.pack(side="left", padx=(0, 8))
        self.cancel_button = ttk.Button(actions, text="Limpar", command=self.clear_form, style="Secondary.TButton")
        self.cancel_button.pack(side="left")

        right = ttk.Frame(content, style="App.TFrame")
        right.pack(fill="both", expand=True)

        toolbar = ttk.Frame(right, style="App.TFrame")
        toolbar.pack(fill="x", pady=(0, 8))

        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(toolbar, textvariable=self.search_var, width=30)
        search_entry.pack(side="left", fill="x", expand=True)
        self.search_var.trace_add("write", lambda *_: self.refresh_table())

        import_btn = ttk.Button(toolbar, text="Importar planilha", command=self.import_excel, style="Secondary.TButton")
        import_btn.pack(side="left", padx=(8, 0))

        export_btn = ttk.Button(toolbar, text="Exportar Excel", command=self.export_excel, style="Gold.TButton")
        export_btn.pack(side="left", padx=(8, 0))

        table_wrap = ttk.Frame(right, style="App.TFrame")
        table_wrap.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(table_wrap, columns=("COD", "CLIENTES", "PACIENTE", "SUP/INF", "SAIDA", "STATUS"), show="headings")
        self.tree.heading("COD", text="COD")
        self.tree.heading("CLIENTES", text="CLIENTES")
        self.tree.heading("PACIENTE", text="PACIENTE")
        self.tree.heading("SUP/INF", text="SUP/INF")
        self.tree.heading("SAIDA", text="SAIDA")
        self.tree.heading("STATUS", text="STATUS")

        self.tree.column("COD", width=110, anchor="center")
        self.tree.column("CLIENTES", width=180, anchor="w")
        self.tree.column("PACIENTE", width=180, anchor="w")
        self.tree.column("SUP/INF", width=100, anchor="center")
        self.tree.column("SAIDA", width=120, anchor="center")
        self.tree.column("STATUS", width=120, anchor="center")

        self.tree.tag_configure("even", background="#f8faf9")
        self.tree.tag_configure("odd", background="#ffffff")
        self.tree.tag_configure("status_no_prazo", background=SUCCESS_BG, foreground=SUCCESS_TEXT)
        self.tree.tag_configure("status_atrasado", background=ALERT_BG, foreground=ALERT_TEXT)
        self.tree.tag_configure("status_entregue", background=INFO_BG, foreground=INFO_TEXT)

        self.tree.bind("<ButtonRelease-1>", self.on_tree_select)
        self.tree.pack(fill="both", expand=True, side="left")

        scroll = ttk.Scrollbar(table_wrap, orient="vertical", command=self.tree.yview)
        scroll.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scroll.set)

        self.empty_state = ttk.Label(right, text="Nenhum trabalho em andamento. Bora cadastrar o próximo caso?", style="Muted.TLabel")
        self.empty_state.pack_forget()

        footer = ttk.Label(main, text="Feito para acompanhar cada etapa com o mesmo cuidado que você põe no trabalho.", style="Muted.TLabel")
        footer.pack(fill="x", pady=(10, 0))

        self.tree.bind("<Double-1>", self.on_tree_double_click)

        self.set_status_display()

    def _update_form_scrollregion(self, _event: tk.Event) -> None:
        self.form_canvas.configure(scrollregion=self.form_canvas.bbox("all"))

    def _resize_form_content(self, event: tk.Event) -> None:
        self.form_canvas.itemconfigure(self.form_canvas_window, width=event.width)

    def _on_form_mousewheel(self, event: tk.Event) -> str | None:
        widget = self.root.winfo_containing(event.x_root, event.y_root)
        while widget is not None and widget is not self.form_panel:
            widget = getattr(widget, "master", None)
        if widget is None:
            return None
        self.form_canvas.yview_scroll(-int(event.delta / 120), "units")
        return "break"

    def create_summary_card(self, master: ttk.Frame, title: str, value: str, background: str, foreground: str, style_name: str) -> ttk.Frame:
        self.style.configure(style_name, background=background)
        card = ttk.Frame(master, style=style_name, padding=(12, 10), width=180)
        card.pack(side="left", padx=(0, 12), fill="y")
        ttk.Label(card, text=title, font=("Segoe UI", 10, "bold"), background=background, foreground="#4a5a56").pack(anchor="w")
        label = ttk.Label(card, text=value, font=("Segoe UI", 18, "bold"), background=background, foreground=foreground)
        label.pack(anchor="w", pady=(4, 0))
        return label

    def set_loading(self, loading: bool, message: str = "Carregando...") -> None:
        self.root.config(cursor="wait" if loading else "arrow")
        if loading:
            self.loading_label = ttk.Label(self.root, text=message, background=BG, foreground=TEAL_DEEP, font=("Segoe UI", 10, "bold"))
            self.loading_label.place(relx=1.0, rely=1.0, anchor="se", x=-18, y=-16)
        else:
            if hasattr(self, "loading_label"):
                self.loading_label.destroy()

    def clear_form(self) -> None:
        for key in self.form_vars:
            if isinstance(self.form_vars[key], tk.BooleanVar):
                self.form_vars[key].set(False)
            else:
                self.form_vars[key].set("")
        self.form_vars["SUP/INF"].set("Superior")
        self.selected_record = None
        self.set_status_display()

    def set_status_display(self) -> None:
        status = "No prazo"
        try:
            data = {
                "COD": self.form_vars["COD"].get(),
                "DATA ENTRADA": self.form_vars["DATA ENTRADA"].get(),
                "CLIENTES": self.form_vars["CLIENTES"].get(),
                "PACIENTE": self.form_vars["PACIENTE"].get(),
                "SUP/INF": self.form_vars["SUP/INF"].get(),
                "SAIDA": self.form_vars["SAIDA"].get(),
                "OBSERVAÇÕES": self.form_vars["OBSERVAÇÕES"].get(),
                "ENTREGUE": self.form_vars["ENTREGUE"].get(),
            }
            status = prepare_record(data)["STATUS"]
        except Exception:
            pass
        self.status_var.set(status)
        style_name = {
            "No prazo": "Success.TLabel",
            "Atrasado": "Danger.TLabel",
            "Entregue": "Info.TLabel",
        }.get(status, "Muted.TLabel")
        self.status_label.configure(style=style_name)

    def save_work(self) -> None:
        payload = {
            "COD": self.form_vars["COD"].get(),
            "DATA ENTRADA": self.form_vars["DATA ENTRADA"].get(),
            "CLIENTES": self.form_vars["CLIENTES"].get(),
            "PACIENTE": self.form_vars["PACIENTE"].get(),
            "SUP/INF": self.form_vars["SUP/INF"].get(),
            "SAIDA": self.form_vars["SAIDA"].get(),
            "OBSERVAÇÕES": self.form_vars["OBSERVAÇÕES"].get(),
            "ENTREGUE": self.form_vars["ENTREGUE"].get(),
        }

        try:
            prepared = prepare_record(payload)
        except ValueError as exc:
            messagebox.showerror("Dados inválidos", str(exc))
            return

        def worker() -> None:
            try:
                if self.selected_record is None:
                    records = list(self.records)
                    records.append(prepared)
                else:
                    records = []
                    updated = False
                    for item in self.records:
                        if item.get("COD") == self.selected_record.get("COD") and item.get("PACIENTE") == self.selected_record.get("PACIENTE"):
                            records.append(prepared)
                            updated = True
                        else:
                            records.append(item)
                    if not updated:
                        records.append(prepared)
                save_records(records, self.data_path)
                self.records = records
                self.root.after(0, self.clear_form)
                self.root.after(0, self.refresh_table)
                self.root.after(0, self.refresh_summary)
                self.root.after(0, lambda: messagebox.showinfo("Sucesso", "Trabalho salvo com sucesso."))
            except Exception as exc:
                self.root.after(0, lambda: messagebox.showerror("Erro ao salvar", str(exc)))
            finally:
                self.root.after(0, lambda: self.set_loading(False))

        self.set_loading(True, "Salvando trabalho...")
        threading.Thread(target=worker, daemon=True).start()

    def delete_selected(self) -> None:
        if self.selected_record is None:
            return
        confirmed = messagebox.askyesno("Excluir trabalho", f"Deseja excluir o trabalho {self.selected_record.get('COD', '')}?")
        if not confirmed:
            return

        self.records = [item for item in self.records if item != self.selected_record]
        save_records(self.records, self.data_path)
        self.clear_form()
        self.refresh_table()
        self.refresh_summary()

    def load_selected_record(self, record: Dict[str, Any]) -> None:
        self.selected_record = record
        self.form_vars["COD"].set(record.get("COD", ""))
        self.form_vars["DATA ENTRADA"].set(record.get("DATA ENTRADA", ""))
        self.form_vars["CLIENTES"].set(record.get("CLIENTES", ""))
        self.form_vars["PACIENTE"].set(record.get("PACIENTE", ""))
        self.form_vars["SUP/INF"].set(record.get("SUP/INF", "Superior"))
        self.form_vars["SAIDA"].set(record.get("SAIDA", ""))
        self.form_vars["OBSERVAÇÕES"].set(record.get("OBSERVAÇÕES", ""))
        self.form_vars["ENTREGUE"].set(bool(record.get("ENTREGUE", False)))
        self.set_status_display()

    def on_tree_select(self, _event: Any) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        item = self.tree.item(selection[0], "values")
        for record in self.records:
            if record.get("COD") == item[0] and record.get("PACIENTE") == item[2]:
                self.load_selected_record(record)
                break

    def on_tree_double_click(self, _event: Any) -> None:
        self.on_tree_select(_event)

    def refresh_table(self) -> None:
        search = self.search_var.get().strip().lower()
        filtered = []
        for record in self.records:
            haystack = " ".join([
                str(record.get("COD", "")),
                str(record.get("CLIENTES", "")),
                str(record.get("PACIENTE", "")),
            ]).lower()
            if not search or search in haystack:
                filtered.append(record)

        self.tree.delete(*self.tree.get_children())
        if not filtered:
            self.empty_state.pack(fill="x", pady=(10, 0))
            return
        self.empty_state.pack_forget()

        for index, record in enumerate(filtered):
            prepared = prepare_record(record)
            status = prepared.get("STATUS", "No prazo")
            tag = "even" if index % 2 == 0 else "odd"
            if status == "Atrasado":
                tag = "status_atrasado"
            elif status == "Entregue":
                tag = "status_entregue"
            elif status == "No prazo":
                tag = "status_no_prazo"

            self.tree.insert(
                "",
                "end",
                values=(
                    prepared.get("COD", ""),
                    prepared.get("CLIENTES", ""),
                    prepared.get("PACIENTE", ""),
                    prepared.get("SUP/INF", ""),
                    prepared.get("SAIDA", ""),
                    status,
                ),
                tags=(tag,),
            )

    def refresh_summary(self) -> None:
        total = len(self.records)
        atrasados = sum(1 for item in self.records if prepare_record(item).get("STATUS") == "Atrasado")
        entregues = sum(1 for item in self.records if prepare_record(item).get("STATUS") == "Entregue")

        self.summary_total.configure(text=str(total))
        self.summary_delay.configure(text=str(atrasados))
        self.summary_done.configure(text=str(entregues))

    def import_excel(self) -> None:
        path = filedialog.askopenfilename(title="Importar planilha", filetypes=[("Excel", "*.xlsx *.xls")])
        if not path:
            return

        def worker() -> None:
            try:
                records = import_excel_file(path)
                self.records.extend(records)
                save_records(self.records, self.data_path)
                self.root.after(0, self.refresh_table)
                self.root.after(0, self.refresh_summary)
                self.root.after(0, lambda: messagebox.showinfo("Importação concluída", f"{len(records)} trabalho(s) importado(s) com sucesso."))
            except Exception as exc:
                self.root.after(0, lambda: messagebox.showerror("Erro na importação", str(exc)))
            finally:
                self.root.after(0, lambda: self.set_loading(False))

        self.set_loading(True, "Importando planilha...")
        threading.Thread(target=worker, daemon=True).start()

    def export_excel(self) -> None:
        destination = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")], title="Exportar Excel")
        if not destination:
            return

        def worker() -> None:
            try:
                export_excel_file(self.records, destination)
                self.root.after(0, lambda: messagebox.showinfo("Exportação concluída", f"Arquivo salvo em:\n{destination}"))
            except Exception as exc:
                self.root.after(0, lambda: messagebox.showerror("Erro na exportação", str(exc)))
            finally:
                self.root.after(0, lambda: self.set_loading(False))

        self.set_loading(True, "Exportando Excel...")
        threading.Thread(target=worker, daemon=True).start()

    def run(self) -> None:
        self.root.mainloop()


def main() -> None:
    try:
        from ctypes import windll
        windll.shell32.SetCurrentProcessExplicitAppUserModelID("labmarcos.desktop.app.1.0")
    except Exception:
        pass

    root = tk.Tk()
    icon_path = get_icon_path()
    if icon_path and icon_path.exists():
        try:
            root.iconbitmap(str(icon_path))
        except Exception:
            pass
    app = LaboratorioApp(root)
    app.run()


if __name__ == "__main__":
    main()
