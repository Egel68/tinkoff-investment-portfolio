"""GUI на CustomTkinter."""

import os
import threading
from decimal import Decimal
from tkinter import filedialog, messagebox

import customtkinter as ctk

from api_client import AccountInfo, AggregatedPosition, TinkoffApiClient
from config import EXCEL_FILENAME, TOKEN
from excel_export import ExcelExporter

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class PortfolioApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("📈 Т-Инвестиции — Анализатор портфеля")
        self.geometry("1200x800")
        self.minsize(1000, 650)

        self.client: TinkoffApiClient | None = None
        self.accounts: list[AccountInfo] = []
        self.aggregated: list[AggregatedPosition] = []
        self.total_capital = Decimal("0")

        self._build()
        if TOKEN:
            self.token_entry.insert(0, TOKEN)

    def _build(self):
        # === Токен ===
        top = ctk.CTkFrame(self, corner_radius=10)
        top.pack(fill="x", padx=12, pady=(12, 4))

        ctk.CTkLabel(
            top, text="🔑 Токен:", font=ctk.CTkFont(size=14, weight="bold")
        ).pack(side="left", padx=(12, 4), pady=8)

        self.token_entry = ctk.CTkEntry(
            top, show="•", width=420, placeholder_text="Вставьте токен Т-Инвестиций"
        )
        self.token_entry.pack(side="left", padx=4, pady=8, fill="x", expand=True)

        self.show_var = ctk.BooleanVar()
        ctk.CTkCheckBox(
            top, text="👁", variable=self.show_var, command=self._toggle_show, width=50
        ).pack(side="left", padx=4)

        self.btn_connect = ctk.CTkButton(
            top,
            text="Подключиться",
            width=140,
            fg_color="#27AE60",
            hover_color="#2ECC71",
            command=self._connect,
        )
        self.btn_connect.pack(side="right", padx=12, pady=8)

        # === Панель действий ===
        act = ctk.CTkFrame(self, corner_radius=10)
        act.pack(fill="x", padx=12, pady=4)

        self.lbl_status = ctk.CTkLabel(
            act, text="⏸ Не подключено", font=ctk.CTkFont(size=13), text_color="#E74C3C"
        )
        self.lbl_status.pack(side="left", padx=12, pady=8)

        self.btn_export = ctk.CTkButton(
            act,
            text="📥 Excel",
            command=self._export,
            state="disabled",
            width=130,
            fg_color="#2980B9",
            hover_color="#3498DB",
        )
        self.btn_export.pack(side="right", padx=4, pady=8)

        self.btn_refresh = ctk.CTkButton(
            act, text="🔄 Обновить", command=self._load, state="disabled", width=110
        )
        self.btn_refresh.pack(side="right", padx=4, pady=8)

        # === Капитал ===
        cap = ctk.CTkFrame(self, corner_radius=10)
        cap.pack(fill="x", padx=12, pady=4)
        self.lbl_capital = ctk.CTkLabel(
            cap, text="💰 Общий капитал: —", font=ctk.CTkFont(size=16, weight="bold")
        )
        self.lbl_capital.pack(pady=8)

        # === Табы ===
        self.tabs = ctk.CTkTabview(self, corner_radius=10)
        self.tabs.pack(fill="both", expand=True, padx=12, pady=(4, 12))
        self.tabs.add("Обзор")
        ctk.CTkLabel(
            self.tabs.tab("Обзор"), text="Подключитесь к API", font=ctk.CTkFont(size=14)
        ).pack(expand=True)

        # === Прогресс ===
        self.progress = ctk.CTkProgressBar(self, mode="indeterminate")

    def _toggle_show(self):
        self.token_entry.configure(show="" if self.show_var.get() else "•")

    def _set_busy(self, busy, text=""):
        if busy:
            self.progress.pack(fill="x", padx=12, pady=(0, 4))
            self.progress.start()
            if text:
                self.lbl_status.configure(text=text, text_color="#F39C12")
            for b in (self.btn_connect, self.btn_refresh, self.btn_export):
                b.configure(state="disabled")
        else:
            self.progress.stop()
            self.progress.pack_forget()
            self.btn_connect.configure(state="normal")
            if self.client:
                self.btn_refresh.configure(state="normal")
                self.btn_export.configure(state="normal")

    # === Подключение ===

    def _connect(self):
        token = self.token_entry.get().strip()
        if not token:
            messagebox.showwarning("Токен", "Введите токен API")
            return
        self.client = TinkoffApiClient(token)
        self._set_busy(True, "⏳ Проверка…")

        def job():
            ok, msg = self.client.test_connection()
            result_ok, result_msg = ok, msg
            self.after(0, lambda: self._on_connected(result_ok, result_msg))

        threading.Thread(target=job, daemon=True).start()

    def _on_connected(self, ok, msg):
        self._set_busy(False)
        if ok:
            self.lbl_status.configure(text=f"✅ {msg}", text_color="#27AE60")
            self._load()
        else:
            self.lbl_status.configure(text=f"❌ {msg}", text_color="#E74C3C")
            messagebox.showerror("Ошибка", msg)

    # === Загрузка ===

    def _load(self):
        self._set_busy(True, "⏳ Загрузка портфеля…")

        def job():
            try:
                accs, tot = self.client.get_all_accounts()
                agg = TinkoffApiClient.aggregate_positions(accs, tot)
                self.after(0, lambda: self._on_loaded(accs, tot, agg))
            except Exception as e:
                err_msg = str(e)
                self.after(0, lambda: self._on_load_err(err_msg))

        threading.Thread(target=job, daemon=True).start()

    def _on_loaded(self, accs, total, aggregated):
        self._set_busy(False)
        self.accounts = accs
        self.total_capital = total
        self.aggregated = aggregated
        self.lbl_capital.configure(text=f"💰 Общий капитал: {total:,.2f} ₽")
        n_pos = sum(len(a.positions) for a in accs)
        self.lbl_status.configure(
            text=f"✅ {len(accs)} счетов, {n_pos} позиций, {len(aggregated)} уникальных бумаг",
            text_color="#27AE60",
        )
        self._fill_tabs()

    def _on_load_err(self, err):
        self._set_busy(False)
        self.lbl_status.configure(text=f"❌ {err}", text_color="#E74C3C")
        messagebox.showerror("Ошибка", err)

    # === Вкладки ===

    def _fill_tabs(self):
        for name in list(self.tabs._tab_dict.keys()):
            self.tabs.delete(name)

        # 1) Сводка
        t0 = self.tabs.add("📊 Сводка")
        self._fill_summary_tab(t0)

        # 2) Все позиции
        t_all = self.tabs.add("📦 Все позиции")
        self._fill_aggregated_tab(t_all)

        # 3) По счетам
        for acc in self.accounts:
            tab = self.tabs.add(f"🏦 {acc.name}"[:30])
            self._fill_account_tab(tab, acc)

    def _fill_summary_tab(self, parent):
        sf = ctk.CTkScrollableFrame(parent)
        sf.pack(fill="both", expand=True, padx=4, pady=4)

        for acc in self.accounts:
            share = (
                (acc.total_value / self.total_capital * 100)
                if self.total_capital
                else 0
            )
            fr = ctk.CTkFrame(sf, corner_radius=8)
            fr.pack(fill="x", padx=4, pady=4)
            ctk.CTkLabel(
                fr,
                text=f"🏦 {acc.name} ({acc.acc_type})",
                font=ctk.CTkFont(size=14, weight="bold"),
            ).pack(anchor="w", padx=12, pady=(8, 2))
            ctk.CTkLabel(
                fr,
                text=(
                    f"💵 {acc.total_value:,.2f} ₽  ·  Доля: {share:.1f}%  ·  "
                    f"Позиций: {len(acc.positions)}  ·  {acc.opened_date}"
                ),
                font=ctk.CTkFont(size=12),
                text_color="#BDC3C7",
            ).pack(anchor="w", padx=12, pady=(2, 8))

    def _fill_aggregated_tab(self, parent):
        """Вкладка со всеми позициями, агрегированными по бумагам."""
        sf = ctk.CTkScrollableFrame(parent)
        sf.pack(fill="both", expand=True, padx=4, pady=4)

        if not self.aggregated:
            ctk.CTkLabel(sf, text="📭 Нет позиций", font=ctk.CTkFont(size=14)).pack(
                expand=True, pady=40
            )
            return

        # Статистика сверху
        total_pnl = sum(p.profit_loss for p in self.aggregated)
        total_avg = sum(p.total_avg_cost for p in self.aggregated)
        total_pnl_pct = (total_pnl / total_avg * 100) if total_avg else Decimal(0)
        pnl_sign = "+" if total_pnl >= 0 else ""
        pnl_color = "#27AE60" if total_pnl >= 0 else "#E74C3C"

        stats_frame = ctk.CTkFrame(sf, corner_radius=8, fg_color="#1a1a2e")
        stats_frame.pack(fill="x", padx=4, pady=(4, 8))

        ctk.CTkLabel(
            stats_frame,
            text=(
                f"📦 Уникальных бумаг: {len(self.aggregated)}  ·  "
                f"💰 Капитал: {self.total_capital:,.2f} ₽  ·  "
                f"P&L: {pnl_sign}{total_pnl:,.2f} ₽ ({pnl_sign}{total_pnl_pct:.2f}%)"
            ),
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=pnl_color,
        ).pack(padx=12, pady=10)

        # Заголовок таблицы
        hdr = ctk.CTkFrame(sf, corner_radius=4, fg_color="#1F4E79")
        hdr.pack(fill="x", padx=2, pady=(2, 0))

        cols = [
            ("№", 35),
            ("Тикер", 80),
            ("Название", 160),
            ("Тип", 70),
            ("Кол-во", 65),
            ("Ср.взв.\nцена", 90),
            ("Текущая", 90),
            ("P&L", 100),
            ("P&L %", 65),
            ("Доля\nкап.%", 55),
            ("Счета", 140),
        ]
        for txt, w in cols:
            ctk.CTkLabel(
                hdr,
                text=txt,
                width=w,
                font=ctk.CTkFont(size=10, weight="bold"),
                text_color="white",
            ).pack(side="left", padx=2, pady=4)

        # Строки
        type_ru = {
            "share": "Акция",
            "bond": "Облиг.",
            "etf": "ETF",
            "currency": "Валюта",
            "futures": "Фьюч.",
            "option": "Опцион",
        }

        for i, p in enumerate(self.aggregated):
            bg = "#2C3E50" if i % 2 == 0 else "#34495E"
            row = ctk.CTkFrame(sf, corner_radius=3, fg_color=bg)
            row.pack(fill="x", padx=2, pady=1)

            pnl_c = "#27AE60" if p.profit_loss >= 0 else "#E74C3C"
            s = "+" if p.profit_loss >= 0 else ""
            accounts_str = ", ".join(p.accounts)

            data = [
                (f"{i + 1}", 35, "#8899AA"),
                (p.ticker, 80, "white"),
                (p.name[:20], 160, "#ECF0F1"),
                (type_ru.get(p.instrument_type, p.instrument_type), 70, "#BDC3C7"),
                (f"{p.total_quantity:.0f}", 65, "white"),
                (f"{p.weighted_avg_price:,.2f}", 90, "white"),
                (f"{p.current_price:,.2f}", 90, "white"),
                (f"{s}{p.profit_loss:,.2f}", 100, pnl_c),
                (f"{s}{p.profit_loss_pct:.1f}%", 65, pnl_c),
                (f"{p.share_of_total:.1f}%", 55, "#3498DB"),
                (accounts_str[:20], 140, "#8899AA"),
            ]
            for txt, w, clr in data:
                ctk.CTkLabel(
                    row, text=txt, width=w, font=ctk.CTkFont(size=10), text_color=clr
                ).pack(side="left", padx=2, pady=3)

    def _fill_account_tab(self, parent, acc: AccountInfo):
        sf = ctk.CTkScrollableFrame(parent)
        sf.pack(fill="both", expand=True, padx=4, pady=4)

        if not acc.positions:
            ctk.CTkLabel(sf, text="📭 Нет позиций", font=ctk.CTkFont(size=14)).pack(
                expand=True, pady=40
            )
            return

        # Заголовок
        hdr = ctk.CTkFrame(sf, corner_radius=4, fg_color="#1F4E79")
        hdr.pack(fill="x", padx=2, pady=(2, 0))

        cols = [
            ("Тикер", 80),
            ("Название", 170),
            ("Кол-во", 65),
            ("Ср.цена", 95),
            ("Текущая", 95),
            ("P&L", 105),
            ("P&L %", 65),
            ("Доля сч.", 60),
            ("Доля кап.", 60),
        ]
        for txt, w in cols:
            ctk.CTkLabel(
                hdr,
                text=txt,
                width=w,
                font=ctk.CTkFont(size=11, weight="bold"),
                text_color="white",
            ).pack(side="left", padx=2, pady=4)

        for i, p in enumerate(acc.positions):
            bg = "#2C3E50" if i % 2 == 0 else "#34495E"
            row = ctk.CTkFrame(sf, corner_radius=3, fg_color=bg)
            row.pack(fill="x", padx=2, pady=1)

            pnl_c = "#27AE60" if p.profit_loss >= 0 else "#E74C3C"
            s = "+" if p.profit_loss >= 0 else ""

            data = [
                (p.ticker, 80, "white"),
                (p.name[:20], 170, "#ECF0F1"),
                (f"{p.quantity:.0f}", 65, "white"),
                (f"{p.average_price:,.2f}", 95, "white"),
                (f"{p.current_price:,.2f}", 95, "white"),
                (f"{s}{p.profit_loss:,.2f}", 105, pnl_c),
                (f"{s}{p.profit_loss_pct:.1f}%", 65, pnl_c),
                (f"{p.share_of_account:.1f}%", 60, "#3498DB"),
                (f"{p.share_of_total:.1f}%", 60, "#9B59B6"),
            ]
            for txt, w, clr in data:
                ctk.CTkLabel(
                    row, text=txt, width=w, font=ctk.CTkFont(size=11), text_color=clr
                ).pack(side="left", padx=2, pady=3)

    # === Excel ===

    def _export(self):
        if not self.accounts:
            messagebox.showwarning("Нет данных", "Сначала загрузите портфель")
            return
        fp = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=EXCEL_FILENAME,
            title="Сохранить отчёт",
        )
        if not fp:
            return
        self._set_busy(True, "⏳ Создание Excel…")

        def job():
            try:
                p = ExcelExporter().export(
                    self.accounts,
                    self.total_capital,
                    fp,
                    aggregated=self.aggregated,
                )
                self.after(0, lambda: self._on_exported(p))
            except Exception as e:
                err_msg = str(e)
                self.after(0, lambda: self._on_export_err(err_msg))

        threading.Thread(target=job, daemon=True).start()

    def _on_exported(self, path):
        self._set_busy(False)
        self.lbl_status.configure(
            text=f"✅ {os.path.basename(path)}", text_color="#27AE60"
        )
        if messagebox.askyesno("Готово", f"Сохранено:\n{path}\n\nОткрыть?"):
            if os.name == "nt":
                os.startfile(path)
            else:
                os.system(f'xdg-open "{path}" 2>/dev/null &')

    def _on_export_err(self, err):
        self._set_busy(False)
        messagebox.showerror("Ошибка", err)


def run_gui():
    app = PortfolioApp()
    app.mainloop()


if __name__ == "__main__":
    run_gui()
