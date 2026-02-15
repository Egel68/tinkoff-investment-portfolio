"""GUI на CustomTkinter."""

import math
import os
import threading
from decimal import Decimal
from tkinter import filedialog, messagebox

import customtkinter as ctk

from api_client import (
    AccountInfo,
    AggregatedPosition,
    TinkoffApiClient,
    TypeAllocation,
)
from config import EXCEL_FILENAME, TOKEN
from excel_export import ExcelExporter

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class PieChart(ctk.CTkCanvas):
    """Круговая donut-диаграмма."""

    def __init__(self, master, size=320, **kwargs):
        super().__init__(
            master,
            width=size,
            height=size,
            bg="#2b2b2b",
            highlightthickness=0,
            **kwargs,
        )
        self.size = size
        self.segments: list[tuple[float, str, str]] = []

    def set_data(self, segments: list[tuple[float, str, str]]):
        self.segments = segments
        self._draw()

    def _draw(self):
        self.delete("all")
        if not self.segments:
            return

        cx = self.size / 2
        cy = self.size / 2
        radius = self.size / 2 - 20
        inner_radius = radius * 0.55

        start_angle = 90

        for share_pct, color, label in self.segments:
            extent = share_pct / 100 * 360
            if extent < 0.3:
                start_angle -= extent
                continue

            self.create_arc(
                cx - radius,
                cy - radius,
                cx + radius,
                cy + radius,
                start=start_angle,
                extent=-extent,
                fill=color,
                outline="#2b2b2b",
                width=2,
            )

            if share_pct >= 3:
                mid_angle = math.radians(start_angle - extent / 2)
                label_r = (radius + inner_radius) / 2
                lx = cx + label_r * math.cos(mid_angle)
                ly = cy - label_r * math.sin(mid_angle)
                self.create_text(
                    lx,
                    ly,
                    text=f"{share_pct:.1f}%",
                    fill="white",
                    font=("Calibri", 10, "bold"),
                )

            start_angle -= extent

        self.create_oval(
            cx - inner_radius,
            cy - inner_radius,
            cx + inner_radius,
            cy + inner_radius,
            fill="#2b2b2b",
            outline="#2b2b2b",
        )

        self.create_text(
            cx,
            cy - 8,
            text="Распределение",
            fill="#ECF0F1",
            font=("Calibri", 11, "bold"),
        )
        self.create_text(
            cx,
            cy + 12,
            text="по типам",
            fill="#BDC3C7",
            font=("Calibri", 10),
        )


class PortfolioApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("📈 Т-Инвестиции — Анализатор портфеля")
        self.geometry("1200x800")
        self.minsize(1000, 650)

        self.client: TinkoffApiClient | None = None
        self.accounts: list[AccountInfo] = []
        self.aggregated: list[AggregatedPosition] = []
        self.allocations: list[TypeAllocation] = []
        self.total_capital = Decimal("0")

        self._build()
        if TOKEN:
            self.token_entry.insert(0, TOKEN)

    def _build(self):
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

        cap = ctk.CTkFrame(self, corner_radius=10)
        cap.pack(fill="x", padx=12, pady=4)
        self.lbl_capital = ctk.CTkLabel(
            cap, text="💰 Общий капитал: —", font=ctk.CTkFont(size=16, weight="bold")
        )
        self.lbl_capital.pack(pady=8)

        self.tabs = ctk.CTkTabview(self, corner_radius=10)
        self.tabs.pack(fill="both", expand=True, padx=12, pady=(4, 12))
        self.tabs.add("Обзор")
        ctk.CTkLabel(
            self.tabs.tab("Обзор"), text="Подключитесь к API", font=ctk.CTkFont(size=14)
        ).pack(expand=True)

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

    def _load(self):
        self._set_busy(True, "⏳ Загрузка портфеля…")

        def job():
            try:
                accs, tot = self.client.get_all_accounts()
                agg = TinkoffApiClient.aggregate_positions(accs, tot)
                alloc = TinkoffApiClient.calc_type_allocation(accs, tot)
                self.after(0, lambda: self._on_loaded(accs, tot, agg, alloc))
            except Exception as e:
                err_msg = str(e)
                self.after(0, lambda: self._on_load_err(err_msg))

        threading.Thread(target=job, daemon=True).start()

    def _on_loaded(self, accs, total, aggregated, allocations):
        self._set_busy(False)
        self.accounts = accs
        self.total_capital = total
        self.aggregated = aggregated
        self.allocations = allocations
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

    def _fill_tabs(self):
        for name in list(self.tabs._tab_dict.keys()):
            self.tabs.delete(name)

        t0 = self.tabs.add("📊 Сводка")
        self._fill_summary_tab(t0)

        t_alloc = self.tabs.add("🎯 По типам")
        self._fill_allocation_tab(t_alloc)

        t_all = self.tabs.add("📦 Все позиции")
        self._fill_aggregated_tab(t_all)

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

    def _fill_allocation_tab(self, parent):
        """Вкладка: диаграмма + легенда + таблица, деньги отдельно."""
        main_frame = ctk.CTkFrame(parent)
        main_frame.pack(fill="both", expand=True, padx=4, pady=4)

        # --- Разделяем на ценные бумаги и деньги ---
        securities = [a for a in self.allocations if a.instrument_type != "cash"]
        cash_items = [a for a in self.allocations if a.instrument_type == "cash"]

        total_securities = sum(a.total_market_cost for a in securities)
        total_cash = sum(a.total_market_cost for a in cash_items)

        # Левая часть — диаграмма
        left = ctk.CTkFrame(main_frame, width=380)
        left.pack(side="left", fill="y", padx=(8, 4), pady=8)
        left.pack_propagate(False)

        ctk.CTkLabel(
            left,
            text="📊 Распределение по типам",
            font=ctk.CTkFont(size=15, weight="bold"),
        ).pack(pady=(12, 4))

        # Подпись: ценные бумаги vs деньги
        sec_pct = (
            float(total_securities / self.total_capital * 100)
            if self.total_capital
            else 0
        )
        cash_pct = (
            float(total_cash / self.total_capital * 100) if self.total_capital else 0
        )
        ctk.CTkLabel(
            left,
            text=f"📈 Бумаги: {sec_pct:.1f}%  ·  💵 Деньги: {cash_pct:.1f}%",
            font=ctk.CTkFont(size=12),
            text_color="#BDC3C7",
        ).pack(pady=(0, 8))

        chart = PieChart(left, size=320)
        chart.pack(pady=8)

        # Данные для диаграммы — все типы включая cash
        segments = []
        for alloc in self.allocations:
            segments.append(
                (
                    float(alloc.share_of_total),
                    alloc.color,
                    alloc.type_name_ru,
                )
            )
        chart.set_data(segments)

        # Легенда под диаграммой
        legend_frame = ctk.CTkFrame(left, corner_radius=8, fg_color="#1a1a2e")
        legend_frame.pack(fill="x", padx=8, pady=(4, 8))

        for alloc in self.allocations:
            row = ctk.CTkFrame(legend_frame, fg_color="transparent")
            row.pack(fill="x", padx=8, pady=1)

            color_box = ctk.CTkCanvas(
                row,
                width=14,
                height=14,
                bg="#1a1a2e",
                highlightthickness=0,
            )
            color_box.pack(side="left", padx=(0, 6))
            color_box.create_rectangle(1, 1, 13, 13, fill=alloc.color, outline="")

            # Иконка для денег
            icon = "💵" if alloc.instrument_type == "cash" else "📄"

            ctk.CTkLabel(
                row,
                text=f"{icon} {alloc.type_name_ru} — {alloc.share_of_total:.1f}%",
                font=ctk.CTkFont(size=11),
                text_color="#ECF0F1" if alloc.instrument_type != "cash" else "#F39C12",
            ).pack(side="left")

        # Правая часть — детальная таблица
        right = ctk.CTkScrollableFrame(main_frame)
        right.pack(side="left", fill="both", expand=True, padx=(4, 8), pady=8)

        # --- Секция: Ценные бумаги ---
        ctk.CTkLabel(
            right,
            text="📈 Ценные бумаги",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(anchor="w", padx=8, pady=(8, 4))

        self._draw_allocation_table(right, securities)

        # --- Секция: Денежные средства ---
        if cash_items:
            sep = ctk.CTkFrame(right, height=2, fg_color="#555555")
            sep.pack(fill="x", padx=8, pady=12)

            ctk.CTkLabel(
                right,
                text="💵 Денежные средства",
                font=ctk.CTkFont(size=14, weight="bold"),
                text_color="#F39C12",
            ).pack(anchor="w", padx=8, pady=(0, 4))

            # Детали по деньгам
            cash_detail = ctk.CTkFrame(right, corner_radius=8, fg_color="#1a1a2e")
            cash_detail.pack(fill="x", padx=4, pady=(0, 8))

            total_cash_val = sum(a.total_market_cost for a in cash_items)
            cash_share = (
                float(total_cash_val / self.total_capital * 100)
                if self.total_capital
                else 0
            )

            ctk.CTkLabel(
                cash_detail,
                text=f"Всего денежных средств: {total_cash_val:,.2f} ₽  ({cash_share:.1f}% от капитала)",
                font=ctk.CTkFont(size=12, weight="bold"),
                text_color="#F39C12",
            ).pack(padx=12, pady=(10, 4))

            # Разбивка по валютам (из total_currencies всех счетов)
            all_currencies: dict[str, Decimal] = {}
            for acc in self.accounts:
                for cur, amt in acc.total_currencies.items():
                    all_currencies[cur] = all_currencies.get(cur, Decimal("0")) + amt

            if all_currencies:
                for cur, amt in sorted(all_currencies.items(), key=lambda x: -x[1]):
                    ctk.CTkLabel(
                        cash_detail,
                        text=f"  {cur}: {amt:,.2f}",
                        font=ctk.CTkFont(size=11),
                        text_color="#BDC3C7",
                    ).pack(anchor="w", padx=20, pady=1)

            ctk.CTkLabel(cash_detail, text="").pack(pady=2)

            self._draw_allocation_table(right, cash_items)

        # --- Итого ---
        sep2 = ctk.CTkFrame(right, height=2, fg_color="#555555")
        sep2.pack(fill="x", padx=8, pady=12)

        total_mkt = sum(a.total_market_cost for a in self.allocations)
        total_pnl = sum(a.profit_loss for a in self.allocations)
        pnl_sign = "+" if total_pnl >= 0 else ""
        pnl_color = "#27AE60" if total_pnl >= 0 else "#E74C3C"

        totals_frame = ctk.CTkFrame(right, corner_radius=8, fg_color="#1a1a2e")
        totals_frame.pack(fill="x", padx=4, pady=4)

        ctk.CTkLabel(
            totals_frame,
            text=(
                f"📊 ИТОГО:  Стоимость: {total_mkt:,.2f} ₽  |  "
                f"P&L: {pnl_sign}{total_pnl:,.2f} ₽"
            ),
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=pnl_color,
        ).pack(padx=12, pady=10)

    def _draw_allocation_table(self, parent, items: list[TypeAllocation]):
        """Рисуем таблицу распределения для группы типов."""
        if not items:
            return

        hdr = ctk.CTkFrame(parent, corner_radius=4, fg_color="#1F4E79")
        hdr.pack(fill="x", padx=4, pady=(2, 0))

        cols = [
            ("Тип", 130),
            ("Позиций", 60),
            ("Стоимость", 120),
            ("P&L", 110),
            ("P&L %", 65),
            ("Доля %", 60),
        ]
        for txt, w in cols:
            ctk.CTkLabel(
                hdr,
                text=txt,
                width=w,
                font=ctk.CTkFont(size=11, weight="bold"),
                text_color="white",
            ).pack(side="left", padx=2, pady=4)

        for i, alloc in enumerate(items):
            bg = "#2C3E50" if i % 2 == 0 else "#34495E"
            row = ctk.CTkFrame(parent, corner_radius=3, fg_color=bg)
            row.pack(fill="x", padx=4, pady=1)

            pnl_c = "#27AE60" if alloc.profit_loss >= 0 else "#E74C3C"
            s = "+" if alloc.profit_loss >= 0 else ""

            data = [
                (f"● {alloc.type_name_ru}", 130, alloc.color),
                (f"{alloc.positions_count}", 60, "white"),
                (f"{alloc.total_market_cost:,.0f} ₽", 120, "white"),
                (f"{s}{alloc.profit_loss:,.0f} ₽", 110, pnl_c),
                (f"{s}{alloc.profit_loss_pct:.1f}%", 65, pnl_c),
                (f"{alloc.share_of_total:.1f}%", 60, "#3498DB"),
            ]
            for txt, w, clr in data:
                ctk.CTkLabel(
                    row, text=txt, width=w, font=ctk.CTkFont(size=11), text_color=clr
                ).pack(side="left", padx=2, pady=4)

    def _fill_aggregated_tab(self, parent):
        sf = ctk.CTkScrollableFrame(parent)
        sf.pack(fill="both", expand=True, padx=4, pady=4)

        if not self.aggregated:
            ctk.CTkLabel(sf, text="📭 Нет позиций", font=ctk.CTkFont(size=14)).pack(
                expand=True, pady=40
            )
            return

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

            # Помечаем деньги
            name_display = p.name[:20]
            if p.is_cash:
                name_display = f"💵 {name_display}"

            data = [
                (f"{i + 1}", 35, "#8899AA"),
                (p.ticker, 80, "white"),
                (name_display, 160, "#F39C12" if p.is_cash else "#ECF0F1"),
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

            name_display = p.name[:20]
            if p.is_cash:
                name_display = f"💵 {name_display}"

            data = [
                (p.ticker, 80, "white"),
                (name_display, 170, "#F39C12" if p.is_cash else "#ECF0F1"),
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
                    allocations=self.allocations,
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
