"""
Клиент для работы с API Т-Инвестиций.
Библиотека: t-tech-investments (модуль t_tech.invest)
"""

from dataclasses import dataclass, field
from decimal import Decimal

from t_tech.invest import Client
from t_tech.invest.schemas import (
    Account,
    InstrumentIdType,
    MoneyValue,
    PortfolioPosition,
    Quotation,
)


def quotation_to_decimal(q) -> Decimal:
    """Quotation / MoneyValue -> Decimal."""
    if q is None:
        return Decimal("0")
    units = getattr(q, "units", 0) or 0
    nano = getattr(q, "nano", 0) or 0
    return Decimal(str(units)) + Decimal(str(nano)) / Decimal("1_000_000_000")


@dataclass
class PositionInfo:
    """Одна позиция в портфеле."""

    figi: str
    name: str
    ticker: str
    isin: str
    instrument_type: str
    quantity: Decimal
    average_price: Decimal
    current_price: Decimal
    currency: str
    average_cost: Decimal
    market_cost: Decimal
    expected_yield: Decimal
    profit_loss: Decimal
    profit_loss_pct: Decimal
    share_of_account: Decimal
    share_of_total: Decimal
    lot_size: int = 1
    sector: str = ""
    country: str = ""
    exchange: str = ""


@dataclass
class AggregatedPosition:
    """Агрегированная позиция по всем счетам."""

    figi: str
    name: str
    ticker: str
    instrument_type: str
    currency: str
    sector: str
    country: str
    total_quantity: Decimal
    total_avg_cost: Decimal  # сумма средних стоимостей
    total_market_cost: Decimal  # сумма рыночных стоимостей
    weighted_avg_price: Decimal  # средневзвешенная цена покупки
    current_price: Decimal
    profit_loss: Decimal
    profit_loss_pct: Decimal
    share_of_total: Decimal  # доля от всего капитала
    accounts: list[str] = field(default_factory=list)  # на каких счетах есть


@dataclass
class AccountInfo:
    """Один брокерский счёт."""

    account_id: str
    name: str
    acc_type: str
    status: str
    opened_date: str
    positions: list[PositionInfo] = field(default_factory=list)
    total_value: Decimal = Decimal("0")
    total_currencies: dict[str, Decimal] = field(default_factory=dict)


@dataclass
class TypeAllocation:
    """Распределение по типу инструмента."""

    instrument_type: str
    type_name_ru: str
    total_market_cost: Decimal
    total_avg_cost: Decimal
    profit_loss: Decimal
    profit_loss_pct: Decimal
    share_of_total: Decimal
    positions_count: int
    color: str  # цвет для диаграммы


class TinkoffApiClient:
    """Обёртка над t_tech.invest SDK."""

    ACCOUNT_TYPE_MAP = {
        0: "Не определён",
        1: "Брокерский",
        2: "ИИС",
        3: "Инвесткопилка",
        4: "ИИС нового типа",
    }

    ACCOUNT_STATUS_MAP = {
        0: "Не определён",
        1: "Новый",
        2: "Открыт",
        3: "Закрыт",
    }

    def __init__(self, token: str):
        self.token = token
        self._instr_cache: dict[str, dict] = {}

    def test_connection(self) -> tuple[bool, str]:
        try:
            with Client(self.token) as cl:
                accs = cl.users.get_accounts().accounts
                active = [a for a in accs if self._status_val(a) != 3]
                return True, f"Подключено. Счетов: {len(active)}"
        except Exception as e:
            return False, f"Ошибка: {e}"

    @staticmethod
    def _status_val(acc: Account) -> int:
        s = acc.status
        return s if isinstance(s, int) else getattr(s, "value", s)

    @staticmethod
    def _type_val(acc: Account) -> int:
        t = acc.type
        return t if isinstance(t, int) else getattr(t, "value", t)

    def _fetch_instrument(self, cl, figi: str) -> dict:
        if figi in self._instr_cache:
            return self._instr_cache[figi]

        info = dict(
            name="—",
            ticker="",
            isin="",
            lot_size=1,
            sector="",
            country="",
            exchange="",
            instrument_type="unknown",
        )

        try:
            resp = cl.instruments.get_instrument_by(
                id_type=InstrumentIdType.INSTRUMENT_ID_TYPE_FIGI,
                id=figi,
            )
            instr = resp.instrument
            info.update(
                name=instr.name or "—",
                ticker=instr.ticker or "",
                isin=getattr(instr, "isin", "") or "",
                lot_size=getattr(instr, "lot", 1) or 1,
                sector=getattr(instr, "sector", "") or "",
                country=getattr(instr, "country_of_risk_name", "") or "",
                exchange=getattr(instr, "exchange", "") or "",
                instrument_type=getattr(instr, "instrument_type", "unknown")
                or "unknown",
            )
            self._instr_cache[figi] = info
            return info
        except Exception:
            pass

        for method_name in (
            "share_by",
            "bond_by",
            "etf_by",
            "currency_by",
            "future_by",
        ):
            method = getattr(cl.instruments, method_name, None)
            if method is None:
                continue
            try:
                resp2 = method(
                    id_type=InstrumentIdType.INSTRUMENT_ID_TYPE_FIGI,
                    id=figi,
                )
                instr2 = resp2.instrument
                info.update(
                    name=instr2.name or "—",
                    ticker=instr2.ticker or "",
                    isin=getattr(instr2, "isin", "") or "",
                    lot_size=getattr(instr2, "lot", 1) or 1,
                    sector=getattr(instr2, "sector", "") or "",
                    country=getattr(instr2, "country_of_risk_name", "") or "",
                    exchange=getattr(instr2, "exchange", "") or "",
                    instrument_type=method_name.replace("_by", ""),
                )
                break
            except Exception:
                continue

        self._instr_cache[figi] = info
        return info

    def get_all_accounts(self) -> tuple[list[AccountInfo], Decimal]:
        accounts_out: list[AccountInfo] = []

        with Client(self.token) as cl:
            raw_accounts = cl.users.get_accounts().accounts

            portfolios = {}
            total_capital = Decimal("0")

            for acc in raw_accounts:
                if self._status_val(acc) == 3:
                    continue
                try:
                    pf = cl.operations.get_portfolio(account_id=acc.id)
                    portfolios[acc.id] = pf
                    total_capital += quotation_to_decimal(pf.total_amount_portfolio)
                except Exception as e:
                    print(f"⚠ Портфель {acc.id}: {e}")

            for acc in raw_accounts:
                if acc.id not in portfolios:
                    continue

                pf = portfolios[acc.id]
                acc_type_str = self.ACCOUNT_TYPE_MAP.get(
                    self._type_val(acc), f"Тип {self._type_val(acc)}"
                )
                acc_status_str = self.ACCOUNT_STATUS_MAP.get(
                    self._status_val(acc), f"Статус {self._status_val(acc)}"
                )

                opened = "—"
                if acc.opened_date:
                    try:
                        opened = acc.opened_date.strftime("%d.%m.%Y")
                    except Exception:
                        opened = str(acc.opened_date)

                account_total = quotation_to_decimal(pf.total_amount_portfolio)

                acc_info = AccountInfo(
                    account_id=acc.id,
                    name=acc.name or f"Счёт …{acc.id[-4:]}",
                    acc_type=acc_type_str,
                    status=acc_status_str,
                    opened_date=opened,
                    total_value=account_total,
                )

                for pos in pf.positions:
                    figi = pos.figi or ""

                    instr = (
                        self._fetch_instrument(cl, figi)
                        if figi
                        else dict(
                            name="—",
                            ticker="",
                            isin="",
                            lot_size=1,
                            sector="",
                            country="",
                            exchange="",
                            instrument_type="unknown",
                        )
                    )

                    quantity = quotation_to_decimal(pos.quantity)
                    avg_price = quotation_to_decimal(
                        getattr(pos, "average_position_price", None)
                    )
                    current_price = quotation_to_decimal(
                        getattr(pos, "current_price", None)
                    )

                    currency = ""
                    avg_money = getattr(pos, "average_position_price", None)
                    if avg_money and hasattr(avg_money, "currency"):
                        currency = avg_money.currency or ""

                    expected_yield = quotation_to_decimal(
                        getattr(pos, "expected_yield", None)
                    )

                    avg_cost = avg_price * quantity
                    market_cost = current_price * quantity

                    profit_loss = market_cost - avg_cost
                    profit_loss_pct = (
                        (profit_loss / avg_cost * Decimal("100"))
                        if avg_cost != 0
                        else Decimal("0")
                    )
                    share_account = (
                        (market_cost / account_total * Decimal("100"))
                        if account_total != 0
                        else Decimal("0")
                    )
                    share_total = (
                        (market_cost / total_capital * Decimal("100"))
                        if total_capital != 0
                        else Decimal("0")
                    )

                    position = PositionInfo(
                        figi=figi,
                        name=instr["name"],
                        ticker=instr["ticker"],
                        isin=instr["isin"],
                        instrument_type=instr["instrument_type"],
                        quantity=quantity,
                        average_price=avg_price,
                        current_price=current_price,
                        currency=currency,
                        average_cost=avg_cost,
                        market_cost=market_cost,
                        expected_yield=expected_yield,
                        profit_loss=profit_loss,
                        profit_loss_pct=profit_loss_pct,
                        share_of_account=share_account,
                        share_of_total=share_total,
                        lot_size=instr["lot_size"],
                        sector=instr["sector"],
                        country=instr["country"],
                        exchange=instr["exchange"],
                    )
                    acc_info.positions.append(position)

                try:
                    pos_resp = cl.operations.get_positions(account_id=acc.id)
                    for m in pos_resp.money:
                        curr_code = (m.currency or "???").upper()
                        acc_info.total_currencies[curr_code] = quotation_to_decimal(m)
                except Exception:
                    pass

                acc_info.positions.sort(key=lambda p: p.market_cost, reverse=True)
                accounts_out.append(acc_info)

        return accounts_out, total_capital

    @staticmethod
    def aggregate_positions(
        accounts: list[AccountInfo], total_capital: Decimal
    ) -> list[AggregatedPosition]:
        """
        Агрегация одинаковых бумаг со всех счетов.
        Группировка по FIGI (или ticker+name если figi пустой).
        """
        aggregated: dict[str, dict] = {}

        for acc in accounts:
            for pos in acc.positions:
                key = pos.figi if pos.figi else f"{pos.ticker}_{pos.name}"

                if key not in aggregated:
                    aggregated[key] = {
                        "figi": pos.figi,
                        "name": pos.name,
                        "ticker": pos.ticker,
                        "instrument_type": pos.instrument_type,
                        "currency": pos.currency,
                        "sector": pos.sector,
                        "country": pos.country,
                        "current_price": pos.current_price,
                        "total_quantity": Decimal("0"),
                        "total_avg_cost": Decimal("0"),
                        "total_market_cost": Decimal("0"),
                        "accounts": [],
                    }

                agg = aggregated[key]
                agg["total_quantity"] += pos.quantity
                agg["total_avg_cost"] += pos.average_cost
                agg["total_market_cost"] += pos.market_cost
                agg["current_price"] = pos.current_price  # обновляем до последней

                if acc.name not in agg["accounts"]:
                    agg["accounts"].append(acc.name)

        result = []
        for key, agg in aggregated.items():
            total_qty = agg["total_quantity"]
            total_avg = agg["total_avg_cost"]
            total_mkt = agg["total_market_cost"]

            weighted_avg_price = (
                (total_avg / total_qty) if total_qty != 0 else Decimal("0")
            )

            pnl = total_mkt - total_avg
            pnl_pct = (
                (pnl / total_avg * Decimal("100")) if total_avg != 0 else Decimal("0")
            )

            share = (
                (total_mkt / total_capital * Decimal("100"))
                if total_capital != 0
                else Decimal("0")
            )

            result.append(
                AggregatedPosition(
                    figi=agg["figi"],
                    name=agg["name"],
                    ticker=agg["ticker"],
                    instrument_type=agg["instrument_type"],
                    currency=agg["currency"],
                    sector=agg["sector"],
                    country=agg["country"],
                    total_quantity=total_qty,
                    total_avg_cost=total_avg,
                    total_market_cost=total_mkt,
                    weighted_avg_price=weighted_avg_price,
                    current_price=agg["current_price"],
                    profit_loss=pnl,
                    profit_loss_pct=pnl_pct,
                    share_of_total=share,
                    accounts=agg["accounts"],
                )
            )

        result.sort(key=lambda p: p.total_market_cost, reverse=True)
        return result

    TYPE_COLORS = {
        "share": "#3498DB",
        "bond": "#2ECC71",
        "etf": "#E67E22",
        "currency": "#9B59B6",
        "futures": "#E74C3C",
        "option": "#1ABC9C",
        "sp": "#F39C12",
        "unknown": "#95A5A6",
    }

    TYPE_NAMES_RU = {
        "share": "Акции",
        "bond": "Облигации",
        "etf": "Фонды (ETF)",
        "currency": "Валюта",
        "futures": "Фьючерсы",
        "option": "Опционы",
        "sp": "Структ. продукты",
        "unknown": "Прочее",
    }

    @classmethod
    def calc_type_allocation(
        cls,
        accounts: list[AccountInfo],
        total_capital: Decimal,
    ) -> list[TypeAllocation]:
        """Распределение портфеля по типам инструментов."""
        buckets: dict[str, dict] = {}

        for acc in accounts:
            for pos in acc.positions:
                t = pos.instrument_type or "unknown"

                if t not in buckets:
                    buckets[t] = {
                        "total_market_cost": Decimal("0"),
                        "total_avg_cost": Decimal("0"),
                        "count": 0,
                    }

                buckets[t]["total_market_cost"] += pos.market_cost
                buckets[t]["total_avg_cost"] += pos.average_cost
                buckets[t]["count"] += 1

        result = []
        # Предопределённый порядок цветов для нестандартных типов
        extra_colors = [
            "#D35400",
            "#8E44AD",
            "#16A085",
            "#C0392B",
            "#2980B9",
            "#7F8C8D",
        ]
        extra_idx = 0

        for t, data in buckets.items():
            mkt = data["total_market_cost"]
            avg = data["total_avg_cost"]
            pnl = mkt - avg
            pnl_pct = (pnl / avg * Decimal("100")) if avg != 0 else Decimal("0")
            share = (
                (mkt / total_capital * Decimal("100"))
                if total_capital != 0
                else Decimal("0")
            )

            color = cls.TYPE_COLORS.get(t)
            if not color:
                color = extra_colors[extra_idx % len(extra_colors)]
                extra_idx += 1

            result.append(
                TypeAllocation(
                    instrument_type=t,
                    type_name_ru=cls.TYPE_NAMES_RU.get(t, t),
                    total_market_cost=mkt,
                    total_avg_cost=avg,
                    profit_loss=pnl,
                    profit_loss_pct=pnl_pct,
                    share_of_total=share,
                    positions_count=data["count"],
                    color=color,
                )
            )

        result.sort(key=lambda x: x.total_market_cost, reverse=True)
        return result
