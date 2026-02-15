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
        """Проверка токена."""
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
        """Получение информации по FIGI с кешированием."""
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

        # Способ 1: универсальный get_instrument_by
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

        # Способ 2: перебор по типам
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
        """Загрузка всех счетов с позициями."""
        accounts_out: list[AccountInfo] = []

        with Client(self.token) as cl:
            raw_accounts = cl.users.get_accounts().accounts

            # 1) Собираем портфели, считаем общий капитал
            portfolios = {}
            total_capital = Decimal("0")

            for acc in raw_accounts:
                if self._status_val(acc) == 3:  # закрытый
                    continue
                try:
                    pf = cl.operations.get_portfolio(account_id=acc.id)
                    portfolios[acc.id] = pf
                    total_capital += quotation_to_decimal(pf.total_amount_portfolio)
                except Exception as e:
                    print(f"⚠ Портфель {acc.id}: {e}")

            # 2) Обработка каждого счёта
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

                # 3) Позиции
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

                # 4) Денежные остатки
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
