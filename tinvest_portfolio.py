#!/usr/bin/env python3
"""
Скрипт для экспорта портфеля Т-Инвестиций в Excel.

Требования:
- Python 3.8+
- Установить SDK: pip install t-tech-investments --index-url https://opensource.tbank.ru/api/v4/projects/238/packages/pypi/simple
- Установить openpyxl: pip install openpyxl

Использование:
1. Получите токен доступа на https://www.tbank.ru/invest/settings/
2. Установите переменную окружения TINKOFF_INVEST_TOKEN или передайте токен при запуске
3. Запустите скрипт: python tinvest_portfolio.py --token ВАШ_ТОКЕН --output portfolio.xlsx
"""

import argparse
import os
import sys
from datetime import datetime
from decimal import Decimal
from typing import Optional

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("Ошибка: Не установлен пакет openpyxl")
    print("Установите: pip install openpyxl")
    sys.exit(1)

try:
    from tinkoff.invest import (
        AccountType,
        Client,
        InstrumentType,
        MoneyValue,
        PortfolioPosition,
        Quotation,
    )
    from tinkoff.invest.utils import money_to_decimal, quotation_to_decimal
except ImportError:
    print("Ошибка: Не установлен SDK Т-Инвестиций")
    print(
        "Установите: pip install t-tech-investments --index-url https://opensource.tbank.ru/api/v4/projects/238/packages/pypi/simple"
    )
    sys.exit(1)


def get_accounts(client: Client) -> list:
    """Получить список всех счетов пользователя."""
    accounts = []
    response = client.users.get_accounts()
    for account in response.accounts:
        accounts.append(
            {
                "id": account.id,
                "name": account.name,
                "type": account.type,
                "status": account.status,
            }
        )
    return accounts


def get_portfolio(client: Client, account_id: str) -> dict:
    """Получить портфель по конкретному счету."""
    portfolio = client.operations.get_portfolio(account_id=account_id)
    return portfolio


def get_instrument_info(
    client: Client, figi: str, instrument_type: InstrumentType
) -> dict:
    """Получить информацию об инструменте."""
    try:
        if instrument_type == InstrumentType.INSTRUMENT_TYPE_SHARE:
            instrument = client.instruments.get_instrument(
                id_type=1,  # FIGI
                id=figi,
            ).instrument
            return {
                "ticker": instrument.ticker,
                "name": instrument.name,
                "lot": instrument.lot,
                "currency": instrument.currency,
                "exchange": instrument.exchange,
            }
        elif instrument_type == InstrumentType.INSTRUMENT_TYPE_BOND:
            instrument = client.instruments.get_instrument(
                id_type=1, id=figi
            ).instrument
            return {
                "ticker": instrument.ticker,
                "name": instrument.name,
                "lot": instrument.lot,
                "currency": instrument.currency,
                "exchange": instrument.exchange,
            }
        elif instrument_type == InstrumentType.INSTRUMENT_TYPE_ETF:
            instrument = client.instruments.get_instrument(
                id_type=1, id=figi
            ).instrument
            return {
                "ticker": instrument.ticker,
                "name": instrument.name,
                "lot": instrument.lot,
                "currency": instrument.currency,
                "exchange": instrument.exchange,
            }
        elif instrument_type == InstrumentType.INSTRUMENT_TYPE_CURRENCY:
            return {
                "ticker": figi,
                "name": f"Валюта {figi}",
                "lot": 1,
                "currency": "RUB",
                "exchange": "",
            }
        else:
            return {
                "ticker": figi,
                "name": "Неизвестный инструмент",
                "lot": 1,
                "currency": "RUB",
                "exchange": "",
            }
    except Exception as e:
        print(f"Ошибка при получении информации об инструменте {figi}: {e}")
        return {
            "ticker": figi,
            "name": f"Инструмент {figi}",
            "lot": 1,
            "currency": "RUB",
            "exchange": "",
        }


def get_current_prices(client: Client, figi_list: list) -> dict:
    """Получить текущие рыночные цены для списка инструментов."""
    prices = {}
    try:
        # Получаем цены через get_last_prices
        response = client.market_data.get_last_prices(figi=figi_list)
        for price_info in response.last_prices:
            if price_info.price:
                prices[price_info.figi] = quotation_to_decimal(price_info.price)
    except Exception as e:
        print(f"Ошибка при получении цен: {e}")
    return prices


def decimal_to_float(value) -> float:
    """Безопасно конвертировать Decimal в float."""
    if value is None:
        return 0.0
    try:
        return float(value)
    except:
        return 0.0


def money_to_float(money: Optional[MoneyValue]) -> float:
    """Конвертировать MoneyValue в float."""
    if money is None:
        return 0.0
    try:
        return float(money.units) + float(money.nano) / 1_000_000_000
    except:
        return 0.0


def quotation_to_float(quot: Optional[Quotation]) -> float:
    """Конвертировать Quotation в float."""
    if quot is None:
        return 0.0
    try:
        return float(quot.units) + float(quot.nano) / 1_000_000_000
    except:
        return 0.0


def process_portfolio(client: Client, account_id: str) -> tuple:
    """
    Обработать портфель и вернуть данные о позициях и общую сумму.

    Возвращает:
        - positions: список словарей с информацией о позициях
        - total_value: общая стоимость портфеля в рублях
    """
    portfolio = get_portfolio(client, account_id)

    positions = []
    total_value_rub = 0.0

    # Собираем все FIGI для получения цен
    figi_list = [pos.figi for pos in portfolio.positions if pos.figi]

    # Получаем текущие рыночные цены
    current_prices = get_current_prices(client, figi_list)

    # Обрабатываем каждую позицию
    for pos in portfolio.positions:
        try:
            # Информация об инструменте
            instrument_info = get_instrument_info(client, pos.figi, pos.instrument_type)

            # Количество в лотах
            quantity = quotation_to_float(pos.quantity)
            quantity_lots = (
                quotation_to_float(pos.quantity_lots) if pos.quantity_lots else quantity
            )

            # Средняя цена покупки (за единицу)
            average_price = money_to_float(pos.average_position_price)

            # Текущая рыночная цена
            current_price = decimal_to_float(current_prices.get(pos.figi, 0))
            if current_price == 0:
                current_price = money_to_float(pos.current_price)

            # Стоимость позиции по текущей цене
            market_value = money_to_float(pos.position_price)

            # Средняя стоимость позиции в портфеле
            average_value = quantity * average_price if quantity > 0 else 0

            # Валюта
            currency = (
                pos.average_position_price.currency
                if pos.average_position_price
                else "RUB"
            )

            # Тип инструмента
            instrument_type = (
                pos.instrument_type.name if pos.instrument_type else "UNKNOWN"
            )

            positions.append(
                {
                    "figi": pos.figi,
                    "ticker": instrument_info["ticker"],
                    "name": instrument_info["name"],
                    "instrument_type": instrument_type,
                    "quantity": quantity,
                    "quantity_lots": quantity_lots,
                    "lot": instrument_info["lot"],
                    "average_price": average_price,
                    "current_price": current_price,
                    "market_value": market_value,
                    "average_value": average_value,
                    "currency": currency,
                    "exchange": instrument_info["exchange"],
                }
            )

            total_value_rub += market_value

        except Exception as e:
            print(f"Ошибка при обработке позиции {pos.figi}: {e}")
            continue

    return positions, total_value_rub


def create_excel_report(accounts_data: list, output_path: str, total_capital: float):
    """Создать Excel отчёт с листами для каждого счёта."""
    wb = Workbook()

    # Удаляем пустой лист, созданный по умолчанию
    default_sheet = wb.active

    # Стили
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(
        start_color="1F4E79", end_color="1F4E79", fill_type="solid"
    )
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    data_alignment = Alignment(horizontal="right", vertical="center")
    text_alignment = Alignment(horizontal="left", vertical="center")

    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    # Заголовки таблицы
    headers = [
        "Тикер",
        "Название",
        "Тип",
        "Количество",
        "Лотов",
        "Валюта",
        "Ср. цена покупки",
        "Текущая цена",
        "Ср. стоимость",
        "Рын. стоимость",
        "Доля от счета, %",
        "Доля от капитала, %",
    ]

    for account_data in accounts_data:
        account_name = account_data["account_name"]
        account_id = account_data["account_id"]
        positions = account_data["positions"]
        account_total = account_data["total_value"]

        # Создаём лист для счёта
        # Очищаем имя листа от недопустимых символов
        sheet_name = account_name if account_name else f"Счет_{account_id[:8]}"
        sheet_name = (
            sheet_name.replace("/", "_")
            .replace("\\", "_")
            .replace("?", "_")
            .replace("*", "_")
            .replace("[", "_")
            .replace("]", "_")[:31]
        )

        ws = wb.create_sheet(title=sheet_name)

        # Заголовок листа
        ws.merge_cells("A1:L1")
        ws["A1"] = f"Портфель: {account_name or account_id}"
        ws["A1"].font = Font(bold=True, size=14)
        ws["A1"].alignment = Alignment(horizontal="center")

        # Информация о счёте
        ws.merge_cells("A2:L2")
        ws["A2"] = (
            f"ID счёта: {account_id} | Общая стоимость: {account_total:,.2f} руб. | Дата: {datetime.now().strftime('%d.%m.%Y %H:%M')}"
        )
        ws["A2"].alignment = Alignment(horizontal="center")
        ws["A2"].font = Font(italic=True)

        # Заголовки таблицы (строка 4)
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=4, column=col)
            cell.value = header
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
            cell.border = thin_border

        # Данные
        for row_idx, pos in enumerate(positions, 5):
            # Доля от счета
            share_account = (
                (pos["market_value"] / account_total * 100) if account_total > 0 else 0
            )
            # Доля от всего капитала
            share_capital = (
                (pos["market_value"] / total_capital * 100) if total_capital > 0 else 0
            )

            row_data = [
                pos["ticker"],
                pos["name"],
                pos["instrument_type"],
                pos["quantity"],
                pos["quantity_lots"],
                pos["currency"],
                pos["average_price"],
                pos["current_price"],
                pos["average_value"],
                pos["market_value"],
                share_account,
                share_capital,
            ]

            for col, value in enumerate(row_data, 1):
                cell = ws.cell(row=row_idx, column=col)
                cell.value = value
                cell.border = thin_border

                # Форматирование
                if col in [7, 8, 9, 10]:  # Цены и стоимости
                    cell.number_format = "#,##0.00"
                    cell.alignment = data_alignment
                elif col in [11, 12]:  # Проценты
                    cell.number_format = "0.00%"
                    cell.alignment = data_alignment
                    cell.value = value / 100  # Для правильного отображения процентов
                elif col in [4, 5]:  # Количество
                    cell.number_format = "#,##0.00"
                    cell.alignment = data_alignment
                else:
                    cell.alignment = text_alignment

        # Итоговая строка
        total_row = len(positions) + 5
        ws.merge_cells(f"A{total_row}:J{total_row}")
        ws[f"A{total_row}"] = "ИТОГО:"
        ws[f"A{total_row}"].font = Font(bold=True)
        ws[f"A{total_row}"].alignment = Alignment(horizontal="right")
        ws[f"K{total_row}"] = 1.0  # 100%
        ws[f"K{total_row}"].number_format = "0.00%"
        ws[f"L{total_row}"] = account_total / total_capital if total_capital > 0 else 0
        ws[f"L{total_row}"].number_format = "0.00%"

        # Настраиваем ширину столбцов
        column_widths = [12, 40, 15, 12, 10, 8, 15, 15, 15, 15, 15, 15]
        for i, width in enumerate(column_widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = width

        # Закрепляем заголовок
        ws.freeze_panes = "A5"

    # Создаём итоговый лист
    summary_sheet = wb.create_sheet(title="Сводка", index=0)

    # Заголовок
    summary_sheet.merge_cells("A1:E1")
    summary_sheet["A1"] = "СВОДНЫЙ ОТЧЁТ ПО ПОРТФЕЛЮ Т-ИНВЕСТИЦИЙ"
    summary_sheet["A1"].font = Font(bold=True, size=16)
    summary_sheet["A1"].alignment = Alignment(horizontal="center")

    summary_sheet["A2"] = (
        f"Дата формирования: {datetime.now().strftime('%d.%m.%Y %H:%M')}"
    )
    summary_sheet["A2"].font = Font(italic=True)

    # Заголовки таблицы счетов
    summary_headers = ["Счёт", "ID", "Тип", "Стоимость, руб.", "Доля от капитала, %"]
    for col, header in enumerate(summary_headers, 1):
        cell = summary_sheet.cell(row=4, column=col)
        cell.value = header
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border

    # Данные по счетам
    for row_idx, account_data in enumerate(accounts_data, 5):
        share = (
            (account_data["total_value"] / total_capital * 100)
            if total_capital > 0
            else 0
        )
        row_data = [
            account_data["account_name"] or "Без названия",
            account_data["account_id"],
            account_data["account_type"],
            account_data["total_value"],
            share / 100,
        ]

        for col, value in enumerate(row_data, 1):
            cell = summary_sheet.cell(row=row_idx, column=col)
            cell.value = value
            cell.border = thin_border
            if col == 4:
                cell.number_format = "#,##0.00"
                cell.alignment = data_alignment
            elif col == 5:
                cell.number_format = "0.00%"
                cell.alignment = data_alignment
            else:
                cell.alignment = text_alignment

    # Итоговая строка
    total_row = len(accounts_data) + 5
    summary_sheet.merge_cells(f"A{total_row}:C{total_row}")
    summary_sheet[f"A{total_row}"] = "ОБЩИЙ КАПИТАЛ:"
    summary_sheet[f"A{total_row}"].font = Font(bold=True)
    summary_sheet[f"A{total_row}"].alignment = Alignment(horizontal="right")
    summary_sheet[f"D{total_row}"] = total_capital
    summary_sheet[f"D{total_row}"].number_format = "#,##0.00"
    summary_sheet[f"D{total_row}"].font = Font(bold=True)
    summary_sheet[f"E{total_row}"] = 1.0
    summary_sheet[f"E{total_row}"].number_format = "0.00%"
    summary_sheet[f"E{total_row}"].font = Font(bold=True)

    # Ширина столбцов
    summary_widths = [30, 40, 20, 20, 20]
    for i, width in enumerate(summary_widths, 1):
        summary_sheet.column_dimensions[get_column_letter(i)].width = width

    # Удаляем пустой лист
    if default_sheet and default_sheet.title == "Sheet":
        wb.remove(default_sheet)

    # Сохраняем файл
    wb.save(output_path)
    print(f"\n✓ Отчёт успешно сохранён в файл: {output_path}")


def main():
    parser = argparse.ArgumentParser(
        description="Экспорт портфеля Т-Инвестиций в Excel",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Примеры использования:
  python tinvest_portfolio.py --token ВАШ_ТОКЕН
  python tinvest_portfolio.py --token ВАШ_ТОКЕН --output my_portfolio.xlsx
  export TINKOFF_INVEST_TOKEN=ВАШ_ТОКЕН && python tinvest_portfolio.py
        """,
    )
    parser.add_argument("--token", type=str, help="Токен доступа Т-Инвестиций")
    parser.add_argument(
        "--output",
        type=str,
        default="tinvest_portfolio.xlsx",
        help="Путь к выходному Excel файлу (по умолчанию: tinvest_portfolio.xlsx)",
    )

    args = parser.parse_args()

    # Получаем токен
    token = args.token or os.environ.get("TINKOFF_INVEST_TOKEN")
    if not token:
        print("Ошибка: Не указан токен доступа!")
        print("Получите токен на https://www.tbank.ru/invest/settings/")
        print(
            "И передайте его через аргумент --token или переменную окружения TINKOFF_INVEST_TOKEN"
        )
        sys.exit(1)

    print("=" * 60)
    print("  ЭКСПОРТ ПОРТФЕЛЯ Т-ИНВЕСТИЦИЙ В EXCEL")
    print("=" * 60)

    try:
        with Client(token) as client:
            # Получаем список счетов
            print("\n→ Получение списка счетов...")
            accounts = get_accounts(client)
            print(f"  Найдено счетов: {len(accounts)}")

            if not accounts:
                print("Ошибка: Не найдено ни одного счёта!")
                sys.exit(1)

            # Обрабатываем каждый счёт
            accounts_data = []
            total_capital = 0.0

            for account in accounts:
                print(f"\n→ Обработка счёта: {account['name'] or account['id']}")

                positions, account_total = process_portfolio(client, account["id"])

                account_type_name = {
                    1: "Брокерский",
                    2: "ИИС",
                    3: "Инвесткопилка",
                }.get(account["type"], "Другой")

                accounts_data.append(
                    {
                        "account_id": account["id"],
                        "account_name": account["name"],
                        "account_type": account_type_name,
                        "positions": positions,
                        "total_value": account_total,
                    }
                )

                total_capital += account_total
                print(f"  Позиций: {len(positions)}")
                print(f"  Стоимость: {account_total:,.2f} руб.")

            # Создаём Excel отчёт
            print(f"\n→ Создание Excel отчёта...")
            print(f"  Общий капитал: {total_capital:,.2f} руб.")

            create_excel_report(accounts_data, args.output, total_capital)

            print("\n" + "=" * 60)
            print("  ГОТОВО!")
            print("=" * 60)

    except Exception as e:
        print(f"\nОшибка: {e}")
        import traceback

        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
