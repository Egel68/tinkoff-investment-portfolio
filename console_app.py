"""Консольный интерфейс."""

import sys
from getpass import getpass

from api_client import TinkoffApiClient
from config import EXCEL_FILENAME, TOKEN
from excel_export import ExcelExporter


def banner():
    print("=" * 60)
    print("  📈 Т-Инвестиции — Анализатор портфеля")
    print("=" * 60)


def menu():
    print("\n  1. Проверить подключение")
    print("  2. Показать портфель (по счетам)")
    print("  3. Показать все позиции (агрегированные)")
    print("  4. Распределение по типам")
    print("  5. Экспортировать в Excel")
    print("  6. Всё: портфель + позиции + типы + Excel")
    print("  7. Сменить токен")
    print("  0. Выход\n")


def show_portfolio(accounts, total):
    print(f"\n{'=' * 80}")
    print(f"  💰 Общий капитал: {total:,.2f} ₽    Счетов: {len(accounts)}")
    print(f"{'=' * 80}")

    for acc in accounts:
        print(f"\n{'─' * 80}")
        print(f"  🏦 {acc.name} ({acc.acc_type}) — {acc.status}")
        print(f"  📅 Открыт: {acc.opened_date}   💵 Портфель: {acc.total_value:,.2f} ₽")

        if acc.total_currencies:
            parts = ", ".join(f"{c}: {a:,.2f}" for c, a in acc.total_currencies.items())
            print(f"  💰 Деньги: {parts}")

        if not acc.positions:
            print("  📭 Нет позиций")
            continue

        print(
            f"\n  {'№':>3} {'Тикер':<10} {'Название':<25} {'Кол':>6} "
            f"{'Ср.цена':>11} {'Текущая':>11} {'P&L':>12} {'P&L%':>8} {'Доля':>6}"
        )
        print(f"  {'─' * 100}")

        for i, p in enumerate(acc.positions, 1):
            s = "+" if p.profit_loss >= 0 else ""
            print(
                f"  {i:>3} {p.ticker:<10} {p.name[:25]:<25} "
                f"{p.quantity:>6.0f} {p.average_price:>11,.2f} "
                f"{p.current_price:>11,.2f} {s}{p.profit_loss:>11,.2f} "
                f"{s}{p.profit_loss_pct:>7.2f}% {p.share_of_account:>5.1f}%"
            )


def show_aggregated(aggregated, total):
    print(f"\n{'=' * 100}")
    print(f"  📦 Все позиции (агрегированные)    Уникальных бумаг: {len(aggregated)}")
    print(f"{'=' * 100}")

    print(
        f"\n  {'№':>3} {'Тикер':<10} {'Название':<25} {'Кол':>7} "
        f"{'Ср.взв.':>11} {'Текущая':>11} {'P&L':>12} {'P&L%':>8} "
        f"{'Доля%':>6} {'Счета'}"
    )
    print(f"  {'─' * 110}")

    for i, p in enumerate(aggregated, 1):
        s = "+" if p.profit_loss >= 0 else ""
        accs = ", ".join(p.accounts)
        print(
            f"  {i:>3} {p.ticker:<10} {p.name[:25]:<25} "
            f"{p.total_quantity:>7.0f} {p.weighted_avg_price:>11,.2f} "
            f"{p.current_price:>11,.2f} {s}{p.profit_loss:>11,.2f} "
            f"{s}{p.profit_loss_pct:>7.2f}% {p.share_of_total:>5.1f}% "
            f"{accs}"
        )


def show_allocation(allocations, total):
    print(f"\n{'=' * 80}")
    print(f"  🎯 Распределение по типам инструментов")
    print(f"{'=' * 80}")

    print(
        f"\n  {'Тип':<20} {'Позиций':>8} {'Стоимость':>15} "
        f"{'P&L':>12} {'P&L%':>8} {'Доля':>7}"
    )
    print(f"  {'─' * 75}")

    for a in allocations:
        s = "+" if a.profit_loss >= 0 else ""
        bar_len = int(float(a.share_of_total) / 2)
        bar = "█" * bar_len
        print(
            f"  {a.type_name_ru:<20} {a.positions_count:>8} "
            f"{a.total_market_cost:>15,.2f} "
            f"{s}{a.profit_loss:>11,.2f} "
            f"{s}{a.profit_loss_pct:>7.2f}% "
            f"{a.share_of_total:>6.1f}%  {bar}"
        )


def run_console():
    banner()
    token = TOKEN
    if not token:
        print("\n⚠️  Токен не задан (TINKOFF_TOKEN).")
        token = getpass("Введите токен: ").strip()

    client = TinkoffApiClient(token)

    while True:
        menu()
        ch = input("▶ ").strip()

        if ch == "1":
            print("⏳ Проверка…")
            ok, msg = client.test_connection()
            print(f"  {'✅' if ok else '❌'} {msg}")

        elif ch == "2":
            print("⏳ Загрузка…")
            try:
                accs, tot = client.get_all_accounts()
                show_portfolio(accs, tot)
            except Exception as e:
                print(f"  ❌ {e}")

        elif ch == "3":
            print("⏳ Загрузка…")
            try:
                accs, tot = client.get_all_accounts()
                agg = TinkoffApiClient.aggregate_positions(accs, tot)
                show_aggregated(agg, tot)
            except Exception as e:
                print(f"  ❌ {e}")

        elif ch == "4":
            print("⏳ Загрузка…")
            try:
                accs, tot = client.get_all_accounts()
                alloc = TinkoffApiClient.calc_type_allocation(accs, tot)
                show_allocation(alloc, tot)
            except Exception as e:
                print(f"  ❌ {e}")

        elif ch == "5":
            print("⏳ Экспорт…")
            try:
                accs, tot = client.get_all_accounts()
                agg = TinkoffApiClient.aggregate_positions(accs, tot)
                alloc = TinkoffApiClient.calc_type_allocation(accs, tot)
                path = ExcelExporter().export(
                    accs,
                    tot,
                    EXCEL_FILENAME,
                    aggregated=agg,
                    allocations=alloc,
                )
                print(f"  ✅ Сохранено: {path}")
            except Exception as e:
                print(f"  ❌ {e}")

        elif ch == "6":
            print("⏳ Загрузка…")
            try:
                accs, tot = client.get_all_accounts()
                agg = TinkoffApiClient.aggregate_positions(accs, tot)
                alloc = TinkoffApiClient.calc_type_allocation(accs, tot)
                show_portfolio(accs, tot)
                print()
                show_aggregated(agg, tot)
                print()
                show_allocation(alloc, tot)
                path = ExcelExporter().export(
                    accs,
                    tot,
                    EXCEL_FILENAME,
                    aggregated=agg,
                    allocations=alloc,
                )
                print(f"\n  ✅ Excel: {path}")
            except Exception as e:
                print(f"  ❌ {e}")

        elif ch == "7":
            token = getpass("Новый токен: ").strip()
            client = TinkoffApiClient(token)
            print("  ✅ Обновлён")

        elif ch == "0":
            print("👋 Выход")
            sys.exit(0)
        else:
            print("  ⚠️ Неверный ввод")


if __name__ == "__main__":
    run_console()
