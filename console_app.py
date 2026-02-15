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
    print("  2. Показать портфель")
    print("  3. Экспортировать в Excel")
    print("  4. Портфель + Excel")
    print("  5. Сменить токен")
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
            print("⏳ Экспорт…")
            try:
                accs, tot = client.get_all_accounts()
                path = ExcelExporter().export(accs, tot, EXCEL_FILENAME)
                print(f"  ✅ Сохранено: {path}")
            except Exception as e:
                print(f"  ❌ {e}")

        elif ch == "4":
            print("⏳ Загрузка…")
            try:
                accs, tot = client.get_all_accounts()
                show_portfolio(accs, tot)
                path = ExcelExporter().export(accs, tot, EXCEL_FILENAME)
                print(f"\n  ✅ Excel: {path}")
            except Exception as e:
                print(f"  ❌ {e}")

        elif ch == "5":
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
