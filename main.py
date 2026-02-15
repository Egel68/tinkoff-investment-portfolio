"""Точка входа."""

import sys


def main():
    print("=" * 50)
    print("  📈 Т-Инвестиции — Анализатор портфеля")
    print("=" * 50)

    if len(sys.argv) > 1:
        arg = sys.argv[1].lower()
        if arg in ("--gui", "-g"):
            mode = "1"
        elif arg in ("--console", "-c"):
            mode = "2"
        else:
            mode = None
    else:
        mode = None

    if mode is None:
        print("\n  1. 🖥  GUI")
        print("  2. ⌨️  Консоль")
        mode = input("\nВыбор (1/2): ").strip()

    if mode == "1":
        try:
            from gui_app import run_gui

            run_gui()
        except ImportError as e:
            print(f"\n❌ customtkinter не установлен: {e}")
            print("   Запускаю консоль…\n")
            from console_app import run_console

            run_console()
    elif mode == "2":
        from console_app import run_console

        run_console()
    else:
        print("❌ Неверный ввод")
        sys.exit(1)


if __name__ == "__main__":
    main()
