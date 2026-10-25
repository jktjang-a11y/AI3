# -*- coding: utf-8 -*-
"""국토교통부 실거래 API 수집자료 전용 주택동향보고서 실행기."""

from pathlib import Path

from report_generator_unpacked import App


def main() -> None:
    output_root = Path(__file__).resolve().parent / "OUTPUT" / "주택동향_API기반"
    App(input_source="api", output_root=output_root).mainloop()


if __name__ == "__main__":
    main()
