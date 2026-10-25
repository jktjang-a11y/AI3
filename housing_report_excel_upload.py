# -*- coding: utf-8 -*-
"""사용자 지정 매매·전월세 엑셀 원장 전용 주택동향보고서 실행기."""

from pathlib import Path

from report_generator_unpacked import App


def main() -> None:
    output_root = Path(__file__).resolve().parent / "OUTPUT" / "주택동향_엑셀지정기반"
    App(input_source="excel_upload", output_root=output_root).mainloop()


if __name__ == "__main__":
    main()
