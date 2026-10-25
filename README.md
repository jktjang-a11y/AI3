# AI3 — 부동산 AI 동향 리포터 (원본 소스)

`C:\AI 3 제출` 원본 폴더의 소스와 관련 설정 파일입니다. 한국부동산원 장규태 작성, v1.72.20.

- 메인 실행 파일: `realestate_api_report_unified_v1_72_20_rent_compare_raw_first_FIXED.py`
- 실행 배치: `주택동향_API기반_실행.bat`, `주택동향_엑셀지정기반_실행.bat`
- KOSIS·국토교통부·실거래가(RTMS)·R-ONE API로 자료를 모아 보고서를 만듭니다.

## API 키
공개 저장소라 소스 안의 키는 모두 가려 두었습니다. 실행 전에 아래 자리표시자를 본인 키로 바꾸세요.

- `__KOSIS_API_KEY_REDACTED__` (KOSIS)
- `__API_KEY_REDACTED__` (R-ONE)
- `__MOLIT_API_KEY_REDACTED__` (국토교통부)
- `__RTMS_SERVICE_KEY_REDACTED__` (실거래가)

`kosis_app_config.json`에도 같은 자리표시자가 들어 있습니다.
