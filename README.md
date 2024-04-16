# 홍동달력

홍동마을 실무자들이 함께 만드는 공유달력을 위한 Google Apps Script입니다. 스프레드시트의 일정 데이터를 캘린더로 동기화합니다. 시트에서 추가, 수정, 삭제를 일괄 실행할 수 있습니다.

## Spreadsheet

\* required

| status \*         | title \* | date \* | start time | end time | description \* | link   | location | UID    | creators | update date |
| ----------------- | -------- | ------- | ---------- | -------- | -------------- | ------ | -------- | ------ | -------- | ----------- |
| string (dropdown) | string   | Date    | Date       | Date     | string         | string | string   | string | string   | Date        |

1. 실행할 때는 상태에서 `추가`, `수정`, `삭제` 작업을 선택할 수 있습니다. 실행 후에는 결과 아이콘으로 자동 변경됩니다.

- 발행 완료 `✅`
- 삭제 완료 `☑️`
- 오류 경고 `⚠️`

2. 상태, 제목, 날짜, 설명은 필수항목입니다.
3. UID, 편집자, 업데이트 날짜는 자동생성됩니다. 시트에서 변경할 수 없습니다.
4. 삭제는 캘린더에서만 삭제됩니다. 삭제 후에 다시 발행 상태로 변경할 때는 추가로 선택해야 합니다.

## Google Apps Script

[https://developers.google.com/apps-script/reference/calendar](https://developers.google.com/apps-script/reference/calendar)

### license

No License 2024 홍성 마을활력소, 홍동커먼즈
