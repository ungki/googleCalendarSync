/*!
 * @fileoverview 홍동달력
 * 홍동마을 실무자들이 함께 만드는 공유달력을 위한 Google Apps Script입니다. 스프레드시트의 일정 데이터를 캘린더로 동기화합니다. 시트에서 추가, 수정, 삭제를 일괄 실행할 수 있습니다.
 *
 * @version 0.9.1
 * @author ungki https://github.com/ungki/googleCalendarSync
 * @see https://developers.google.com/apps-script/reference/calendar
 *
 * @license
 * No License 2024 홍성 마을활력소, 홍동커먼즈
 */

function main() {
  const activesheetByName = 'example.env.ACTIVE_SHEET_BY_NAME';
  const sheet = SpreadsheetApp.getActive().getSheetByName(activesheetByName);
  const range = sheet.getDataRange();
  const values = range.getValues();

  syncCalendar(values, range);

  /**
   * getLastRow(): Integer
   * 시트의 열 계산
   */
  // let countRows = sheet.getLastRow();
  // Logger.log("countRows : ", countRows);

  /**
   * getLastColumn(): Integer
   * 시트의 행 계산
   */
  // let countColumns = sheet.getLastColumn();
  // Logger.log("countColumns : ", countColumns);
}

/**
 * 스프레드시트를 캘린더로 동기화
 *
 * @see https://developers.google.com/apps-script/reference/spreadsheet
 *
 * @param {object} values from active sheet data
 * @param {object} range A range consisting of all the data in the spreadsheet
 */
function syncCalendar(values, range) {
  const calendarId = 'example.env.CALENDAR_ID';
  const hongdongCalendar = CalendarApp.getCalendarById(calendarId);
  hongdongCalendar.setTimeZone('Asia/Seoul');

  for (let i = 1; i < values.length; i++) {
    const entry = values[i];
    const title = entry[1].toString().trim();
    const start = createDateTime(entry[2], entry[3]);
    const end = createDateTime(entry[2], entry[4]);
    const description = createDescription(entry[5], entry[6]);
    const location = entry[7].toString().trim();
    const uidPattern = /([a-zA-Z0-9][a-zA-Z0-9-_]*\.)*@google.com(.*)/g;

    /**
     * Class Calendar
     * Class CalendarEvent
     *
     * @see https://developers.google.com/apps-script/reference/calendar/calendar
     * @see https://developers.google.com/apps-script/reference/calendar/calendar-event
     */
    if (
      entry[0] === '추가' &&
      title !== undefined &&
      description !== undefined &&
      start instanceof Date &&
      end instanceof Date
    ) {
      try {
        const existingEvent = hongdongCalendar.getEvents(start, end, {
          search: title,
        });
        if (existingEvent.length) {
          entry[0] = '⚠️';
        } else if (existingEvent.length === 0) {
          const event = hongdongCalendar.createEvent(title, start, end, {
            description,
            location,
          });
          entry[0] = '✅';
          entry[8] = event.getId();
          entry[9] = event.getCreators();
          entry[10] = event.getLastUpdated();
        }
      } catch (e) {
        entry[0] = '⚠️';
        console.error('Publish Error: ' + e);
      }
    }

    if (
      entry[0] === '수정' &&
      entry[8] !== '' &&
      title !== undefined &&
      description !== undefined &&
      start instanceof Date &&
      end instanceof Date
    ) {
      try {
        if (uidPattern.test(entry[8])) {
          // Delete Event
          const deletedEvent = hongdongCalendar.getEventById(entry[8]);
          deletedEvent.deleteEvent();

          // Edit Event
          const editEvent = hongdongCalendar.createEvent(title, start, end, {
            description,
            location,
          });
          entry[0] = '✅';
          entry[8] = editEvent.getId();
          entry[9] = editEvent.getCreators();
          entry[10] = editEvent.getLastUpdated();
        } else {
          entry[0] = '⚠️';
        }
      } catch (e) {
        entry[0] = '⚠️';
        console.error('Edit Error: ' + e);
      }
    }

    if (
      entry[0] === '삭제' &&
      title !== undefined &&
      description !== undefined &&
      start instanceof Date &&
      end instanceof Date
    ) {
      try {
        const event = hongdongCalendar.getEventById(entry[8]);
        event.deleteEvent();
        entry[0] = '☑️';
        entry[8] = event.getId();
        entry[9] = event.getCreators();
        entry[10] = event.getLastUpdated();
      } catch (e) {
        entry[0] = '⚠️';
        console.error('Delete Error: ' + e);
      }
    }
  }

  /**
   * 시트 값 확인 Log
   */
  // Logger.log(values);

  range.setValues(values);
}

/**
 * 날짜 date와 시작, 종료 time을 Date 객체로 결합
 *
 * @param {Date} date A Date object
 * @param {Date} time A Date object
 * @return {Date} combined date and time | default 12:00
 */
function createDateTime(date, time) {
  date = new Date(date);

  if (!isNaN(new Date(time))) {
    date.setHours(time.getHours());
    date.setMinutes(time.getMinutes());
  } else {
    date.setHours(12);
    date.setMinutes(0);
  }

  return date;
}

/**
 * 링크 확인 후 내용과 결합
 *
 * @param {string} descripion
 * @param {string} link
 * @return {string} combined descripion and link or descripion only
 */
function createDescription(desc, link) {
  const linkPattern =
    /(https:\/\/www\.|http:\/\/www\.|https:\/\/|http:\/\/)?[a-zA-Z]{2,}(\.[a-zA-Z]{2,})(\.[a-zA-Z]{2,})?\/[a-zA-Z0-9]{2,}|((https:\/\/www\.|http:\/\/www\.|https:\/\/|http:\/\/)?[a-zA-Z]{2,}(\.[a-zA-Z]{2,})(\.[a-zA-Z]{2,})?)|(https:\/\/www\.|http:\/\/www\.|https:\/\/|http:\/\/)?[a-zA-Z0-9]{2,}\.[a-zA-Z0-9]{2,}\.[a-zA-Z0-9]{2,}(\.[a-zA-Z0-9]{2,})?/g;

  if (linkPattern.test(link)) {
    return (
      desc.toString().trim() + `\n\n\uD83D\uDD17 ` + link.toString().trim()
    );
  } else {
    return desc.toString().trim();
  }
}

/**
 * 스프레드시트 커스텀 메뉴
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('example.env.MENU_NAME')
    .addItem('동기화', 'main')
    .addToUi();
}
