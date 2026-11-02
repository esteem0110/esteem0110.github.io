// 실시간 버스 위치 받는 구글 Apps Script
// 정류장 구글시트에서 [확장 프로그램 > Apps Script]로 열어 붙여넣으세요.

const KEY = '여기에_비밀번호'; // driver.html의 KEY와 똑같이 맞추세요 (영문·숫자 추천)

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.key !== KEY) return out({ ok: false });
    const sh = sheet();
    sh.getRange(2, 1, 1, 4).setValues([[Date.now(), d.lat, d.lng, d.acc || '']]);
    return out({ ok: true });
  } catch (err) {
    return out({ ok: false, e: String(err) });
  }
}

// 홈페이지가 읽어가는 부분 (위치 정보만 내보냄)
function doGet() {
  const v = sheet().getRange(2, 1, 1, 3).getValues()[0];
  return out({ t: v[0], lat: v[1], lng: v[2] });
}

function sheet() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('실시간위치');
  if (!sh) {
    sh = ss.insertSheet('실시간위치');
    sh.getRange(1, 1, 1, 4).setValues([['시각(ms)', '위도', '경도', '정확도(m)']]);
  }
  return sh;
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
