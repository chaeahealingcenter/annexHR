/**
 * 아넥스HR교육원 — 교육 문의 폼 수신기 (Google Apps Script)
 *
 * 하는 일
 *   1) 문의 내용을 스프레드시트에 한 줄씩 쌓고
 *   2) 지정한 주소로 메일을 보냅니다.
 *
 * 설치
 *   1. 구글 드라이브에서 새 스프레드시트를 하나 만듭니다.
 *   2. 확장 프로그램 → Apps Script → 이 파일 내용을 전부 붙여넣습니다.
 *   3. 아래 NOTIFY_TO 를 받을 주소로 바꿉니다.
 *   4. 배포 → 새 배포 → 유형 "웹 앱"
 *        실행 계정: 나
 *        액세스 권한: 모든 사용자            ← 반드시 이것
 *   5. 나오는 /exec 주소를 index.html 의 form action 에 넣습니다.
 *
 * 비용 0원. 하루 메일 100통까지 (일반 Gmail 계정 기준).
 */

var NOTIFY_TO = 'annexhredu@gmail.com';   // ← 문의를 받을 주소
var SHEET_NAME = '문의';

function doPost(e) {
  try {
    var data = (e && e.parameter) ? e.parameter : {};

    // 스팸 허니팟 — 사람이 아니면 조용히 버린다
    if (data._gotcha) return _ok();

    var sheet = _sheet();
    var now = new Date();
    var row = [
      now,
      data['기관·기업명'] || '',
      data['담당자명']   || '',
      data['연락처']     || '',
      data['이메일']     || '',
      data['희망일정']   || '',
      data['진행지역']   || '',
      data['예상인원']   || '',
      data['희망시간']   || '',
      data['관심포맷']   || '',
      data['희망패키지'] || '',
      data['예상예산']   || '',
      data['기타요청']   || '',
      data['개인정보동의'] || ''
    ];
    sheet.appendRow(row);

    MailApp.sendEmail({
      to: NOTIFY_TO,
      subject: '[아넥스HR교육원] ' + (data['기관·기업명'] || '신규') + ' 교육 문의',
      replyTo: data['이메일'] || NOTIFY_TO,
      htmlBody: _mailBody(data, now)
    });

    return _ok();
  } catch (err) {
    // 실패해도 메일로 알려 문의가 유실되지 않게 한다
    try {
      MailApp.sendEmail(NOTIFY_TO, '[아넥스HR교육원] 폼 처리 오류',
        err + '\n\n' + JSON.stringify(e && e.parameter));
    } catch (_) {}
    return _ok();
  }
}

function doGet() {
  return ContentService.createTextOutput('ok');
}

function _sheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['접수일시','기관·기업명','담당자명','연락처','이메일',
                  '희망일정','진행지역','예상인원','희망시간','관심포맷',
                  '희망패키지','예상예산','기타요청','개인정보동의']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function _mailBody(d, now) {
  function row(k, v) {
    if (!v) return '';
    return '<tr><th align="left" style="padding:6px 14px 6px 0;color:#616B79;' +
           'font-weight:600;white-space:nowrap;vertical-align:top">' + k +
           '</th><td style="padding:6px 0;color:#2A303B">' +
           String(v).replace(/\n/g, '<br>') + '</td></tr>';
  }
  return '' +
    '<div style="font-family:-apple-system,BlinkMacSystemFont,\'Malgun Gothic\',sans-serif;' +
    'font-size:14px;line-height:1.7;max-width:640px">' +
      '<div style="background:#22314B;color:#fff;padding:16px 20px;border-radius:10px 10px 0 0">' +
        '<b style="font-size:16px">새 교육 문의</b><br>' +
        '<span style="color:#96A3B8;font-size:12.5px">' +
          Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd HH:mm') +
        '</span>' +
      '</div>' +
      '<div style="border:1px solid #E2E5EA;border-top:0;border-radius:0 0 10px 10px;padding:18px 20px">' +
        '<table cellpadding="0" cellspacing="0">' +
          row('기관·기업명', d['기관·기업명']) +
          row('담당자',     d['담당자명']) +
          row('연락처',     d['연락처']) +
          row('이메일',     d['이메일']) +
          row('희망 일정',  d['희망일정']) +
          row('진행 지역',  d['진행지역']) +
          row('예상 인원',  d['예상인원']) +
          row('희망 시간',  d['희망시간']) +
          row('관심 포맷',  d['관심포맷']) +
          row('희망 패키지', d['희망패키지']) +
          row('예상 예산',  d['예상예산']) +
          row('기타 요청',  d['기타요청']) +
        '</table>' +
        '<p style="margin-top:16px;padding-top:14px;border-top:1px solid #E2E5EA;' +
        'color:#616B79;font-size:12.5px">' +
          '이 메일에 그대로 답장하면 문의하신 분에게 전달됩니다. ' +
          '전체 내역은 연결된 스프레드시트에서 확인하세요.' +
        '</p>' +
      '</div>' +
    '</div>';
}

function _ok() {
  return ContentService.createTextOutput(JSON.stringify({ result: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}
