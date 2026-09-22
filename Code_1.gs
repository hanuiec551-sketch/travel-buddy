/**
 * 엄마용 하루 10분 이야기 - 서버 (Google Apps Script)
 * - 구글 시트에서 오늘의 주제/기본지침을 읽음
 * - Gemini API 호출 (API 키는 스크립트 속성에만 저장, 사이트에는 노출 안 됨)
 * - 로그 시트에 AI 요약 기록
 *
 * 스크립트 속성(프로젝트 설정 > 스크립트 속성)에 등록할 값:
 *   GEMINI_KEY : Google AI Studio에서 받은 API 키
 *   ACCESS_KEY : 엄마 링크에 붙일 접속 코드 (아무 문자열, 예: mom-2026-abc)
 */
const MODEL = 'gemini-2.5-flash';   // 모델 이름이 바뀌면 여기만 수정
const TZ = 'Asia/Seoul';
const TARGET_MIN = 10;              // 이 시간이 지나면 AI가 마무리 모드로

function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k); }
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function sheet_(name) {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!s) throw new Error('시트 탭을 찾을 수 없어요: ' + name);
  return s;
}

// ---------- 오늘의 주제 ----------
function getTopic_() {
  const rows = sheet_('주제').getDataRange().getValues().slice(1);
  const now = new Date();
  const today = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');
  const dayKo = ['일', '월', '화', '수', '목', '금', '토'][Number(Utilities.formatDate(now, TZ, 'u')) % 7];
  let byDay = null, byDate = null;
  rows.forEach(r => {
    let a = r[0];
    if (a instanceof Date) a = Utilities.formatDate(a, TZ, 'yyyy-MM-dd');
    a = String(a).trim();
    if (!a) return;
    if (a === today) byDate = r;                 // 특정 날짜 지정이 우선
    else if (a.charAt(0) === dayKo && a.length <= 3) byDay = r; // 월 / 월요일
  });
  const r = byDate || byDay;
  if (!r) return { title: '자유 이야기 시간', guide: '어머니가 편하게 하고 싶은 이야기를 자유롭게 나눠주세요.', date: today };
  return { title: String(r[1]), guide: String(r[2]), date: today };
}

// ---------- 엔드포인트 ----------
function doGet(e) {
  if ((e.parameter.k || '') !== prop_('ACCESS_KEY')) return out_({ error: 'auth' });
  const t = getTopic_();
  return out_({ title: t.title, date: t.date });   // 지침 전문은 브라우저로 보내지 않음
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.k !== prop_('ACCESS_KEY')) return out_({ error: 'auth' });
    if (body.action === 'chat') return out_(chat_(body));
    if (body.action === 'end') return out_(end_(body));
    return out_({ error: 'bad action' });
  } catch (err) {
    return out_({ error: String(err) });
  }
}

// ---------- 대화 ----------
function chat_(b) {
  const t = getTopic_();
  const base = String(sheet_('기본지침').getRange('A1').getValue());
  const min = Math.round(b.elapsedMin || 0);
  let system = base +
    '\n\n# 오늘의 주제\n제목: ' + t.title + '\n' + t.guide +
    '\n\n# 음성 대화 규칙 (매우 중요)\n' +
    '- 소리 내어 읽힐 글입니다. 한 번에 1~3문장, 짧게 말하세요.\n' +
    '- 마크다운, 목록, 이모지, 괄호 설명을 쓰지 마세요.\n' +
    '- 질문은 한 번에 하나만 하세요.\n' +
    '- 엄마 말씀이 짧거나 멈칫하면 재촉하지 말고 쉬운 단서를 하나 주세요.\n' +
    '\n# 현재 상태\n대화 경과 시간: ' + min + '분.';
  if (min >= TARGET_MIN) {
    system += ' 목표 시간이 지났습니다. 지금 하던 이야기를 자연스럽게 마무리하고, 오늘 이야기한 것 1~2개를 짧게 정리한 뒤 따뜻하게 인사하세요.';
  }
  const text = gemini_(system, b.messages, 400, 0.8);
  logProgress_(b, t);
  return { reply: text.trim() };
}

function end_(b) {
  const t = getTopic_();
  const transcript = (b.messages || []).map(m =>
    (m.role === 'user' ? '엄마' : 'AI') + ': ' + m.text).join('\n');
  let summary = '(대화가 거의 없었음)';
  if ((b.messages || []).filter(m => m.role === 'user').length > 1) {
    const sys = '당신은 경도인지장애가 있는 어머니의 일상 대화를 보호자(자녀)에게 요약해주는 도우미입니다. ' +
      '진단이나 평가는 하지 말고, 관찰한 사실만 중립적으로 한국어 3줄 이내로 쓰세요. ' +
      '형식: 1) 오늘 나눈 이야기 2) 잘 떠올리신 것 3) 머뭇거리거나 도움이 필요했던 부분(없으면 "특이사항 없음"). ' +
      '기분이나 대화 참여도가 눈에 띄면 짧게 덧붙이세요.';
    summary = gemini_(sys, [{ role: 'user', text: transcript }], 500, 0.3).trim();
  }
  logProgress_(b, t, summary, true);
  return { ok: true, summary: summary };
}

// ---------- Gemini ----------
function gemini_(system, msgs, maxTokens, temp) {
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + MODEL + ':generateContent';
  const payload = {
    systemInstruction: { parts: [{ text: system }] },
    contents: msgs.map(m => ({ role: m.role, parts: [{ text: m.text }] })),
    generationConfig: { maxOutputTokens: maxTokens, temperature: temp, thinkingConfig: { thinkingBudget: 0 } }
  };
  const opt = {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'x-goog-api-key': prop_('GEMINI_KEY') },
    payload: JSON.stringify(payload)
  };
  let res;
  for (let i = 0; i < 2; i++) {
    res = UrlFetchApp.fetch(url, opt);
    const code = res.getResponseCode();
    if (code === 200) break;
    if (code !== 429 && code < 500) break;
    Utilities.sleep(1200);
  }
  const j = JSON.parse(res.getContentText());
  if (res.getResponseCode() !== 200) throw new Error('Gemini ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 200));
  const parts = j.candidates && j.candidates[0] && j.candidates[0].content && j.candidates[0].content.parts;
  return parts ? parts.map(p => p.text || '').join('') : '';
}

// ---------- 로그 ----------
// 열: 시작시각 | 날짜 | 주제 | 대화(분) | 엄마 발화 수 | AI 요약 | 종료여부 | 세션ID
function logProgress_(b, t, summary, ended) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    let s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('로그');
    if (!s) {
      s = SpreadsheetApp.getActiveSpreadsheet().insertSheet('로그');
      s.appendRow(['시작시각', '날짜', '주제', '대화(분)', '엄마 발화 수', 'AI 요약', '종료여부', '세션ID']);
      s.setFrozenRows(1);
    }
    const turns = (b.messages || []).filter(m => m.role === 'user').length - 1; // 첫 시작 신호 제외
    const mins = Math.round((b.elapsedMin || 0) * 10) / 10;
    const found = s.getRange('H:H').createTextFinder(String(b.sid)).matchEntireCell(true).findNext();
    if (found) {
      const r = found.getRow();
      s.getRange(r, 4, 1, 2).setValues([[mins, Math.max(turns, 0)]]);
      if (ended) s.getRange(r, 6, 1, 2).setValues([[summary, '종료']]);
    } else {
      s.appendRow([Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'), t.date, t.title,
        mins, Math.max(turns, 0), summary || '', ended ? '종료' : '진행중/중단', String(b.sid)]);
    }
  } finally {
    lock.releaseLock();
  }
}
