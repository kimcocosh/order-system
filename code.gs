// ============================================================
// 옥수주조 주문 관리 시스템 - Google Apps Script
// ============================================================

const SHEET_ID = '1TFW2-hbXTNR_AljzO68__eXlxiwqPdoJ6QEhjYSuzXo';
const ADMIN_PASSWORD = 'admin1234';

// ─────────────────────────────────────────────
// JSON 응답 헬퍼
// ─────────────────────────────────────────────
function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────────
// doGet
// ─────────────────────────────────────────────
function doGet(e) {
  try {
    const action = e.parameter.action;
    const ss = SpreadsheetApp.openById(SHEET_ID);

    if (action === 'clients')  return getClients(ss);
    if (action === 'orders')   return getOrders(ss);
    if (action === 'stock')    return getStock(ss);
    if (action === 'products') return getProducts(ss);
    if (action === 'prices')   return getPrices(ss);

    return jsonResponse({ error: 'Unknown action' });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// ─────────────────────────────────────────────
// doPost
// ─────────────────────────────────────────────
function doPost(e) {
  try {
    const action = e.parameter.action;
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const body = JSON.parse(e.postData.contents);

    if (action === 'submitOrder')    return submitOrder(ss, body);
    if (action === 'updateStatus')   return updateStatus(ss, body);
    if (action === 'updateStock')    return updateStock(ss, body);
    if (action === 'updateProducts') return updateProducts(ss, body);
    if (action === 'deleteOrder')    return deleteOrder(ss, body);

    return jsonResponse({ error: 'Unknown action' });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// ─────────────────────────────────────────────
// 거래처 목록
// ─────────────────────────────────────────────
function getClients(ss) {
  const sheet = ss.getSheetByName('거래처');
  if (!sheet) return jsonResponse({ clients: [] });
  const data = sheet.getDataRange().getValues();

  const clients = [];
  // 헤더가 6~7행, 데이터는 8행부터 → index 7(0-based)부터 시작
  // B열=업체명, C열=연락처, D열=주소
  for (let i = 7; i < data.length; i++) {
    const row = data[i];
    const name = String(row[1] || '').trim(); // B열
    if (!name) continue;

    const phone   = String(row[2] || '').trim(); // C열
    const address = String(row[3] || '').trim(); // D열
    const bizNum  = String(row[8] || '').trim(); // I열 사업자번호
    const key = generateKey(name);

    clients.push({ name, phone, address, bizNum, key });
  }

  return jsonResponse({ clients });
}

function generateKey(name) {
  // 단순 해시: 문자 코드 합산 + 이름 길이 조합
  let hash = 0;
  for (let i = 0; i < name.length; i++) {
    hash = (hash * 31 + name.charCodeAt(i)) & 0x7fffffff;
  }
  return hash.toString(36) + name.length.toString(36);
}

// ─────────────────────────────────────────────
// 주문 목록
// ─────────────────────────────────────────────
function getOrders(ss) {
  const sheet = ss.getSheetByName('주문접수');
  if (!sheet) return jsonResponse({ orders: [] });

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return jsonResponse({ orders: [] });

  const headers = data[0];
  const orders = data.slice(1).map((row, idx) => {
    const obj = { rowIndex: idx + 2 };
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });

  return jsonResponse({ orders });
}

// ─────────────────────────────────────────────
// 재고 현황
// ─────────────────────────────────────────────
function getStock(ss) {
  const sheet = ss.getSheetByName('재고');
  if (!sheet) return jsonResponse({ stock: [] });

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return jsonResponse({ stock: [] });

  const stock = data.slice(1).map(row => ({
    name:        String(row[0] || '').trim(),
    quantity:    Number(row[1]) || 0,
    lastUpdated: row[2] ? String(row[2]) : ''
  })).filter(s => s.name);

  return jsonResponse({ stock });
}

// ─────────────────────────────────────────────
// 상품 목록 (Y인 것만)
// ─────────────────────────────────────────────
function getProducts(ss) {
  const sheet = ss.getSheetByName('상품목록');
  if (!sheet) return jsonResponse({ products: [] });

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return jsonResponse({ products: [] });

  const products = data.slice(1)
    .filter(row => String(row[1]).trim().toUpperCase() === 'Y' && String(row[0]).trim())
    .map(row => ({ name: String(row[0]).trim() }));

  return jsonResponse({ products });
}

// ─────────────────────────────────────────────
// 주문 접수
// ─────────────────────────────────────────────
function submitOrder(ss, body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const stockSheet = ss.getSheetByName('재고');
    const orderSheet = ss.getSheetByName('주문접수');

    if (!stockSheet || !orderSheet) {
      return jsonResponse({ success: false, error: '시트를 찾을 수 없습니다.' });
    }

    // 재고 재확인
    const stockData = stockSheet.getDataRange().getValues();
    const stockMap = {};
    for (let i = 1; i < stockData.length; i++) {
      const name = String(stockData[i][0]).trim();
      if (name) stockMap[name] = { qty: Number(stockData[i][1]) || 0, row: i + 1 };
    }

    const items = body.items || {}; // { "망고": 10, "패션후르츠": 5 }
    const insufficientItems = [];
    const validItems = {};

    for (const [name, qty] of Object.entries(items)) {
      if (!qty || qty <= 0) continue;
      if (stockMap[name] === undefined) {
        validItems[name] = qty;
        continue;
      }
      if (stockMap[name].qty < qty) {
        insufficientItems.push({ name, requested: qty, available: stockMap[name].qty });
      } else {
        validItems[name] = qty;
      }
    }

    // 재고 부족이라도 경고와 함께 접수 (관리자가 판단)
    // 내품명 문자열
    const allItems = { ...validItems };
    insufficientItems.forEach(i => { allItems[i.name] = i.requested; });
    const itemStr = Object.entries(allItems)
      .filter(([, q]) => q > 0)
      .map(([n, q]) => `${n} ${q}`)
      .join(' ');

    const timestamp = new Date();
    const status = '미처리';

    // 주문 시트에 행 추가
    orderSheet.appendRow([
      timestamp,
      body.clientName  || '',
      body.phone       || '',
      body.address     || '',
      itemStr,
      body.note        || '',
      status
    ]);

    // 재고 차감 (validItems만)
    for (const [name, qty] of Object.entries(validItems)) {
      if (stockMap[name]) {
        const newQty = stockMap[name].qty - qty;
        stockSheet.getRange(stockMap[name].row, 2).setValue(newQty);
        stockSheet.getRange(stockMap[name].row, 3).setValue(new Date());
      }
    }

    // 이메일 알림
    sendOrderAlert(ss, body, allItems, insufficientItems, timestamp);

    const newRowIndex = orderSheet.getLastRow();
    return jsonResponse({ success: true, insufficientItems, rowIndex: newRowIndex });
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────
// Gmail 알림
// ─────────────────────────────────────────────
function sendOrderAlert(ss, body, allItems, insufficientItems, timestamp) {
  try {
    const notiSheet = ss.getSheetByName('알림수신자');
    if (!notiSheet) return;

    const notiData = notiSheet.getDataRange().getValues();
    const emails = notiData.slice(1)
      .map(r => String(r[0]).trim())
      .filter(e => e && e.includes('@'));

    if (!emails.length) return;

    const hasInsufficient = insufficientItems.length > 0;
    const subject = hasInsufficient
      ? `[주문알림] ${body.clientName} ⚠️ 재고 확인 필요`
      : `[주문알림] ${body.clientName} 주문 접수`;

    const acceptedLines = Object.entries(allItems)
      .filter(([, q]) => q > 0 && !insufficientItems.find(i => i.name === Object.keys(allItems)[0]))
      .map(([n, q]) => `- ${n} ${q}개`).join('\n');

    const validItemLines = Object.entries(allItems)
      .filter(([name]) => !insufficientItems.some(i => i.name === name))
      .map(([n, q]) => `- ${n} ${q}개`).join('\n') || '(없음)';

    const insufficientLines = insufficientItems
      .map(i => `- ${i.name}: 요청 ${i.requested}개 / 재고 ${i.available}개`)
      .join('\n');

    const ts = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    let body_text = `${body.clientName}님 주문이 접수되었습니다.\n\n`;
    body_text += `✅ 접수 품목\n${validItemLines}\n`;
    if (hasInsufficient) {
      body_text += `\n⚠️ 재고 부족 품목\n${insufficientLines}\n`;
    }
    body_text += `\n요청사항: ${body.note || '없음'}`;
    body_text += `\n주문시각: ${ts}`;

    emails.forEach(email => {
      GmailApp.sendEmail(email, subject, body_text);
    });
  } catch (err) {
    Logger.log('이메일 발송 오류: ' + err.message);
  }
}

// ─────────────────────────────────────────────
// 주문 삭제
// ─────────────────────────────────────────────
function deleteOrder(ss, body) {
  const sheet = ss.getSheetByName('주문접수');
  if (!sheet) return jsonResponse({ success: false, error: '시트 없음' });
  sheet.deleteRow(body.rowIndex);
  return jsonResponse({ success: true });
}

// ─────────────────────────────────────────────
// 처리상태 업데이트
// ─────────────────────────────────────────────
function updateStatus(ss, body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const sheet = ss.getSheetByName('주문접수');
    if (!sheet) return jsonResponse({ success: false, error: '시트 없음' });

    const rowIndex = body.rowIndex;
    const newStatus = body.status;

    // 처리상태 (G열 = 7번째)
    sheet.getRange(rowIndex, 7).setValue(newStatus);

    // 출고일 (H열 = 8번째): 처리완료 시 오늘 날짜(한국시간) 자동 기록
    if (newStatus === '처리완료') {
      const today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
      sheet.getRange(rowIndex, 8).setValue(today);
      Logger.log(`[출고일 자동 입력] 행: ${rowIndex}, 날짜: ${today}`);
    } else {
      sheet.getRange(rowIndex, 8).setValue('');
    }

    return jsonResponse({ success: true });
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────
// 재고 업데이트
// ─────────────────────────────────────────────
function updateStock(ss, body) {
  const sheet = ss.getSheetByName('재고');
  if (!sheet) return jsonResponse({ success: false, error: '시트 없음' });

  const items = body.items || []; // [{ name, quantity }]
  const data = sheet.getDataRange().getValues();
  const nameToRow = {};
  for (let i = 1; i < data.length; i++) {
    nameToRow[String(data[i][0]).trim()] = i + 1;
  }

  const now = new Date();
  items.forEach(item => {
    const row = nameToRow[item.name];
    if (row) {
      sheet.getRange(row, 2).setValue(item.quantity);
      sheet.getRange(row, 3).setValue(now);
    } else {
      sheet.appendRow([item.name, item.quantity, now]);
    }
  });

  return jsonResponse({ success: true });
}

// ─────────────────────────────────────────────
// 상품목록 업데이트
// ─────────────────────────────────────────────
function updateProducts(ss, body) {
  const sheet = ss.getSheetByName('상품목록');
  if (!sheet) return jsonResponse({ success: false, error: '시트 없음' });

  const products = body.products || [];

  // 상품목록 전체 재작성
  const existing = sheet.getDataRange().getValues();
  const header = existing[0];
  sheet.clearContents();
  sheet.appendRow(header);
  products.forEach(p => {
    sheet.appendRow([p.name, p.active ? 'Y' : 'N']);
  });

  // 재고 시트에 없는 품목 자동 추가
  const stockSheet = ss.getSheetByName('재고');
  if (stockSheet) {
    const stockData = stockSheet.getDataRange().getValues();
    const existingStock = new Set(stockData.slice(1).map(r => String(r[0]).trim()));
    const now = new Date();
    products.forEach(p => {
      if (p.name && !existingStock.has(p.name)) {
        stockSheet.appendRow([p.name, 0, now]);
      }
    });
  }

  return jsonResponse({ success: true });
}

// ─────────────────────────────────────────────
// 단가표
// ─────────────────────────────────────────────
function getPrices(ss) {
  const sheet = ss.getSheetByName('단가표');
  if (!sheet) return jsonResponse({ prices: [] });
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return jsonResponse({ prices: [] });
  const prices = data.slice(1).map(row => ({
    client:  String(row[0] || '').trim(),
    product: String(row[1] || '').trim(),
    price:   Number(row[2]) || 0
  })).filter(p => p.product);
  return jsonResponse({ prices });
}

// 단가 계산 헬퍼 (서버 내부 사용)
function resolvePrice(priceList, clientName, productName) {
  // 업체별 예외 단가 우선
  const specific = priceList.find(p => p.client === clientName && p.product === productName);
  if (specific) return specific.price;
  // 기본 단가 (업체명 빈칸)
  const def = priceList.find(p => p.client === '' && p.product === productName);
  if (def) return def.price;
  return 0;
}

// ─────────────────────────────────────────────
// 초기 시트 세팅 (수동 실행용)
// ─────────────────────────────────────────────
function initSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // 1. 주문접수
  let s1 = ss.getSheetByName('주문접수');
  if (!s1) {
    s1 = ss.insertSheet('주문접수');
    s1.appendRow(['타임스탬프', '업체명', '전화번호', '주소', '내품명', '요청사항', '처리상태', '출고일']);
  }

  // 2. 재고
  let s2 = ss.getSheetByName('재고');
  if (!s2) {
    s2 = ss.insertSheet('재고');
    s2.appendRow(['품목명', '현재재고', '마지막업데이트']);
    const items = ['오리지널(달콤)', '새콤', '옥수수', '고구마', '바나나', '망고',
      '패션후르츠', '블러드오렌지', '말차', '토마토', '블루베리',
      '라이트&드라이', '옥수수7도', '생쌀9도'];
    items.forEach(name => s2.appendRow([name, 0, new Date()]));
  }

  // 3. 알림수신자
  let s3 = ss.getSheetByName('알림수신자');
  if (!s3) {
    s3 = ss.insertSheet('알림수신자');
    s3.appendRow(['이메일']);
  }

  // 5. 단가표
  let s5 = ss.getSheetByName('단가표');
  if (!s5) {
    s5 = ss.insertSheet('단가표');
    s5.appendRow(['업체명(비우면 기본단가)', '품목명', '단가']);
    // 기본 단가 (업체명 없음 = 전체 기본)
    const defaults = [
      ['', '오리지널(달콤)', 4000], ['', '새콤', 4000], ['', '옥수수', 4000],
      ['', '고구마', 4000], ['', '바나나', 4000], ['', '망고', 4000],
      ['', '패션후르츠', 4000], ['', '블러드오렌지', 4000], ['', '말차', 4000],
      ['', '토마토', 4000], ['', '블루베리', 4000],
      ['', '라이트&드라이', 2800], ['', '옥수수7도', 2800], ['', '생쌀9도', 3500],
      ['', '이밤에취해', 17000], ['', '옥주15-감', 17000],
      ['', '옥주15-신', 17000], ['', '옥주15-무', 17000],
    ];
    defaults.forEach(r => s5.appendRow(r));
  }

  // 4. 상품목록
  let s4 = ss.getSheetByName('상품목록');
  if (!s4) {
    s4 = ss.insertSheet('상품목록');
    s4.appendRow(['품목명', '사용여부']);
    const items = ['오리지널(달콤)', '새콤', '옥수수', '고구마', '바나나', '망고',
      '패션후르츠', '블러드오렌지', '말차', '토마토', '블루베리',
      '라이트&드라이', '옥수수7도', '생쌀9도'];
    items.forEach(name => s4.appendRow([name, 'Y']));
  }

  Logger.log('시트 초기화 완료');
}
