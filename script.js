/* =========================================================================
 * 공통 유틸 - 느슨한 매칭용 문자열 정규화
 * ---------------------------------------------------------------------
 * 공백/괄호/하이픈/언더바 등 구분자와 발음기호(Nguyễn -> NGUYEN)를 모두 제거해
 * "2525419128홍길동", "2525419128 홍길동", "홍길동 (홍길동)" 을 같은 값으로 만든다.
 * ========================================================================= */
function norm(s) {
    return String(s == null ? '' : s)
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')   // 발음기호 제거 (한글 자모는 이 범위 밖이라 안전)
        .normalize('NFC')
        .replace(/[^0-9A-Za-z가-힣]/g, '')
        .toUpperCase();
}

// 학번은 4자 이상 영숫자일 때만 매칭 키로 쓴다.
// '-', ' ' 같은 값이 모든 이름에 걸려버리는 사고를 막기 위함.
function idKey(sid) {
    const k = norm(sid);
    return /^[0-9A-Z]{4,}$/.test(k) ? k : null;
}

// 한글 이름은 Zoom이 성을 뒤로 보내는 경우가 있다. (이정희 -> "정희 이")
function nameVariants(name) {
    const n = norm(name);
    if (n.length < 2) return [];
    const out = new Set([n]);
    if (/^[가-힣]{2,4}$/.test(n)) {
        out.add(n.slice(1) + n.charAt(0));            // 이정희 -> 정희이
        out.add(n.charAt(n.length - 1) + n.slice(0, -1)); // 정희이 -> 이정희
    }
    return Array.from(out);
}

// 영문 이름은 어순이 자주 뒤바뀐다. (명단 "Hueppe Kathrin" vs Zoom "Kathrin Hueppe-2525409001")
// 모든 토큰이 순서 상관없이 들어 있으면 매칭. 오탐 방지로 4글자 이상 토큰을 최소 1개 요구.
function latinTokens(name) {
    const t = String(name == null ? '' : name)
        .split(/[\s,()/_.·\-]+/)
        .map(norm)
        .filter(x => /^[A-Z]{2,}$/.test(x));
    if (t.length >= 2 && Math.max.apply(null, t.map(x => x.length)) >= 4) return t;
    return null;
}

// 헤더 이름을 정규화해서 느슨하게 찾는다. ('기간(분)', 'Duration (Minutes)' 모두 인식)
function findKey(keys, patterns) {
    for (let i = 0; i < patterns.length; i++) {
        const p = patterns[i];
        const hit = keys.find(k => norm(k).indexOf(p) !== -1);
        if (hit) return hit;
    }
    return null;
}

function toMinutes(row, durKey, joinKey, leaveKey) {
    if (durKey) {
        const v = parseInt(String(row[durKey]).replace(/[^0-9-]/g, ''), 10);
        if (!isNaN(v)) return v;
    }
    // 기간 열이 없는 형식이면 참가/나간 시간으로 직접 계산
    if (joinKey && leaveKey) {
        const a = Date.parse(String(row[joinKey]).replace(/-/g, '/'));
        const b = Date.parse(String(row[leaveKey]).replace(/-/g, '/'));
        if (!isNaN(a) && !isNaN(b) && b > a) return Math.round((b - a) / 60000);
    }
    return 0;
}

/* =========================================================================
 * 파일 읽기
 * ========================================================================= */

// Zoom 로그는 UTF-8이 원본이지만, 엑셀에서 한 번 열었다 저장하면 CP949(EUC-KR)가 된다.
// UTF-8로 읽어 깨짐(U+FFFD)이 보이면 EUC-KR로 다시 디코딩한다.
function decodeText(buffer) {
    let text = new TextDecoder('utf-8').decode(buffer);
    if (text.indexOf('\uFFFD') !== -1) {
        try {
            const alt = new TextDecoder('euc-kr').decode(buffer);
            if (alt.indexOf('\uFFFD') === -1) text = alt;
        } catch (e) { /* 브라우저가 euc-kr을 모르면 원본 유지 */ }
    }
    return text.charCodeAt(0) === 0xFEFF ? text.slice(1) : text;
}

function readFileAsBuffer(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = e => resolve(e.target.result);
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

// 출석 명단을 2차원 배열(AOA)로 읽는다. 원본 열을 그대로 보존하기 위해 객체가 아닌 AOA를 쓴다.
function readRosterAoa(file) {
    return readFileAsBuffer(file).then(buffer => {
        if (/\.csv$/i.test(file.name)) {
            const parsed = Papa.parse(decodeText(buffer), { header: false, skipEmptyLines: true });
            return parsed.data;
        }
        const wb = XLSX.read(new Uint8Array(buffer), { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        return XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: false, defval: '' });
    });
}

const NAME_HEADERS = ['성명', '이름', '학생명', 'NAME'];
const ID_HEADERS = ['학번', '학생번호', '수험번호', 'STUDENTID', 'ID'];

// 헤더 행 위치와 성명/학번 열 위치를 찾는다. 못 찾으면 null.
function locateColumns(aoa) {
    for (let i = 0; i < Math.min(aoa.length, 20); i++) {
        const row = (aoa[i] || []).map(norm);
        const nameCol = row.findIndex(c => NAME_HEADERS.indexOf(c) !== -1);
        const idCol = row.findIndex(c => ID_HEADERS.indexOf(c) !== -1);
        if (nameCol !== -1 || idCol !== -1) return { headerRow: i, nameCol, idCol };
    }
    return null;
}

// 헤더가 중복되거나 비어 있어도 열 순서를 유지할 수 있도록 고유한 이름을 만든다.
function uniqueHeaders(row, width) {
    const used = Object.create(null);
    const out = [];
    for (let i = 0; i < width; i++) {
        let h = String((row && row[i]) != null ? row[i] : '').trim() || ('열' + (i + 1));
        if (used[h]) { used[h]++; h = h + '_' + used[h]; } else { used[h] = 1; }
        out.push(h);
    }
    return out;
}

/* =========================================================================
 * 샘플 명단
 * ========================================================================= */
document.getElementById('downloadSampleBtn').addEventListener('click', () => {
    const sampleData = [
        { "성명": "홍길동", "학번": "20240001" },
        { "성명": "김철수", "학번": "20240002" },
    ];
    const ws = XLSX.utils.json_to_sheet(sampleData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "출석명단");
    XLSX.writeFile(wb, "출석명단_샘플.xlsx");
});

/* =========================================================================
 * 출석 처리
 * ========================================================================= */
document.getElementById('processBtn').addEventListener('click', () => {
    const attendanceFile = document.getElementById('attendanceFile').files[0];
    const zoomLogFile = document.getElementById('zoomLogFile').files[0];
    const dateColumn = document.getElementById('dateColumn').value.trim();
    const minMinutesInput = document.getElementById('minMinutes');
    const minMinutes = Math.max(0, parseInt((minMinutesInput && minMinutesInput.value) || '90', 10) || 0);
    const resultDiv = document.getElementById('result');

    resultDiv.innerHTML = '';

    if (!attendanceFile || !zoomLogFile || !dateColumn) {
        resultDiv.innerHTML = `<div class="alert alert-danger">모든 필드를 채워주세요: 출석 명단 파일, Zoom 로그 파일, 그리고 날짜 열 이름.</div>`;
        return;
    }
    if (!/\.(csv|xlsx)$/i.test(attendanceFile.name)) {
        resultDiv.innerHTML = `<div class="alert alert-danger">출석 명단 파일은 .csv 또는 .xlsx 형식이어야 합니다.</div>`;
        return;
    }

    const zoomLogPromise = readFileAsBuffer(zoomLogFile)
        .then(buffer => Papa.parse(decodeText(buffer), { header: true, skipEmptyLines: true }).data);

    Promise.all([readRosterAoa(attendanceFile), zoomLogPromise]).then(([aoa, zoomLogRaw]) => {
        // ---------- 1. 출석 명단 파싱 (원본 열 전체 보존) ----------
        const located = locateColumns(aoa);
        let headerRow, nameCol, idCol, dataRows;

        if (located) {
            headerRow = located.headerRow;
            nameCol = located.nameCol;
            idCol = located.idCol;
            dataRows = aoa.slice(headerRow + 1);
        } else if (/\.csv$/i.test(attendanceFile.name)) {
            // 헤더를 못 찾은 CSV는 기존 동작(3행부터, E열=성명 / F열=학번)으로 되돌린다.
            headerRow = -1; nameCol = 4; idCol = 5;
            dataRows = aoa.slice(2);
        } else {
            headerRow = 0; nameCol = 0; idCol = 1;
            dataRows = aoa.slice(1);
        }
        if (nameCol === -1 && idCol === -1) throw new Error("명단에서 '성명' 또는 '학번' 열을 찾지 못했습니다.");

        const width = Math.max(
            headerRow >= 0 ? (aoa[headerRow] || []).length : 0,
            nameCol + 1, idCol + 1,
            ...dataRows.map(r => (r || []).length)
        );
        const headers = uniqueHeaders(headerRow >= 0 ? aoa[headerRow] : null, width);

        // 성명과 학번 중 하나만 있어도 명단에 남긴다. (학번 공란이라고 학생을 버리지 않는다)
        const students = [];
        dataRows.forEach(row => {
            const name = String((nameCol >= 0 && row[nameCol] != null) ? row[nameCol] : '').trim();
            const sid = String((idCol >= 0 && row[idCol] != null) ? row[idCol] : '').trim();
            if (!name && !sid) return;
            students.push({
                row: row,
                name: name,
                sid: sid,
                idk: idKey(sid),
                variants: nameVariants(name),
                tokens: latinTokens(name),
                minutes: 0,
                hits: []
            });
        });
        if (!students.length) throw new Error('출석 명단에서 학생을 한 명도 읽지 못했습니다.');

        // ---------- 2. Zoom 로그 파싱 ----------
        const zoomKeys = zoomLogRaw.length ? Object.keys(zoomLogRaw[0]) : [];
        const nameKey = findKey(zoomKeys, ['이름', 'NAME']);
        const durKey = findKey(zoomKeys, ['기간', 'DURATION']);
        const joinKey = findKey(zoomKeys, ['참가시간', 'JOINTIME']);
        const leaveKey = findKey(zoomKeys, ['나간시간', 'LEAVETIME']);
        if (!nameKey) throw new Error("Zoom 로그에서 이름 열을 찾지 못했습니다. (파일 인코딩이 깨졌을 수 있습니다)");

        const logs = zoomLogRaw.map(row => {
            const display = String(row[nameKey] == null ? '' : row[nameKey]).trim();
            return { display: display, key: norm(display), minutes: toMinutes(row, durKey, joinKey, leaveKey) };
        }).filter(log => log.display);

        // ---------- 3. 매칭 ----------
        // 학생마다 독립적으로 자기 로그를 전부 합산한다.
        // (로그 1건을 학생 1명에게만 배정하면, 먼저 걸린 학생이 남의 출석을 가로챈다)
        const claimed = new Array(logs.length).fill(false);

        students.forEach(student => {
            logs.forEach((log, i) => {
                let how = null;
                if (student.idk && log.key.indexOf(student.idk) !== -1) {
                    how = '학번';
                } else if (student.variants.some(v => log.key.indexOf(v) !== -1)) {
                    how = '이름';
                } else if (student.tokens && student.tokens.every(t => log.key.indexOf(t) !== -1)) {
                    how = '영문이름';
                }
                if (how) {
                    student.minutes += log.minutes;
                    student.hits.push(log.display);
                    claimed[i] = true;
                }
            });
        });

        // ---------- 4. 출석 시트 (원본 열 + 날짜 열 + 근거) ----------
        const MIN_COL = '총 체류(분)';
        const SRC_COL = '매칭된 Zoom 이름';
        // 근거 열(총 체류/매칭 이름)은 매 실행마다 갱신되므로 항상 맨 뒤로 보낸다.
        // 지난 회차의 날짜 열은 그대로 보존되고, 새 날짜 열만 그 앞에 추가된다.
        const outHeaders = headers.filter(h => h !== MIN_COL && h !== SRC_COL);
        if (outHeaders.indexOf(dateColumn) === -1) outHeaders.push(dateColumn);
        outHeaders.push(MIN_COL, SRC_COL);

        const finalAttendance = students.map(student => {
            const obj = {};
            headers.forEach((h, i) => {
                obj[h] = student.row[i] != null ? student.row[i] : '';
            });
            obj[dateColumn] = student.minutes >= minMinutes ? '출석' : '';
            obj[MIN_COL] = student.minutes;
            obj[SRC_COL] = Array.from(new Set(student.hits)).join(' / ');
            return obj;
        });

        // ---------- 5. 확인 필요 시트 (아무 학생과도 매칭되지 않은 로그) ----------
        const unmatched = {};
        logs.forEach((log, i) => {
            if (claimed[i]) return;
            if (!unmatched[log.display]) {
                unmatched[log.display] = {
                    'Zoom 표시 이름': log.display,
                    '총 체류 시간(분)': 0,
                    '접속 횟수': 0,
                    '비고': '이름/학번 모두 미일치'
                };
            }
            unmatched[log.display]['총 체류 시간(분)'] += log.minutes;
            unmatched[log.display]['접속 횟수'] += 1;
        });
        const unmatchedSummary = Object.values(unmatched)
            .sort((a, b) => b['총 체류 시간(분)'] - a['총 체류 시간(분)']);

        // ---------- 6. 저장 ----------
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(finalAttendance, { header: outHeaders }), "출석 체크");
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(unmatchedSummary), "확인 필요");
        XLSX.writeFile(wb, "출석결과_느슨한매칭_자동.xlsx");

        const presentCount = finalAttendance.filter(r => r[dateColumn] === '출석').length;
        resultDiv.innerHTML =
            `<div class="alert alert-success mb-2">✅ 엑셀 저장 완료: 출석결과_느슨한매칭_자동.xlsx</div>` +
            `<ul class="small text-muted mb-0">` +
            `<li>명단 <strong>${students.length}명</strong> 중 <strong>${presentCount}명</strong> 출석 (${minMinutes}분 이상)</li>` +
            `<li>매칭되지 않은 Zoom 참가자 <strong>${unmatchedSummary.length}명</strong> → '확인 필요' 시트</li>` +
            `<li>'매칭된 Zoom 이름' 열에서 잘못 붙은 사람이 없는지 확인해 주세요.</li>` +
            `</ul>`;

    }).catch(error => {
        console.error(error);
        resultDiv.innerHTML = `<div class="alert alert-danger">파일 처리 중 오류가 발생했습니다: ${error.message}</div>`;
    });
});
