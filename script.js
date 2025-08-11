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

document.getElementById('processBtn').addEventListener('click', () => {
    const attendanceFileInput = document.getElementById('attendanceFile');
    const zoomLogFileInput = document.getElementById('zoomLogFile');
    const dateColumnInput = document.getElementById('dateColumn');
    const resultDiv = document.getElementById('result');

    const attendanceFile = attendanceFileInput.files[0];
    const zoomLogFile = zoomLogFileInput.files[0];
    const dateColumn = dateColumnInput.value.trim();

    resultDiv.innerHTML = '';

    if (!attendanceFile || !zoomLogFile || !dateColumn) {
        resultDiv.innerHTML = `<div class="alert alert-danger">모든 필드를 채워주세요: 출석 명단 파일, Zoom 로그 파일, 그리고 날짜 열 이름.</div>`;
        return;
    }

    const zoomLogPromise = new Promise(resolve => Papa.parse(zoomLogFile, { header: true, skipEmptyLines: true, complete: result => resolve(result.data) }));

    const attendancePromise = new Promise((resolve, reject) => {
        if (attendanceFile.name.endsWith('.csv')) {
            Papa.parse(attendanceFile, {
                header: false,
                complete: result => {
                    const attendanceData = result.data.slice(2).map(row => ({
                        "성명": (row[4] || '').toString().trim(),
                        "학번": (row[5] || '').toString().trim(),
                    })).filter(row => row["성명"] && row["학번"]);
                    resolve(attendanceData);
                }
            });
        } else if (attendanceFile.name.endsWith('.xlsx')) {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const sheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[sheetName];
                    const attendanceData = XLSX.utils.sheet_to_json(worksheet).map(row => ({
                         "성명": (row["성명"] || '').toString().trim(),
                         "학번": (row["학번"] || '').toString().trim(),
                    })).filter(row => row["성명"] && row["학번"]);
                    resolve(attendanceData);
                } catch (err) {
                    reject(err);
                }
            };
            reader.onerror = (err) => reject(err);
            reader.readAsArrayBuffer(attendanceFile);
        } else {
            reject(new Error("출석 명단 파일은 .csv 또는 .xlsx 형식이어야 합니다."));
        }
    });


    Promise.all([attendancePromise, zoomLogPromise]).then(([attendanceClean, zoomLogRaw]) => {
        try {
            // Zoom 로그 정리 (기존 로직과 동일)
            const zoomLogClean = zoomLogRaw.map(row => ({
                ...row,
                "zoom_이름전체": (row['이름(원래 이름)'] || '').toString(),
                "기간(분)": parseInt(row['기간(분)'], 10) || 0
            }));
            const zoomNameList = zoomLogClean.map(row => row.zoom_이름전체);

            // 출석 명단과 Zoom 이름 매칭
            const matchedStudents = new Set();
            attendanceClean.forEach(student => {
                const studentId = student["학번"];
                const studentName = student["성명"];
                if (zoomNameList.some(zoomName => zoomName.includes(studentId) || zoomName.includes(studentName))) {
                    matchedStudents.add(`${studentId}|${studentName}`);
                }
            });

            // 매칭된 학생 기준으로 Zoom 로그에 식별자 추가
            const studentList = Array.from(matchedStudents).map(item => {
                const [id, name] = item.split('|');
                return { id, name };
            });

            zoomLogClean.forEach(log => {
                const matchedStudent = studentList.find(student => 
                    log.zoom_이름전체.includes(student.id) || log.zoom_이름전체.includes(student.name)
                );
                if (matchedStudent) {
                    log["식별자"] = `${matchedStudent.id}|${matchedStudent.name}`;
                }
            });

            // 식별자 기준으로 시간 합산 및 출석 처리
            const timeSum = zoomLogClean
                .filter(log => log["식별자"])
                .reduce((acc, log) => {
                    if (!acc[log["식별자"]]) {
                        acc[log["식별자"]] = 0;
                    }
                    acc[log["식별자"]] += log["기간(분)"];
                    return acc;
                }, {});

            const finalAttendance = attendanceClean.map(student => {
                const identifier = `${student["학번"]}|${student["성명"]}`;
                const totalMinutes = timeSum[identifier] || 0;
                const attendanceStatus = totalMinutes >= 90 ? '출석' : '';
                return {
                    "성명": student["성명"],
                    "학번": student["학번"],
                    [dateColumn]: attendanceStatus
                };
            });

            // 확인 필요 명단 생성
            const matchedLogIdentifiers = new Set(Object.keys(timeSum));
            const unmatchedLogs = zoomLogClean.filter(log => !matchedLogIdentifiers.has(log["식별자"]));
            
            const unmatchedSummary = Object.values(unmatchedLogs.reduce((acc, log) => {
                const name = log.zoom_이름전체;
                if (!acc[name]) {
                    acc[name] = {
                        'Zoom 표시 이름': name,
                        '총 체류 시간(분)': 0,
                        '비고': '이름/학번 모두 미일치'
                    };
                }
                acc[name]['총 체류 시간(분)'] += log["기간(분)"];
                return acc;
            }, {}));

            // 엑셀 파일 생성
            const wb = XLSX.utils.book_new();
            const ws1 = XLSX.utils.json_to_sheet(finalAttendance);
            const ws2 = XLSX.utils.json_to_sheet(unmatchedSummary);

            XLSX.utils.book_append_sheet(wb, ws1, "출석 체크");
            XLSX.utils.book_append_sheet(wb, ws2, "확인 필요");

            XLSX.writeFile(wb, "출석결과_느슨한매칭_자동.xlsx");

            resultDiv.innerHTML = `<div class="alert alert-success">✅ 엑셀 저장 완료: 출석결과_느슨한매칭_자동.xlsx</div>`;

        } catch (error) {
            console.error(error);
            resultDiv.innerHTML = `<div class="alert alert-danger">오류가 발생했습니다: ${error.message}</div>`;
        }
    }).catch(error => {
        console.error(error);
        resultDiv.innerHTML = `<div class="alert alert-danger">파일 처리 중 오류가 발생했습니다: ${error.message}</div>`;
    });
});
