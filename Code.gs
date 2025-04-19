// =============================================================================
// 1. 상수 및 설정
// =============================================================================
const CONSTANTS = {
  // 시트 관련
  SHEET_NAME_FORMAT: '%d월_출퇴근기록',
  EMPLOYEE_LIST_SHEET: '직원목록',
  
  // 헤더 및 행 레이블
  SUMMARY_HEADER: '근무 요약 정보',
  TOTAL_HOURS_ROW: '총시간',
  
  // 날짜 및 시간 포맷
  DATE_FORMAT: 'yyyy-MM-dd',
  TIME_FORMAT: 'HH:mm',
  
  // 근무 관련
  DEFAULT_HOURLY_RATE: 9860, // 최저시급 기본값
  MIN_HOURS_FOR_BONUS: 15    // 주휴수당 기준 최소 시간
};

// =============================================================================
// 2. 유틸리티 함수 모듈
// =============================================================================
const Utils = {
  // 열 번호를 Excel 열 문자로 변환 (1 → A, 2 → B, 27 → AA, ...)
  getColumnLetter: function(columnNumber) {
    try {
      if (!columnNumber || columnNumber <= 0) return 'A';
      
      let columnLetter = '';
      let tempColumnNumber = columnNumber;
      
      while (tempColumnNumber > 0) {
        let modulo = (tempColumnNumber - 1) % 26;
        columnLetter = String.fromCharCode(65 + modulo) + columnLetter;
        tempColumnNumber = Math.floor((tempColumnNumber - modulo - 1) / 26);
      }
      
      return columnLetter;
    } catch (e) {
      Logger.log(`getColumnLetter 오류: ${e}`);
      return 'A'; // 오류 발생시 기본값 반환
    }
  },
  
  // 시간을 10분 단위로 반올림/내림
  roundTimeToTenMinutes: function(date) {
    const newDate = new Date(date);
    const minutes = newDate.getMinutes();
    const remainder = minutes % 10;
    
    if (remainder < 5) {
      // 1~4분은 내림 (예: 9:11 → 9:10)
      newDate.setMinutes(minutes - remainder);
    } else {
      // 5~9분은 올림 (예: 9:15 → 9:20)
      newDate.setMinutes(minutes + (10 - remainder));
    }
    
    newDate.setSeconds(0);
    newDate.setMilliseconds(0);
    
    return newDate;
  },
  
  // 근무시간 계산 수식 생성
  createWorkingHoursFormula: function(checkInCell, checkOutCell) {
    return `=IF(AND(NOT(ISBLANK(${checkInCell})), NOT(ISBLANK(${checkOutCell})), ${checkInCell}<>"미출근", ${checkOutCell}<>"미퇴근"), 
      IF((HOUR(${checkOutCell}) + MINUTE(${checkOutCell})/60) < (HOUR(${checkInCell}) + MINUTE(${checkInCell})/60),
        (HOUR(${checkOutCell}) + MINUTE(${checkOutCell})/60) + 24 - (HOUR(${checkInCell}) + MINUTE(${checkInCell})/60),
        (HOUR(${checkOutCell}) + MINUTE(${checkOutCell})/60) - (HOUR(${checkInCell}) + MINUTE(${checkInCell})/60)
      ),
      0)`;
  },
  
  // 근무 시간 계산 (문자열 변환)
  calculateHours: function(checkInTime, checkOutTime) {
    try {
      if (!checkInTime || !checkOutTime) return null;
      if (checkInTime === "미출근" || checkOutTime === "미퇴근") return null;
      
      // 시간 문자열을 분으로 변환 (HH:MM 형식)
      function timeToMinutes(timeStr) {
        const parts = String(timeStr).trim().split(':');
        return parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10);
      }
      
      // 출퇴근 시간을 분으로 변환
      const checkInMinutes = timeToMinutes(checkInTime);
      const checkOutMinutes = timeToMinutes(checkOutTime);
      
      // 근무 시간 계산 (분)
      let workMinutes = checkOutMinutes - checkInMinutes;
      
      // 퇴근 시간이 출근 시간보다 이른 경우 (다음날 퇴근 가정)
      if (workMinutes < 0) {
        workMinutes += 24 * 60;
      }
      
      // 시간과 분으로 변환
      const hours = Math.floor(workMinutes / 60);
      const minutes = workMinutes % 60;
      
      return `${hours}시간 ${minutes}분`;
    } catch (e) {
      Logger.log(`근무 시간 계산 오류: ${e}`);
      return null;
    }
  },
  
  // 오류 로깅
  logError: function(functionName, error) {
    Logger.log(`${functionName} 오류: ${error.stack || error.toString()}`);
  }
};

// =============================================================================
// 3. 시트 관리 모듈
// =============================================================================
const SheetManager = {
  // 스프레드시트 인스턴스 가져오기
  getSpreadsheet: function() {
    return SpreadsheetApp.getActiveSpreadsheet();
  },
  
  // 직원목록 시트 가져오기
  getEmployeeSheet: function() {
    return this.getSpreadsheet().getSheetByName(CONSTANTS.EMPLOYEE_LIST_SHEET);
  },
  
  // 월별 출퇴근 시트 가져오기
  getAttendanceSheet: function(targetMonth) {
    try {
      if (!targetMonth) {
        Logger.log("targetMonth가 null 또는 undefined입니다");
        return null;
      }
      
      // 문자열 변환 및 숫자 확인
      const monthNumber = Number(targetMonth);
      if (isNaN(monthNumber) || monthNumber < 1 || monthNumber > 12) {
        Logger.log(`유효하지 않은 월: ${targetMonth}`);
        return null;
      }
      
      const monthStr = String(monthNumber).trim();
      const sheetName = CONSTANTS.SHEET_NAME_FORMAT.replace('%d', monthStr);
      Logger.log(`검색 중인 시트: ${sheetName}`);
      
      const ss = this.getSpreadsheet();
      const sheet = ss.getSheetByName(sheetName);
      
      if (sheet) {
        Logger.log(`시트 찾음: ${sheetName}`);
        return sheet;
      }
      
      Logger.log(`시트를 찾을 수 없음: ${sheetName}`);
      return null;
    } catch (e) {
      Utils.logError('getAttendanceSheet', e);
      return null;
    }
  },
  
  // 시트 구조 분석
  analyzeSheetStructure: function(sheet) {
    // 헤더 행 가져오기
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colHeaderRow = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    const existingEmployees = [];
    const employeeColumns = {}; // 직원 이름 -> 시작 열 번호 매핑
    
    // 직원 열 찾기
    for (let col = 1; col < headerRow.length; col++) {
      const empName = headerRow[col];
      if (empName && !['날짜', '요일', '주차'].includes(empName)) {  // '주차' 추가
        const colHeader = colHeaderRow[col];
        if (colHeader === '출근시간') {
          existingEmployees.push(empName);
          employeeColumns[empName] = col + 1; // 1-based
        }
      }
    }
    
    // 총시간 행과 요약 행 찾기 로직
    let totalRow = -1;
    let summaryRow = -1;
    const allData = sheet.getDataRange().getValues();
    
    // 총시간 행 찾기
    for (let row = 0; row < allData.length; row++) {
      if (allData[row][0] === CONSTANTS.TOTAL_HOURS_ROW) {
        totalRow = row + 1; // 1-based
        break;
      }
    }
    
    // 총시간 행을 찾지 못했을 때 대체 로직
    if (totalRow === -1) {
      for (let row = allData.length - 1; row >= 0; row--) {
        if (allData[row][0] instanceof Date) {
          totalRow = row + 2; // 날짜 다음 행
          break;
        }
      }
      
      if (totalRow === -1) {
        totalRow = allData.length > 3 ? allData.length : 34; // 한 달 기준
      }
    }
    
    // 요약 정보 시작 행 찾기
    for (let row = 0; row < allData.length; row++) {
      if (allData[row][0] === CONSTANTS.SUMMARY_HEADER) {
        summaryRow = row + 1;
        break;
      }
    }
    
    if (summaryRow === -1) {
      summaryRow = totalRow + 2;
    }
    
    return {
      existingEmployees,
      employeeColumns, 
      totalRow,
      summaryRow
    };
  },
  
  // 기본 시트 구조 설정
  setupBasicSheetStructure: function(sheet, date) {

  // 날짜/요일/주차 헤더 설정
  sheet.getRange(1, 1).setValue('날짜');
  sheet.getRange(1, 2).setValue('요일');
  sheet.getRange(1, 3).setValue('주차'); // 새로운 주차 열 추가
  sheet.getRange(1, 1, 2, 1).merge();
  sheet.getRange(1, 2, 2, 1).merge();
  sheet.getRange(1, 3, 2, 1).merge();
  sheet.getRange(1, 1, 2, 3).setBackground('#f3f3f3'); // 헤더 배경색

  // 열 너비 설정
  sheet.setColumnWidth(1, 100); // 날짜 열
  sheet.setColumnWidth(2, 60);  // 요일 열
  sheet.setColumnWidth(3, 60);  // 주차 열

  // 날짜 채우기
  const lastDay = new Date(date.getFullYear(), date.getMonth() + 1, 0).getDate();
  const dates = [];
  for (let i = 1; i <= lastDay; i++) {
    const currentDate = new Date(date.getFullYear(), date.getMonth(), i);
    const dayNames = ['일', '월', '화', '수', '목', '금', '토'];
    const dayOfWeek = dayNames[currentDate.getDay()];
    
    // 날짜, 요일만 데이터에 추가 (주차는 나중에 수식으로)
    dates.push([
      currentDate,
      dayOfWeek
    ]);
  }


    // 날짜 데이터 입력 및 서식 설정
    const dateRange = sheet.getRange(3, 1, dates.length, 2);
    dateRange.setValues(dates);
    sheet.getRange(3, 1, dates.length, 1).setNumberFormat(CONSTANTS.DATE_FORMAT);
    sheet.getRange(3, 2, dates.length, 1).setHorizontalAlignment('center');

    // 주차 열에 WEEKNUM 수식 추가
    for (let i = 0; i < dates.length; i++) {
      const rowNum = i + 3;  // 데이터 시작 행 (3부터 시작)
      
      // 주차 계산 수식: 해당 날짜의 주차 
      const formula = `=WEEKNUM(A${rowNum},2)-WEEKNUM(DATE(YEAR(A${rowNum}),MONTH(A${rowNum}),1),2)+1 & "주"`;
      
      // 주차 열에 수식 설정
      sheet.getRange(rowNum, 3).setFormula(formula);
      sheet.getRange(rowNum, 3).setHorizontalAlignment('center');
    }

    // 총 근무시간 행 추가
    const totalRow = dates.length + 3;
    sheet.getRange(totalRow, 1).setValue(CONSTANTS.TOTAL_HOURS_ROW);
    sheet.getRange(totalRow, 1, 1, 3).merge(); // 3열까지 병합 (날짜, 요일, 주차)
    sheet.getRange(totalRow, 1).setFontWeight('bold');

    
    return {
      lastDay: lastDay,
      totalRow: totalRow,
      summaryRow: totalRow + 2
    };
  },
  
  // 데이터 유효성 검사 규칙 설정
  setupValidationRules: function(sheet, daysInMonth) {
    try {
      const employees = EmployeeManager.getEmployees();
      const dataStartRow = 3; // 실제 데이터 시작 행
      
      // daysInMonth가 유효한 값인지 확인
      if (!daysInMonth || daysInMonth <= 0) {
        Logger.log('유효하지 않은 daysInMonth 값: ' + daysInMonth);
        daysInMonth = 31; // 기본값 설정
      }
      
      employees.forEach((emp, index) => {
        const startCol = 4 + (index * 3); // 각 직원의 출근시간 열
        
        // 현재 열의 A1 표기법 얻기
        const checkInCol = Utils.getColumnLetter(startCol);
        const checkOutCol = Utils.getColumnLetter(startCol + 1);
        
        try {
          // 출근 시간 확인 규칙
          const checkInRule = SpreadsheetApp.newDataValidation()
            .requireFormulaSatisfied(`=OR(ISBLANK(${checkInCol}${dataStartRow}), ${checkInCol}${dataStartRow}="미출근", REGEXMATCH(TEXT(${checkInCol}${dataStartRow},"HH:MM"), "^[0-2][0-9]:[0-5][0-9]$"))`)
            .setAllowInvalid(false)
            .setHelpText('올바른 시간 형식을 입력하세요 (HH:MM) 또는 "미출근"')
            .build();
          
          // 퇴근 시간 확인 규칙
          const checkOutRule = SpreadsheetApp.newDataValidation()
            .requireFormulaSatisfied(`=OR(ISBLANK(${checkOutCol}${dataStartRow}), ${checkOutCol}${dataStartRow}="미퇴근", REGEXMATCH(TEXT(${checkOutCol}${dataStartRow},"HH:MM"), "^[0-2][0-9]:[0-5][0-9]$"))`)
            .setAllowInvalid(false)
            .setHelpText('올바른 시간 형식을 입력하세요 (HH:MM) 또는 "미퇴근"')
            .build();
          
          // 규칙 적용 - 출퇴근 기록 부분에만 적용 (요약 정보 섹션 제외)
          sheet.getRange(dataStartRow, startCol, daysInMonth, 1).setDataValidation(checkInRule);
          sheet.getRange(dataStartRow, startCol + 1, daysInMonth, 1).setDataValidation(checkOutRule);
        } catch (e) {
          Logger.log(`유효성 검사 규칙 설정 오류 (직원 ${emp.name}): ${e}`);
        }
      });
      
      Logger.log('유효성 검사 규칙이 성공적으로 설정되었습니다.');
      return true;
    } catch (e) {
      Utils.logError('setupValidationRules', e);
      return false;
    }
  },
  
  // 두 달 전 시트 삭제
  deleteOldSheets: function() {
    const ss = this.getSpreadsheet();
    const sheets = ss.getSheets();
    
    const now = new Date();
    const currentMonth = now.getMonth() + 1; // 1-based (1~12)
    
    let deletedCount = 0;
    
    sheets.forEach(function(sheet) {
      const sheetName = sheet.getName();
      // "N월_출퇴근기록" 형식의 시트 찾기
      const match = sheetName.match(/^(\d+)월_출퇴근기록$/);
      
      if (match) {
        const sheetMonth = parseInt(match[1]);
        
        // 두 달 전 시트 계산 (현재가 4월이면 2월 이전 시트)
        let twoMonthsAgo = currentMonth - 2;
        
        // 1월, 2월인 경우 조정
        if (twoMonthsAgo <= 0) {
          twoMonthsAgo += 12;
        }
        
        // 두 달 전 이전 시트만 삭제 (현재 4월이면 1, 2월 시트 삭제)
        if (sheetMonth <= twoMonthsAgo) {
          // 바로 지우지 말고 확인 로그 먼저 출력
          Logger.log(`삭제 예정 시트: ${sheetName} (현재: ${currentMonth}월, 기준: ${twoMonthsAgo}월 이전)`);
          
          try {
            ss.deleteSheet(sheet);
            deletedCount++;
            Logger.log(`시트 삭제됨: ${sheetName}`);
          } catch (e) {
            Logger.log(`시트 삭제 오류: ${sheetName}, 오류: ${e.toString()}`);
          }
        }
      }
    });
    
    return `${deletedCount}개의 이전 시트가 삭제되었습니다.`;
  }
};

// =============================================================================
// 4. 직원 관리 모듈
// =============================================================================
const EmployeeManager = {
  // 직원 목록 가져오기
  getEmployees: function() {
    try {
      const sheet = SheetManager.getEmployeeSheet();
      if (!sheet) return [];
      
      const data = sheet.getDataRange().getValues();
      const employees = [];
      
      if (data.length <= 1) return employees;
      
      // 헤더 확인
      const headers = data[0];
      const hourlyRateColumn = headers.indexOf('시급');
      const passwordColumn = headers.indexOf('비밀번호');
      
      // 헤더를 제외한 행 처리
      for (let i = 1; i < data.length; i++) {
        if (data[i][0]) { // 이름이 있는 경우만 추가
          const employee = {
            name: data[i][0],
            id: i
          };
          
          // 시급 정보가 있으면 추가
          if (hourlyRateColumn !== -1 && data[i][hourlyRateColumn]) {
            employee.hourlyRate = parseInt(data[i][hourlyRateColumn]);
          } else {
            employee.hourlyRate = CONSTANTS.DEFAULT_HOURLY_RATE;
          }
          
          // 보안을 위해 비밀번호는 클라이언트에게 전송하지 않음
          employees.push(employee);
        }
      }
      
      return employees;
    } catch (e) {
      Utils.logError('getEmployees', e);
      return [];
    }
  },
  
  // 비밀번호 검증 함수
  verifyEmployeePassword: function(employeeName, password) {
    try {
      const sheet = SheetManager.getEmployeeSheet();
      if (!sheet) {
        return {
          success: false,
          message: '직원목록 시트를 찾을 수 없습니다.'
        };
      }
      
      const data = sheet.getDataRange().getValues();
      
      if (data.length <= 1) {
        return {
          success: false,
          message: '직원 정보가 없습니다.'
        };
      }
      
      // 헤더 확인
      const headers = data[0];
      const nameColumn = headers.indexOf('이름') !== -1 ? headers.indexOf('이름') : 0;
      const passwordColumn = headers.indexOf('비밀번호');
      
      if (passwordColumn === -1) {
        return {
          success: false,
          message: '비밀번호 정보가 시트에 없습니다.'
        };
      }
      
      // 직원 찾기
      let found = false;
      let passwordCorrect = false;
      
      for (let i = 1; i < data.length; i++) {
        if (data[i][nameColumn] === employeeName) {
          found = true;
          // 비밀번호 검증 (문자열로 변환하여 비교)
          const storedPassword = String(data[i][passwordColumn]);
          if (storedPassword === String(password)) {
            passwordCorrect = true;
          }
          break;
        }
      }
      
      if (!found) {
        return {
          success: false,
          message: '해당 이름의 직원을 찾을 수 없습니다.'
        };
      }
      
      if (!passwordCorrect) {
        return {
          success: false,
          message: '비밀번호가 일치하지 않습니다.'
        };
      }
      
      return {
        success: true,
        message: '비밀번호 확인 완료'
      };
      
    } catch (e) {
      Utils.logError('verifyEmployeePassword', e);
      return {
        success: false,
        message: `비밀번호 검증 중 오류가 발생했습니다: ${e.message}`
      };
    }
  },
  
  // 직원 열 설정 모듈
  setupEmployeeColumns: function(sheet, employee, startCol, dataRows, totalRow) {
    // 1. 직원 이름 헤더 설정 - 병합 적용
    sheet.getRange(1, startCol, 1, 3).merge();
    sheet.getRange(1, startCol).setValue(employee.name);
    sheet.getRange(1, startCol, 1, 3).setHorizontalAlignment('center');
    sheet.getRange(1, startCol, 1, 3).setFontWeight('bold');
    sheet.getRange(1, startCol, 1, 3).setBorder(true, true, true, true, true, true);
    sheet.getRange(1, startCol, 1, 3).setBackground('#f3f3f3'); // 헤더 회색 배경 추가
    
    // 2. 열 헤더 명확하게 설정
    sheet.getRange(2, startCol).setValue('출근시간').setFontWeight('bold').setBackground('#f3f3f3').setBorder(true, true, true, true, true, true);
    sheet.getRange(2, startCol + 1).setValue('퇴근시간').setFontWeight('bold').setBackground('#f3f3f3').setBorder(true, true, true, true, true, true);
    sheet.getRange(2, startCol + 2).setValue('근무시간').setFontWeight('bold').setBackground('#f3f3f3').setBorder(true, true, true, true, true, true);
    
    // 3. 출퇴근 시간 열 서식 설정
    sheet.getRange(3, startCol, dataRows, 2).setNumberFormat('HH:mm');
    
    // 4. 근무시간 열에 수식 추가
    for (let row = 3; row < totalRow; row++) {
      const checkInCell = sheet.getRange(row, startCol).getA1Notation();
      const checkOutCell = sheet.getRange(row, startCol + 1).getA1Notation();
      
      const formula = Utils.createWorkingHoursFormula(checkInCell, checkOutCell);
      sheet.getRange(row, startCol + 2).setFormula(formula);
      sheet.getRange(row, startCol + 2).setNumberFormat('0.00');
    }
    
    // 5. 총 근무시간 행 수식 추가
    const totalHoursFormula = `=SUM(${Utils.getColumnLetter(startCol + 2)}3:${Utils.getColumnLetter(startCol + 2)}${totalRow - 1})`;
    sheet.getRange(totalRow, startCol + 2).setFormula(totalHoursFormula);
    sheet.getRange(totalRow, startCol + 2).setFontWeight('bold');
    sheet.getRange(totalRow, startCol + 2).setNumberFormat('0.00');
    
    return startCol + 3; // 다음 직원의 시작 열 반환
  }
};
// =============================================================================
// 신입 등록
// =============================================================================

  // 직원 동기화 함수
function syncEmployeesForMonth(month) {
    try {
      // 월 파라미터가 없으면 현재 월 사용
      if (!month) {
        const now = new Date();
        month = now.getMonth() + 1;
      }
      
      Logger.log(`${month}월 시트 직원 동기화 시작...`);
      
      // 직원 목록 가져오기
      const employees = EmployeeManager.getEmployees();
      
      // 해당 월 시트 찾기
      const sheetName = CONSTANTS.SHEET_NAME_FORMAT.replace('%d', month);
      const sheet = SheetManager.getAttendanceSheet(month);
      
      if (!sheet) {
        Logger.log(`${sheetName} 시트를 찾을 수 없습니다.`);
        return `${sheetName} 시트를 찾을 수 없습니다.`;
      }
      
      // 현재 시트에 있는 직원 및 구조 분석
      const { existingEmployees, employeeColumns, totalRow, summaryRow } = SheetManager.analyzeSheetStructure(sheet);
      
      // 추가할 직원 목록 생성 (기존에 없는 직원만)
      const employeesToAdd = employees.filter(emp => !existingEmployees.includes(emp.name));
      
      // 변경 사항이 없으면 종료
      if (employeesToAdd.length === 0) {
        return `${month}월 시트에 변경 사항이 없습니다.`;
      }
      
      // 마지막 직원 열 위치 찾기
      let lastColumn = 3; // 날짜, 요일 열
      for (const empName in employeeColumns) {
        const startCol = employeeColumns[empName];
        lastColumn = Math.max(lastColumn, startCol + 2); // 각 직원은 3열 사용
      }
      
      // 새 직원 추가
      for (const employee of employeesToAdd) {
        // 다음 직원 위치는 마지막 열 다음
        const newStartCol = lastColumn + 1;
        
        // 직원 열 설정
        EmployeeManager.setupEmployeeColumns(sheet, employee, newStartCol, totalRow - 3, totalRow);
        
        // 요약 섹션 추가
        SummaryManager.createEmployeeSummarySection(sheet, employee, summaryRow, newStartCol);
        
        // 다음 직원을 위해 마지막 열 위치 업데이트
        lastColumn = newStartCol + 2;
      }
      
      return `${month}월 시트 직원 동기화 완료! ${employeesToAdd.length}명의 직원이 추가되었습니다.`;
      
    } catch (e) {
      const errorMsg = `${month}월 시트 동기화 오류: ${e.toString()}`;
      Logger.log(errorMsg);
      return errorMsg;
    }
  }

// =============================================================================
// 5. 출퇴근 기록 모듈
// =============================================================================
const AttendanceManager = {
  // 출퇴근 기록 처리 (비밀번호 검증 포함)
  recordAttendanceWithPassword: function(employeeName, password, type) {
    try {
      // 비밀번호 검증
      const verification = EmployeeManager.verifyEmployeePassword(employeeName, password);
      
      if (!verification.success) {
        return verification; // 비밀번호 검증 실패 결과 반환
      }
      
      // 비밀번호 검증 성공 시 출퇴근 기록 처리
      return this.recordAttendance(employeeName, type);
      
    } catch (e) {
      Utils.logError('recordAttendanceWithPassword', e);
      return {
        success: false,
        message: `출퇴근 처리 중 오류가 발생했습니다: ${e.message}`,
        timestamp: new Date().toString()
      };
    }
  },
  
  // 출퇴근 기록 처리
  recordAttendance: function(employeeName, type) {
    try {
      const now = new Date();
      const currentMonth = now.getMonth() + 1;
      const currentDay = now.getDate();
      const currentHour = now.getHours();
      
      // 현재 월의 출퇴근기록 시트 가져오기
      let sheet = SheetManager.getAttendanceSheet(currentMonth);
      
      if (!sheet) {
        return {
          success: false,
          message: '출퇴근기록 시트를 찾을 수 없습니다.'
        };
      }

      // 10분 단위로 반올림/내림
      const roundedTime = Utils.roundTimeToTenMinutes(now);
      const currentTime = Utilities.formatDate(roundedTime, 'Asia/Seoul', CONSTANTS.TIME_FORMAT);
      
      // 헤더 데이터 가져오기 (직원 이름 행과 열 헤더 행)
      const headerRows = sheet.getRange(1, 1, 2, sheet.getLastColumn()).getValues();
      const employeeRow = headerRows[0];
      const columnTypesRow = headerRows[1];

      // 직원의 열 찾기
      let employeeStartCol = -1;
      for(let i = 3; i < employeeRow.length; i += 3) {
        if(employeeRow[i] === employeeName) {
          employeeStartCol = i + 1;  // +1 because getRange is 1-based
          break;
        }
      }

      if(employeeStartCol === -1) {
        return {
          success: false,
          message: '직원 정보를 찾을 수 없습니다.'
        };
      }

      // 날짜 열 데이터 가져오기
      const dateColumn = sheet.getRange(3, 1, sheet.getLastRow(), 1).getValues();
      
      // 현재 날짜 포맷
      const today = Utilities.formatDate(now, 'Asia/Seoul', CONSTANTS.DATE_FORMAT);
      
      // 오늘 날짜의 행 찾기
      let todayRow = -1;
      for(let i = 0; i < dateColumn.length; i++) {
        if(dateColumn[i][0] instanceof Date) {
          const rowDate = Utilities.formatDate(dateColumn[i][0], 'Asia/Seoul', CONSTANTS.DATE_FORMAT);
          if(rowDate === today) {
            todayRow = i + 3;  // +3 because we start from row 3 and getRange is 1-based
            break;
          }
        }
      }

      if(todayRow === -1) {
        return {
          success: false,
          message: '오늘 날짜를 시트에서 찾을 수 없습니다.'
        };
      }

      // 출근/퇴근 처리
      if(type === '출근') {
        // 이미 출근했는지 확인
        const checkInValue = sheet.getRange(todayRow, employeeStartCol).getValue();
        if(checkInValue && checkInValue !== '미출근') {
          return {
            success: false,
            message: `${employeeName}님은 오늘 이미 출근 처리되었습니다.`,
            timestamp: roundedTime.toString()
          };
        }

        // 출근 시간 입력
        sheet.getRange(todayRow, employeeStartCol).setValue(currentTime);
        
        return {
          success: true,
          message: `${employeeName}님 ${type} 처리 완료 (${currentTime})`,
          timestamp: roundedTime.toString()
        };
        
      } else if(type === '퇴근') {
        // 퇴근 처리 중 월 전환 특별 케이스 확인 (1일 새벽 00시~04시)
        if (currentDay === 1 && currentHour >= 0 && currentHour < 4) {
          // 이전 월 계산
          const prevMonth = currentMonth === 1 ? 12 : currentMonth - 1;
          const prevMonthSheet = SheetManager.getAttendanceSheet(prevMonth);
          
          // 이전 월 시트가 있으면 처리
          if (prevMonthSheet) {
            Logger.log(`월 전환 퇴근 감지: ${prevMonth}월 시트 확인`);
            
            // 이전 월의 마지막 날 계산
            const lastDayOfPrevMonth = new Date(
              currentMonth === 1 ? now.getFullYear() - 1 : now.getFullYear(), 
              currentMonth === 1 ? 11 : currentMonth - 2, 
              0
            ).getDate();
            
            // 이전 월 시트에서 직원 열 찾기
            const prevHeaderRows = prevMonthSheet.getRange(1, 1, 2, prevMonthSheet.getLastColumn()).getValues();
            const prevEmployeeRow = prevHeaderRows[0];
            
            let prevEmployeeStartCol = -1;
            for(let i = 3; i < prevEmployeeRow.length; i += 3) {
              if(prevEmployeeRow[i] === employeeName) {
                prevEmployeeStartCol = i + 1;
                break;
              }
            }
            
            if(prevEmployeeStartCol === -1) {
              Logger.log(`이전 월(${prevMonth}월) 시트에서 직원 정보를 찾을 수 없음`);
            } else {
              // 이전 월 마지막 날의 데이터 찾기
              const prevDateColumn = prevMonthSheet.getRange(3, 1, prevMonthSheet.getLastRow(), 1).getValues();
              let lastDayRow = -1;
              
              for(let i = 0; i < prevDateColumn.length; i++) {
                if(prevDateColumn[i][0] instanceof Date) {
                  const rowDay = prevDateColumn[i][0].getDate();
                  if(rowDay === lastDayOfPrevMonth) {
                    lastDayRow = i + 3;  // +3 because we start from row 3
                    break;
                  }
                }
              }
              
              if(lastDayRow === -1) {
                Logger.log(`이전 월 시트에서 마지막 날(${lastDayOfPrevMonth}일)을 찾을 수 없음`);
              } else {
                // 이전 월 마지막 날의 출근 기록 확인
                const checkInValue = prevMonthSheet.getRange(lastDayRow, prevEmployeeStartCol).getValue();
                const checkOutValue = prevMonthSheet.getRange(lastDayRow, prevEmployeeStartCol + 1).getValue();
                
                // 출근 기록이 있고 퇴근 기록이 없으면 이전 월 시트에 퇴근 처리
                if(checkInValue && checkInValue !== '미출근' && 
                  (!checkOutValue || checkOutValue === '' || checkOutValue === '미퇴근')) {
                  
                  // 이미 퇴근 처리되었는지 다시 확인
                  if(checkOutValue && checkOutValue !== '미퇴근') {
                    return {
                      success: false,
                      message: `${employeeName}님은 이미 퇴근 처리되었습니다.`,
                      timestamp: roundedTime.toString()
                    };
                  }
                  
                  // 퇴근 시간 입력
                  prevMonthSheet.getRange(lastDayRow, prevEmployeeStartCol + 1).setValue(currentTime);
                  
                  return {
                    success: true,
                    message: `${employeeName}님 ${type} 처리 완료 (${currentTime}) - ${prevMonth}월 ${lastDayOfPrevMonth}일 기록에 추가되었습니다.`,
                    timestamp: roundedTime.toString()
                  };
                }
              }
            }
            // 이전 월 마지막 날에 출근 기록이 없거나 이미 퇴근 처리됨 - 일반 처리로 진행
            Logger.log(`이전 월 마지막 날에 출근 기록 없음 또는 이미 퇴근 처리됨 - 일반 처리로 진행`);
          }
        }
        
        // 일반적인 퇴근 처리 (이전 월 처리가 안 된 경우 포함)
        let targetRow = todayRow;
        
        // 00시~04시 사이라면 전날 데이터 확인 (월 전환이 아닌 경우)
        if(currentHour >= 0 && currentHour < 4 && currentDay !== 1) {
          // 전날 날짜 계산
          const yesterday = new Date(now);
          yesterday.setDate(yesterday.getDate() - 1);
          const yesterdayStr = Utilities.formatDate(yesterday, 'Asia/Seoul', CONSTANTS.DATE_FORMAT);
          
          // 전날 행 찾기
          let yesterdayRow = -1;
          for(let i = 0; i < dateColumn.length; i++) {
            if(dateColumn[i][0] instanceof Date) {
              const rowDate = Utilities.formatDate(dateColumn[i][0], 'Asia/Seoul', CONSTANTS.DATE_FORMAT);
              if(rowDate === yesterdayStr) {
                yesterdayRow = i + 3;
                break;
              }
            }
          }
          
          // 전날에 출근 기록이 있는지 확인
          if(yesterdayRow > 0) {
            const yesterdayCheckIn = sheet.getRange(yesterdayRow, employeeStartCol).getValue();
            const yesterdayCheckOut = sheet.getRange(yesterdayRow, employeeStartCol + 1).getValue();
            
            // 전날 출근 기록이 있고 퇴근 기록이 없으면 전날 행에 퇴근 처리
            if(yesterdayCheckIn && yesterdayCheckIn !== '미출근' && 
              (!yesterdayCheckOut || yesterdayCheckOut === '' || yesterdayCheckOut === '미퇴근')) {
              targetRow = yesterdayRow;
            }
          }
        }
        
        // 최종 결정된 행에서 출퇴근 상태 확인
        const checkInValue = sheet.getRange(targetRow, employeeStartCol).getValue();
        const checkOutValue = sheet.getRange(targetRow, employeeStartCol + 1).getValue();

        // 출근 기록 없는 경우
        if(!checkInValue || checkInValue === '미출근') {
          // 출근 기록이 없는 경우
          sheet.getRange(targetRow, employeeStartCol).setValue('미출근');
          sheet.getRange(targetRow, employeeStartCol + 1).setValue(currentTime);
          return {
            success: true,
            message: `${employeeName}님 퇴근 처리 완료 (${currentTime}) - 출근 기록이 없습니다.`,
            timestamp: roundedTime.toString()
          };
        }

        // 이미 퇴근한 경우
        if(checkOutValue && checkOutValue !== '미퇴근') {
          return {
            success: false,
            message: `${employeeName}님은 이미 퇴근 처리되었습니다.`,
            timestamp: roundedTime.toString()
          };
        }

        // 퇴근 시간 입력
        sheet.getRange(targetRow, employeeStartCol + 1).setValue(currentTime);
        
        // 00시 이후에 전날 행에 기록한 경우 특별 메시지
        if(targetRow !== todayRow) {
          return {
            success: true,
            message: `${employeeName}님 ${type} 처리 완료 (${currentTime}) - 전날 기록에 추가되었습니다.`,
            timestamp: roundedTime.toString()
          };
        }
        
        return {
          success: true,
          message: `${employeeName}님 ${type} 처리 완료 (${currentTime})`,
          timestamp: roundedTime.toString()
        };
      }

      // 알 수 없는 처리 타입
      return {
        success: false,
        message: `알 수 없는 처리 타입: ${type}`
      };

    } catch (e) {
      Utils.logError('recordAttendance', e);
      return {
        success: false,
        message: '출퇴근 기록 중 오류가 발생했습니다: ' + e.message
      };
    }
  },
  
  // 전날 미퇴근 처리 함수
  checkAndMarkMissingCheckouts: function() {
    try {
      // 1. 현재 날짜 정보 확인
      const now = new Date();
      const currentMonth = now.getMonth() + 1;
      const currentDay = now.getDate();
      
      let updatedCount = 0;
      
      // 2. 현재 월 시트 처리 (기존 로직)
      const sheet = SheetManager.getAttendanceSheet(currentMonth);
      
      if (sheet) {
        // 전날 날짜 계산
        const yesterday = new Date(now);
        yesterday.setDate(yesterday.getDate() - 1);
        const yesterdayStr = Utilities.formatDate(yesterday, 'Asia/Seoul', CONSTANTS.DATE_FORMAT);
        
        Logger.log(`전날 미퇴근 처리 시작: ${yesterdayStr}`);
        
        // 현재 월 시트 데이터 가져오기
        const data = sheet.getDataRange().getValues();
        
        // 전날 행 찾기
        let yesterdayRow = -1;
        for (let row = 2; row < data.length; row++) {
          if (data[row][0] instanceof Date) {
            const rowDateStr = Utilities.formatDate(data[row][0], 'Asia/Seoul', CONSTANTS.DATE_FORMAT);
            if (rowDateStr === yesterdayStr) {
              yesterdayRow = row;
              break;
            }
          }
        }
        
        if (yesterdayRow === -1) {
          Logger.log(`${yesterdayStr} 데이터를 찾을 수 없음`);
        } else {
          // 직원별 출퇴근 열 찾기
          const employeeColumns = {};
          for (let col = 3; col < data[0].length; col += 3) {
            const empName = data[0][col];
            if (empName && !['날짜', '요일', '주차'].includes(empName)) {
              employeeColumns[empName] = {
                checkInCol: col, 
                checkOutCol: col + 1
              };
            }
          }
          
          // 각 직원의 미퇴근 확인
          for (const empName in employeeColumns) {
            const checkInCol = employeeColumns[empName].checkInCol;
            const checkOutCol = employeeColumns[empName].checkOutCol;
            
            // 출근 기록은 있고 퇴근 기록이 없는 경우
            const checkInValue = data[yesterdayRow][checkInCol];
            const checkOutValue = data[yesterdayRow][checkOutCol];
            
            if (checkInValue && checkInValue !== '미출근' && 
                (!checkOutValue || checkOutValue === '' || checkOutValue === '미퇴근')) {
              
              // 미퇴근 처리
              Logger.log(`${empName} 미퇴근 처리 (행: ${yesterdayRow+1})`);
              sheet.getRange(yesterdayRow + 1, checkOutCol + 1).setValue('미퇴근');
              updatedCount++;
            }
          }
        }
      }
      
      // 3. 날짜가 1일이면 이전 월 마지막 날도 처리 (추가된 로직)
      if (currentDay === 1) {
        // 이전 월 계산
        const prevMonth = currentMonth === 1 ? 12 : currentMonth - 1;
        const prevYear = currentMonth === 1 ? now.getFullYear() - 1 : now.getFullYear();
        
        const prevMonthSheet = SheetManager.getAttendanceSheet(prevMonth);
        
        if (prevMonthSheet) {
          // 이전 월의 마지막 날 계산
          const lastDayOfPrevMonth = new Date(prevYear, currentMonth === 1 ? 11 : currentMonth - 2, 0);
          const lastDayStr = Utilities.formatDate(lastDayOfPrevMonth, 'Asia/Seoul', CONSTANTS.DATE_FORMAT);
          
          Logger.log(`월 전환 감지: 이전 월(${prevMonth}월) 마지막 날(${lastDayStr}) 확인`);
          
          // 이전 월 시트 데이터 가져오기
          const prevData = prevMonthSheet.getDataRange().getValues();
          
          // 마지막 날 행 찾기
          let lastDayRow = -1;
          for (let row = 2; row < prevData.length; row++) {
            if (prevData[row][0] instanceof Date) {
              const rowDateStr = Utilities.formatDate(prevData[row][0], 'Asia/Seoul', CONSTANTS.DATE_FORMAT);
              if (rowDateStr === lastDayStr) {
                lastDayRow = row;
                break;
              }
            }
          }
          
          if (lastDayRow === -1) {
            Logger.log(`이전 월 마지막 날(${lastDayStr}) 데이터를 찾을 수 없음`);
          } else {
            // 직원별 출퇴근 열 찾기 (이전 월 시트)
            const prevEmployeeColumns = {};
            for (let col = 3; col < prevData[0].length; col += 3) {
              const empName = prevData[0][col];
              if (empName && !['날짜', '요일', '주차'].includes(empName)) {
                prevEmployeeColumns[empName] = {
                  checkInCol: col, 
                  checkOutCol: col + 1
                };
              }
            }
            
            // 각 직원의 미퇴근 확인 (이전 월 시트)
            for (const empName in prevEmployeeColumns) {
              const checkInCol = prevEmployeeColumns[empName].checkInCol;
              const checkOutCol = prevEmployeeColumns[empName].checkOutCol;
              
              // 출근 기록은 있고 퇴근 기록이 없는 경우
              const checkInValue = prevData[lastDayRow][checkInCol];
              const checkOutValue = prevData[lastDayRow][checkOutCol];
              
              if (checkInValue && checkInValue !== '미출근' && 
                  (!checkOutValue || checkOutValue === '' || checkOutValue === '미퇴근')) {
                
                // 미퇴근 처리
                Logger.log(`${empName} 미퇴근 처리 (이전 월 마지막 날, 행: ${lastDayRow+1})`);
                prevMonthSheet.getRange(lastDayRow + 1, checkOutCol + 1).setValue('미퇴근');
                updatedCount++;
              }
            }
          }
        }
      }
      
      Logger.log(`미퇴근 처리 완료: ${updatedCount}건`);
      return updatedCount > 0;
      
    } catch (e) {
      Utils.logError('checkAndMarkMissingCheckouts', e);
      return false;
    }
  }
};

// =============================================================================
// 6. 요약 정보 모듈
// =============================================================================
const SummaryManager = {
  // 요약 섹션 수식 설정 함수
  setupSummaryFormulas: function(sheet, employee, startRow, startCol) {
    try {
      // 주차별 시간 계산 수식 설정
      for (let week = 1; week <= 5; week++) {
        const weekRow = startRow + 2 + week;
        
        // 중요: 근무시간 열은 startCol + 2입니다 (출근시간, 퇴근시간, 근무시간 순서이므로)
        const hourColumn = startCol + 2;
        const hourColumnLetter = Utils.getColumnLetter(hourColumn);
        
        // 주차별 근무시간 합계 계산 (해당 주차의 데이터만 합산)
        const hoursFormula = `=SUMIFS(${hourColumnLetter}3:${hourColumnLetter}${startRow - 3}, C3:C${startRow - 3}, "${week}주")`;
        sheet.getRange(weekRow, startCol + 1).setFormula(hoursFormula);
        sheet.getRange(weekRow, startCol + 1).setNumberFormat('0.0');
        
        // 금액 계산 수식 (시간 * 시급)
        const salaryFormula = `=${Utils.getColumnLetter(startCol + 1)}${weekRow}*${Utils.getColumnLetter(startCol + 2)}${startRow + 1}`;
        sheet.getRange(weekRow, startCol + 2).setFormula(salaryFormula);
        sheet.getRange(weekRow, startCol + 2).setNumberFormat('#,##0');
        
        // 주차별 주휴수당 및 세금 계산
        const taxRow = startRow + 8 + week;
        
        const bonusFormula = `=IF(${Utils.getColumnLetter(startCol + 1)}${weekRow}>=15, 
        ${Utils.getColumnLetter(startCol + 1)}${weekRow} / 5 * ${Utils.getColumnLetter(startCol + 2)}${startRow + 1}, 
        0)`;
        sheet.getRange(taxRow, startCol).setFormula(bonusFormula);
        sheet.getRange(taxRow, startCol).setNumberFormat('#,##0');

        // 주차별 급여 참조
        sheet.getRange(taxRow, startCol + 1).setFormula(`=${Utils.getColumnLetter(startCol + 2)}${weekRow}`);
        sheet.getRange(taxRow, startCol + 1).setNumberFormat('#,##0');
        
        // 3.3% 세금 계산
        sheet.getRange(taxRow, startCol + 2).setFormula(`=(${Utils.getColumnLetter(startCol)}${taxRow}+${Utils.getColumnLetter(startCol + 1)}${taxRow})*0.033`);
        sheet.getRange(taxRow, startCol + 2).setNumberFormat('#,##0');
      }
      
      // 주유합 계산 (모든 주차 주휴수당의 합)
      const bonusSum = `=SUM(${Utils.getColumnLetter(startCol)}${startRow + 9}:${Utils.getColumnLetter(startCol)}${startRow + 13})`;
      sheet.getRange(startRow + 15, startCol).setFormula(bonusSum);
      sheet.getRange(startRow + 15, startCol).setNumberFormat('#,##0');
      
      // 총합 계산 (주차별 합계의 총액)
      const totalSum = `=SUM(${Utils.getColumnLetter(startCol + 1)}${startRow + 9}:${Utils.getColumnLetter(startCol + 1)}${startRow + 13})`;
      sheet.getRange(startRow + 15, startCol + 1).setFormula(totalSum);
      sheet.getRange(startRow + 15, startCol + 1).setNumberFormat('#,##0');
      
      // 3.3% 세금 총액
      const taxSum = `=SUM(${Utils.getColumnLetter(startCol + 2)}${startRow + 9}:${Utils.getColumnLetter(startCol + 2)}${startRow + 13})`;
      sheet.getRange(startRow + 15, startCol + 2).setFormula(taxSum);
      sheet.getRange(startRow + 15, startCol + 2).setNumberFormat('#,##0');
      
      // 최종 지급액 계산 (주휴합 + 총합 - 3.3% 세금)
      const finalAmount = `=${Utils.getColumnLetter(startCol)}${startRow + 15} + ${Utils.getColumnLetter(startCol + 1)}${startRow + 15} - ${Utils.getColumnLetter(startCol + 2)}${startRow + 15}`;
      sheet.getRange(startRow + 17, startCol + 2).setFormula(finalAmount);
      sheet.getRange(startRow + 17, startCol + 2).setNumberFormat('#,##0');
      sheet.getRange(startRow + 17, startCol + 2).setFontWeight('bold');
      
      return true;
    } catch (e) {
      Utils.logError('setupSummaryFormulas', e);
      return false;
    }
  },
  
  // 직원 요약 섹션 생성
  createEmployeeSummarySection: function(sheet, employee, startRow, startCol) {
    try {
      // 직원 이름 헤더 - 병합 적용
      sheet.getRange(startRow, startCol, 1, 3).merge();
      sheet.getRange(startRow, startCol).setValue(employee.name);
      sheet.getRange(startRow, startCol, 1, 3).setHorizontalAlignment('center');
      sheet.getRange(startRow, startCol).setFontWeight('bold');
      sheet.getRange(startRow, startCol, 1, 3).setBorder(true, true, true, true, true, true);
      
      // 각 행 정보 설정 (시급, 구분, 주차별 데이터 등)
      const rows = [
        ['시급', '', employee.hourlyRate || CONSTANTS.DEFAULT_HOURLY_RATE],
        ['구분', '시간', '금액'],
        ['1주차', 0, '₩0'],
        ['2주차', 0, '₩0'],
        ['3주차', 0, '₩0'],
        ['4주차', 0, '₩0'],
        ['5주차', 0, '₩0'],
        ['주차별 주휴', '주차별 합계', '3.3%'],
        ['', '₩0', '₩0'],
        ['', '₩0', '₩0'],
        ['', '₩0', '₩0'],
        ['', '₩0', '₩0'],
        ['', '₩0', '₩0'],
        ['주휴합', '총합', '3.3%합'],
        ['₩0', '₩0', '₩0'],
        ['', '', '지급액'],
        ['', '', '₩0']
      ];
      
      // 데이터 입력
      for (let i = 0; i < rows.length; i++) {
        for (let j = 0; j < rows[i].length; j++) {
          if (rows[i][j] !== '') {
            sheet.getRange(startRow + i + 1, startCol + j).setValue(rows[i][j]);
          }
        }
      }
      
      // 구분 행 강조
      sheet.getRange(startRow + 2, startCol, 1, 3).setBackground('#f3f3f3');
      
      // 주차별 주휴 행 배경색 설정
      sheet.getRange(startRow + 8, startCol, 1, 3).setBackground('#FFFF00');
      
      // 주유합/총합 행 배경색 설정
      sheet.getRange(startRow + 14, startCol, 1, 3).setBackground('#FFFF00');
      
      // 지급액 행 배경색 설정
      sheet.getRange(startRow + 16, startCol + 2).setBackground('#FFFF00');
      
      // 시급 셀 서식
      sheet.getRange(startRow + 1, startCol + 2).setNumberFormat('#,##0');
      
      // 수식 설정
      this.setupSummaryFormulas(sheet, employee, startRow, startCol);
      
    } catch (e) {
      Utils.logError('createEmployeeSummarySection', e);
    }
  },
  
  // 시트에서 직원 요약 정보 가져오기 - 급여 정보 제외 버전
  getEmployeeSummaryFromSheet: function(monthNumber, employeeId) {
    Logger.log(`getEmployeeSummaryFromSheet 호출: month=${monthNumber}, employeeId=${employeeId}`);
    try {
      const sheet = SheetManager.getAttendanceSheet(monthNumber);
      
      if (!sheet) {
        return { error: `${monthNumber}월_출퇴근기록 시트를 찾을 수 없습니다.` };
      }

      // 직원 정보 가져오기
      const employees = EmployeeManager.getEmployees();
      const employee = employees.find(emp => emp.id == employeeId);
      if (!employee) {
        Logger.log(`직원 ID ${employeeId}에 해당하는 직원을 찾을 수 없습니다.`);
        return { error: `직원 정보를 찾을 수 없습니다.` };
      }

      // 시트의 모든 데이터 가져오기
      const data = sheet.getDataRange().getValues();
      
      // 1. 출퇴근 데이터 수집 (날짜별)
      const attendance = {};
      
      // 직원 열 찾기
      let employeeCol = -1;
      for (let col = 3; col < data[0].length; col += 3) {
        if (data[0][col] === employee.name) {
          employeeCol = col;
          break;
        }
      }
      
      if (employeeCol === -1) {
        return { error: `해당 월 시트에서 ${employee.name}님의 출퇴근 기록을 찾을 수 없습니다.` };
      }
      
      // 날짜 데이터 시작 위치 (일반적으로 3행부터)
      let dateStartRow = -1;
      for (let row = 0; row < data.length; row++) {
        if (data[row][0] instanceof Date) {
          dateStartRow = row;
          break;
        }
      }
      
      if (dateStartRow === -1) {
        return { error: `시트에서 날짜 데이터를 찾을 수 없습니다.` };
      }
      
      // 출퇴근 기록 수집 (날짜별)
      let workingDays = 0;
      let totalHours = 0;
      
      for (let row = dateStartRow; row < data.length; row++) {
        // 근무 요약 정보 섹션 시작되면 종료
        if (data[row][0] === CONSTANTS.SUMMARY_HEADER || !data[row][0]) break;
        
        if (data[row][0] instanceof Date) {
          const date = data[row][0];
          const dateStr = Utilities.formatDate(date, 'Asia/Seoul', CONSTANTS.DATE_FORMAT);
          
          let checkIn;
          if (data[row][employeeCol] instanceof Date) {
              const checkInDate = new Date(data[row][employeeCol]);
              checkInDate.setHours(checkInDate.getHours() + 1); // 한 시간 추가
              checkInDate.setMinutes(checkInDate.getMinutes() - 27); // 27분 빼기
              checkIn = Utilities.formatDate(checkInDate, 'Asia/Seoul', CONSTANTS.TIME_FORMAT);
          } else {
              checkIn = data[row][employeeCol];
          }

          let checkOut;
          if (data[row][employeeCol + 1] instanceof Date) {
              const checkOutDate = new Date(data[row][employeeCol + 1]);
              checkOutDate.setHours(checkOutDate.getHours() + 1); // 한 시간 추가
              checkOutDate.setMinutes(checkOutDate.getMinutes() - 27); // 27분 빼기
              checkOut = Utilities.formatDate(checkOutDate, 'Asia/Seoul', CONSTANTS.TIME_FORMAT);
          } else {
              checkOut = data[row][employeeCol + 1];
          }

          // 근무 시간 값 가져오기
          const workHours = (data[row][employeeCol + 2] instanceof Date)
              ? Utilities.formatDate(data[row][employeeCol + 2], 'Asia/Seoul', CONSTANTS.TIME_FORMAT)
              : data[row][employeeCol + 2];

          attendance[dateStr] = {
            checkIn: checkIn || null,
            checkOut: checkOut || null,
            workHours: workHours || 0
          };
          
          // 근무일 및 시간 계산
          if (workHours > 0) {
            workingDays++;
            totalHours += parseFloat(workHours);
          }
        }
      }
      
      Logger.log(`총 근무일: ${workingDays}, 총 근무시간: ${totalHours}`);

      // 요약 정보 섹션 찾기
      let summaryStartRow = -1;
      for (let row = 0; row < data.length; row++) {
        if (data[row][0] === CONSTANTS.SUMMARY_HEADER) {
          summaryStartRow = row;
          break;
        }
      }
      
      // 요약 정보가 없으면 기본 정보만 반환
      if (summaryStartRow === -1) {
        return { 
          attendance,
          summary: {
            employeeName: employee.name,
            totalHours: totalHours,
            workingDays: workingDays
          }
        };
      }
      
      // 직원 요약 정보 찾기
      let empSummaryCol = -1;
      for (let col = 3; col < data[summaryStartRow].length; col += 3) {
        if (data[summaryStartRow][col] === employee.name) {
          empSummaryCol = col;
          break;
        }
      }
      
      // 직원 요약 정보가 없으면 기본 정보만 반환
      if (empSummaryCol === -1) {
        return { 
          attendance,
          summary: {
            employeeName: employee.name,
            totalHours: totalHours,
            workingDays: workingDays
          }
        };
      }
      
      // 주차별 시간 정보 찾기
      let weeklyHours = {
        week1: 0,
        week2: 0,
        week3: 0,
        week4: 0,
        week5: 0
      };
      
      // 주차별 근무시간 찾기
      let weeklyRowStart = -1;
      for (let row = summaryStartRow; row < data.length; row++) {
        if (data[row][empSummaryCol] === '구분' && data[row][empSummaryCol + 1] === '시간') {
          weeklyRowStart = row + 1;
          break;
        }
      }
      
      // 주차별 시간 데이터 수집
      if (weeklyRowStart !== -1) {
        for (let week = 0; week < 5; week++) {
          const row = weeklyRowStart + week;
          if (row < data.length && data[row][empSummaryCol] === `${week + 1}주차`) {
            weeklyHours[`week${week + 1}`] = data[row][empSummaryCol + 1] || 0;
          }
        }
      }
      
      // 결과 반환 - 시간 관련 정보만 포함
      const result = {
        attendance,
        summary: {
          employeeName: employee.name,
          totalHours: totalHours,
          workingDays: workingDays,
          weeklyHours: weeklyHours
        }
      };
      
      Logger.log("결과 객체 생성 완료 (급여 정보 제외)");
      
      return result;

    } catch (e) {
      Utils.logError('getEmployeeSummaryFromSheet', e);
      return { error: `데이터 처리 중 오류가 발생했습니다: ${e.message}` };
    }
  },
  
  // 요약 정보 가져오기 (API)
  getSummaryData: function(year, month, employeeId) {
    try {
      // 매개변수 로깅
      Logger.log(`getSummaryData 호출: year=${year}, month=${month}, employeeId=${employeeId}`);
      
      // 월 계산 (JavaScript는 0부터 시작)
      const monthNumber = Number(month) + 1;
      Logger.log(`계산된 월: ${monthNumber}`);
      
      // 직원 요약 정보 가져오기
      const result = this.getEmployeeSummaryFromSheet(monthNumber, employeeId);
      
      // 결과 확인
      if (!result) {
        Logger.log("getEmployeeSummaryFromSheet에서 null 반환됨");
        return { error: "데이터를 가져올 수 없습니다." };
      }
      
      Logger.log(`getSummaryData 반환 데이터: 출석=${Object.keys(result.attendance || {}).length}개 항목`);
      return result;
    } catch (e) {
      Utils.logError('getSummaryData', e);
      return { 
        error: `데이터 처리 중 오류 발생: ${e.message}`,
        attendance: {},
        summary: { message: "오류 발생" } 
      };
    }
  }
};

// =============================================================================
// 7. 자동화 트리거 모듈
// =============================================================================
const AutomationManager = {
  // 자동 미퇴근 처리 (트리거용 함수)
  autoProcessMissingCheckouts: function() {
    const result = AttendanceManager.checkAndMarkMissingCheckouts();
    Logger.log(`자동 미퇴근 처리 실행 결과: ${result ? '성공' : '없음'}`);
    return result;
  },
  
  // 월별 시트 생성 함수
  createMonthlySheet: function() {
    try {
      const ss = SheetManager.getSpreadsheet();
      const today = new Date();
      const targetMonth = today.getMonth() + 1;
      
      // 시트 이름 생성
      const sheetName = CONSTANTS.SHEET_NAME_FORMAT.replace('%d', targetMonth);
      
      // 이미 존재하는지 확인
      if (ss.getSheetByName(sheetName)) {
        Logger.log(`${sheetName} 시트가 이미 존재합니다.`);
        return;
      }
      
      // 새 시트 생성
      Logger.log(`${sheetName} 시트 생성 시작...`);
      const newSheet = ss.insertSheet(sheetName);
      
      // 충분한 열 확보
      newSheet.insertColumnsAfter(newSheet.getMaxColumns(), 100);

      // 기본 시트 구조 설정
      const { lastDay, totalRow, summaryRow } = SheetManager.setupBasicSheetStructure(newSheet, today);
      
      // 직원 목록 가져오기
      const employees = EmployeeManager.getEmployees();
      
      // 직원별 열 추가
      let nextCol = 4; // 날짜, 요일 다음부터 시작
      for (const employee of employees) {
        nextCol = EmployeeManager.setupEmployeeColumns(newSheet, employee, nextCol, lastDay, totalRow);
      }
      
      // 요약 정보 섹션 헤더 추가
      newSheet.getRange(summaryRow, 1).setValue(CONSTANTS.SUMMARY_HEADER);
      
      // 직원별 요약 섹션 추가
      nextCol = 4; // 다시 처음부터
      for (const employee of employees) {
        SummaryManager.createEmployeeSummarySection(newSheet, employee, summaryRow, nextCol);
        nextCol += 3;
      }
      
      // 데이터 유효성 검사 규칙 설정
      SheetManager.setupValidationRules(newSheet, lastDay);
      
      Logger.log(`${sheetName} 시트가 성공적으로 생성되었습니다.`);
      
    } catch (error) {
      Utils.logError('createMonthlySheet', error);
    }
  },
  
  // 이전/다음 월 시트 존재 여부 확인
  checkAdjacentMonthSheets: function(currentMonth) {
    const prevMonth = currentMonth === 1 ? 12 : currentMonth - 1;
    const nextMonth = currentMonth === 12 ? 1 : currentMonth + 1;
    
    const prevSheet = SheetManager.getAttendanceSheet(prevMonth);
    const nextSheet = SheetManager.getAttendanceSheet(nextMonth);
    
    return {
      hasPrevMonth: prevSheet !== null,
      hasNextMonth: nextSheet !== null
    };
  }
};

// =============================================================================
// 8. 웹 앱 인터페이스
// =============================================================================

// 웹 앱으로 접근 시 실행되는 함수
function doGet(e) {
  try {
    Logger.log("e 파라미터: " + JSON.stringify(e));
    
    // 파라미터 체크
    const page = e && e.parameter && e.parameter.page ? e.parameter.page : '';
    Logger.log("페이지 값: " + page);
    
    if (page === 'worklog') {
      // HtmlOutput으로 변경
      return HtmlService.createHtmlOutputFromFile('Dashboard')
        .setTitle('근무 대시보드')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    } else {
      return HtmlService.createHtmlOutputFromFile('Index')
        .setTitle('옥계노루 출퇴근')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  } catch (error) {
    Logger.log("오류: " + error.message);
    // 오류 화면 표시
    return HtmlService.createHtmlOutput(
      "<h1>오류가 발생했습니다</h1>" +
      "<p>상세 정보: " + error.message + "</p>"
    );
  }
}
// 로고 1v1ijhnNz6YHZKsUN5jweT-yuYcHzWIPX
// 노루 1SgnLFH1Bu23sUiQ20BOygL7UnbMAAB6X
function getLogoImage() {
  try {
    var logoFileId = "1SgnLFH1Bu23sUiQ20BOygL7UnbMAAB6X"; // 실제 파일 ID로 변경
    
    var file = DriveApp.getFileById(logoFileId);
    var contentType = file.getMimeType();
    var blob = file.getBlob();
    
    return {
      data: Utilities.base64Encode(blob.getBytes()),
      contentType: contentType,
      name: file.getName()
    };
  } catch (e) {
    Logger.log("오류 발생: " + e.toString());
    return { error: e.toString() };
  }
}

// 출퇴근 기록 함수 (클라이언트 -> 서버)
function recordAttendanceWithPassword(employeeName, password, type) {
  return AttendanceManager.recordAttendanceWithPassword(employeeName, password, type);
}

// 직원 목록 가져오기 (클라이언트 -> 서버)
function getEmployees() {
  return EmployeeManager.getEmployees();
}

// 직원 요약 정보 가져오기 (클라이언트 -> 서버)
function getEmployeeSummaryFromSheet(month, employeeId) {
  return SummaryManager.getEmployeeSummaryFromSheet(month, employeeId);
}

// 로그 확인용 함수
function checkLogs() {
  const now = new Date();
  Logger.log(`현재 서버 시간: ${now}`);
  Logger.log(`CONSTANTS 확인: ${JSON.stringify(CONSTANTS)}`);
  
  return "로그 출력 완료. Apps Script 편집기의 '실행 > 로그 보기' 메뉴에서 확인하세요.";
}

// 자동화 처리 함수들 (트리거 함수들)
function autoProcessMissingCheckouts() {
  return AutomationManager.autoProcessMissingCheckouts();
}

function createMonthlySheet() {
  return AutomationManager.createMonthlySheet();
}

function deleteOldSheets() {
  return SheetManager.deleteOldSheets();
}
