/**
 * 주님의교회 증명서 신청 — Google Apps Script 서버 코드
 *
 * 저장 위치
 *  - 신청 내용: 내 구글 드라이브의 "증명서 신청 접수" 스프레드시트 → "접수 목록" 탭
 *  - 첨부파일: 내 구글 드라이브의 "증명서 신청 첨부파일" 폴더 (비공개)
 *  - 주민등록번호(기부금영수증): 시트에 남기지 않고 이 스크립트의 비공개 저장소(스크립트 속성)에만 보관
 *                              → "발급 완료"로 바꾸는 순간 자동 삭제
 *  - 관리자 비밀번호: 스크립트 속성 ADMIN_PASSWORD (코드와 시트에는 남지 않음)
 *  - 발급대장: 같은 스프레드시트의 "발급대장" 탭 (주님의교회 TLC-연도-번호, LEVITES MINISTRIES LM-연도-번호)
 *  - 직접 발급 공문·승인서: "직접 발급" 탭
 *  - 발급본 PDF: 내 드라이브 TLC › 문서 › 연도 › 발급 주체 › 구분 폴더
 *
 * 처음 한 번: 편집기 위쪽에서 setupAll 을 선택하고 [실행]
 *   → 스프레드시트와 첨부파일 폴더가 자동으로 만들어집니다.
 */

// ─────────────────────────────────────────────
// 기본값
// ─────────────────────────────────────────────

var TZ = 'Asia/Seoul';

// 깃허브에 올린 신청 화면 주소 (알림 메일의 "관리 화면 열기" 버튼에 쓰입니다)
var SITE_URL = 'https://tlc-sites.github.io/certificate-request/';

var SPREADSHEET_NAME = '증명서 신청 접수';
var FOLDER_NAME = '증명서 신청 첨부파일';
var DEFAULT_NOTIFY_EMAIL = 'dohyun4work@gmail.com';

var SHEET_REQUESTS = '접수 목록';
var SHEET_COPY = '화면 문구';
var SHEET_MAIL = '알림 메일';
var SHEET_LEDGER = '발급대장';
var SHEET_COPY_BACKUP = '화면 문구 (이전)';

// 화면 문구 탭의 판(버전). 이 숫자를 올리면 다음 접속 때 화면 문구 탭을 새 기본값으로 다시 만들고,
// 이전 탭은 "화면 문구 (이전)"으로 남겨 둡니다.
var COPY_VERSION = 3;

// ── 발급대장 (주님의교회·LEVITES MINISTRIES 공식 문서 통합 대장)
var LEDGER_HEADERS = [
  '문서번호', '발급일', '발급 주체', '구분', '문서명', '수신·발급 대상', '대표자', '상태', '파일',
  '접수번호·문서 ID', '비고', '발급 언어', '생년월일', '신청자', '신청 대상', '수령 방법', '발급 부수', '유효기한', '파일 ID'
];
var LCOL = {
  number: 1, date: 2, issuer: 3, category: 4, title: 5, recipient: 6, rep: 7, status: 8, file: 9,
  source: 10, note: 11, lang: 12, birth: 13, applicant: 14, relation: 15, delivery: 16, copies: 17, validUntil: 18, fileId: 19
};
var LEDGER_ISSUED = '발급 완료';
var LEDGER_CANCELLED = '취소';

// 발급 주체 — 문서번호 앞자리(prefix)는 대장 번호가 섞이지 않도록 코드에서만 바꿉니다
var ISSUERS = {
  tlc: { prefix: 'TLC', folder: 'TLC' },
  lm: { prefix: 'LM', folder: 'LEVITES MINISTRIES' }
};
var DOC_CATEGORIES = ['증명서', '추천서', '승인서', '공문', '기타'];

// ── 직접 발급 (온라인 신청 없이 관리자가 바로 작성하는 공문·승인서)
var SHEET_DIRECT = '직접 발급';
var DIRECT_HEADERS = ['문서 ID', '상태', '문서번호', '발급 주체', '구분', '문서명', '수신', '수정일시', '내용 (고치지 마세요)'];
var DCOL = { id: 1, status: 2, number: 3, issuer: 4, category: 5, title: 6, recipient: 7, updatedAt: 8, data: 9 };
var DIRECT_DRAFT = '작성 중';
var DIRECT_FIELDS = {
  issuer: 20, category: 20, heading: 30, docTitle: 200, recipient: 200, reference: 200, subject: 300,
  body: 12000, representativeTitle: 50, representative: 100, note: 500
};

// ── 발급본 PDF 보관 위치: 내 드라이브 TLC › 문서 (없으면 "문서" 폴더를 새로 만듭니다)
var DOCS_ROOT_FOLDER_ID = '15qgqo1eakLmF5jBE18xfDYMvSfoOeqck';
var MAX_PDF_BYTES = 15 * 1024 * 1024;

// 증명서 효력 기간 (개월)
var VALID_MONTHS = 3;

var REQUEST_HEADERS = [
  '접수번호', '접수일시', '처리 상태', '문서 종류', '신청자', '휴대전화번호', '생년월일',
  '소속 부서', '신청 대상', '수령 방법', '희망 발급일', '발급 부수', '신청 내용',
  '첨부파일', '주민등록번호', '수정일시', '발급일시', '상세 데이터 (고치지 마세요)', '첨부파일 ID'
];
var COL = {
  reference: 1, createdAt: 2, status: 3, documentName: 4, name: 5, phone: 6, birthDate: 7,
  department: 8, relation: 9, delivery: 10, requestedDate: 11, copies: 12, summary: 13,
  attachment: 14, rrn: 15, updatedAt: 16, issuedAt: 17, data: 18, attachmentId: 19
};

var STATUSES = ['접수 완료', '담당자 확인 중', '추가 정보 필요', '발급 준비 중', '발급 완료'];
var STATUS_ISSUED = '발급 완료';

var DOCUMENT_IDS = ['attendance', 'membership', 'dedication', 'baptism', 'donation', 'seniorPastor', 'minister',
  'memberRecommendation', 'employment', 'career', 'retirement'];
var RECOMMENDATION_TYPES = ['seniorPastor', 'minister', 'memberRecommendation'];
var CEREMONY_TYPES = ['dedication', 'baptism'];

var ALLOWED_KEYS = [
  'documentType', 'applicantName', 'birthDate', 'phone', 'affiliatedDepartment',
  'ministerName', 'applicantRelation', 'subjectName', 'subjectBirthDate',
  'attendanceStart', 'attendancePeriod', 'purpose', 'submissionTo',
  'registrationDate', 'registeredDepartment', 'ceremonyTime', 'applicationArea', 'recommendingMember',
  'recommenderRelation', 'deadline', 'submissionMethod', 'recommendationPurpose',
  'referenceInfo', 'department', 'position', 'employmentStart', 'employmentEnd',
  'duties', 'certificateLanguage', 'englishName', 'deliveryMethod', 'email',
  'recipientOrganization', 'recipientPerson', 'recipientEmail', 'requestedDate', 'copies', 'notes',
  // 기부금영수증 (지금은 신청 화면에서 숨김)
  'donorName', 'residentNumber', 'registeredAddress', 'taxYear', 'issueType'
];

// 관리자가 발급할 때 적는 항목
var ISSUANCE_KEYS = ['recommendationBody', 'donationAmount', 'donationPeriod', 'donationType',
  'confirmedStart', 'confirmedDate', 'officiant'];
var MAX_EN_KEYS = 20;

// 관리 화면과 시트의 "신청 내용" 칸에 쓰는 항목 이름
var FIELD_LABELS = {
  applicantName: '신청자 이름', birthDate: '생년월일', phone: '휴대전화번호',
  affiliatedDepartment: '소속 부서', ministerName: '담당 교역자', applicantRelation: '신청 대상',
  subjectName: '증명 대상자 이름', subjectBirthDate: '증명 대상자 생년월일', subjectRelation: '증명 대상자와의 관계',
  attendanceStart: '출석 시작 시점 (신청자 기재)', attendancePeriod: '증명할 출석 기간', purpose: '발급 목적',
  ceremonyTime: '헌아·침례 시기 (신청자 기재)', certificateLanguage: '발급 언어', englishName: '영문 이름',
  recipientEmail: '기관 이메일',
  submissionTo: '제출처', deliveryMethod: '수령 방법', email: '이메일', requestedDate: '희망 발급일',
  copies: '발급 부수', notes: '추가 요청사항', templateFileName: '첨부파일명',
  registrationDate: '교회 등록일', registeredDepartment: '등록 부서', donorName: '기부자 이름',
  registeredAddress: '등본상 주소', taxYear: '귀속연도', issueType: '발급 구분',
  applicationArea: '지원·신청 분야', recommendingMember: '추천인', recommenderRelation: '추천인과의 관계',
  deadline: '제출 기한', submissionMethod: '제출 방법', recommendationPurpose: '추천서가 필요한 이유',
  referenceInfo: '추천서 작성 참고사항', department: '소속 부서', position: '직책',
  employmentStart: '재직 시작일', employmentEnd: '재직 종료일·퇴직일', duties: '담당 업무',
  recipientOrganization: '수신 기관', recipientPerson: '기관 담당자', recipientAddress: '발송 주소'
};

var MAX_FILE_BYTES = 10 * 1024 * 1024;
var ALLOWED_FILE_EXT = ['pdf', 'doc', 'docx', 'hwp', 'hwpx', 'jpg', 'jpeg', 'png'];
var ADMIN_SESSION_SECONDS = 6 * 60 * 60; // 6시간

// ─────────────────────────────────────────────
// 화면 문구 기본값 (원래 사이트의 설정을 그대로 옮겼습니다)
// ─────────────────────────────────────────────

function f_(label, placeholder) { return { label: label, placeholder: placeholder || '' }; }

function recommendationFields_(lastLabel) {
  return {
    submissionTo: f_('제출 기관'),
    applicationArea: f_('지원·신청 분야'),
    deadline: f_('제출 기한'),
    submissionMethod: f_('제출 방법', '예: 이메일, 온라인 제출'),
    recommendationPurpose: f_('추천서가 필요한 이유'),
    referenceInfo: f_(lastLabel, '추천서 작성에 참고할 내용을 입력해 주세요.')
  };
}

function employmentDocument_(name, description, sectionTitle, endLabel) {
  return {
    enabled: true, name: name, description: description, sectionTitle: sectionTitle, notice: '',
    fields: {
      department: f_('소속 부서'),
      position: f_('직책'),
      employmentStart: f_('재직 시작일'),
      employmentEnd: f_(endLabel),
      duties: f_('담당 업무', '증명서에 기재할 주요 업무'),
      purpose: f_('발급 목적', '예: 기관 제출'),
      submissionTo: f_('제출처', '제출 기관명')
    }
  };
}

function outputTemplate_(title, statement, labels, sections) {
  return { title: title, statement: statement, recipientSuffix: '귀중', signerLabel: '담임목사', sections: sections || {}, labels: labels };
}

function ceremonyDocument_(name, description, sectionTitle, timeLabel) {
  return {
    enabled: true, name: name, description: description, sectionTitle: sectionTitle,
    notice: '정확한 날짜를 모르셔도 괜찮습니다. 대략의 시기를 적어 주시면 교회 기록을 확인해 기재해 드립니다.',
    fields: {
      ceremonyTime: f_(timeLabel, '예: 2010년 5월쯤, 2015년 부활절'),
      purpose: f_('발급 목적', '예: 기관 제출'),
      submissionTo: f_('제출처', '제출 기관명')
    }
  };
}

// 영문 출력 양식
function outputTemplateEn_(title, statement, labels) {
  return { title: title, statement: statement, recipientPrefix: 'To:', signerLabel: 'Senior Pastor', labels: labels };
}
function employmentOutputLabelsEn_() {
  return { name: 'Name', birthDate: 'Date of Birth', department: 'Department', position: 'Position', period: 'Period of Employment', duties: 'Duties', purpose: 'Purpose' };
}
function recommendationTemplateEn_(title, signerLabel) {
  var t = outputTemplateEn_(title, 'Please write the letter in the admin page.',
    { name: 'Name', birthDate: 'Date of Birth', department: 'Department', applicationArea: 'Applying For' });
  t.signerLabel = signerLabel;
  return t;
}
function ceremonyLabelsEn_(dateLabel) {
  return { name: 'Name', birthDate: 'Date of Birth', date: dateLabel, officiant: 'Officiant', submissionTo: 'Submitted To', purpose: 'Purpose' };
}

function recommendationTemplate_(title, signerLabel) {
  return {
    title: title, statement: '추천서 본문을 관리 화면에서 작성해 주세요.', recipientSuffix: '귀중',
    signerLabel: signerLabel, sections: {},
    labels: { name: '성명', birthDate: '생년월일', department: '소속 부서', applicationArea: '지원·신청 분야' }
  };
}

function employmentOutputLabels_() {
  return { name: '성명', birthDate: '생년월일', department: '소속 부서', position: '직책', period: '재직 기간', duties: '담당 업무', purpose: '발급 목적' };
}

function defaultSettings_() {
  return {
    brand: { churchName: '주님의교회', churchEnglish: "THE LORD'S CHURCH", badge: '온라인 행정' },
    appearance: { primaryColor: '#1a3150', accentColor: '#bd8c48', showLogo: true, logoSize: '18', showIntroDescription: true, showSteps: true },
    certificate: {
      address: '경기도 성남시 분당구 새마을로 43', phone: '031-706-0191', seniorPastor: '스캇 브래너', organizationNumber: '',
      churchNameEn: "The Lord's Church", addressEn: '43 Saemaeul-ro, Bundang-gu, Seongnam-si, Gyeonggi-do, Republic of Korea',
      phoneEn: '+82-31-706-0191', seniorPastorEn: 'Scott Brenner'
    },
    issuers: {
      tlc: {
        name: '주님의교회', subName: "THE LORD'S CHURCH", representativeTitle: '담임목사', representative: '스캇 브래너',
        address: '경기도 성남시 분당구 새마을로 43', phone: '031-706-0191', email: ''
      },
      lm: {
        name: 'LEVITES MINISTRIES', subName: '리바이츠 미니스트리즈', representativeTitle: '대표', representative: '스캇 브래너 (Scott Brenner)',
        address: '', phone: '', email: ''
      }
    },
    direct: { closing: '끝.', sealLabel: '직인' },
    intro: { title: '증명서 신청', description: '필요한 문서를 선택하고 신청 내용을 작성해 주세요. 담당자가 확인한 후 발급 절차를 안내해 드립니다.' },
    steps: { document: '문서 선택', details: '정보 입력', review: '신청 확인' },
    categories: { member: '교인 · 신앙', recommendation: '추천서', employment: '재직 · 경력' },
    documentPrompt: {
      title: '어떤 문서가 필요하신가요?', description: '한 번에 한 종류의 문서를 신청할 수 있습니다.', button: '신청서 작성하기',
      donationNotice: '기부금영수증은 이 화면에서 신청하실 수 없습니다. 발급 방법은 교회 사무실로 문의해 주세요.',
      otherNotice: '목록에 없는 서류가 필요하시면 교회 사무실로 문의해 주세요.'
    },
    common: {
      selectPlaceholder: '선택해 주세요',
      selectedDocumentLabel: '선택한 문서',
      basicTitle: '기본정보',
      sectionTitle: '신청자 정보를 입력해 주세요',
      sectionHelp: '* 표시는 필수 항목입니다.',
      applicantName: f_('이름', '신청자 이름'),
      birthDate: f_('생년월일', '예: 19560312'),
      phone: f_('휴대전화번호', '010-0000-0000'),
      department: f_('소속 부서', '예: 청년부, 예배부'),
      minister: f_('담당 교역자', '알고 계신 경우 입력'),
      relation: f_('신청 대상'),
      relationSelf: '본인',
      relationGuardian: '미성년 자녀의 보호자',
      relationHelp: '증명서는 본인이 신청하시거나, 대상자가 미성년자(만 19세 미만)인 경우 보호자가 신청하실 수 있습니다.',
      subjectTitle: '증명 대상자 (자녀)',
      subjectName: f_('자녀 이름', '증명서에 기재될 이름'),
      subjectBirthDate: f_('자녀 생년월일', '예: 20150312'),
      subjectAdultError: '증명 대상자가 성인인 경우 본인이 직접 신청해 주세요.',
      agreementTitle: '신청 동의',
      previousButton: '이전',
      reviewButton: '신청 내용 확인'
    },
    documents: {
      attendance: {
        enabled: true, name: '교회 출석증명서', description: '교회 출석 및 신앙생활 기간을 확인합니다.',
        sectionTitle: '출석증명 정보',
        notice: '출석 시작 시점은 대략 적어 주시면 됩니다. 교회 기록을 확인한 뒤 증명서에는 정확한 날짜로 기재해 드립니다.',
        fields: {
          attendanceStart: f_('출석을 시작한 시점 (대략)', '예: 2015년쯤, 2018년 봄'),
          attendancePeriod: f_('증명이 필요한 기간', '예: 2024년 1월부터 현재까지'),
          purpose: f_('발급 목적', '예: 기관 제출, 비자 신청'),
          submissionTo: f_('제출처', '제출 기관명')
        }
      },
      membership: {
        enabled: true, name: '등록교인 확인서', description: '주님의교회 등록교인임을 확인합니다.',
        sectionTitle: '등록교인 확인 정보', notice: '',
        fields: {
          registrationDate: f_('교회 등록일'),
          registeredDepartment: f_('소속 부서', '예: 청년부, 예배부'),
          purpose: f_('발급 목적', '예: 기관 제출'),
          submissionTo: f_('제출처', '제출 기관명')
        }
      },
      dedication: ceremonyDocument_('헌아 확인 증명서', '주님의교회에서 헌아한 사실을 확인합니다.', '헌아 확인 정보', '헌아한 시기 (대략)'),
      baptism: ceremonyDocument_('침례 확인 증명서', '주님의교회에서 침례 받은 사실을 확인합니다.', '침례 확인 정보', '침례 받은 시기 (대략)'),
      donation: {
        enabled: false, name: '기부금영수증', description: '연말정산 및 소득세 신고용 영수증을 신청합니다.',
        sectionTitle: '기부금영수증 정보',
        notice: '기부금영수증 발급을 위해 주민등록번호 전체와 등본상 주소를 정확하게 입력해 주세요.',
        fields: {
          donorName: f_('기부자 이름'),
          residentNumber: f_('주민등록번호 전체', '000000-0000000'),
          registeredAddress: f_('등본상 주소', '도로명 주소와 상세주소'),
          taxYear: f_('귀속연도'),
          issueType: f_('발급 구분'),
          purpose: f_('발급 목적')
        }
      },
      seniorPastor: {
        enabled: true, name: '담임목사 추천서', description: '외부 기관 제출용 담임목사 추천서를 신청합니다.',
        sectionTitle: '추천서 정보',
        notice: '추천서는 작성과 확인에 시간이 필요합니다. 제출 기한을 고려하여 여유 있게 신청해 주세요.',
        fields: recommendationFields_('담임목사님께 전달할 내용')
      },
      minister: {
        enabled: true, name: '담당교역자 추천서', description: '담당 교역자의 확인과 추천이 필요한 경우 신청합니다.',
        sectionTitle: '추천서 정보', notice: '담당 교역자의 확인 후 추천서 작성이 진행됩니다.',
        fields: recommendationFields_('담당 교역자에게 전달할 내용')
      },
      memberRecommendation: {
        enabled: true, name: '교인 추천서', description: '교인의 추천이 필요한 기관·지원 분야 제출용 추천서를 신청합니다.',
        sectionTitle: '추천서 정보', notice: '교인의 추천이 필요한 기관이나 지원 분야에 제출할 추천서를 신청합니다.',
        fields: {
          submissionTo: f_('제출 기관'),
          applicationArea: f_('지원·신청 분야'),
          recommendingMember: f_('추천서를 작성할 교인', '정해진 경우 이름을 입력해 주세요.'),
          recommenderRelation: f_('추천인과의 관계', '예: 같은 셀 구성원, 사역팀 동료'),
          deadline: f_('제출 기한'),
          submissionMethod: f_('제출 방법', '예: 이메일, 온라인 제출'),
          recommendationPurpose: f_('추천서가 필요한 이유'),
          referenceInfo: f_('추천인에게 전달할 참고 내용', '추천서 작성에 참고할 내용을 입력해 주세요.')
        }
      },
      employment: employmentDocument_('재직증명서', '현재 교회 재직 사실을 증명합니다.', '재직 정보', '재직 종료일'),
      career: employmentDocument_('경력증명서', '근무 또는 사역 경력과 담당 업무를 확인합니다.', '경력 정보', '경력 종료일'),
      retirement: employmentDocument_('퇴직증명서', '교회 근무 및 퇴직 사실을 증명합니다.', '퇴직 정보', '퇴직일')
    },
    outputCommon: {
      documentNumberLabel: '문서번호', issueDateLabel: '발급일', sealLabel: '직인\n위치', signatureSealLabel: '서명/인',
      validityNote: '이 증명서는 발급일로부터 3개월간 유효합니다.'
    },
    outputCommonEn: {
      documentNumberLabel: 'Document No.', issueDateLabel: 'Date of Issue', sealLabel: 'SEAL', signatureSealLabel: 'Signature',
      validityNote: 'This certificate is valid for three (3) months from the date of issue.',
      sinceSuffix: 'to present'
    },
    outputs: {
      attendance: outputTemplate_('교회 출석증명서', '위 사람은 주님의교회에 등록하여 출석하고 있음을 증명합니다.', {
        name: '성명', birthDate: '생년월일', department: '소속 부서', period: '출석 기간', submissionTo: '제출처', purpose: '발급 목적'
      }),
      membership: outputTemplate_('등록교인 확인서', '위 사람은 주님의교회에 등록된 교인임을 증명합니다.', {
        name: '성명', birthDate: '생년월일', department: '소속 부서', registrationDate: '교회 등록일', submissionTo: '제출처', purpose: '발급 목적'
      }),
      dedication: outputTemplate_('헌아 확인 증명서', '위 사람은 주님의교회에서 헌아하였음을 확인합니다.', {
        name: '성명', birthDate: '생년월일', date: '헌아일', officiant: '집례자', submissionTo: '제출처', purpose: '발급 목적'
      }),
      baptism: outputTemplate_('침례 확인 증명서', '위 사람은 주님의교회에서 침례를 받았음을 확인합니다.', {
        name: '성명', birthDate: '생년월일', date: '침례일', officiant: '집례자', submissionTo: '제출처', purpose: '발급 목적'
      }),
      donation: outputTemplate_('기부금영수증', '위와 같이 기부금을 수령하였음을 확인합니다.', {
        name: '성명', residentNumber: '주민등록번호', address: '주소', taxYear: '귀속연도', period: '기부 기간', donationType: '기부금 유형', amount: '기부금액',
        organizationName: '단체명', organizationNumber: '고유번호', organizationAddress: '소재지', representative: '대표자', phone: '전화번호'
      }, { donor: '1. 기부자', donation: '2. 기부 내용', organization: '3. 기부금 단체' }),
      seniorPastor: recommendationTemplate_('담임목사 추천서', '담임목사'),
      minister: recommendationTemplate_('담당교역자 추천서', '담당교역자'),
      memberRecommendation: recommendationTemplate_('교인 추천서', '추천인'),
      employment: outputTemplate_('재직증명서', '위 사람은 주님의교회에 재직하고 있음을 증명합니다.', employmentOutputLabels_()),
      career: outputTemplate_('경력증명서', '위 사람은 주님의교회에서 위와 같이 근무하였음을 증명합니다.', employmentOutputLabels_()),
      retirement: outputTemplate_('퇴직증명서', '위 사람은 주님의교회에서 위와 같이 근무하고 퇴직하였음을 증명합니다.', employmentOutputLabels_())
    },
    outputsEn: {
      attendance: outputTemplateEn_('Certificate of Church Attendance',
        "This is to certify that the person named above is a registered member of The Lord's Church and has been attending the church.",
        { name: 'Name', birthDate: 'Date of Birth', department: 'Department', period: 'Period of Attendance', submissionTo: 'Submitted To', purpose: 'Purpose' }),
      membership: outputTemplateEn_('Certificate of Church Membership',
        "This is to certify that the person named above is a registered member of The Lord's Church.",
        { name: 'Name', birthDate: 'Date of Birth', department: 'Department', registrationDate: 'Date of Registration', submissionTo: 'Submitted To', purpose: 'Purpose' }),
      dedication: outputTemplateEn_('Certificate of Child Dedication',
        "This is to certify that the person named above was dedicated at The Lord's Church.", ceremonyLabelsEn_('Date of Dedication')),
      baptism: outputTemplateEn_('Certificate of Baptism',
        "This is to certify that the person named above was baptized at The Lord's Church.", ceremonyLabelsEn_('Date of Baptism')),
      seniorPastor: recommendationTemplateEn_('Letter of Recommendation', 'Senior Pastor'),
      minister: recommendationTemplateEn_('Letter of Recommendation', 'Minister'),
      memberRecommendation: recommendationTemplateEn_('Letter of Recommendation', 'Recommender'),
      employment: outputTemplateEn_('Certificate of Employment',
        "This is to certify that the person named above is currently employed at The Lord's Church.", employmentOutputLabelsEn_()),
      career: outputTemplateEn_('Certificate of Career',
        "This is to certify that the person named above has worked at The Lord's Church as stated above.", employmentOutputLabelsEn_()),
      retirement: outputTemplateEn_('Certificate of Retirement',
        "This is to certify that the person named above worked at The Lord's Church and has retired as stated above.", employmentOutputLabelsEn_())
    },
    attachment: {
      title: '지정 양식 및 참고자료',
      description: '제출 기관에서 지정한 양식이 있는 경우 파일을 첨부해 주세요.',
      button: '파일 선택',
      help: 'PDF, Word, 한글, 이미지 파일 · 최대 10MB',
      attachedLabel: '첨부됨',
      fileSizeError: '10MB 이하의 파일을 선택해 주세요.'
    },
    language: {
      title: '발급 언어',
      ko: '한국어',
      en: '영어',
      enNotice: '영문 증명서는 담당자가 내용을 번역해 발급하므로 시간이 조금 더 걸릴 수 있습니다. 증명서에 들어갈 영문 이름을 여권과 같게 적어 주세요.',
      englishNameLabel: '영문 이름 (여권과 같게)',
      englishNamePlaceholder: '예: HONG GILDONG',
      englishNameError: '영문 이름은 영어 알파벳으로 적어 주세요.'
    },
    delivery: {
      title: '수령 방법',
      email: '이메일 수령',
      visit: '교회 방문 수령',
      direct: '제출 기관에 이메일로 발송',
      emailNotice: '이메일로 받기를 원하시는 경우, 문서를 받으실 이메일 주소를 입력해 주세요.',
      emailLabel: '수신 이메일',
      emailConfirmLabel: '이메일 재확인',
      emailPlaceholder: 'name@example.com',
      emailConfirmPlaceholder: '동일한 주소를 다시 입력',
      directNotice: '제출 기관으로는 이메일로만 보내 드립니다. 우편 발송은 하지 않습니다.',
      organizationLabel: '수신 기관명',
      contactLabel: '기관 담당자',
      recipientEmailLabel: '기관 이메일 주소',
      validityNotice: '증명서는 발급일로부터 3개월간 유효합니다. 제출 일정에 맞춰 신청해 주세요.',
      requestedDateLabel: '희망 발급일',
      copiesLabel: '발급 부수',
      copyOptions: '1부\n2부\n3부\n4부 이상',
      requestedDateFallback: '담당자와 협의'
    },
    options: {
      taxYears: '2026\n2025\n2024',
      issueTypes: '처음 발급\n재발급',
      donationPurposes: '연말정산\n종합소득세 신고\n기타'
    },
    validation: {
      documentRequired: '신청할 문서를 선택해 주세요.',
      requiredFields: '필수 항목을 확인해 주세요.',
      emailMismatch: '이메일 주소가 서로 일치하지 않습니다.',
      recipientEmailError: '제출 기관의 이메일 주소를 확인해 주세요.',
      noDocuments: '현재 신청 가능한 증명서가 없습니다.'
    },
    notes: { title: '추가 요청사항', placeholder: '문서에 포함되어야 할 내용이나 담당자에게 전달할 사항을 입력해 주세요.' },
    consent: {
      privacyTitle: '개인정보 수집·이용에 동의합니다.',
      privacyHelp: '수집 목적: 증명서 신청 확인 및 발급\n수집 항목: 이름, 생년월일, 휴대전화번호, 소속 부서, 담당 교역자, 자녀 이름·생년월일(보호자 신청 시), 영문 이름(영문 발급 시), 이메일(이메일 수령 시), 수신 기관명·이메일(기관 발송 시)\n보관 기간: 발급 완료 후 6개월이 지나면 삭제합니다. 다만 발급 사실(발급번호, 발급일, 문서 종류, 대상자)은 발급대장에 기록해 보관합니다.\n증명서 효력: 발급일로부터 3개월간 유효합니다.\n동의하지 않으실 수 있으며, 이 경우 온라인 신청이 어렵습니다. 교회 사무실로 문의해 주세요.',
      sensitiveTitle: '주민등록번호 수집·이용에 동의합니다.',
      sensitiveHelp: '수집 목적: 기부금영수증 발급(소득세법에 따라 영수증에 기재가 필요합니다)\n수집 항목: 주민등록번호\n보관 기간: 영수증 발급 완료 후 즉시 삭제합니다.\n동의하지 않으실 수 있으며, 이 경우 온라인으로 기부금영수증을 신청하기 어렵습니다.'
    },
    review: {
      title: '신청 내용을 확인해 주세요',
      description: '잘못 입력된 내용이 있다면 이전 단계에서 수정할 수 있습니다.',
      noticeTitle: '안내',
      noticeBody: '추천서는 확인과 작성 과정에 시간이 필요할 수 있으며, 신청 내용에 따라 담당자가 추가 자료를 요청할 수 있습니다.',
      backButton: '내용 수정',
      submitButton: '신청하기'
    },
    success: {
      eyebrow: '접수 완료',
      title: '신청이 잘 접수되었습니다',
      description: '담당자가 신청 내용을 확인한 뒤 발급 절차를 안내해 드리겠습니다. 아래 접수번호는 신청 내역을 조회할 때 쓰실 수 있습니다.',
      resetButton: '다른 문서 신청하기'
    },
    footer: { contact: '증명서 발급과 관련한 문의는 교회 행정 담당자에게 연락해 주세요.' }
  };
}

// "화면 문구" 탭의 설명 칸을 만들 때 쓰는 이름표
var GROUP_LABELS = {
  brand: '헤더와 교회명', appearance: '색상과 로고', certificate: '발급 문서 정보', outputCommon: '출력 양식 공통',
  outputCommonEn: '영문 출력 양식 공통', language: '발급 언어', direct: '직접 발급 공문',
  intro: '첫 화면', steps: '신청 단계 이름', categories: '문서 분류 이름', documentPrompt: '문서 선택 안내',
  common: '신청자 기본정보', attachment: '파일 첨부', delivery: '수령 방법', options: '선택 항목',
  validation: '안내·오류 문구', notes: '추가 요청사항', consent: '개인정보 동의', review: '신청 확인 화면',
  success: '접수 완료 화면', footer: '맨 아래'
};
var KEY_LABELS = {
  churchName: '교회명', churchEnglish: '영문 교회명', badge: '맨 위 오른쪽 표시',
  address: '교회 주소', phone: '대표 전화번호', seniorPastor: '담임목사 이름', organizationNumber: '교회 고유번호 (기부금영수증에 필요)',
  churchNameEn: '영문 교회명 (영문 증명서)', addressEn: '영문 주소 (영문 증명서)', phoneEn: '영문 전화번호 표기', seniorPastorEn: '담임목사 영문 이름',
  donationNotice: '기부금영수증 안내 (문서 목록 아래)', otherNotice: '기타 서류 안내 (문서 목록 아래)',
  relationGuardian: '신청 대상 선택지: 보호자', relationHelp: '신청 대상 안내', subjectAdultError: '대상자가 성인일 때 안내',
  ko: '한국어 선택지', en: '영어 선택지', enNotice: '영문 발급 안내', englishNameLabel: '영문 이름 칸 제목',
  englishNamePlaceholder: '영문 이름 칸 안 흐린 안내', englishNameError: '영문 이름 오류 안내',
  directNotice: '기관 발송 안내', recipientEmailLabel: '기관 이메일 칸 제목', validityNotice: '효력 기간 안내',
  recipientEmailError: '기관 이메일 오류 안내', validityNote: '효력 기간 문구', sinceSuffix: '"현재까지" 영문 표기',
  recipientPrefix: '수신 기관 앞 문구', closing: '공문 본문 끝 표시', email: '이메일',
  subName: '부제 (이름 아래 작은 글씨)', representativeTitle: '대표자 직함', representative: '대표자 이름',
  title: '제목', description: '설명', document: '1단계', details: '2단계', review: '3단계',
  member: '교인·헌금 분류', recommendation: '추천서 분류', employment: '재직·경력 분류',
  button: '버튼 문구', sectionTitle: '입력 구역 제목', sectionHelp: '필수 항목 안내', subjectTitle: '증명 대상자 구역 제목',
  email: '이메일 수령 버튼', visit: '교회 방문 수령 버튼', direct: '제출 기관 발송 버튼', emailNotice: '이메일 안내',
  emailLabel: '이메일 칸 제목', emailConfirmLabel: '이메일 재확인 칸 제목', organizationLabel: '수신 기관 칸 제목',
  contactLabel: '기관 담당자 칸 제목', addressLabel: '주소 칸 제목', requestedDateLabel: '희망 발급일 칸 제목', copiesLabel: '발급 부수 칸 제목',
  privacyTitle: '개인정보 동의 제목', privacyHelp: '개인정보 동의 설명', sensitiveTitle: '주민등록번호 동의 제목', sensitiveHelp: '주민등록번호 동의 설명',
  noticeTitle: '안내 제목', noticeBody: '안내 내용', backButton: '수정 버튼', submitButton: '신청 버튼',
  resetButton: '다시 신청 버튼', contact: '문의 안내', notice: '문서별 안내 문구', name: '문서 이름', placeholder: '칸 안 흐린 안내',
  selectedDocumentLabel: '선택 문서 표시', basicTitle: '기본정보 구역 제목', agreementTitle: '동의 구역 제목',
  previousButton: '이전 버튼', reviewButton: '확인 버튼', attachedLabel: '첨부 완료 표시', help: '도움말',
  selectPlaceholder: '선택 칸 안내', relationSelf: '신청 대상 선택지: 본인',
  fileSizeError: '파일 용량 초과 안내', emailPlaceholder: '이메일 칸 안 흐린 안내', emailConfirmPlaceholder: '이메일 재확인 칸 안 흐린 안내',
  copyOptions: '발급 부수 선택지 (한 줄에 하나)', requestedDateFallback: '희망일 미입력 시 표시',
  taxYears: '귀속연도 선택지 (한 줄에 하나)', issueTypes: '발급 구분 선택지 (한 줄에 하나)', donationPurposes: '기부금 발급 목적 선택지 (한 줄에 하나)',
  primaryColor: '대표 색상 (예: #1a3150)', accentColor: '강조 색상 (예: #bd8c48)', logoSize: '출력 문서 로고 크기 (mm, 숫자)',
  showLogo: '출력 문서에 로고 표시 (예/아니요)', showIntroDescription: '첫 화면 설명 보이기 (예/아니요)', showSteps: '단계 표시 보이기 (예/아니요)',
  enabled: '신청 화면에 보이기 (예/아니요)',
  documentNumberLabel: '문서번호 표기', issueDateLabel: '발급일 표기', sealLabel: '직인 자리 표기', signatureSealLabel: '서명 자리 표기',
  statement: '증명 문구', recipientSuffix: '수신 기관 뒤 문구', signerLabel: '서명자 직함',
  documentRequired: '문서 미선택 안내', requiredFields: '필수 항목 안내', emailMismatch: '이메일 불일치 안내', noDocuments: '신청 가능 문서 없음 안내'
};

// ─────────────────────────────────────────────
// 처음 설정 (편집기에서 실행)
// ─────────────────────────────────────────────

function setupAll() {
  var ss = ss_(true);

  // 접수 목록
  var req = ss.getSheetByName(SHEET_REQUESTS) || ss.insertSheet(SHEET_REQUESTS, 0);
  if (req.getLastRow() === 0) {
    req.getRange(1, 1, 1, REQUEST_HEADERS.length).setValues([REQUEST_HEADERS])
      .setFontWeight('bold').setBackground('#1a3150').setFontColor('#ffffff');
    req.setFrozenRows(1);
    req.getRange('F:G').setNumberFormat('@');
    req.getRange('K:K').setNumberFormat('@');
    req.getRange('B:B').setNumberFormat('yyyy-mm-dd hh:mm');
    req.getRange('P:Q').setNumberFormat('yyyy-mm-dd hh:mm');
    req.getRange('M:M').setWrap(true);
    req.setColumnWidth(1, 130); req.setColumnWidth(4, 140); req.setColumnWidth(13, 360);
    req.hideColumns(COL.data, 2);
  }
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).setAllowInvalid(false).build();
  req.getRange(2, COL.status, Math.max(req.getMaxRows() - 1, 1), 1).setDataValidation(rule);

  // 화면 문구 (없는 항목만 아래에 추가하므로 여러 번 실행해도 괜찮습니다)
  if (ss.getSheetByName(SHEET_COPY)) ensureCopyVersion_();
  else PropertiesService.getScriptProperties().setProperty('COPY_VERSION', String(COPY_VERSION));
  fillCopySheet_(ss.getSheetByName(SHEET_COPY) || ss.insertSheet(SHEET_COPY));

  // 발급대장 · 직접 발급
  getLedgerSheet_();
  getDirectSheet_();

  // 알림 메일
  var mail = ss.getSheetByName(SHEET_MAIL) || ss.insertSheet(SHEET_MAIL);
  if (mail.getLastRow() === 0) {
    mail.getRange(1, 1).setValue('새 신청 알림을 받을 이메일 (한 칸에 한 명)').setFontWeight('bold').setBackground('#E6EEF2');
    mail.getRange(2, 1).setValue(DEFAULT_NOTIFY_EMAIL);
    mail.setColumnWidth(1, 320);
  }

  var blank = ss.getSheetByName('시트1') || ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(blank);

  var folder = getFolder_();
  var hasPassword = !!PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  console.log('처음 설정이 끝났습니다.');
  console.log('스프레드시트: ' + ss.getUrl());
  console.log('첨부파일 폴더: ' + folder.getUrl());
  console.log(hasPassword ? '관리자 비밀번호 설정됨' : '관리자 비밀번호(ADMIN_PASSWORD)를 스크립트 속성에 추가해 주세요.');
}

// 화면 문구 탭에 없는 항목을 기본값으로 채웁니다
function fillCopySheet_(copy) {
  if (copy.getLastRow() === 0) {
    copy.getRange(1, 1, 1, 3).setValues([['항목 (바꾸지 마세요)', '화면 어디에 나오나요', '문구 (여기만 고쳐 주세요)']])
      .setFontWeight('bold').setBackground('#E6EEF2');
    copy.setFrozenRows(1);
    copy.setColumnWidth(1, 250); copy.setColumnWidth(2, 300); copy.setColumnWidth(3, 460);
  }
  var existing = {};
  if (copy.getLastRow() > 1) {
    copy.getRange(2, 1, copy.getLastRow() - 1, 1).getValues().forEach(function (r) { existing[String(r[0]).trim()] = true; });
  }
  var rows = flattenSettings_(defaultSettings_()).filter(function (r) { return !existing[r[0]]; })
    .map(function (r) { return [r[0], describeKey_(r[0]), toCell_(r[1])]; });
  if (rows.length) {
    var start = copy.getLastRow() + 1;
    copy.getRange(start, 3, rows.length, 1).setNumberFormat('@');
    copy.getRange(start, 1, rows.length, 3).setValues(rows);
    copy.getRange(start, 1, rows.length, 2).setFontColor('#8a97a2');
    copy.getRange(start, 3, rows.length, 1).setWrap(true);
  }
  if (!copy.getProtections(SpreadsheetApp.ProtectionType.RANGE).length && copy.getLastRow() > 1) {
    copy.getRange(2, 1, copy.getLastRow() - 1, 2).protect().setDescription('항목 이름은 바꾸지 마세요').setWarningOnly(true);
  }
}

// 화면 문구 판(COPY_VERSION)이 바뀌었으면 한 번만 새로 만듭니다. 이전 탭은 "화면 문구 (이전)"으로 남습니다.
function ensureCopyVersion_() {
  var props = PropertiesService.getScriptProperties();
  if (Number(props.getProperty('COPY_VERSION') || 1) >= COPY_VERSION) return;
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var current = Number(props.getProperty('COPY_VERSION') || 1);
    if (current >= COPY_VERSION) return;
    var ss = ss_();
    var old = ss.getSheetByName(SHEET_COPY);
    // 판 2부터는 고쳐 둔 문구를 그대로 두고, 새로 생긴 항목만 아래에 덧붙입니다
    if (current >= 2 && old) {
      fillCopySheet_(old);
      props.setProperty('COPY_VERSION', String(COPY_VERSION));
      SpreadsheetApp.flush();
      return;
    }
    if (old) {
      var prevBackup = ss.getSheetByName(SHEET_COPY_BACKUP);
      if (prevBackup) ss.deleteSheet(prevBackup);
      old.setName(SHEET_COPY_BACKUP);
    }
    fillCopySheet_(ss.insertSheet(SHEET_COPY, old ? old.getIndex() - 1 : ss.getSheets().length));
    props.setProperty('COPY_VERSION', String(COPY_VERSION));
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
}

function getLedgerSheet_() {
  var ss = ss_();
  var sheet = ss.getSheetByName(SHEET_LEDGER);
  if (sheet) return sheet;
  sheet = ss.insertSheet(SHEET_LEDGER, 1);
  sheet.getRange(1, 1, 1, LEDGER_HEADERS.length).setValues([LEDGER_HEADERS])
    .setFontWeight('bold').setBackground('#1a3150').setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  ['A:A', 'J:J', 'M:M'].forEach(function (a) { sheet.getRange(a).setNumberFormat('@'); });
  ['B:B', 'R:R'].forEach(function (a) { sheet.getRange(a).setNumberFormat('yyyy-mm-dd'); });
  sheet.setColumnWidth(1, 130); sheet.setColumnWidth(3, 150); sheet.setColumnWidth(5, 320); sheet.setColumnWidth(6, 170);
  sheet.hideColumns(LCOL.fileId);
  return sheet;
}

function getDirectSheet_() {
  var ss = ss_();
  var sheet = ss.getSheetByName(SHEET_DIRECT);
  if (sheet) return sheet;
  sheet = ss.insertSheet(SHEET_DIRECT, 2);
  sheet.getRange(1, 1, 1, DIRECT_HEADERS.length).setValues([DIRECT_HEADERS])
    .setFontWeight('bold').setBackground('#1a3150').setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  sheet.getRange('A:A').setNumberFormat('@');
  sheet.getRange('C:C').setNumberFormat('@');
  sheet.getRange('H:H').setNumberFormat('yyyy-mm-dd hh:mm');
  sheet.setColumnWidth(6, 320);
  sheet.hideColumns(DCOL.data);
  return sheet;
}

function sendTestMail() {
  var to = getNotifyEmails_();
  if (!to.length) { console.log('"알림 메일" 탭에 이메일 주소를 먼저 적어 주세요.'); return; }
  MailApp.sendEmail({
    to: to.join(','),
    subject: '[증명서 신청] 알림 메일 시험 발송',
    htmlBody: '증명서 신청 알림 메일이 잘 도착했습니다.<br>새 신청이 들어오면 이 주소로 알려 드립니다.',
    name: '주님의교회 증명서 신청'
  });
  console.log(to.join(', ') + ' 로 시험 메일을 보냈습니다.');
}

// ─────────────────────────────────────────────
// 깃허브 화면과 주고받는 통로
// ─────────────────────────────────────────────

function doGet() {
  return ContentService.createTextOutput('주님의교회 증명서 신청 서버입니다. 신청은 교회에서 안내한 신청 화면에서 해 주세요.');
}

var API_FUNCTIONS = {
  getPublicSettings: function () { return { settings: publicSettings_(getSettings_()) }; },
  submitApplication: submitApplication,
  lookupStatus: lookupStatus,
  adminLogin: adminLogin,
  adminLogout: adminLogout,
  adminList: adminList,
  adminDetail: adminDetail,
  adminUpdateStatus: adminUpdateStatus,
  adminSaveIssuance: adminSaveIssuance,
  adminRevealResident: adminRevealResident,
  adminAttachment: adminAttachment,
  adminLedger: adminLedger,
  adminLedgerCancel: adminLedgerCancel,
  adminLedgerUpload: adminLedgerUpload,
  adminLedgerFile: adminLedgerFile,
  adminDirectList: adminDirectList,
  adminDirectGet: adminDirectGet,
  adminDirectSave: adminDirectSave,
  adminDirectIssue: adminDirectIssue,
  adminDirectDelete: adminDirectDelete
};

function doPost(e) {
  var out;
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var fn = API_FUNCTIONS[req.fn];
    if (!fn) throw new Error('알 수 없는 요청입니다.');
    var args = Array.isArray(req.args) ? req.args : [];
    out = { ok: true, data: fn.apply(null, args) };
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────────
// 신청 접수
// ─────────────────────────────────────────────

function submitApplication(input, attachment) {
  if (!input || typeof input !== 'object' || Array.isArray(input)) throw new Error('입력 내용을 확인해 주세요.');
  var settings = getSettings_();

  var payload = {};
  ALLOWED_KEYS.forEach(function (key) {
    var v = input[key];
    if (typeof v !== 'string') return;
    var max = (key === 'notes' || /Purpose$/.test(key) || key === 'referenceInfo' || key === 'duties') ? 4000 : 500;
    v = clean_(v, max);
    if (v) payload[key] = v;
  });

  var type = payload.documentType;
  if (DOCUMENT_IDS.indexOf(type) < 0 || !settings.documents[type].enabled) throw new Error('신청할 문서를 다시 선택해 주세요.');
  if (!payload.applicantName || !payload.phone || payload.phone.replace(/\D/g, '').length < 9) {
    throw new Error('이름과 휴대전화번호를 확인해 주세요.');
  }
  ['birthDate', 'subjectBirthDate'].forEach(function (key) {
    if (payload[key] && !isValidBirthDate_(payload[key])) throw new Error('생년월일 8자리를 확인해 주세요.');
  });
  if (type !== 'donation' && !payload.birthDate) throw new Error('생년월일을 입력해 주세요.');

  // 신청 대상: 본인, 또는 미성년 자녀의 보호자
  if (['self', 'guardian'].indexOf(payload.applicantRelation) < 0) throw new Error('신청 대상을 선택해 주세요.');
  if (payload.applicantRelation === 'guardian') {
    if (!payload.subjectName || !payload.subjectBirthDate) throw new Error('자녀 이름과 생년월일을 입력해 주세요.');
    if (!isMinor_(payload.subjectBirthDate, new Date())) throw new Error(settings.common.subjectAdultError);
  } else {
    delete payload.subjectName;
    delete payload.subjectBirthDate;
  }

  // 발급 언어
  if (payload.certificateLanguage !== 'en') payload.certificateLanguage = 'ko';
  if (payload.certificateLanguage === 'en') {
    payload.englishName = String(payload.englishName || '').replace(/\s+/g, ' ').trim();
    if (!/^[A-Za-z][A-Za-z .,'\-]{0,99}$/.test(payload.englishName)) throw new Error(settings.language.englishNameError);
  } else {
    delete payload.englishName;
  }

  if (['email', 'visit', 'direct'].indexOf(payload.deliveryMethod) < 0) throw new Error('수령 방법을 선택해 주세요.');
  if (payload.deliveryMethod === 'email' && !isEmail_(payload.email)) {
    throw new Error('이메일 주소를 확인해 주세요.');
  }
  if (payload.deliveryMethod === 'direct') {
    if (!payload.recipientOrganization) throw new Error('수신 기관명을 입력해 주세요.');
    if (!isEmail_(payload.recipientEmail)) throw new Error(settings.validation.recipientEmailError);
  }
  if (payload.deliveryMethod !== 'email') delete payload.email;
  if (payload.deliveryMethod !== 'direct') { delete payload.recipientOrganization; delete payload.recipientPerson; delete payload.recipientEmail; }
  if (input.privacyConsent !== true) throw new Error('개인정보 수집·이용에 동의해 주세요.');

  var resident = '';
  if (type === 'donation') {
    resident = payload.residentNumber || '';
    if (!/^\d{6}-\d{7}$/.test(resident)) throw new Error('주민등록번호 13자리를 확인해 주세요.');
    if (input.sensitiveConsent !== true) throw new Error('주민등록번호 수집·이용에 동의해 주세요.');
  }
  delete payload.residentNumber; // 시트에는 남기지 않습니다

  var file = null;
  if (attachment && attachment.data) {
    file = decodeAttachment_(attachment);
    if (!file) throw new Error('첨부파일을 확인해 주세요. (PDF, Word, 한글, 이미지 · 10MB 이하)');
    payload.templateFileName = file.name;
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  var savedFile = null, reference;
  try {
    var sheet = getRequestSheet_();
    var now = new Date();
    reference = makeReference_(sheet, now);
    if (file) savedFile = getFolder_().createFile(file.blob.setName(reference + '_' + file.name));
    if (resident) PropertiesService.getScriptProperties().setProperty('RRN_' + reference, resident);

    var row = [];
    row[COL.reference - 1] = reference;
    row[COL.createdAt - 1] = now;
    row[COL.status - 1] = STATUSES[0];
    row[COL.documentName - 1] = settings.documents[type].name;
    row[COL.name - 1] = safeCell_(payload.applicantName);
    row[COL.phone - 1] = safeCell_(payload.phone);
    row[COL.birthDate - 1] = payload.birthDate || '';
    row[COL.department - 1] = safeCell_(payload.affiliatedDepartment || '');
    row[COL.relation - 1] = relationLabel_(payload.applicantRelation, settings);
    row[COL.delivery - 1] = deliveryLabel_(payload.deliveryMethod, settings);
    row[COL.requestedDate - 1] = payload.requestedDate || '';
    row[COL.copies - 1] = safeCell_(payload.copies || '');
    row[COL.summary - 1] = safeCell_(summaryText_(payload, settings));
    row[COL.attachment - 1] = savedFile ? '=HYPERLINK("' + savedFile.getUrl() + '","' + String(file.name).replace(/"/g, '') + '")' : '';
    row[COL.rrn - 1] = resident ? maskResident_(resident) : '';
    row[COL.updatedAt - 1] = now;
    row[COL.issuedAt - 1] = '';
    row[COL.data - 1] = JSON.stringify(payload);
    row[COL.attachmentId - 1] = savedFile ? savedFile.getId() : '';
    sheet.appendRow(row);
    var newRow = sheet.getLastRow();
    sheet.getRange(newRow, COL.status).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).setAllowInvalid(false).build());
    sheet.getRange(newRow, COL.phone, 1, 2).setNumberFormat('@').setValues([[safeCell_(payload.phone), payload.birthDate || '']]);
    SpreadsheetApp.flush();
  } catch (err) {
    try { if (savedFile) savedFile.setTrashed(true); } catch (e1) {}
    try { if (reference) PropertiesService.getScriptProperties().deleteProperty('RRN_' + reference); } catch (e2) {}
    console.error(err);
    throw new Error('신청을 접수하지 못했습니다. 잠시 후 다시 시도해 주세요.');
  } finally {
    lock.releaseLock();
  }

  try { notifyNewApplication_(reference, settings.documents[type].name, payload, settings); } catch (mailErr) { console.error(mailErr); }
  return { reference: reference };
}

function notifyNewApplication_(reference, documentName, payload, settings) {
  var to = getNotifyEmails_();
  if (!to.length) return;
  var adminUrl = (SITE_URL || '') + '?admin=1';
  var esc = function (s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  var row = function (k, v) { return '<tr><td style="padding:4px 14px 4px 0;color:#64727d">' + k + '</td><td>' + v + '</td></tr>'; };
  var html =
    '<div style="font-family:sans-serif;font-size:15px;line-height:1.7;color:#182536">' +
    '<p>새 증명서 신청이 접수되었습니다.</p>' +
    '<table style="border-collapse:collapse">' +
    row('접수번호', '<b>' + esc(reference) + '</b>') +
    row('문서', esc(documentName)) +
    row('수령 방법', esc(deliveryLabel_(payload.deliveryMethod, settings))) +
    row('발급 언어', esc(languageLabel_(payload.certificateLanguage, settings))) +
    row('희망 발급일', esc(payload.requestedDate || settings.delivery.requestedDateFallback)) +
    row('접수 시각', Utilities.formatDate(new Date(), TZ, 'M월 d일 HH:mm')) +
    '</table>' +
    '<p style="margin-top:18px">신청자 이름·연락처 등 자세한 내용은 관리 화면에서 확인해 주세요.</p>' +
    (SITE_URL ? '<p><a href="' + adminUrl + '" style="display:inline-block;background:#1a3150;color:#fff;padding:10px 18px;border-radius:10px;text-decoration:none;font-weight:bold">관리 화면 열기</a>&nbsp;&nbsp;' : '<p>') +
    '<a href="' + ss_().getUrl() + '">접수 목록 시트</a></p></div>';
  MailApp.sendEmail({
    to: to.join(','),
    subject: '[증명서 신청] 새 접수 ' + reference + ' · ' + documentName,
    htmlBody: html,
    name: '주님의교회 증명서 신청'
  });
}

// ─────────────────────────────────────────────
// 신청 내역 조회 (성도용)
// ─────────────────────────────────────────────

function lookupStatus(input) {
  input = input || {};
  var name = clean_(input.applicantName, 100);
  var phone = String(input.phone || '').replace(/\D/g, '');
  var code = String(input.receiptCode || '').replace(/[^a-zA-Z0-9]/g, '').toUpperCase().slice(0, 14);
  if (!name || phone.length < 9 || (code && code.length < 6)) throw new Error('신청자 이름과 휴대전화번호를 확인해 주세요.');

  var cache = CacheService.getScriptCache();
  var failKey = 'lookup_fail_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, name, Utilities.Charset.UTF_8)).slice(0, 40);
  var fails = Number(cache.get(failKey) || 0);
  if (fails >= 8) throw new Error('여러 번 일치하지 않았습니다. 10분 뒤에 다시 시도하시거나 교회 사무실로 문의해 주세요.');

  var settings = getSettings_();
  var sheet = getRequestSheet_();
  var last = sheet.getLastRow();
  var results = [];
  if (last > 1) {
    sheet.getRange(2, 1, last - 1, REQUEST_HEADERS.length).getValues().forEach(function (r) {
      if (!r[COL.reference - 1]) return;
      if (String(r[COL.name - 1]).replace(/^'/, '').trim() !== name) return;
      if (String(r[COL.phone - 1]).replace(/\D/g, '') !== phone) return;
      var ref = String(r[COL.reference - 1]);
      if (code && ref.replace(/[^a-zA-Z0-9]/g, '').toUpperCase() !== code) return;
      var data = parseData_(r[COL.data - 1]);
      results.push({
        receiptCode: ref,
        documentName: String(r[COL.documentName - 1]),
        statusLabel: STATUSES.indexOf(String(r[COL.status - 1])) >= 0 ? String(r[COL.status - 1]) : STATUSES[0],
        createdAt: dateLabel_(r[COL.createdAt - 1]),
        updatedAt: dateLabel_(r[COL.updatedAt - 1] || r[COL.createdAt - 1]),
        requestedDate: data.requestedDate ? formatKoreanDate_(data.requestedDate) : '별도 지정 없음',
        delivery: deliveryLabel_(data.deliveryMethod, settings) || '담당자 확인 예정'
      });
    });
  }
  if (!results.length) {
    cache.put(failKey, String(fails + 1), 600);
    throw new Error('일치하는 신청 내역을 찾을 수 없습니다. 신청할 때 적은 이름과 휴대전화번호를 확인해 주세요.');
  }
  results.reverse();
  return { applications: results.slice(0, 25) };
}

// ─────────────────────────────────────────────
// 관리자 기능
// ─────────────────────────────────────────────

function adminLogin(password) {
  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get('login_fails') || 0);
  if (fails >= 10) throw new Error('비밀번호를 여러 번 잘못 입력했습니다. 10분 뒤에 다시 시도해 주세요.');
  var saved = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!saved) throw new Error('관리자 비밀번호가 아직 정해지지 않았습니다. Apps Script 프로젝트 설정에서 먼저 정해 주세요.');
  if (String(password || '') !== saved) {
    cache.put('login_fails', String(fails + 1), 600);
    throw new Error('비밀번호가 올바르지 않습니다.');
  }
  var token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
  cache.put('adm_' + token, '1', ADMIN_SESSION_SECONDS);
  return { token: token };
}

function adminLogout(token) {
  if (token) CacheService.getScriptCache().remove('adm_' + token);
  return { ok: true };
}

function adminList(token) {
  requireAdmin_(token);
  var sheet = getRequestSheet_();
  var last = sheet.getLastRow();
  var list = [];
  if (last > 1) {
    sheet.getRange(2, 1, last - 1, REQUEST_HEADERS.length).getValues().forEach(function (r) {
      if (!r[COL.reference - 1]) return;
      var data = parseData_(r[COL.data - 1]);
      var created = toDate_(r[COL.createdAt - 1]);
      list.push({
        reference: String(r[COL.reference - 1]),
        createdAt: created ? created.toISOString() : '',
        dateLabel: created ? Utilities.formatDate(created, TZ, 'yyyy년 M월 d일') : '',
        documentType: data.documentType || '',
        documentName: String(r[COL.documentName - 1]),
        applicantName: String(r[COL.name - 1]).replace(/^'/, ''),
        status: STATUSES.indexOf(String(r[COL.status - 1])) >= 0 ? String(r[COL.status - 1]) : STATUSES[0],
        requestedDate: data.requestedDate || ''
      });
    });
  }
  list.sort(function (a, b) { return a.createdAt < b.createdAt ? 1 : -1; });
  return { applications: list, statuses: STATUSES, sheetUrl: ss_().getUrl(), settings: getSettings_() };
}

function adminDetail(token, reference) {
  requireAdmin_(token);
  var sheet = getRequestSheet_();
  var row = findRow_(sheet, reference);
  if (!row) throw new Error('신청 내역을 찾을 수 없습니다.');
  var r = sheet.getRange(row, 1, 1, REQUEST_HEADERS.length).getValues()[0];
  var data = parseData_(r[COL.data - 1]);
  var resident = data.documentType === 'donation' ? PropertiesService.getScriptProperties().getProperty('RRN_' + reference) : null;
  var created = toDate_(r[COL.createdAt - 1]);
  return {
    reference: String(r[COL.reference - 1]),
    createdAt: created ? created.toISOString() : '',
    dateLabel: created ? Utilities.formatDate(created, TZ, 'yyyy년 M월 d일 HH:mm') : '',
    status: STATUSES.indexOf(String(r[COL.status - 1])) >= 0 ? String(r[COL.status - 1]) : STATUSES[0],
    documentName: String(r[COL.documentName - 1]),
    payload: data,
    labels: FIELD_LABELS,
    residentMasked: data.documentType === 'donation' ? (resident ? maskResident_(resident) : '삭제됨') : null,
    hasAttachment: !!r[COL.attachmentId - 1],
    documentNumber: data.issuedNumber || '',
    issuedAt: data.issuedAt || ''
  };
}

function adminUpdateStatus(token, reference, status) {
  requireAdmin_(token);
  if (STATUSES.indexOf(status) < 0) throw new Error('처리 상태를 확인해 주세요.');
  var settings = getSettings_(); // 잠금 전에 읽어 둡니다
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  var deleted = false, issued = null;
  try {
    var sheet = getRequestSheet_();
    var row = findRow_(sheet, reference);
    if (!row) throw new Error('신청 내역을 찾을 수 없습니다.');
    var now = new Date();
    if (status === STATUS_ISSUED) issued = issueApplication_(sheet, row, now, settings);
    sheet.getRange(row, COL.status).setValue(status);
    sheet.getRange(row, COL.updatedAt).setValue(now);
    sheet.getRange(row, COL.issuedAt).setValue(status === STATUS_ISSUED ? (toDate_(issued.issuedAt) || now) : '');
    if (status === STATUS_ISSUED) {
      var props = PropertiesService.getScriptProperties();
      if (props.getProperty('RRN_' + reference)) {
        props.deleteProperty('RRN_' + reference);
        sheet.getRange(row, COL.rrn).setValue('삭제됨');
        deleted = true;
      }
    }
  } finally {
    lock.releaseLock();
  }
  return { ok: true, residentDeleted: deleted, documentNumber: issued ? issued.number : '', issuedAt: issued ? issued.issuedAt : '' };
}

// 신청 건을 "발급 완료"로 확정합니다: 문서번호를 붙이고 발급대장에 한 줄 기록합니다.
// 이미 번호가 있으면 그대로 쓰고, 그 번호가 대장에서 "취소"된 경우에만 새 번호를 붙입니다. (잠금 안에서 부르세요)
function issueApplication_(sheet, row, now, settings) {
  var r = sheet.getRange(row, 1, 1, REQUEST_HEADERS.length).getValues()[0];
  var reference = String(r[COL.reference - 1]);
  var cell = sheet.getRange(row, COL.data);
  var p = parseData_(cell.getValue());
  var ledger = getLedgerSheet_();
  if (p.issuedNumber) {
    var lrow = findLedgerRow_(ledger, p.issuedNumber);
    if (lrow && String(ledger.getRange(lrow, LCOL.status).getValue()) !== LEDGER_CANCELLED) {
      return { number: p.issuedNumber, issuedAt: p.issuedAt || '' };
    }
  }
  var number = nextDocNumber_(ledger, ISSUERS.tlc.prefix, now);
  p.issuedNumber = number;
  p.issuedAt = now.toISOString();
  cell.setValue(JSON.stringify(p));

  var type = p.documentType;
  var isRec = RECOMMENDATION_TYPES.indexOf(type) >= 0;
  var category = type === 'donation' ? '기타' : isRec ? '추천서' : '증명서';
  var subject = p.subjectName || p.applicantName || '';
  if (p.certificateLanguage === 'en' && p.englishName) subject += ' (' + p.englishName + ')';
  var rep = isRec
    ? (type === 'seniorPastor' ? '담임목사 ' + settings.certificate.seniorPastor
      : type === 'minister' ? '담당교역자 ' + (p.ministerName || '') : '추천인 ' + (p.recommendingMember || ''))
    : '담임목사 ' + settings.certificate.seniorPastor;
  var row2 = [];
  row2[LCOL.number - 1] = number;
  row2[LCOL.date - 1] = now;
  row2[LCOL.issuer - 1] = settings.issuers.tlc.name;
  row2[LCOL.category - 1] = category;
  row2[LCOL.title - 1] = String(r[COL.documentName - 1]) + (p.certificateLanguage === 'en' ? ' (영문)' : '');
  row2[LCOL.recipient - 1] = safeCell_(subject);
  row2[LCOL.rep - 1] = safeCell_(rep.trim());
  row2[LCOL.status - 1] = LEDGER_ISSUED;
  row2[LCOL.file - 1] = '';
  row2[LCOL.source - 1] = reference;
  row2[LCOL.note - 1] = safeCell_((p.recipientOrganization || p.submissionTo) ? '제출처: ' + (p.recipientOrganization || p.submissionTo) : '');
  row2[LCOL.lang - 1] = languageLabel_(p.certificateLanguage, settings);
  row2[LCOL.birth - 1] = p.subjectBirthDate || p.birthDate || '';
  row2[LCOL.applicant - 1] = safeCell_(p.applicantName || '');
  row2[LCOL.relation - 1] = relationLabel_(p.applicantRelation, settings);
  row2[LCOL.delivery - 1] = deliveryLabel_(p.deliveryMethod, settings);
  row2[LCOL.copies - 1] = safeCell_(p.copies || '');
  row2[LCOL.validUntil - 1] = isRec ? '' : addMonths_(now, VALID_MONTHS);
  row2[LCOL.fileId - 1] = '';
  appendLedgerRow_(ledger, row2);
  return { number: number, issuedAt: p.issuedAt };
}

function appendLedgerRow_(ledger, row) {
  for (var i = 0; i < LEDGER_HEADERS.length; i++) if (row[i] === undefined) row[i] = '';
  ledger.appendRow(row);
  var n = ledger.getLastRow();
  ledger.getRange(n, LCOL.number).setNumberFormat('@').setValue(row[LCOL.number - 1]);
  ledger.getRange(n, LCOL.birth).setNumberFormat('@').setValue(row[LCOL.birth - 1]);
  ledger.getRange(n, LCOL.source).setNumberFormat('@').setValue(row[LCOL.source - 1]);
}

// 발급 주체별·연도별 일련번호: TLC-2026-001, LM-2026-001 …
function nextDocNumber_(ledger, prefix, now) {
  var year = Utilities.formatDate(now, TZ, 'yyyy');
  var re = new RegExp('^' + prefix + '-' + year + '-(\\d+)$');
  var max = 0, last = ledger.getLastRow();
  if (last > 1) {
    ledger.getRange(2, LCOL.number, last - 1, 1).getValues().forEach(function (r) {
      var m = re.exec(String(r[0]).trim());
      if (m) max = Math.max(max, Number(m[1]));
    });
  }
  var n = String(max + 1);
  while (n.length < 3) n = '0' + n;
  return prefix + '-' + year + '-' + n;
}

function findLedgerRow_(ledger, number) {
  var last = ledger.getLastRow();
  if (last < 2 || !number) return 0;
  var vals = ledger.getRange(2, LCOL.number, last - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) if (String(vals[i][0]).trim() === String(number)) return i + 2;
  return 0;
}

function addMonths_(date, months) {
  var d = new Date(date.getTime());
  var day = d.getDate();
  d.setMonth(d.getMonth() + months);
  if (d.getDate() !== day) d.setDate(0); // 예: 1월 31일 + 1개월 → 2월 말일
  d.setDate(d.getDate() - 1);           // 발급일 포함 3개월
  return d;
}

function languageLabel_(v, s) { return v === 'en' ? s.language.en : s.language.ko; }

function adminSaveIssuance(token, reference, values) {
  requireAdmin_(token);
  values = values || {};
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = getRequestSheet_();
    var row = findRow_(sheet, reference);
    if (!row) throw new Error('신청 내역을 찾을 수 없습니다.');
    var cell = sheet.getRange(row, COL.data);
    var data = parseData_(cell.getValue());
    ISSUANCE_KEYS.forEach(function (key) {
      if (values[key] === undefined) return;
      var v = clean_(values[key], key === 'recommendationBody' ? 6000 : 500);
      if (key === 'donationAmount') v = v.replace(/[^0-9]/g, '');
      data[key] = v;
    });
    // 영문 발급 내용 (번역한 항목들)
    if (values.en && typeof values.en === 'object' && !Array.isArray(values.en)) {
      var en = {};
      Object.keys(values.en).slice(0, MAX_EN_KEYS).forEach(function (k) {
        if (!/^[a-zA-Z]{1,40}$/.test(k)) return;
        var v = clean_(values.en[k], k === 'body' ? 6000 : 1000);
        if (v) en[k] = v;
      });
      data.en = en;
    }
    cell.setValue(JSON.stringify(data));
    sheet.getRange(row, COL.updatedAt).setValue(new Date());
  } finally {
    lock.releaseLock();
  }
  return { saved: true };
}

function adminRevealResident(token, reference) {
  requireAdmin_(token);
  var value = PropertiesService.getScriptProperties().getProperty('RRN_' + reference);
  return value ? { residentNumber: value } : { deleted: true };
}

function adminAttachment(token, reference) {
  requireAdmin_(token);
  var sheet = getRequestSheet_();
  var row = findRow_(sheet, reference);
  if (!row) throw new Error('신청 내역을 찾을 수 없습니다.');
  var id = sheet.getRange(row, COL.attachmentId).getValue();
  if (!id) throw new Error('첨부파일이 없습니다.');
  var file = DriveApp.getFileById(String(id));
  var blob = file.getBlob();
  var data = parseData_(sheet.getRange(row, COL.data).getValue());
  return { name: data.templateFileName || file.getName(), dataUrl: 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes()) };
}

// ─────────────────────────────────────────────
// 발급대장
// ─────────────────────────────────────────────

function issuerKeyOf_(number) {
  var pre = String(number || '').split('-')[0];
  for (var k in ISSUERS) if (ISSUERS[k].prefix === pre) return k;
  return '';
}

function adminLedger(token) {
  requireAdmin_(token);
  var settings = getSettings_();
  var ss = ss_();
  var ledger = ss.getSheetByName(SHEET_LEDGER);
  var rows = [];
  if (ledger && ledger.getLastRow() > 1) {
    ledger.getRange(2, 1, ledger.getLastRow() - 1, LEDGER_HEADERS.length).getValues().forEach(function (r) {
      var number = String(r[LCOL.number - 1]).trim();
      if (!number) return;
      var date = toDate_(r[LCOL.date - 1]);
      var until = toDate_(r[LCOL.validUntil - 1]);
      var source = String(r[LCOL.source - 1]).trim();
      rows.push({
        number: number,
        issuerKey: issuerKeyOf_(number),
        date: date ? Utilities.formatDate(date, TZ, 'yyyy-MM-dd') : '',
        issuer: String(r[LCOL.issuer - 1]),
        category: String(r[LCOL.category - 1]),
        title: String(r[LCOL.title - 1]),
        recipient: String(r[LCOL.recipient - 1]).replace(/^'/, ''),
        rep: String(r[LCOL.rep - 1]).replace(/^'/, ''),
        status: String(r[LCOL.status - 1]) || LEDGER_ISSUED,
        hasFile: !!String(r[LCOL.fileId - 1]),
        source: source,
        sourceType: /^D/.test(source) ? 'direct' : source ? 'application' : '',
        note: String(r[LCOL.note - 1]).replace(/^'/, ''),
        lang: String(r[LCOL.lang - 1]),
        validUntil: until ? Utilities.formatDate(until, TZ, 'yyyy-MM-dd') : ''
      });
    });
  }
  rows.reverse();
  return { rows: rows, categories: DOC_CATEGORIES, issuers: settings.issuers, sheetUrl: ss.getUrl() };
}

// 발급한 문서는 지우지 않고 "취소"로 남깁니다
function adminLedgerCancel(token, number, reason) {
  requireAdmin_(token);
  reason = clean_(reason, 300);
  if (!reason) throw new Error('취소 사유를 적어 주세요.');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ledger = getLedgerSheet_();
    var row = findLedgerRow_(ledger, number);
    if (!row) throw new Error('발급대장에서 문서를 찾을 수 없습니다.');
    if (String(ledger.getRange(row, LCOL.status).getValue()) === LEDGER_CANCELLED) throw new Error('이미 취소된 문서입니다.');
    var note = String(ledger.getRange(row, LCOL.note).getValue() || '');
    var stamp = '[취소 ' + Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd') + '] ' + reason;
    ledger.getRange(row, LCOL.status).setValue(LEDGER_CANCELLED);
    ledger.getRange(row, LCOL.note).setValue(safeCell_(note ? note + '\n' + stamp : stamp));
    var source = String(ledger.getRange(row, LCOL.source).getValue()).trim();
    if (/^D/.test(source)) {
      var ds = getDirectSheet_(), drow = findRow_(ds, source);
      if (drow) ds.getRange(drow, DCOL.status).setValue(LEDGER_CANCELLED);
    }
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

// 발급본 PDF를 드라이브 문서/연도/발급 주체/구분 폴더에 저장하고 대장에 링크를 남깁니다
function adminLedgerUpload(token, number, file) {
  requireAdmin_(token);
  var m = /^data:application\/pdf;base64,([A-Za-z0-9+/=]+)$/.exec(String((file && file.data) || ''));
  if (!m || !/\.pdf$/i.test(String(file.name || ''))) throw new Error('PDF 파일만 올릴 수 있습니다.');
  var bytes = Utilities.base64Decode(m[1]);
  if (!bytes.length || bytes.length > MAX_PDF_BYTES) throw new Error('15MB 이하의 PDF 파일을 올려 주세요.');
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ledger = getLedgerSheet_();
    var row = findLedgerRow_(ledger, number);
    if (!row) throw new Error('발급대장에서 문서를 찾을 수 없습니다.');
    var r = ledger.getRange(row, 1, 1, LEDGER_HEADERS.length).getValues()[0];
    var date = toDate_(r[LCOL.date - 1]) || new Date();
    var key = issuerKeyOf_(number) || 'tlc';
    var category = DOC_CATEGORIES.indexOf(String(r[LCOL.category - 1])) >= 0 ? String(r[LCOL.category - 1]) : '기타';
    var folder = subFolder_(subFolder_(subFolder_(getDocsRoot_(), Utilities.formatDate(date, TZ, 'yyyy')), ISSUERS[key].folder), category);
    var title = String(r[LCOL.title - 1]).replace(/[\\\/:*?"<>|\r\n]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 80);
    var saved = folder.createFile(Utilities.newBlob(bytes, 'application/pdf', number + '_' + title + '.pdf'));
    var oldId = String(r[LCOL.fileId - 1] || '');
    if (oldId) { try { DriveApp.getFileById(oldId).setTrashed(true); } catch (e) {} }
    ledger.getRange(row, LCOL.file).setValue('=HYPERLINK("' + saved.getUrl() + '","PDF")');
    ledger.getRange(row, LCOL.fileId).setValue(saved.getId());
    return { ok: true, replaced: !!oldId };
  } finally {
    lock.releaseLock();
  }
}

function adminLedgerFile(token, number) {
  requireAdmin_(token);
  var ledger = getLedgerSheet_();
  var row = findLedgerRow_(ledger, number);
  if (!row) throw new Error('발급대장에서 문서를 찾을 수 없습니다.');
  var id = String(ledger.getRange(row, LCOL.fileId).getValue() || '');
  if (!id) throw new Error('아직 올린 발급본이 없습니다.');
  var f = DriveApp.getFileById(id);
  return { name: f.getName(), url: f.getUrl(), dataUrl: 'data:application/pdf;base64,' + Utilities.base64Encode(f.getBlob().getBytes()) };
}

function getDocsRoot_() {
  var props = PropertiesService.getScriptProperties();
  var ids = [props.getProperty('DOCS_ROOT_ID'), DOCS_ROOT_FOLDER_ID];
  for (var i = 0; i < ids.length; i++) {
    if (!ids[i]) continue;
    try { var f = DriveApp.getFolderById(ids[i]); if (!f.isTrashed()) return f; } catch (e) {}
  }
  var created = DriveApp.createFolder('문서');
  props.setProperty('DOCS_ROOT_ID', created.getId());
  return created;
}

function subFolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  while (it.hasNext()) { var f = it.next(); if (!f.isTrashed()) return f; }
  return parent.createFolder(name);
}

// ─────────────────────────────────────────────
// 직접 발급 (관리자 전용 공문·승인서)
// ─────────────────────────────────────────────

function adminDirectList(token) {
  requireAdmin_(token);
  var sheet = ss_().getSheetByName(SHEET_DIRECT);
  var list = [];
  if (sheet && sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, DIRECT_HEADERS.length).getValues().forEach(function (r) {
      if (!r[DCOL.id - 1]) return;
      var d = parseData_(r[DCOL.data - 1]);
      var up = toDate_(r[DCOL.updatedAt - 1]);
      list.push({
        id: String(r[DCOL.id - 1]), status: String(r[DCOL.status - 1]), number: String(r[DCOL.number - 1]),
        issuerKey: d.issuer || '', category: String(r[DCOL.category - 1]), title: String(r[DCOL.title - 1]).replace(/^'/, ''),
        recipient: String(r[DCOL.recipient - 1]).replace(/^'/, ''),
        updatedAt: up ? Utilities.formatDate(up, TZ, 'yyyy-MM-dd HH:mm') : ''
      });
    });
  }
  list.reverse();
  return { documents: list, issuers: getSettings_().issuers, categories: DOC_CATEGORIES };
}

function adminDirectGet(token, id) {
  requireAdmin_(token);
  var sheet = getDirectSheet_();
  var row = findRow_(sheet, id);
  if (!row) throw new Error('문서를 찾을 수 없습니다.');
  var r = sheet.getRange(row, 1, 1, DIRECT_HEADERS.length).getValues()[0];
  var data = parseData_(r[DCOL.data - 1]);
  return { id: String(r[DCOL.id - 1]), status: String(r[DCOL.status - 1]), number: String(r[DCOL.number - 1]), doc: data };
}

function cleanDirectDoc_(input) {
  input = input && typeof input === 'object' && !Array.isArray(input) ? input : {};
  var doc = {};
  Object.keys(DIRECT_FIELDS).forEach(function (k) { doc[k] = clean_(input[k], DIRECT_FIELDS[k]); });
  if (!ISSUERS[doc.issuer]) throw new Error('발급 주체를 선택해 주세요.');
  if (DOC_CATEGORIES.indexOf(doc.category) < 0) throw new Error('문서 구분을 선택해 주세요.');
  return doc;
}

function adminDirectSave(token, id, input) {
  requireAdmin_(token);
  getSettingsCache_(getSettings_());
  var doc = cleanDirectDoc_(input);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = getDirectSheet_();
    var row = id ? findRow_(sheet, id) : 0;
    if (id && !row) throw new Error('문서를 찾을 수 없습니다.');
    if (row && String(sheet.getRange(row, DCOL.status).getValue()) !== DIRECT_DRAFT) throw new Error('발급된 문서는 고칠 수 없습니다.');
    var prev = row ? parseData_(sheet.getRange(row, DCOL.data).getValue()) : {};
    Object.keys(doc).forEach(function (k) { prev[k] = doc[k]; });
    var now = new Date();
    if (!row) {
      id = makeDirectId_(sheet, now);
      sheet.appendRow([id, DIRECT_DRAFT, '', '', '', '', '', now, '']);
      row = sheet.getLastRow();
      sheet.getRange(row, DCOL.id).setNumberFormat('@').setValue(id);
    }
    writeDirectRow_(sheet, row, prev, now);
  } finally {
    lock.releaseLock();
  }
  return { id: id };
}

function writeDirectRow_(sheet, row, doc, now) {
  var issuers = getSettingsCache_().issuers;
  sheet.getRange(row, DCOL.issuer, 1, 4).setValues([[
    (issuers[doc.issuer] || {}).name || '', doc.category || '', safeCell_(doc.docTitle || doc.subject || ''), safeCell_(doc.recipient || '')
  ]]);
  sheet.getRange(row, DCOL.updatedAt).setValue(now);
  sheet.getRange(row, DCOL.data).setValue(JSON.stringify(doc));
}

// 발급 확정: 문서번호를 붙이고 발급대장에 기록합니다. 확정 뒤에는 고칠 수 없습니다.
function adminDirectIssue(token, id, input) {
  requireAdmin_(token);
  var settings = getSettings_();
  getSettingsCache_(settings);
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getDirectSheet_();
    var row = findRow_(sheet, id);
    if (!row) throw new Error('문서를 찾을 수 없습니다.');
    if (String(sheet.getRange(row, DCOL.status).getValue()) !== DIRECT_DRAFT) throw new Error('이미 발급된 문서입니다.');
    var doc = parseData_(sheet.getRange(row, DCOL.data).getValue());
    if (input) { var fresh = cleanDirectDoc_(input); Object.keys(fresh).forEach(function (k) { doc[k] = fresh[k]; }); }
    ['recipient', 'subject', 'body', 'docTitle'].forEach(function (k) {
      if (!doc[k]) throw new Error({ recipient: '수신', subject: '제목', body: '본문', docTitle: '문서명' }[k] + '을(를) 입력해 주세요.');
    });
    var now = new Date();
    var ledger = getLedgerSheet_();
    var number = nextDocNumber_(ledger, ISSUERS[doc.issuer].prefix, now);
    var issuer = settings.issuers[doc.issuer];
    doc.number = number;
    doc.issuedAt = now.toISOString();
    if (!doc.representativeTitle) doc.representativeTitle = issuer.representativeTitle;
    if (!doc.representative) doc.representative = issuer.representative;
    writeDirectRow_(sheet, row, doc, now);
    sheet.getRange(row, DCOL.status).setValue(LEDGER_ISSUED);
    sheet.getRange(row, DCOL.number).setNumberFormat('@').setValue(number);

    var lr = [];
    lr[LCOL.number - 1] = number;
    lr[LCOL.date - 1] = now;
    lr[LCOL.issuer - 1] = issuer.name;
    lr[LCOL.category - 1] = doc.category;
    lr[LCOL.title - 1] = safeCell_(doc.docTitle);
    lr[LCOL.recipient - 1] = safeCell_(doc.recipient);
    lr[LCOL.rep - 1] = safeCell_((doc.representativeTitle + ' ' + doc.representative).trim());
    lr[LCOL.status - 1] = LEDGER_ISSUED;
    lr[LCOL.source - 1] = id;
    lr[LCOL.note - 1] = safeCell_(doc.note || '');
    lr[LCOL.lang - 1] = /[가-힣]/.test(doc.body) ? '한국어' : '영어';
    appendLedgerRow_(ledger, lr);
    return { number: number, issuedAt: doc.issuedAt };
  } finally {
    lock.releaseLock();
  }
}

function adminDirectDelete(token, id) {
  requireAdmin_(token);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = getDirectSheet_();
    var row = findRow_(sheet, id);
    if (!row) throw new Error('문서를 찾을 수 없습니다.');
    if (String(sheet.getRange(row, DCOL.status).getValue()) !== DIRECT_DRAFT) throw new Error('발급된 문서는 지울 수 없습니다. 발급대장에서 취소해 주세요.');
    sheet.deleteRow(row);
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

// 문서 ID 예: D261016-K7Q2
function makeDirectId_(sheet, now) {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  var datePart = Utilities.formatDate(now, TZ, 'yyMMdd');
  for (var i = 0; i < 20; i++) {
    var tail = '';
    for (var j = 0; j < 4; j++) tail += chars.charAt(Math.floor(Math.random() * chars.length));
    var id = 'D' + datePart + '-' + tail;
    if (!findRow_(sheet, id)) return id;
  }
  throw new Error('문서 ID를 만들지 못했습니다.');
}

// 잠금 안에서 화면 문구를 다시 읽지 않도록 한 번 읽은 값을 씁니다
var SETTINGS_CACHE_ = null;
function getSettingsCache_(s) {
  if (s) SETTINGS_CACHE_ = s;
  return SETTINGS_CACHE_ || defaultSettings_();
}

// ─────────────────────────────────────────────
// 화면 문구 읽기
// ─────────────────────────────────────────────

function getSettings_() {
  ensureCopyVersion_();
  var settings = defaultSettings_();
  var sheet = ss_().getSheetByName(SHEET_COPY);
  if (!sheet || sheet.getLastRow() < 2) return settings;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues().forEach(function (r) {
    var key = String(r[0]).trim();
    if (!key) return;
    var parts = key.split('.');
    var parent = settings;
    for (var i = 0; i < parts.length - 1; i++) {
      if (!parent || typeof parent[parts[i]] !== 'object') return;
      parent = parent[parts[i]];
    }
    var last = parts[parts.length - 1];
    if (!parent || !parent.hasOwnProperty(last) || typeof parent[last] === 'object') return;
    var raw = r[2];
    if (typeof parent[last] === 'boolean') {
      var t = String(raw).trim().toLowerCase();
      if (['예', '네', 'o', 'yes', 'true', '보이기', '표시'].indexOf(t) >= 0 || raw === true) parent[last] = true;
      else if (['아니요', '아니오', 'x', 'no', 'false', '숨기기'].indexOf(t) >= 0 || raw === false) parent[last] = false;
      return;
    }
    var val = String(raw == null ? '' : raw).replace(/\r\n?/g, '\n').trim().slice(0, 2000);
    // 칸을 비우면 기본 문구를 씁니다 (안내 문구·흐린 안내·설명·고유번호는 비워 둘 수 있습니다)
    if (!val && !/^(notice|placeholder|description|organizationNumber|donationNotice|otherNotice|address|phone|email|subName)$/.test(last)) return;
    if (/Color$/.test(last) && !/^#[0-9a-fA-F]{6}$/.test(val)) return;
    if (last === 'logoSize' && !/^\d{1,2}(\.\d)?$/.test(val)) return;
    parent[last] = val;
  });
  return settings;
}

function publicSettings_(s) {
  var copy = JSON.parse(JSON.stringify(s));
  delete copy.outputs;
  delete copy.outputCommon;
  delete copy.outputsEn;
  delete copy.outputCommonEn;
  return copy;
}

function flattenSettings_(obj, prefix, out) {
  out = out || [];
  Object.keys(obj).forEach(function (key) {
    var path = prefix ? prefix + '.' + key : key;
    var v = obj[key];
    if (v && typeof v === 'object') flattenSettings_(v, path, out);
    else out.push([path, v]);
  });
  return out;
}

function describeKey_(path) {
  var d = defaultSettings_();
  var parts = path.split('.');
  var group, rest;
  if (parts[0] === 'documents') { group = '문서 · ' + d.documents[parts[1]].name; rest = parts.slice(2); }
  else if (parts[0] === 'issuers') {
    var isLabel = { name: '이름', address: '주소', phone: '전화번호' }[parts[2]];
    return '발급 주체 · ' + d.issuers[parts[1]].name + ' · ' + (isLabel || KEY_LABELS[parts[2]] || parts[2]);
  }
  else if (parts[0] === 'outputs') { group = '출력 양식 · ' + d.outputs[parts[1]].title; rest = parts.slice(2); }
  else if (parts[0] === 'outputsEn') { group = '영문 출력 양식 · ' + d.outputs[parts[1]].title; rest = parts.slice(2); }
  else { group = GROUP_LABELS[parts[0]] || parts[0]; rest = parts.slice(1); }

  var lastKey = rest[rest.length - 1];
  var detail;
  if ((lastKey === 'label' || lastKey === 'placeholder') && rest.length >= 2) {
    var node = d;
    for (var i = 0; i < parts.length - 1; i++) node = node[parts[i]];
    detail = '\'' + node.label + '\' ' + (lastKey === 'label' ? '칸 제목' : '칸 안 흐린 안내');
  } else if (rest[0] === 'labels' && rest.length === 2) {
    detail = '\'' + d[parts[0]][parts[1]].labels[rest[1]] + '\' 표기';
  } else if (rest[0] === 'sections' && rest.length === 2) {
    detail = '\'' + d.outputs[parts[1]].sections[rest[1]] + '\' 구역 제목';
  } else {
    detail = KEY_LABELS[lastKey] || lastKey;
  }
  return group + ' · ' + detail;
}

function toCell_(v) {
  if (v === true) return '예';
  if (v === false) return '아니요';
  return String(v);
}

// ─────────────────────────────────────────────
// 내부 도우미
// ─────────────────────────────────────────────

function ss_(createIfMissing) {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SPREADSHEET_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) {
      if (!createIfMissing) throw new Error('접수 시트를 열지 못했습니다. 관리자에게 알려 주세요.');
    }
  }
  if (!createIfMissing) throw new Error('아직 처음 설정(setupAll)이 실행되지 않았습니다.');
  var ss = SpreadsheetApp.create(SPREADSHEET_NAME);
  props.setProperty('SPREADSHEET_ID', ss.getId());
  return ss;
}

function getFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('FOLDER_ID');
  if (id) {
    try {
      var f = DriveApp.getFolderById(id);
      if (!f.isTrashed()) return f;
    } catch (e) {}
  }
  var folder = DriveApp.createFolder(FOLDER_NAME);
  props.setProperty('FOLDER_ID', folder.getId());
  return folder;
}

function requireAdmin_(token) {
  if (!token || !CacheService.getScriptCache().get('adm_' + token)) throw new Error('LOGIN_REQUIRED');
}

function getRequestSheet_() {
  var sheet = ss_().getSheetByName(SHEET_REQUESTS);
  if (!sheet) throw new Error('"접수 목록" 탭이 없습니다. Apps Script에서 setupAll을 먼저 실행해 주세요.');
  return sheet;
}

function getNotifyEmails_() {
  var sheet = ss_().getSheetByName(SHEET_MAIL);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues()
    .map(function (r) { return String(r[0]).trim(); })
    .filter(function (v) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v); });
}

function findRow_(sheet, reference) {
  var last = sheet.getLastRow();
  if (last < 2) return 0;
  var refs = sheet.getRange(2, COL.reference, last - 1, 1).getValues();
  for (var i = 0; i < refs.length; i++) {
    if (String(refs[i][0]) === String(reference)) return i + 2;
  }
  return 0;
}

// 접수번호 예: C261006-K7Q2 (헷갈리기 쉬운 0·O·1·I는 쓰지 않습니다)
function makeReference_(sheet, now) {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  var datePart = Utilities.formatDate(now, TZ, 'yyMMdd');
  for (var i = 0; i < 20; i++) {
    var tail = '';
    for (var j = 0; j < 4; j++) tail += chars.charAt(Math.floor(Math.random() * chars.length));
    var ref = 'C' + datePart + '-' + tail;
    if (!findRow_(sheet, ref)) return ref;
  }
  throw new Error('접수번호를 만들지 못했습니다.');
}

function decodeAttachment_(att) {
  var m = /^data:([\w.+\-\/]*);base64,([A-Za-z0-9+/=]+)$/.exec(String(att.data || ''));
  if (!m) return null;
  var name = String(att.name || '첨부파일').replace(/[\\\/:*?"<>|\r\n]/g, '_').slice(-120);
  var ext = (name.split('.').pop() || '').toLowerCase();
  if (ALLOWED_FILE_EXT.indexOf(ext) < 0) return null;
  var bytes = Utilities.base64Decode(m[2]);
  if (!bytes.length || bytes.length > MAX_FILE_BYTES) return null;
  return { blob: Utilities.newBlob(bytes, m[1] || 'application/octet-stream'), name: name };
}

function summaryText_(p, settings) {
  var skip = ['documentType', 'applicantName', 'phone', 'birthDate', 'affiliatedDepartment', 'applicantRelation',
    'deliveryMethod', 'requestedDate', 'copies', 'templateFileName'];
  var lines = [];
  Object.keys(p).forEach(function (key) {
    if (skip.indexOf(key) >= 0 || !FIELD_LABELS[key]) return;
    var v = key === 'certificateLanguage' ? languageLabel_(p[key], settings) : p[key];
    lines.push(FIELD_LABELS[key] + ': ' + v);
  });
  return lines.join('\n');
}

function relationLabel_(v, s) {
  // 가족·대리·기타는 예전 신청 기록을 위해 남겨 둡니다
  return { self: s.common.relationSelf, guardian: s.common.relationGuardian, family: '가족', staff: '교역자가 대신 신청', other: '기타' }[v] || '';
}

function deliveryLabel_(v, s) {
  return { email: s.delivery.email, visit: s.delivery.visit, direct: s.delivery.direct }[v] || '';
}

function isEmail_(v) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(v || '')); }

// 만 19세 미만인지 (생년월일 8자리 기준)
function isMinor_(birth, now) {
  if (!isValidBirthDate_(birth)) return false;
  var y = Number(birth.slice(0, 4)), m = Number(birth.slice(4, 6)), d = Number(birth.slice(6, 8));
  var today = Utilities.formatDate(now, TZ, 'yyyyMMdd');
  var ty = Number(today.slice(0, 4)), tm = Number(today.slice(4, 6)), td = Number(today.slice(6, 8));
  var age = ty - y - ((tm < m || (tm === m && td < d)) ? 1 : 0);
  return age >= 0 && age < 19;
}

function maskResident_(v) { return String(v).slice(0, 8) + '******'; }

function parseData_(v) {
  try { var o = JSON.parse(String(v || '{}')); return o && typeof o === 'object' ? o : {}; } catch (e) { return {}; }
}

function toDate_(v) {
  var d = v instanceof Date ? v : new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

function dateLabel_(v) {
  var d = toDate_(v);
  return d ? Utilities.formatDate(d, TZ, 'yyyy년 M월 d일') : '';
}

function formatKoreanDate_(v) {
  v = String(v || '');
  var m = /^(\d{4})-?(\d{2})-?(\d{2})$/.exec(v);
  return m ? Number(m[1]) + '년 ' + Number(m[2]) + '월 ' + Number(m[3]) + '일' : v;
}

function isValidBirthDate_(value) {
  if (!/^\d{8}$/.test(value)) return false;
  var y = Number(value.slice(0, 4)), m = Number(value.slice(4, 6)), d = Number(value.slice(6, 8));
  var date = new Date(Date.UTC(y, m - 1, d));
  return date.getUTCFullYear() === y && date.getUTCMonth() === m - 1 && date.getUTCDate() === d && y >= 1900;
}

function clean_(value, max) {
  return String(value == null ? '' : value).replace(/\r\n?/g, '\n').trim().slice(0, max);
}

// 시트에서 수식으로 읽히지 않도록 = + - @ 로 시작하는 글 앞에 ' 를 붙입니다.
function safeCell_(value) {
  value = String(value == null ? '' : value);
  return /^[=+\-@]/.test(value) ? "'" + value : value;
}
