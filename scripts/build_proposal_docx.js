/**
 * 기말 프로젝트 제안서 — 논문 형식 .docx 생성 스크립트
 * 실행: node scripts/build_proposal_docx.js
 *
 * 분량 제약: 과제 공지상 "참고문헌 및 부록 제외 3장 이내".
 * 본문(초록~결론)은 짧고 조밀하게, 구현 상세와 시나리오 표는 부록으로 분리한다.
 */
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType,
  LevelFormat, convertInchesToTwip, PageBreak,
} = require("docx");

const FONT = "Malgun Gothic";
const SIZE = 20; // 10pt
const H_COLOR = "1F3864";

/* ───────────────────────── 도우미 ───────────────────────── */

const T = (text, opts = {}) => new TextRun({ text, font: FONT, size: SIZE, ...opts });

const P = (children, opts = {}) =>
  new Paragraph({
    children: Array.isArray(children) ? children : [T(children)],
    spacing: { after: 110, line: 252 },
    alignment: AlignmentType.JUSTIFIED,
    ...opts,
  });

const H1 = (text, num) =>
  new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 220, after: 90 },
    children: [new TextRun({ text: `${num ? num + ". " : ""}${text}`, bold: true, size: 24, font: FONT, color: H_COLOR })],
  });

const H2 = (text) =>
  new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 140, after: 70 },
    children: [new TextRun({ text, bold: true, size: 21, font: FONT, color: H_COLOR })],
  });

const Bullet = (children) =>
  new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 60, line: 252 },
    children: Array.isArray(children) ? children : [T(children)],
  });

const NumItem = (text) =>
  new Paragraph({
    numbering: { reference: "numbered", level: 0 },
    spacing: { after: 60, line: 252 },
    children: Array.isArray(text) ? text : [T(text)],
  });

const cell = (text, opts = {}) =>
  new TableCell({
    width: { size: opts.width ?? 2000, type: WidthType.DXA },
    shading: opts.head ? { type: ShadingType.CLEAR, fill: "1F3864" } : undefined,
    verticalAlign: "center",
    margins: { top: 40, bottom: 40, left: 100, right: 100 },
    children: [
      new Paragraph({
        alignment: AlignmentType.LEFT,
        children: [new TextRun({
          text, font: FONT, size: 17, bold: !!opts.head,
          color: opts.head ? "FFFFFF" : "000000",
        })],
      }),
    ],
  });

function table(widths, rows) {
  const total = widths.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: widths,
    rows: rows.map((r, i) => new TableRow({ children: r.map((c, j) => cell(c, { width: widths[j], head: i === 0 })) })),
  });
}

const Caption = (text) =>
  new Paragraph({
    spacing: { before: 40, after: 160 },
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text, italics: true, size: 16, font: FONT, color: "555555" })],
  });

const Ref = (n, text) =>
  new Paragraph({
    spacing: { after: 100, line: 240 },
    indent: { left: convertInchesToTwip(0.3), hanging: convertInchesToTwip(0.3) },
    children: [new TextRun({ text: `[${n}] `, font: FONT, size: 18 }), new TextRun({ text, font: FONT, size: 18 })],
  });

/* ───────────────────────── 본문 ───────────────────────── */

const children = [];

// 표지 -----------------------------------------------------------------
children.push(
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 60 },
    children: [new TextRun({ text: "디지털 팬 여권", bold: true, size: 38, font: FONT, color: H_COLOR })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 30 },
    children: [new TextRun({ text: "비양도성 참여 이력에 기반한 팬덤 활동 인증 및 공연 티켓 배분 프로토콜", size: 21, font: FONT, italics: true })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 200 },
    children: [new TextRun({
      text: "Digital Fan Passport: A Soulbound-Record Protocol for Fandom Attestation and Concert Ticket Allocation",
      size: 16, font: FONT, color: "555555",
    })],
  }),
);

children.push(
  P([T("초록  ", { bold: true }), T("국내 공연 티켓 시장은 매크로를 이용한 부정 예매와 고액 재판매가 매년 급증하지만, 사후 신고·단속에 의존하는 현행 체계의 실효 조치율은 6%를 밑돈다. 본 제안서는 이 문제를 두 개의 원안 — 참여 이력의 자기주권적 소유(디지털 팬 여권, 왕예주)와 양도 제한형 재판매 구조(김상미) — 를 하나의 메커니즘으로 결합해 접근한다. 초과수요로 발생하는 지대(rent)는 정가 구조상 소거될 수 없으므로(Becker, 1991), 목표는 지대의 제거가 아니라 귀속 주체의 재분배여야 한다. 본 시스템은 비양도성 토큰(soulbound token)에 실제 입장 이력을 누적시키고, 이를 화폐로 전환해 좌석 일부를 배분함으로써 자본만으로는 참여할 수 없는 배분 시장을 설계한다. 동일한 이력 토큰은 플랫폼 간 통합 이력 증명과 등급별 혜택(5회 이상 참석자 굿즈 선구매권 등, 왕예주 원안)의 근거로도 쓰인다. 이더리움 스마트컨트랙트로 구현하고, 9단계 시나리오·22개 항목의 검증을 통해 암표 차단·반납-재배분·노쇼 보증금·점수 감가가 설계대로 작동함을 확인하였다.")], { spacing: { after: 80 } }),
  P([T("키워드  ", { bold: true }), T("블록체인, 스마트컨트랙트, 소울바운드 토큰, 티켓 재판매, 지대추구, 올페이 경매, 팬덤 플랫폼", { size: 18 })], { spacing: { after: 180 } }),
);

// 1. 서론 -----------------------------------------------------------------
children.push(H1("서론", 1));
children.push(
  P("공연 분야 암표 신고는 2020년 359건에서 2022년 4,224건으로 10배 이상 급증했으나, 2023~2025년(8월 기준) 신고 중 유효신고 인정 비율은 약 5.6%(306건)에 그쳤다[7]. 매크로 사용의 기술적 입증이 어렵고 '부정한 방법'의 법적 정의가 모호하다는 점이 원인으로 지적된다. 즉 현행 대응은 구조적으로 사후적이며, 통계상 작동하지 않고 있다."),
  P("이와 별개로 팬 활동 이력은 인터파크·예스24·위버스 등 플랫폼별로 파편화되어 있고, 특정 플랫폼을 이탈하는 순간 소멸한다. 5년을 따라다닌 팬과 이번이 처음인 관객이 예매창 앞에서 구별되지 않는 이유다."),
  P("본 연구는 (a) 암표의 구조적 방지와 (b) 팬 이력의 파편화·소유권 부재가 하나의 해법으로 묶일 수 있다는 관찰에서 출발한다. 참여 이력을 팬 본인이 소유하는 위조·양도 불가능한 자산으로 만들면, 이는 플랫폼에 독립적인 통합 이력 증명인 동시에 자본으로 살 수 없는 배분 지불수단이 될 수 있다."),
);

// 2. 이론적 배경 ------------------------------------------------------------
children.push(H1("이론적 배경", 2));
children.push(
  P([T("초과수요의 구조적 지속성. ", { bold: true }), T("'정가를 시장청산 수준까지 올리면 된다'는 반론에 대해, Becker(1991)는 소비자 수요가 서로에게 양(+)의 영향을 주는 사회적 상호의존성으로 인기 식당·공연이 초과수요에도 가격을 고정하는 현상을 설명한다. 매진·대기 자체가 상품 가치의 일부이므로, 정가와 실제 지불의사의 격차(지대)는 정책으로 소거할 수 있는 대상이 아니라 구조적 상수가 된다. 문제는 소거가 아니라 귀속이다.")]),
  P([T("재판매의 이중성. ", { bold: true }), T("Courty(2003)는 프로모터가 2차 시장 이익을 직접 흡수하거나 브로커 진입을 억제하지 못하는 구조를 설명하며, Leslie & Sorensen(2014)은 콘서트 1·2차 시장 데이터 분석으로 재판매가 배분 효율을 평균 5% 개선하되 그 이득의 약 1/3이 1차 시장 사재기 경쟁으로 상쇄됨을 실증했다. 재판매의 배분 기능은 보존하되 지대추구 유인만 제거하는 설계가 필요하다.")]),
  P([T("배분 메커니즘. ", { bold: true }), T("Baye, Kovenock & de Vries(1996)의 올페이 경매(all-pay auction)는 낙찰 여부와 무관하게 입찰액이 소모되는 경매로, 경합·지대추구 상황의 표준 모형이다. 본 연구는 이를 자본이 아닌 참여 이력을 화폐로 하는 배분에 적용한다.")]),
  P([T("소울바운드 토큰. ", { bold: true }), T("Weyl, Ohlhaver & Buterin(2022)은 양도 불가능한 사회적 관계를 인코딩하는 소울바운드 토큰(SBT) 개념을 제시했으며, 시빌 저항성 논의의 기반이 된다는 점에서 본 시스템의 '이력은 매매될 수 없다'는 전제와 부합한다.")]),
);

// 3. 관련 아이디어의 결합 -----------------------------------------------------
children.push(H1("관련 아이디어의 통합", 3));
children.push(
  P([T("디지털 팬 여권(왕예주 원안)", { bold: true }), T("은 참여 기록을 블록체인에 누적해 플랫폼 간 이력을 통합하고, 양도 불가능하게 하여 실제 참여를 증명하며, 5회 이상 참석자에게 굿즈 선구매·특별 이벤트 기회를 부여하는 안이다. 원안이 강조하듯 요점은 혜택이 아니라 팬의 이력 자기소유·증명에 있다.")]),
  P([T("양도 제한형 재판매(김상미 원안)", { bold: true }), T("은 티켓을 블록체인에 등록해 소유자를 검증 가능하게 하고, 직접 양도 대신 공식 반납 후 재판매만 허용해 암표의 실효성을 낮추는 구조다.")]),
  P("두 원안은 독립적으로는 공백을 남긴다. 후자는 반납 좌석을 '누구에게' 재배분할지 기준이 없어 선착순 운영 시 매크로 경쟁이 재현되고, 전자는 이력의 용도가 굿즈 선구매 수준에 머물러 암표 문제와 직결되지 않는다. 본 제안은 전자의 '비양도성 이력'을 후자의 '반납 좌석 배분 기준'으로 사용하여, 팬 이력을 화폐화한 경매로 재배분함으로써 두 공백을 서로 메운다."),
);

// 4. 제안 시스템 --------------------------------------------------------
children.push(H1("제안 시스템", 4));
children.push(
  P("시스템은 두 계약으로 구성된다. FanPassport는 팬 1인당 하나씩 발급되어 전체 팬덤에 걸쳐 단일하게 유지되는 비양도성 이력 토큰이며, TicketBox는 공연마다 배포되어 발권·배분·입장을 관장한다. 모든 TicketBox가 공통의 FanPassport를 참조하므로 특정 예매처에 국한되지 않는 이력 통합이 이루어진다(왕예주 원안의 구현)."),
);
children.push(
  Bullet([T("FanPassport — ", { bold: true }), T("전송·승인 함수가 항상 실패하는 소울바운드 설계, 예매가 아닌 실제 입장 시에만 생성되는 기록, 2년에 걸쳐 선형 감가하는 점수(신규 팬 진입 여지 확보), 지갑 분실 시 발급기관에 의한 이관(분할·합산 불가).")]),
  Bullet([T("TicketBox — ", { bold: true }), T("공식 반납처 외 전송이 전면 차단된 티켓. 좌석은 ")
    , T("우선권 트랙", { bold: true }), T("(팬 점수 올페이 경매)과 ")
    , T("일반 트랙", { bold: true }), T("(무작위 추첨)으로 이중 배분한다. 반납은 정가 전액 환불되며 반납분은 다음 라운드 우선권 트랙으로 재편입되어, 재판매의 배분 기능은 유지하되 웃돈 요소만 제거한다.")]),
  Bullet([T("노쇼 보증금 — ", { bold: true }), T("입장·반납 시 전액 환급, 무단 노쇼에 한해 몰수 후 다음 회차 리워드 재원으로 적립.")]),
  Bullet([T("등급 인증 — ", { bold: true }), T("누적 참석 횟수는 별도 연동 계약 없이 제3자가 조회 가능한 공개 함수로 노출되어, '5회 이상 참석자 굿즈 선구매'(왕예주 원안)가 그대로 실현된다.")]),
);

// 5. 왜 블록체인인가 -------------------------------------------------------
children.push(H1("왜 블록체인이어야 하는가", 5));
children.push(
  table(
    [2500, 900, 5800],
    [
      ["요구사항", "중앙 DB", "이유"],
      ["복수 예매처 간 이력 통합", "불가", "경쟁 플랫폼은 서로의 DB를 신뢰하지 않으며, 중립적 공동 원장이 필요하다."],
      ["팬의 이력 자기소유", "불가", "플랫폼 탈퇴·서비스 종료 시 이력이 함께 소멸한다."],
      ["계약 없는 제3자 등급 검증", "불가", "매번 개별 API 연동과 데이터 제공 계약이 필요하다."],
      ["발행 주체도 임의 수정 불가", "불가", "DB 관리자는 구조적으로 수정 권한을 보유한다."],
      ["배분 규칙의 사전공표·사후검증", "불가", "'왜 특정인이 당첨됐는가'를 팬이 직접 재계산할 수 없다."],
    ]
  )
);
children.push(Caption("표 1. 중앙형 DB와 블록체인 기반 설계의 요구사항 비교"));
children.push(
  P("특히 마지막 행이 결정적이다. 티켓팅 불신의 근원은 암표 자체보다 배분 과정의 불투명성에 있다. 스마트컨트랙트는 규칙을 코드로 사전 고정하고 결과를 누구나 재계산·검증할 수 있게 하며, 전송 제한·보증금 몰수는 인적 심사 없이 자동 집행된다. 이는 사후 단속(유효조치율 약 5.6%)을 사전 차단으로 전환하는 지점이다."),
);

// 6. 비교 및 한계 ------------------------------------------------------------
children.push(H1("비교 및 한계", 6));
children.push(
  table(
    [1900, 3400, 3900],
    [
      ["항목", "현행 티켓팅", "제안 시스템"],
      ["암표 대응", "사후 단속 (조치율 약 5.6%)", "소유권 이전 경로 자체가 부재"],
      ["배분 기준", "예매 속도·매크로", "참여 이력(경매) + 무작위(추첨)"],
      ["취소 처리", "수수료 부담 또는 암표 유출", "정가 전액 환불 후 자동 재배분"],
      ["팬 이력", "플랫폼별 파편화, 탈퇴 시 소멸", "팬 본인 소유, 플랫폼 독립적 유지"],
      ["지대의 귀속", "암표상", "실제 참여 팬"],
    ]
  )
);
children.push(Caption("표 2. 현행 시스템과 제안 시스템의 비교"));
children.push(
  P([T("한계  ", { bold: true }), T("① 본인확인은 오프체인 문제로, 신분 위조를 통한 복수 발급 자체는 차단하지 못한다. ② 입장 검표 단말이라는 단일 신뢰점(오라클 문제)이 존재한다. ③ 참석 이력이 지갑 주소에 결부되어 공개되므로 프라이버시 보호(해시 커밋·영지식증명)가 후속 과제다. ④ 블록해시 기반 추첨 난수는 조작 가능성이 있어 VRF 등이 요구된다. ⑤ 지갑·가스비는 진입 장벽이며 계정 추상화가 필요하다. ⑥ 지갑 비밀키 자체를 양도하는 우회 거래는 방지되지 않으나, 이 경우 판매자의 신분증명까지 필요해 거래비용이 크게 상승한다.")]),
);

// 7. 결론 -------------------------------------------------------------------
children.push(H1("결론", 7));
children.push(
  P("본 연구는 암표 방지와 팬 이력의 소유권 부재가 별개 문제가 아니라, 비양도성 참여 이력을 배분의 화폐로 삼는 단일 메커니즘으로 동시에 해소될 수 있음을 논증하고 스마트컨트랙트로 구현·검증하였다(부록 A·B). 이는 왕예주 원안(이력의 자기주권적 소유)과 김상미 원안(양도 제한형 재판매)을 대립적 대안이 아닌 상호 보완 요소로 재해석한 결과다. 향후 영지식증명 기반 프라이버시 보호, 검증가능 난수(VRF) 도입, 실사용자 파일럿을 통한 가스비·UX 검증을 계획한다."),
);

// 참고문헌 -------------------------------------------------------------------
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(H1("참고문헌", null));
children.push(
  Ref(1, "Becker, G. S. (1991). A Note on Restaurant Pricing and Other Examples of Social Influences on Price. Journal of Political Economy, 99(5), 1109–1116."),
  Ref(2, "Leslie, P., & Sorensen, A. (2014). Resale and Rent-Seeking: An Application to Ticket Markets. Review of Economic Studies, 81(1), 266–300."),
  Ref(3, "Courty, P. (2003). Some Economics of Ticket Resale. Journal of Economic Perspectives, 17(2), 85–97."),
  Ref(4, "Baye, M. R., Kovenock, D., & de Vries, C. G. (1996). The all-pay auction with complete information. Economic Theory, 8(2), 291–305."),
  Ref(5, "Weyl, E. G., Ohlhaver, P., & Buterin, V. (2022). Decentralized Society: Finding Web3's Soul. SSRN Working Paper No. 4105763."),
  Ref(6, "ERC-5484: Consensual Soulbound Tokens. https://eips.ethereum.org/EIPS/eip-5484"),
  Ref(7, "박수현 의원실 (2025). 공연분야 암표신고 현황 자료 (2020–2025.8)."),
);

// 부록 A -------------------------------------------------------------------
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(H1("부록 A. 구현 상세", null));
children.push(
  P("본 구현은 두 개의 Solidity 스마트컨트랙트와 하나의 시연 스크립트로 구성되며, 코드 전문은 첨부 파일로 제출한다."),
  P([T("FanPassport.sol  ", { bold: true }), T("비양도성 팬 여권. 1인 1여권 발급(issue), 입장 시점 기록(stamp), 선형 감가가 적용된 점수 조회(points — 적립분과 소모분을 동일한 시간축에서 함께 감가시켜, 소모분만 명목값으로 고정할 경우 발생하는 '전원 0점 수렴' 오류를 방지), 지갑 분실 시 발급기관에 의한 이관(migrate), transferFrom/approve 등 모든 양도 경로의 명시적 차단으로 구성된다.")]),
  P([T("TicketBox.sol  ", { bold: true }), T("공연 1건의 발권·배분 계약. 우선권/일반 이중 트랙 응모(enterDraw)와 추첨 확정(draw), 공식 반납처 외 전송을 전면 차단하는 transferFrom, 정가·보증금 전액 환불 반납(returnTicket), 현장 검표 시 여권 기록과 보증금 환급을 동시 처리하는 체크인(checkIn), 무단 노쇼 보증금 몰수(closeNoShow)로 구성된다.")]),
  P([T("scripts/demo.js  ", { bold: true }), T("인메모리 체인(ganache) 위에 두 계약을 실제 배포하고, 6명의 팬과 1명의 암표 시험 계정으로 부록 B의 9단계 시나리오를 실행·검증한다. "), T("npm install && npm run demo", { italics: true }), T(" 로 재현 가능하다.")]),
);

// 부록 B -------------------------------------------------------------------
children.push(H1("부록 B. 시나리오 기반 검증 결과", null));
children.push(
  table(
    [600, 4100, 4500],
    [
      ["단계", "시나리오", "확인된 사항"],
      ["1", "팬 6명 여권 발급, 과거 참석 이력 부여", "암표 시험 계정도 여권은 발급되나 점수는 0"],
      ["2", "이력 자체를 매매하려는 시도", "전송·승인 함수 모두 거부 (Soulbound)"],
      ["3", "1차 배정: 우선권 2석 + 일반 2석", "부계정 중복 응모 차단, 신규 팬도 추첨 당첨 가능"],
      ["4", "암표 시험 계정의 재판매 시도", "직접 전송·마켓 등록·우회 전송 모두 거부 (NonTransferable)"],
      ["5", "판매 실패 후 공식 반납", "정가·보증금 전액 환불, 차익 0"],
      ["6", "반납 좌석의 재배분", "무점수 계정의 입찰 거부 (InsufficientPoints)"],
      ["7", "공연 당일 입장 및 노쇼 처리", "입장 시에만 기록 생성·보증금 환급, 노쇼 시 몰수"],
      ["8", "1년 경과 후 점수 조회", "선형 감가 확인"],
      ["9", "최종 상태 점검", "암표 시험 계정의 최종 참석 기록 0회"],
    ]
  )
);
children.push(Caption("표 3. 시연 시나리오 9단계 및 검증 결과 (총 22개 항목 전부 통과)"));

/* ───────────────────────── 문서 조립 ───────────────────────── */

const doc = new Document({
  numbering: {
    config: [
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "numbered", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ],
  },
  styles: { default: { document: { run: { font: FONT, size: SIZE } } } },
  sections: [
    {
      properties: { page: { margin: { top: 1020, bottom: 1020, left: 1020, right: 1020 } } }, // 1.8cm
      children,
    },
  ],
});

const outPath = path.join(__dirname, "..", "docs", "기말프로젝트_제안서.docx");
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(outPath, buf);
  console.log("wrote", outPath, buf.length, "bytes");
});
