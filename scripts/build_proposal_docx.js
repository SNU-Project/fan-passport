/**
 * 기말 프로젝트 제안서 — 서울대 팀 프로젝트 보고서 양식(서체·헤딩 스타일·페이지 규격)을
 * 따르되, 과제 공지의 "참고문헌·부록 제외 3장 이내" 규정에 맞춰 표지·목차 없이
 * 제목 → 초록 → 본문으로 곧장 시작하는 압축 버전.
 *
 * 실행: node scripts/build_proposal_docx.js
 */
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType,
  LevelFormat, convertInchesToTwip, PageBreak, PageNumber, Footer,
  ExternalHyperlink, UnderlineType,
} = require("docx");

const FONT = "NanumMyeongjo";
const SIZE = 20; // 10pt — 3장 제한에 맞춰 여백·본문을 함께 압축한다.
const HCOLOR = "1F1F1F";

// 원본 서울대 양식과 같은 판형(≈19.0×26.0cm)은 유지하되, 여백은 좁혀 3장에 맞춘다.
const PAGE_W = 10772, PAGE_H = 14740;
const MARGIN = 780; // ≈1.38cm, 상하좌우 동일

/* ───────────────────────── 도우미 ───────────────────────── */

const T = (text, opts = {}) => new TextRun({ text, font: FONT, size: SIZE, ...opts });

const P = (children, opts = {}) =>
  new Paragraph({
    children: Array.isArray(children) ? children : [T(children)],
    spacing: { after: 60, line: 232 },
    alignment: AlignmentType.JUSTIFIED,
    ...opts,
  });

const H1 = (num, text) =>
  new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [T(`제 ${num} 장  ${text}`, { bold: true, size: 26, color: HCOLOR })],
  });

const H2 = (num2, text) =>
  new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [T(`제 ${num2} 절  ${text}`, { bold: true, size: 22, color: HCOLOR })],
  });

const BareHeading = (text) =>
  new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [T(text, { bold: true, size: 26, color: HCOLOR })],
  });

const Bullet = (children) =>
  new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 40, line: 232 },
    children: Array.isArray(children) ? children : [T(children)],
  });

const NumItem = (children) =>
  new Paragraph({
    numbering: { reference: "numbered", level: 0 },
    spacing: { after: 40, line: 232 },
    children: Array.isArray(children) ? children : [T(children)],
  });

const cell = (text, opts = {}) =>
  new TableCell({
    width: { size: opts.width ?? 2000, type: WidthType.DXA },
    shading: opts.head ? { type: ShadingType.CLEAR, fill: "1F1F1F" } : undefined,
    verticalAlign: "center",
    margins: { top: 40, bottom: 40, left: 90, right: 90 },
    children: [new Paragraph({
      alignment: AlignmentType.LEFT,
      children: [new TextRun({ text, font: FONT, size: 16, bold: !!opts.head, color: opts.head ? "FFFFFF" : "000000" })],
    })],
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
    spacing: { before: 40, after: 140 },
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text, italics: true, size: 15, font: FONT, color: "555555" })],
  });

const Link = (url, label) =>
  new ExternalHyperlink({
    link: url,
    children: [T(label ?? url, { color: "1155CC", underline: { type: UnderlineType.SINGLE } })],
  });

const Ref = (n, text) =>
  new Paragraph({
    spacing: { after: 90, line: 240 },
    indent: { left: convertInchesToTwip(0.32), hanging: convertInchesToTwip(0.32) },
    children: [new TextRun({ text: `[${n}] `, font: FONT, size: 17 }), new TextRun({ text, font: FONT, size: 17 })],
  });

/* ═══════════════════════ 본문 ═══════════════════════ */

const body = [];

// 제목 (표지·목차·초록 없이 곧장 서론으로) --------------------------------------
body.push(
  new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { after: 40 },
    children: [T("디지털 팬 여권", { bold: true, size: 34 })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { after: 180 },
    children: [T("— 콘서트 암표를 막고 팬의 참여 기록을 지켜주는 블록체인 설계 —", { bold: true, size: 20 })],
  }),
);

// 제1장 서론 -----------------------------------------------------------
body.push(H1(1, "서론"));
body.push(
  P("공연 분야 암표 신고는 2020년 359건에서 2022년 4,224건으로 급증했지만, 2023~2025년 8월 접수분 중 유효신고로 인정된 것은 약 5.6%(306건)에 그쳤다[7]. 매크로 사용의 기술적 입증이 어렵고 '부정한 방법'의 법적 정의가 모호해, 지금의 대응은 전부 사후적이며 숫자로는 작동하지 않는다. 이와 별개로 팬의 관람 이력은 예매처마다 흩어져 탈퇴 시 사라지므로, 5년째 따라다닌 팬과 오늘 처음 온 관객이 예매창 앞에서 구별되지 않는다."),
  P("본 제안서는 이 두 문제 — 암표 방지 실패와 팬 이력의 파편화 — 를 하나로 묶어 풀 수 있다는 관찰에서 출발한다. 참여 이력을 팬 본인이 소유하되 양도 불가능하게 만들면, 그 기록은 플랫폼 독립적인 통합 이력 증명이자 돈으로는 살 수 없는 티켓 배분 수단이 된다. 이어지는 장에서 이론적 근거, 시스템 설계, 구현·검증 결과를 차례로 제시한다."),
);

// 제2장 이론적 배경 (압축 — 세 근거를 한 절로) ---------------------------------
body.push(H1(2, "이론적 배경"));
body.push(
  P([T("초과수요의 구조적 지속성. ", { bold: true }), T("맛집이 줄을 세우면서도 값을 안 올리는 현상을, 베커(Becker, 1991)는 수요의 사회적 상호의존성으로 설명한다. 매진 자체가 상품 가치의 일부라서 기획사는 값을 시장청산 수준까지 올리지 않고, 정가와 실제 지불의사의 격차(지대, rent)는 구조적으로 남는다. 문제는 이 격차의 소거가 아니라 귀속 주체다.", {})]),
  P([T("재판매의 이중성. ", { bold: true }), T("레슬리·소렌슨(Leslie & Sorensen, 2014)은 콘서트 1·2차 시장 자료로 재판매가 배분 효율을 평균 5% 높이지만 그 이득의 3분의 1이 1차 시장 사재기로 상쇄됨을 보였다(Courty, 2003). 본 제안의 목표는 재판매의 배분 기능은 남기고 사재기 유인만 제거하는 것이다.", {})]),
  P([T("배분 메커니즘. ", { bold: true }), T("낙찰 여부와 무관하게 입찰액이 소모되는 올페이 경매는 지대추구 상황의 표준 모형이며(Baye, Kovenock & de Vries, 1996), 이를 참여 이력 기반 배분에 적용한다. 양도 불가능한 소울바운드 토큰(SBT, Weyl, Ohlhaver & Buterin, 2022)이 그 이력을 담는다.", {})]),
);

// 제3장 시스템 설계 --------------------------------------------------------
body.push(H1(3, "시스템 설계"));
body.push(
  P("설계 초기의 두 방향 — 이력의 자기소유·증명, 그리고 양도 제한형 반납·재판매 — 은 각각 빈틈이 있다. 후자는 반납 표의 재배분 기준이 없어 선착순 운영 시 매크로 경쟁이 재현되고, 전자는 이력의 용도가 굿즈 선구매 수준에 머물러 암표 문제와 직결되지 않는다. 본 설계는 전자의 '양도 불가능한 이력'을 후자의 '반납 표 배분 기준'으로 사용해 두 빈틈을 상호 보완한다."),
  P("시스템은 두 계약으로 구성된다. 팬 1인당 발급되는 팬 여권(FanPassport)은 전송·승인이 항상 실패하는 소울바운드 토큰으로, 실제 입장 시점에만 기록되고 점수는 2년에 걸쳐 선형 감가한다(신규 팬 진입 여지 확보). 공연별로 배포되는 티켓함(TicketBox)의 표는 공식 반납 창구로만 전송되며, 좌석은 팬 점수 올페이 경매인 우선권 트랙과 무작위 추첨인 일반 트랙으로 배분한다. 반납은 정가 전액 환불 후 다음 라운드 우선권 트랙으로 재편입되고, 보증금은 입장·반납 시 환급되며 무단 노쇼에만 몰수돼 리워드 재원이 된다. 누적 참석 횟수는 계약 없이 제3자가 조회 가능해, 등급별 혜택 활용도 실현된다."),
);

// 제4장 블록체인 도입의 필요성 (채점 배점 항목 — 두껍게 유지) --------------------
body.push(H1(4, "블록체인 도입의 필요성"));
body.push(
  P("'예매처가 서버 DB로 관리하면 되지 않는가'라는 질문에 답하려면, 중앙화된 서버로는 원리적으로 풀리지 않는 지점을 짚어야 한다."),
);

body.push(H2(1, "경쟁 플랫폼 간 신뢰 문제"));
body.push(
  P("인터파크·예스24·위버스 등 경쟁 플랫폼의 이력을 통합하려면 한 회사가 DB를 소유·운영해야 하고, 나머지는 그 서버 로그를 무조건 신뢰해야 한다. 이 때문에 예매처 간 이력 통합은 시도조차 되지 않았다. 블록체인은 이를 '특정 회사에 대한 신뢰'에서 '검증 가능한 규칙에 대한 신뢰'로 전환한다."),
);
body.push(H2(2, "위변조 불가능성"));
body.push(
  P("중앙 DB의 관리자는 구조적으로 언제나 수정 권한을 갖는다. 발행 회사조차 특정 팬의 이력을 지우거나 당첨 확률을 조작할 가능성이 열려 있다. 블록체인은 기록이 해시체인으로 엮여 과거 기록 하나를 바꾸려면 이후 전체를 바꿔야 하므로 사실상 불가능하다. 이 성질은 팬 여권 기록이 발급기관 자신에 의해서도 조작될 수 없다는 것과, 경매 결과가 사후에 바뀔 수 없다는 것을 보장한다."),
);
body.push(H2(3, "배분 결과의 검증가능성"));
body.push(
  P("티켓팅 불신의 뿌리는 암표 자체보다 배분 과정의 불투명성에 있다. 스마트컨트랙트로 규칙을 코드화하면 누구나 이를 읽고, 경매·추첨 결과를 온체인 입력값으로부터 동일하게 재계산해 검증할 수 있다. 중앙 서버는 로그를 공개해도 그 로그의 무결성까지는 증명하지 못한다는 점에서 근본적으로 다르다."),
);
body.push(H2(4, "규칙의 자동 집행"));
body.push(
  P("현행법은 부정 예매와 웃돈 재판매를 모두 입증해야 처벌 가능한 사후 절차이며, 이 때문에 유효조치율이 5.6%에 그친다. 스마트컨트랙트는 전송 제한·반납·보증금 몰수를 조건 충족 시 즉시 실행하고 미충족 시 트랜잭션 자체를 거부해, 암표를 '사후 적발'에서 '사전 차단'으로 전환한다."),
);

body.push(
  table(
    [2400, 800, 5900],
    [
      ["필요한 것", "중앙 DB", "왜 안 되는가"],
      ["경쟁 예매처 간 이력 통합", "불가능", "한 회사가 DB를 관리해야 하고, 경쟁사는 그 회사 서버를 무조건 신뢰해야 한다."],
      ["발행 주체도 못 고치는 기록", "불가능", "관리자는 구조적으로 언제나 수정 권한을 갖는다."],
      ["결과를 발표가 아닌 재계산으로 검증", "불가능", "로그를 공개해도 그 로그가 조작 안 됐음은 증명 못 한다."],
      ["규칙을 심사 없이 즉시 자동 집행", "불가능", "예외·재량이 필요한 로직은 결국 운영자 판단에 맡겨진다."],
      ["팬이 플랫폼과 무관하게 이력 보유", "불가능", "탈퇴·서비스 종료 시 이력도 함께 사라진다."],
    ]
  )
);
body.push(Caption("표 1. 중앙화된 DB로는 원리적으로 풀리지 않는 다섯 가지 요구"));

// 제5장 기존 시스템과의 비교 및 장단점 (채점 배점 항목 — 두껍게 유지) -----------
body.push(H1(5, "기존 시스템과의 비교 및 장단점"));

body.push(H2(1, "장점"));
body.push(
  table(
    [1800, 3400, 4000],
    [
      ["항목", "지금의 티켓팅", "제안하는 시스템"],
      ["암표 대응", "사후 신고·단속 (조치율 약 5.6%)", "표를 넘길 방법 자체가 없음"],
      ["누구에게 표가 가는가", "예매 속도, 매크로 여부", "참여 이력(경매) + 무작위(추첨)"],
      ["못 가게 됐을 때", "수수료 부담 또는 암표로 흘러감", "정가·보증금 전액 환불 후 자동 재배분"],
      ["팬의 이력", "예매처마다 흩어지고 탈퇴 시 소멸", "팬 본인 소유, 플랫폼과 무관하게 유지"],
      ["부족분 이득(지대)의 귀속", "암표상", "실제로 참여한 팬"],
      ["규칙 집행 비용", "신고·조사 인건비", "코드 실행 비용(가스비), 5.2절"],
    ]
  )
);
body.push(Caption("표 2. 지금의 티켓팅과 제안 시스템의 비교"));
body.push(
  P("암표 대응이 '적발률을 높이는 문제'에서 '거래 자체가 성립하지 않는 문제'로 바뀌고, 표 부족분의 이득이 누구에게 가는지가 처음으로 정책 목표로 다뤄진다는 점이 핵심이다."),
);

body.push(H2(2, "단점 및 새로 발생하는 비용"));
body.push(
  Bullet([T("거래 비용. ", { bold: true }), T("계약 실측 결과 이더리움 메인넷 기준(20 gwei, 1 ETH=500만원 가정) 응모 1회 약 8,700~10,500원, 발권 1회 약 12,200원의 수수료가 발생했다. 레이어2 도입으로 이 비용을 낮추는 것이 전제되어야 한다.")]),
  Bullet([T("규칙의 경직성. ", { bold: true }), T("배포 후 코드 수정이 어려워 버그나 예외 상황(공연 취소 등)에 서버처럼 즉시 대응하기 어렵고, 업그레이드 구조를 두면 권한 귀속이라는 새 신뢰 문제가 생긴다.")]),
  Bullet([T("지갑 관리 부담. ", { bold: true }), T("비밀키 분실 시 원칙적으로 복구 불가하다. 발급기관 이관 기능(3장)으로 일부 완화했으나, 이는 중앙화된 신뢰점의 부분적 재도입을 뜻한다.")]),
  Bullet([T("개발·감사 난이도. ", { bold: true }), T("버그의 대가가 자금 손실로 직결될 수 있어 배포 전 보안 검증 비용과 기간이 늘어난다.")]),
);

body.push(H2(3, "잔존하는 한계"));
body.push(
  NumItem("본인 확인은 오프체인 실명 확인에 의존하며, 신분 위조를 통한 복수 발급은 막지 못한다."),
  NumItem("입구 검표 단말을 신뢰해야 하는 오라클 문제가 남는다."),
  NumItem("지갑 주소로 참석 이력이 노출되는 프라이버시 문제가 있어, 해시 커밋·영지식증명이 후속 과제다."),
  NumItem("데모의 추첨 난수는 조작 가능하므로 실서비스는 체인링크 VRF 등이 필요하다."),
  NumItem("비밀키 자체를 통째로 넘기는 우회 거래는 막지 못하나, 판매자 신분증까지 필요해 거래비용이 크게 늘어난다."),
);

// 제6장 결론 -----------------------------------------------------------------
body.push(H1(6, "결론"));
body.push(
  P("암표 방지 실패와 팬 이력의 파편화는 '양도 불가능한 참여 이력을 배분의 화폐로 쓴다'는 하나의 장치로 함께 해소된다. 표 부족분의 이득은 정책으로 소거할 수 없으므로 이를 암표상이 아닌 참여 팬에게 귀속시키는 것이 현실적 목표이며, 스마트컨트랙트로 구현해 9단계 시나리오·22개 검증 항목으로 확인했다(부록 A·B). 향후 프라이버시 보호, 검증가능 난수 생성, 실사용자 시범 운영을 통한 가스비·사용성 검증을 계획한다."),
);

// 참고문헌 -------------------------------------------------------------------
body.push(new Paragraph({ children: [new PageBreak()] }));
body.push(BareHeading("참고문헌"));
body.push(
  Ref(1, "Becker, G. S. (1991). A Note on Restaurant Pricing and Other Examples of Social Influences on Price. Journal of Political Economy, 99(5), 1109–1116."),
  Ref(2, "Leslie, P., & Sorensen, A. (2014). Resale and Rent-Seeking: An Application to Ticket Markets. Review of Economic Studies, 81(1), 266–300."),
  Ref(3, "Courty, P. (2003). Some Economics of Ticket Resale. Journal of Economic Perspectives, 17(2), 85–97."),
  Ref(4, "Baye, M. R., Kovenock, D., & de Vries, C. G. (1996). The all-pay auction with complete information. Economic Theory, 8(2), 291–305."),
  Ref(5, "Weyl, E. G., Ohlhaver, P., & Buterin, V. (2022). Decentralized Society: Finding Web3's Soul. SSRN Working Paper No. 4105763."),
  Ref(6, "ERC-5484: Consensual Soulbound Tokens. https://eips.ethereum.org/EIPS/eip-5484"),
  Ref(7, "박수현 의원실 (2025). 공연분야 암표신고 현황 자료 (2020–2025.8)."),
);

// 부록 A -------------------------------------------------------------------
body.push(new Paragraph({ children: [new PageBreak()] }));
body.push(BareHeading("부록 A. 구현 상세"));
body.push(
  P([T("실제 구현은 두 개의 Solidity 스마트컨트랙트와 하나의 시연 스크립트로 이루어져 있으며, 코드 전문은 첨부 파일로 함께 제출한다. ", {}), T("아래 링크에서 지갑·설치 없이 컨트랙트 로직이 브라우저에서 즉시 실행되는 것을 클릭 몇 번으로 확인할 수 있다: ", {}), Link("https://snu-project.github.io/fan-passport/", "인터랙티브 데모 바로가기"), T(" (소스: ", {}), Link("https://github.com/SNU-Project/fan-passport", "github.com/SNU-Project/fan-passport"), T(")", {})]),
  P([T("FanPassport.sol  ", { bold: true }), T("팔 수 없는 팬 여권을 구현한 계약이다. 한 사람에게 하나씩만 여권을 발급하고(issue), 실제 입장한 시점에만 기록을 남기며(stamp), 시간이 지날수록 선형으로 줄어드는 점수를 계산해 조회할 수 있게 했다(points). 적립된 점수와 이미 써버린 점수를 반드시 같은 속도로 함께 줄여야 하는데, 써버린 점수만 명목값으로 고정해 두면 시간이 지날수록 모든 팬의 점수가 0으로 수렴해 버리는 오류가 생긴다는 점을 구현 중 직접 확인하고 수정했다. 지갑 분실 시에는 발급기관이 본인 확인을 거쳐 기록 전체를 새 지갑으로 옮겨 줄 수 있게 했고(migrate), 전송·승인 등 양도 관련 함수는 모두 명시적으로 실패하도록 만들었다.")]),
  P([T("TicketBox.sol  ", { bold: true }), T("공연 한 건의 발권과 배분을 담당하는 계약이다. 우선권·일반 트랙으로 나누어 응모를 받고(enterDraw) 추첨을 확정하며(draw), 공식 반납 창구가 아니면 어떤 전송도 막는 transferFrom을 두었다. 표를 반납하면 정가와 보증금을 전액 돌려주고(returnTicket), 현장 검표 시 여권에 기록을 남기는 동시에 보증금을 돌려주며(checkIn), 무단 노쇼는 보증금을 몰수한다(closeNoShow).")]),
  P([T("scripts/demo.js  ", { bold: true }), T("인메모리 이더리움(ganache) 위에 두 계약을 실제로 배포하고, 팬 여섯 명과 암표상 역할의 시험 계정 한 명으로 부록 B의 아홉 단계 시나리오를 실행해 검증하는 스크립트다. "), T("npm install && npm run demo", { italics: true }), T(" 로 누구나 재현할 수 있으며, 커밋마다 GitHub Actions가 동일한 검증을 서버에서 재실행한다.")]),
);

// 부록 B -------------------------------------------------------------------
body.push(BareHeading("부록 B. 시나리오로 확인한 결과"));
body.push(
  table(
    [500, 4100, 4600],
    [
      ["단계", "시나리오", "확인된 사항"],
      ["1", "팬 여섯 명에게 여권 발급, 과거 참석 이력 부여", "암표 시험 계정도 여권은 받지만 점수는 0"],
      ["2", "이력 자체를 팔아 보려는 시도", "전송·승인 함수 모두 거부됨 (Soulbound)"],
      ["3", "1차 배정: 우선권 2석 + 일반 2석", "부계정 중복 응모 차단, 신규 팬도 추첨 당첨 가능"],
      ["4", "암표 시험 계정이 표를 되팔려는 시도", "직접 전송·마켓 등록·우회 전송 모두 거부됨 (NonTransferable)"],
      ["5", "되팔지 못해 공식 반납", "정가·보증금 전액 환불, 남는 차익 0"],
      ["6", "반납된 좌석을 다시 배분", "점수 없는 계정의 입찰 거부됨 (InsufficientPoints)"],
      ["7", "공연 당일 입장 및 노쇼 처리", "입장할 때만 기록이 남고 보증금이 돌아오며, 노쇼는 몰수됨"],
      ["8", "1년이 지난 뒤 점수 조회", "선형으로 줄어드는 점수 확인"],
      ["9", "최종 상태 점검", "암표 시험 계정의 최종 참석 기록은 0회"],
    ]
  )
);
body.push(Caption("표 3. 시연 시나리오 아홉 단계와 그 결과 (총 22개 항목 전부 통과)"));

/* ═══════════════════════ 문서 조립 ═══════════════════════ */

const doc = new Document({
  numbering: {
    config: [
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 640, hanging: 320 } } } }] },
      { reference: "numbered", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 640, hanging: 320 } } } }] },
    ],
  },
  styles: {
    default: {
      document: { run: { font: FONT, size: SIZE } },
      heading1: {
        run: { font: FONT, bold: true, size: 26, color: HCOLOR },
        paragraph: { alignment: AlignmentType.CENTER, spacing: { before: 140, after: 90 } },
      },
      heading2: {
        run: { font: FONT, bold: true, size: 22, color: HCOLOR },
        paragraph: { spacing: { before: 90, after: 50 } },
      },
    },
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: PAGE_W, height: PAGE_H },
          margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN, header: 400, footer: 400 },
          pageNumbers: { start: 1 },
        },
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16 })],
          })],
        }),
      },
      children: body,
    },
  ],
});

const outPath = path.join(__dirname, "..", "docs", "기말프로젝트_제안서.docx");
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(outPath, buf);
  console.log("wrote", outPath, buf.length, "bytes");
});
