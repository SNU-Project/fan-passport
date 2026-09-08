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
const MARGIN = 700; // ≈1.23cm, 상하좌우 동일

/* ───────────────────────── 도우미 ───────────────────────── */

const T = (text, opts = {}) => new TextRun({ text, font: FONT, size: SIZE, ...opts });

const P = (children, opts = {}) =>
  new Paragraph({
    children: Array.isArray(children) ? children : [T(children)],
    spacing: { after: 60, line: 222 },
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
    spacing: { after: 40, line: 222 },
    children: Array.isArray(children) ? children : [T(children)],
  });

/** 번호 목록. 같은 reference를 공유하면 번호가 이어지므로 목록마다 다른 ref를 준다. */
const NumItem = (children, ref = "numbered") =>
  new Paragraph({
    numbering: { reference: ref, level: 0 },
    spacing: { after: 40, line: 222 },
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
    // cantSplit: 행이 페이지 경계에서 반쪽으로 잘리지 않도록 통째로 넘긴다.
    rows: rows.map((r, i) => new TableRow({
      cantSplit: true,
      children: r.map((c, j) => cell(c, { width: widths[j], head: i === 0 })),
    })),
  });
}

const Caption = (text) =>
  new Paragraph({
    spacing: { before: 30, after: 110 },
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
  P("공연 분야 암표 신고는 2020년 359건에서 2022년 4,224건으로 급증했으나, 2023~2025년 8월 접수분 중 유효신고로 인정된 것은 약 5.6%(306건)에 그쳤다[7]. 매크로 사용의 입증이 어렵고 '부정한 방법'의 정의가 모호해 현행 대응은 전부 사후적이다. 한편 팬의 관람 이력은 예매처마다 흩어져 탈퇴 시 소멸하므로, 5년째 따라다닌 팬과 오늘 처음 온 관객이 예매창 앞에서 구별되지 않는다. 본 제안서는 이 두 문제를 하나로 묶어 푼다. 참여 이력을 팬 본인이 소유하되 양도 불가능하게 만들면, 그 기록은 플랫폼 독립적인 이력 증명이자 돈으로는 살 수 없는 티켓 배분 수단이 된다."),
);

// 제2장 이론적 배경 (압축 — 세 근거를 한 절로) ---------------------------------
body.push(H1(2, "이론적 배경"));
body.push(
  P([T("지대는 소거되지 않는다. ", { bold: true }), T("매진 자체가 상품 가치의 일부여서 기획사는 값을 시장청산 수준까지 올리지 않는다(Becker, 1991). 따라서 정가와 지불의사의 격차(지대)는 구조적으로 남으며, 문제는 격차의 소거가 아니라 귀속 주체다.", {})]),
  P([T("재판매의 이중성과 배분 수단. ", { bold: true }), T("재판매는 배분 효율을 평균 5% 높이지만 그 이득의 3분의 1이 1차 시장 사재기로 상쇄된다(Leslie & Sorensen, 2014; Courty, 2003). 본 제안은 배분 기능은 남기고 사재기 유인만 제거한다. 수단은 낙찰 여부와 무관하게 입찰액이 소모되는 올페이 경매이며(Baye, Kovenock & de Vries, 1996), 그 지불수단이 되는 참여 이력은 양도 불가능한 소울바운드 토큰(SBT, Weyl, Ohlhaver & Buterin, 2022)에 담는다.", {})]),
);

// 제3장 시스템 설계 --------------------------------------------------------
body.push(H1(3, "시스템 설계"));
body.push(
  P("설계는 세 개의 계약으로 이루어진다. 팬이 공연에 실제로 다녀왔다는 사실을 쌓아 두는 팬 여권, 그 기록을 근거로 좌석을 나눠 주는 티켓함, 그리고 예외적인 양도를 규칙 안에서만 허용하는 공식 양도 계약이다. 여권만 있으면 기록의 쓸모가 굿즈 선구매 정도에 그쳐 암표와 연결되지 않고, 티켓함만 있으면 반납된 표의 배분 기준이 없어 선착순, 즉 매크로 경쟁으로 되돌아간다. 여권의 기록을 티켓함의 배분 기준으로 쓸 때 비로소 두 문제가 함께 풀린다."),
);

body.push(H2(1, "시스템 구성요소"));
body.push(
  P("시스템은 오프체인 본인확인 모듈과 세 개의 스마트컨트랙트로 구성된다. 본인확인은 실명, 중복 계정, 티켓 보유 제한을 확인하되 개인정보 원문은 블록체인에 저장하지 않으며, 체인에는 본인확인 완료 여부와 지갑 주소의 연결 상태만 기록한다."),
);
body.push(
  table(
    [1900, 4600, 2550],
    [
      ["구성요소", "역할", "양도 여부"],
      ["FanPassport", "실제 공연 참석 이력과 리워드 관리", "불가"],
      ["TicketBox", "공연별 좌석·가격·보증금·티켓 상태 관리", "공식 경로만 가능"],
      ["OfficialTransfer", "본인확인·가격상한·수령 동의 확인 후 티켓 이전", "제한적 가능"],
    ]
  )
);
body.push(Caption("표 1. 시스템 구성요소"));
body.push(
  P("FanPassport는 사람당 한 번 발급되는 소울바운드 토큰이다. ERC-5484는 발행 후 양도할 수 없는 SBT 구조를 정의하며, 팬의 참여 이력처럼 특정 개인에게 계속 귀속되어야 하는 기록에 적합하다. 여권은 티켓 구매 시점이 아니라 공연장 검표를 통과한 시점에만 기록을 추가한다. TicketBox는 공연마다 생성되며, 티켓은 일반 지갑 간 전송이나 NFT 마켓 등록이 허용되지 않는다. 다만 OfficialTransfer를 통해 본인확인, 중복 보유 여부, 정가 상한, 양수인의 수령 동의가 모두 확인될 때만 이전할 수 있다."),
  P("리워드는 현금이나 가상자산으로 교환할 수 없고 타인에게 전송할 수도 없는 팬 활동 점수이므로, 되팔아 차익을 얻는 구조와 분리된다. 실제 입장 후에만 확정되어 표를 대량 확보해도 관람하지 않으면 쌓이지 않는다."),
);
body.push(
  table(
    [3000, 1750, 1600, 2700],
    [
      ["상황", "원구매자", "양수인", "조건"],
      ["원구매자가 직접 입장", "300점", "—", "실제 검표 완료"],
      ["공식 양도 후 양수인 입장", "100점", "200점", "양도와 입장 모두 완료"],
      ["공식 양도 후 양수인 노쇼", "0점", "0점", "리워드 미확정"],
      ["공식 반납", "0점", "—", "정가 환불"],
      ["무단 노쇼", "0점", "—", "보증금 몰수"],
    ]
  )
);
body.push(Caption("표 2. 상황별 리워드 확정 규칙"));

body.push(H2(2, "거래·입장 흐름"));
body.push(
  NumItem("이용자는 오프체인 본인확인을 마치고 지갑에 FanPassport를 발급받는다.", "flow"),
  NumItem("TicketBox에서 티켓을 구매하거나 추첨·배정으로 받는다.", "flow"),
  NumItem("관람이 가능하면 검표 단말이 checkIn()을 호출하고, 원구매자에게 300점이 확정된다.", "flow"),
  NumItem("관람이 불가능하면 공식 반납 또는 공식 양도를 선택한다. 공식 양도는 가격 상한과 양수인의 본인확인·동의·중복 보유 여부를 확인한 뒤에만 실행된다.", "flow"),
  NumItem("양수인이 실제 입장하면 원구매자 100점, 양수인 200점이 확정된다. 입장하지 않으면 양도 관련 리워드는 발생하지 않는다.", "flow"),
  NumItem("반납도 입장도 하지 않은 좌석은 공연 후 보증금이 몰수되어 그 공연에 실제로 입장한 팬들에게 분배된다. 반납에는 벌칙이 없고 무단 불참에만 손해가 있으므로 빈 좌석 대신 반납하도록 유도된다.", "flow"),
);
body.push(
  P("확정된 리워드는 굿즈 선구매, 한정 MD 구매 자격, 팬미팅 응모 가중치, 멤버십 등급 혜택에 쓰인다. 예컨대 '최근 24개월 동안 1,000점 이상' 같은 조건을 스마트컨트랙트가 직접 확인할 수 있다."),
);

body.push(H2(3, "비공식 재판매가 차단되는 이유"));
body.push(
  P("표는 일반 지갑으로 전송되지 않는다. 친구에게 직접 보내는 것도 NFT 마켓 등록도 실패한다. 허용되는 경로는 공식 반납과 OfficialTransfer를 거치는 공식 양도뿐이다. 공식 양도는 양수인의 본인확인과 수령 동의, 중복 보유 여부, 정가 상한을 모두 통과해야 실행되므로 웃돈을 붙인 거래는 온체인에서 성립하지 않는다. 양도에 협조해도 원구매자의 리워드는 직접 입장의 3분의 1에 그쳐, 미리 사 두었다가 넘기는 전략은 이득이 되지 않는다. 중요한 것은 구매 기록을 블록체인에 남긴다는 것이 아니라 표 자체가 블록체인 위에 있다는 점이다. 예매처 서버라면 관리자가 소유자 이름 한 줄만 고치면 그만이지만, 이 계약에는 그 한 줄을 임의로 고칠 수 있는 함수가 존재하지 않는다."),
);

body.push(H2(4, "점수가 시간이 지나면 줄어드는 이유"));
body.push(
  P("리워드 점수는 확정 시점부터 730일에 걸쳐 0점으로 줄어든다. 오래전에 다녀온 사람이 영원히 1순위를 차지하는 것을 막기 위해서다. 이때 쌓인 점수만 줄이고 이미 쓴 점수를 명목값으로 두면 차감액이 상대적으로 커져 모든 팬의 점수가 0으로 수렴하므로, 쓴 점수도 같은 속도로 줄여야 한다(구현 중 확인·수정). 한편 누적 참석 횟수는 누구나 조회할 수 있어, 굿즈 판매처가 별도 연동 없이 등급 조건을 직접 확인할 수 있다."),
);

// 제4장 블록체인 도입의 필요성 (채점 배점 항목 — 표는 본문과 중복이라 삭제) ----------
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
  P("중앙 DB의 관리자는 구조적으로 언제나 수정 권한을 갖는다. 발행 회사조차 특정 팬의 이력을 지우거나 구매 내역을 조작할 가능성이 열려 있다. 블록체인은 기록이 해시체인으로 엮여 과거 기록 하나를 바꾸려면 이후 전체를 바꿔야 하므로 사실상 불가능하다. 팬 여권은 예매처가 달라도 실제 입장 이력을 하나의 양도 불가능한 기록으로 연결해, 특정 사업자가 이력을 독점하거나 서비스 종료와 함께 기록이 사라지는 위험을 줄인다."),
);
body.push(H2(3, "배분 결과의 검증가능성"));
body.push(
  P("티켓팅 불신의 뿌리는 암표 자체보다 배분 과정의 불투명성에 있다. 스마트컨트랙트로 규칙을 코드화하면 누구나 이를 읽고, 경매·추첨 결과를 온체인 입력값으로부터 동일하게 재계산해 검증할 수 있다. 중앙 서버는 로그를 공개해도 그 로그의 무결성까지는 증명하지 못한다는 점에서 근본적으로 다르다."),
);
body.push(H2(4, "규칙의 자동 집행"));
body.push(
  P("현행법은 부정 예매와 웃돈 재판매를 모두 입증해야 처벌 가능한 사후 절차이며, 이 때문에 유효조치율이 5.6%에 그친다. 스마트컨트랙트는 전송 제한·반납·보증금 몰수를 조건 충족 시 즉시 실행하고 미충족 시 트랜잭션 자체를 거부해, 암표를 '사후 적발'에서 '사전 차단'으로 전환한다."),
);

// 제5장 기존 시스템과의 비교 및 장단점 (채점 배점 항목) -----------
body.push(H1(5, "기존 시스템과의 비교 및 장단점"));

body.push(H2(1, "장점"));
body.push(
  table(
    [2100, 3400, 3550],
    [
      ["항목", "지금의 티켓팅", "제안하는 시스템"],
      ["티켓 양도", "개인 간 비공식 거래", "본인인증·가격상한을 통과한 공식 양도만"],
      ["암표 대응", "사후 신고·단속 (조치율 약 5.6%)", "일반 전송 차단으로 사전 억제"],
      ["팬의 이력", "예매처별 분절, 탈퇴 시 소멸", "개인 귀속 여권에 누적, 플랫폼과 무관"],
      ["리워드 기준", "구매금액·회원등급", "실제 입장 여부"],
      ["부족분 이득(지대)의 귀속", "암표상", "실제로 참여한 팬"],
    ]
  )
);
body.push(Caption("표 3. 지금의 티켓팅과 제안 시스템의 비교"));
body.push(
  P("암표 대응이 '적발률을 높이는 문제'에서 '웃돈 거래가 온체인에서 성립하지 않는 문제'로 바뀌고, 표 부족분의 이득이 누구에게 가는지가 처음으로 정책 목표로 다뤄진다는 점이 핵심이다."),
);

body.push(H2(2, "단점 및 새로 발생하는 비용"));
body.push(
  Bullet([T("거래 비용. ", { bold: true }), T("계약 실측 결과 이더리움 메인넷 기준(20 gwei, 1 ETH=500만원 가정) 응모 1회 약 8,700~10,500원, 발권 1회 약 12,200원의 수수료가 발생했다. 레이어2 도입으로 이 비용을 낮추는 것이 전제되어야 한다.")]),
  Bullet([T("규칙의 경직성. ", { bold: true }), T("배포 후 코드 수정이 어려워 공연 취소 같은 예외에 즉시 대응하기 어렵고, 업그레이드 구조를 두면 권한 귀속이라는 새 신뢰 문제가 생긴다.")]),
  Bullet([T("지갑 관리 부담. ", { bold: true }), T("비밀키 분실 시 복구가 불가하다. 발급기관 이관 기능으로 완화했으나 이는 중앙화된 신뢰점의 부분적 재도입이다.")]),
  Bullet([T("개발·감사 난이도. ", { bold: true }), T("버그가 자금 손실로 직결되어 배포 전 검증 비용과 기간이 늘어난다.")]),
);

body.push(H2(3, "잔존하는 한계"));
body.push(
  NumItem("본인 확인은 오프체인 실명 확인에 의존하며, 신분 위조를 통한 복수 발급은 막지 못한다."),
  NumItem("공식 양도는 온체인 결제액만 정가 이하로 강제하므로 계좌이체 같은 오프체인 웃돈은 차단하지 못한다. 다만 양수인 본인확인과 리워드 차등(300점 대 100점)으로 유인을 줄인다."),
  NumItem("입구 검표 단말을 신뢰해야 하는 오라클 문제가 남는다."),
  NumItem("지갑 주소로 참석 이력이 노출되는 프라이버시 문제가 있어, 해시 커밋·영지식증명이 후속 과제다."),
  NumItem("데모의 추첨 난수는 조작 가능하므로 실서비스는 체인링크 VRF 등이 필요하다."),
);

// 제6장 결론 -----------------------------------------------------------------
body.push(H1(6, "결론"));
body.push(
  P("암표 방지 실패와 팬 이력의 파편화는 '양도 불가능한 참여 이력을 배분의 화폐로 쓴다'는 하나의 장치로 함께 해소된다. 부족분의 이득을 암표상이 아닌 참여 팬에게 귀속시키는 것이 현실적 목표이며, 스마트컨트랙트로 구현해 22개 항목으로 검증했다(부록 A·B). 프라이버시 보호와 검증가능 난수 생성이 후속 과제다."),
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
  P([T("실제 구현은 FanPassport와 TicketBox 두 계약, 그리고 시연 스크립트로 이루어져 있으며 코드 전문은 첨부 파일로 함께 제출한다. OfficialTransfer는 설계안으로 제시했으며 구현은 후속 과제다. ", {}), T("아래 링크에서 지갑·설치 없이 컨트랙트 로직이 브라우저에서 즉시 실행되는 것을 클릭 몇 번으로 확인할 수 있다: ", {}), Link("https://snu-project.github.io/fan-passport/", "인터랙티브 데모 바로가기"), T(" (소스: ", {}), Link("https://github.com/SNU-Project/fan-passport", "github.com/SNU-Project/fan-passport"), T(")", {})]),
  P([T("FanPassport.sol  ", { bold: true }), T("팔 수 없는 팬 여권을 구현한 계약이다. 한 사람에게 하나씩만 여권을 발급하고(issue), 실제 입장한 시점에만 기록을 남기며(stamp), 시간이 지날수록 선형으로 줄어드는 점수를 계산해 조회할 수 있게 했다(points). 적립된 점수와 이미 써버린 점수를 반드시 같은 속도로 함께 줄여야 하는데, 써버린 점수만 명목값으로 고정해 두면 시간이 지날수록 모든 팬의 점수가 0으로 수렴해 버리는 오류가 생긴다는 점을 구현 중 직접 확인하고 수정했다. 지갑 분실 시에는 발급기관이 본인 확인을 거쳐 기록 전체를 새 지갑으로 옮겨 줄 수 있게 했고(migrate), 전송·승인 등 양도 관련 함수는 모두 명시적으로 실패하도록 만들었다.")]),
  P([T("TicketBox.sol  ", { bold: true }), T("공연 한 건의 발권과 배분을 담당하는 계약이다. 우선권·일반 트랙으로 나누어 응모를 받고(enterDraw) 추첨을 확정하며(draw), 공식 반납 창구가 아니면 어떤 전송도 막는 transferFrom을 두었다. 표를 반납하면 정가와 보증금을 전액 돌려주고(returnTicket), 현장 검표 시 여권에 기록을 남기는 동시에 보증금을 돌려주며(checkIn), 무단 노쇼는 보증금을 몰수하고(closeNoShow), 몰수분은 실제 입장자에게 균등 분배한다(distributeReward).")]),
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
body.push(Caption("표 4. 시연 시나리오 아홉 단계와 그 결과 (총 22개 항목 전부 통과)"));

/* ═══════════════════════ 문서 조립 ═══════════════════════ */

const doc = new Document({
  numbering: {
    config: [
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 640, hanging: 320 } } } }] },
      { reference: "numbered", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 640, hanging: 320 } } } }] },
      { reference: "flow", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 640, hanging: 320 } } } }] },
    ],
  },
  styles: {
    default: {
      document: { run: { font: FONT, size: SIZE } },
      heading1: {
        run: { font: FONT, bold: true, size: 26, color: HCOLOR },
        paragraph: { alignment: AlignmentType.CENTER, spacing: { before: 110, after: 70 } },
      },
      heading2: {
        run: { font: FONT, bold: true, size: 22, color: HCOLOR },
        paragraph: { spacing: { before: 70, after: 40 } },
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
