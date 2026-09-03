/**
 * 디지털 팬 여권 — 시연 스크립트
 *
 * 인메모리 이더리움(ganache) 위에 계약을 실제로 배포하고, 암표상이 어디서
 * 막히는지를 순서대로 보여준다. 외부 네트워크도 지갑도 필요 없다.
 *
 *   npm install && npm run demo
 */

const fs = require('fs');
const path = require('path');
const solc = require('solc');
const ganache = require('ganache');
const { ethers } = require('ethers');

/* ─────────────────────────── 컴파일 ─────────────────────────── */

const DIR = path.join(__dirname, '..', 'contracts');
const read = (f) => fs.readFileSync(path.join(DIR, f), 'utf8');

const input = {
  language: 'Solidity',
  sources: {
    'FanPassport.sol': { content: read('FanPassport.sol') },
    'TicketBox.sol':   { content: read('TicketBox.sol') },
  },
  settings: {
    optimizer: { enabled: true, runs: 200 },
    outputSelection: { '*': { '*': ['abi', 'evm.bytecode.object'] } },
  },
};

const out = JSON.parse(solc.compile(JSON.stringify(input)));
const errs = (out.errors || []).filter((e) => e.severity === 'error');
if (errs.length) {
  errs.forEach((e) => console.error(e.formattedMessage));
  process.exit(1);
}
const PASSPORT = out.contracts['FanPassport.sol']['FanPassport'];
const TICKETBOX = out.contracts['TicketBox.sol']['TicketBox'];

/* ─────────────────────────── 출력 도우미 ─────────────────────────── */

process.removeAllListeners('unhandledRejection');
process.on('unhandledRejection', () => {});

/** revert 사유를 사람이 읽을 수 있는 이름으로. 계약이 여러 개면 ABI도 여러 개 뒤져야 한다. */
const IFACES = [];
function reasonOf(e) {
  if (e?.revert?.name) return e.revert.name;

  // TicketBox 호출이 내부의 FanPassport에서 실패하면, TicketBox의 ABI만으로는 해석되지 않는다.
  const data = e?.data ?? e?.info?.error?.data?.data ?? e?.info?.error?.data;
  if (typeof data === 'string' && data.startsWith('0x') && data.length >= 10) {
    for (const iface of IFACES) {
      try {
        const parsed = iface.parseError(data);
        if (parsed) return parsed.name;
      } catch (_) { /* 다음 ABI */ }
    }
  }
  return e?.info?.error?.data?.reason ?? e?.shortMessage ?? e?.reason ?? e?.message ?? '';
}

let pass = 0, fail = 0;

const act = (n, title) => {
  console.log('\n' + '━'.repeat(64));
  console.log(`  ${n}.  ${title}`);
  console.log('━'.repeat(64));
};
const say = (s) => console.log('     ' + s);
const ok = (m) => { pass++; console.log('  ✓  ' + m); };
const no = (m) => { fail++; console.log('  ✗  ' + m); };

/** 반드시 실패해야 하는 트랜잭션. 막히는 것이 곧 기능이다. */
async function mustRevert(promise, expect, label) {
  try {
    await promise;
    no(`${label} — 막히지 않음! (설계 오류)`);
  } catch (e) {
    const r = reasonOf(e);
    if (!expect || r.includes(expect)) ok(`${label}  → 거부됨 [${r}]`);
    else no(`${label} — 다른 이유로 거부: ${r}`);
  }
}

const ETH = (n) => ethers.parseEther(String(n));
const fmt = (w) => Number(ethers.formatEther(w)).toFixed(3);

/* ─────────────────────────── 시연 ─────────────────────────── */

(async () => {
  const provider = new ethers.BrowserProvider(
    ganache.provider({ logging: { quiet: true }, wallet: { totalAccounts: 12 } })
  );
  const accts = await provider.listAccounts();
  const S = (i) => provider.getSigner(accts[i].address);

  const issuer   = await S(0);   // 발권사 컨소시엄 (오프체인 본인확인 담당)
  const promoter = await S(1);   // 기획사
  const gate     = await S(2);   // 공연장 입장 검표 단말
  const jimin    = await S(3);   // 팬 A — 5년째 따라다닌 고인물
  const suhyun   = await S(4);   // 팬 B — 작년부터 2회
  const haneul   = await S(5);   // 팬 C — 이번이 처음
  const yerin    = await S(6);   // 팬 D — 3회, 1차 배정을 놓침
  const scalper  = await S(7);   // 암표상
  const alt      = await S(8);   // 암표상 부계정 (본인확인 미통과)
  const buyerX   = await S(9);   // 웃돈 주고라도 사려는 사람

  const who = {
    [jimin.address]:   '지민(5회)',
    [suhyun.address]:  '수현(2회)',
    [haneul.address]:  '하늘(신규)',
    [yerin.address]:   '예린(3회)',
    [scalper.address]: '암표상',
    [buyerX.address]:  '구매자X',
  };
  const nm = (a) => who[a] ?? a.slice(0, 8);
  /** 한글은 2칸을 차지하므로 표시폭 기준으로 채운다. */
  const pad = (s, w) => {
    let width = 0;
    for (const ch of s) width += /[\u1100-\u11FF\u3130-\u318F\uAC00-\uD7AF]/.test(ch) ? 2 : 1;
    return s + ' '.repeat(Math.max(0, w - width));
  };

  /* ── 배포 ─────────────────────────────────────────────── */

  act('0', '계약 배포 — 여권 발급기관과 공연 1건');

  const pf = new ethers.ContractFactory(PASSPORT.abi, PASSPORT.evm.bytecode.object, issuer);
  const passport = await pf.deploy();
  await passport.waitForDeployment();
  say('FanPassport 배포  ' + (await passport.getAddress()));

  const now = (await provider.getBlock('latest')).timestamp;
  const SHOW_TIME = now + 30 * 24 * 3600;
  const FACE = ETH(0.1);      // 정가 (11만원 상당이라고 치자)
  const DEP  = ETH(0.02);     // 노쇼 보증금
  const CAP  = 4;             // 좌석 4석 (시연용 소극장)

  const bf = new ethers.ContractFactory(TICKETBOX.abi, TICKETBOX.evm.bytecode.object, promoter);
  const box = await bf.deploy(
    await passport.getAddress(), 1, SHOW_TIME, FACE, DEP, CAP, gate.address
  );
  await box.waitForDeployment();
  say('TicketBox 배포    ' + (await box.getAddress()));
  say(`좌석 ${CAP}석 · 정가 ${fmt(FACE)}ETH · 보증금 ${fmt(DEP)}ETH`);

  await (await passport.setAuthorized(await box.getAddress(), true)).wait();
  IFACES.push(passport.interface, box.interface);

  /* ── 여권 발급 + 과거 이력 ──────────────────────────────── */

  act('1', '팬 여권 발급 — 본인확인을 통과한 사람만');

  for (const f of [jimin, suhyun, haneul, yerin, scalper, buyerX]) {
    await (await passport.issue(f.address)).wait();
  }
  ok(`여권 6개 발급 (암표상도 실명확인은 통과했으므로 발급된다)`);
  say('부계정(alt)은 본인확인을 통과하지 못해 여권이 없다.');

  // 과거 공연 이력을 심는다. 실제로는 지난 공연들의 TicketBox가 찍은 스탬프다.
  await (await passport.setAuthorized(accts[0].address, true)).wait();
  const seed = async (f, k) => {
    for (let i = 0; i < k; i++) await (await passport.stamp(f.address, 900 + i)).wait();
  };
  await seed(jimin, 5);
  await seed(suhyun, 2);
  await seed(yerin, 3);
  await (await passport.setAuthorized(accts[0].address, false)).wait();

  console.log();
  for (const f of [jimin, suhyun, yerin, haneul, scalper]) {
    say(`${pad(nm(f.address), 14)} 참석 ${await passport.attendanceCount(f.address)}회 · 점수 ${await passport.points(f.address)}`);
  }
  say('암표상의 점수는 0이다. 그리고 이 점수는 어디서도 살 수 없다.');

  /* ── 여권은 팔 수 없다 ──────────────────────────────────── */

  act('2', '“팬 이력을 통째로 사겠다” — 여권 자체의 거래 시도');

  await mustRevert(
    passport.connect(jimin).transferFrom.staticCall(jimin.address, scalper.address, 0),
    'Soulbound', '암표상이 지민의 여권을 사려 함'
  );
  await mustRevert(
    passport.connect(jimin).approve.staticCall(scalper.address, 0),
    'Soulbound', '거래소에 여권 매물 등록 시도'
  );
  say('이력이 시장에 나올 수 없으므로, 팬 점수는 돈으로 환산되지 않는다.');

  /* ── 1차 배정 ──────────────────────────────────────────── */

  act('3', '1차 배정 — 우선권 2석(점수 경매) + 일반 2석(추첨)');

  await (await box.connect(promoter).openRound(2, 2)).wait();
  say('좌석을 두 갈래로 나눈다. 점수만으로 나누면 신규 팬이 영원히 못 들어온다.');
  console.log();

  await (await box.connect(jimin).enterDraw(300)).wait();  ok('지민   300점 입찰 (보유 500)');
  await (await box.connect(suhyun).enterDraw(150)).wait(); ok('수현   150점 입찰 (보유 200)');
  await (await box.connect(haneul).enterDraw(0)).wait();   ok('하늘   0점 — 추첨 트랙만 응모');
  await (await box.connect(scalper).enterDraw(0)).wait();  ok('암표상 0점 — 추첨 트랙만 응모');

  console.log();
  await mustRevert(
    box.connect(alt).enterDraw.staticCall(0),
    'NoPassport', '암표상이 부계정으로 중복 응모'
  );
  await mustRevert(
    box.connect(jimin).enterDraw.staticCall(100),
    'AlreadyEntered', '지민이 한 번 더 응모'
  );

  console.log();
  say('입찰한 점수는 이 시점에 이미 소모됐다(올페이). 떨어져도 돌아오지 않는다.');
  say(`지민 잔여점수 ${await passport.points(jimin.address)} · 수현 잔여점수 ${await passport.points(suhyun.address)}`);

  await (await box.connect(promoter).draw()).wait();
  console.log();
  const w1 = [];
  for (let i = 0; i < (await box.winnersCount()); i++) w1.push(await box.winners(i));
  say('당첨: ' + w1.map(nm).join(', '));
  say('앞 2명은 점수 순, 뒤 2명은 추첨이다. 암표상도 추첨으로는 들어올 수 있다.');

  for (const f of [jimin, suhyun, haneul, scalper]) {
    await (await box.connect(f).claim({ value: FACE + DEP })).wait();
  }
  console.log();
  ok(`4명 모두 정가 ${fmt(FACE)} + 보증금 ${fmt(DEP)} 지불 후 티켓 수령`);
  for (let id = 1; id <= 4; id++) say(`  #${id} → ${nm(await box.ownerOf(id))}`);

  /* ── 암표 시도 ─────────────────────────────────────────── */

  act('4', '암표 시도 — 여기가 이 프로젝트의 핵심');

  let scalperTicket;
  for (let id = 1; id <= 4; id++) {
    if ((await box.ownerOf(id)) === scalper.address) scalperTicket = id;
  }
  say(`암표상이 보유한 티켓: #${scalperTicket}`);
  say('구매자X가 정가의 5배를 현금으로 주기로 했다. 이제 넘기기만 하면 된다.');
  console.log();

  await mustRevert(
    box.connect(scalper).transferFrom.staticCall(scalper.address, buyerX.address, scalperTicket),
    'NonTransferable', '구매자X에게 직접 전송'
  );
  await mustRevert(
    box.connect(scalper).approve.staticCall(buyerX.address, scalperTicket),
    'NonTransferable', 'NFT 마켓플레이스에 매물 등록'
  );
  await mustRevert(
    box.connect(scalper).safeTransferFrom.staticCall(scalper.address, buyerX.address, scalperTicket),
    'NonTransferable', '중개 계약을 통한 우회 전송'
  );

  console.log();
  say('현금은 이미 오갔다고 치자. 그래도 공연장에서는:');
  say(`  티켓 #${scalperTicket}의 온체인 소유자 = ${nm(await box.ownerOf(scalperTicket))}`);
  say(`  입장 대기자 = 구매자X`);
  say('  → 신분증과 소유자가 불일치. 입장 불가.');
  say('사후 적발이 아니라 사전 불가능이다. 단속률이 아니라 구조의 문제로 바뀐다.');

  /* ── 공식 반납 ─────────────────────────────────────────── */

  act('5', '공식 반납 — 재판매의 순기능만 남기기');

  const rcpt = await (await box.connect(scalper).returnTicket(scalperTicket)).wait();
  const before = await provider.getBalance(scalper.address, rcpt.blockNumber - 1);
  const after  = await provider.getBalance(scalper.address, rcpt.blockNumber);
  const gas = rcpt.gasUsed * rcpt.gasPrice;
  ok(`암표상이 티켓 #${scalperTicket} 반납 → 정가+보증금 전액 환불 (+${fmt(after - before + gas)}ETH)`);
  say('차익이 0이므로 애초에 사재기할 이유가 없다.');
  say(`반납 대기 좌석: ${await box.returnPoolSize()}석`);

  /* ── 반납분 재배분 ─────────────────────────────────────── */

  act('6', '반납분 재배분 — 돈이 아니라 점수로 경쟁한다');

  await (await box.connect(promoter).openRound(1, 0)).wait();
  say('반납된 1석을 우선권 트랙으로 되돌린다.');
  console.log();

  await (await box.connect(yerin).enterDraw(200)).wait();
  ok('예린   200점 입찰 (3회 참석, 1차 배정을 놓쳤다)');
  await (await box.connect(buyerX).enterDraw(0)).wait();
  ok('구매자X 0점 — 입찰할 점수가 없다');

  console.log();
  await mustRevert(
    box.connect(scalper).enterDraw.staticCall(50),
    'InsufficientPoints', '암표상이 50점을 입찰 (보유 0점)'
  );
  say('여기가 결정적이다. 암표상은 자본이 아무리 많아도');
  say('이 경매의 지불수단을 구할 방법이 없다. 공연에 직접 가는 것 말고는.');

  await (await box.connect(promoter).draw()).wait();
  console.log();
  say('낙찰: ' + nm(await box.winners(0)));
  await (await box.connect(yerin).claim({ value: FACE + DEP })).wait();
  ok(`예린이 정가 그대로 티켓 수령 (웃돈 0원)`);

  /* ── 입장과 보증금 ─────────────────────────────────────── */

  act('7', '공연 당일 — 입장해야만 기록이 쌓인다');

  for (const f of [jimin, suhyun, yerin]) {
    let id;
    for (let i = 1; i <= 4; i++) if ((await box.ownerOf(i)) === f.address) id = i;
    await (await box.connect(gate).checkIn(id)).wait();
    ok(`${nm(f.address)} 입장 → 보증금 ${fmt(DEP)} 환급 + 여권 스탬프`);
  }
  say('하늘은 오지 않았다 (무단 노쇼).');

  await provider.send('evm_increaseTime', [31 * 24 * 3600]);
  await provider.send('evm_mine', []);

  let haneulId;
  for (let i = 1; i <= 4; i++) if ((await box.ownerOf(i)) === haneul.address) haneulId = i;
  await (await box.connect(promoter).closeNoShow(haneulId)).wait();
  console.log();
  ok(`하늘의 보증금 ${fmt(DEP)} 몰수 → 팬 리워드 풀 (현재 ${fmt(await box.rewardPool())}ETH)`);
  say('반납에는 페널티가 없고 무단 노쇼만 손해다 → 빈 좌석 대신 반납을 유도한다.');

  /* ── 감가 ─────────────────────────────────────────────── */

  act('8', '점수 감가 — 팬덤이 닫히지 않게');

  say(`현재  지민 ${await passport.points(jimin.address)}점 · 예린 ${await passport.points(yerin.address)}점`);
  await provider.send('evm_increaseTime', [365 * 24 * 3600]);
  await provider.send('evm_mine', []);
  say(`1년 뒤 지민 ${await passport.points(jimin.address)}점 · 예린 ${await passport.points(yerin.address)}점`);
  say('오래된 출석은 가치가 줄어든다. 계속 오는 사람이 유리하되, 영구 독점은 없다.');

  /* ── 최종 ─────────────────────────────────────────────── */

  act('9', '최종 상태');

  const label = ['', '보유중', '반납됨', '입장완료', '노쇼몰수'];
  for (let id = 1; id <= 4; id++) {
    const st = Number(await box.stateOf(id));
    const o = await box.ownerOf(id);
    say(`티켓 #${id}  ${pad(label[st], 10)}  ${o === ethers.ZeroAddress ? '-' : nm(o)}`);
  }
  console.log();
  for (const f of [jimin, suhyun, haneul, yerin, scalper]) {
    say(`${pad(nm(f.address), 14)} 참석 ${await passport.attendanceCount(f.address)}회 · 점수 ${await passport.points(f.address)}`);
  }
  console.log();
  say(`암표상의 최종 참석 기록: ${await passport.attendanceCount(scalper.address)}회 — 티켓을 샀지만 아무것도 남지 않았다.`);

  console.log('\n' + '━'.repeat(64));
  console.log(`  통과 ${pass} · 실패 ${fail}`);
  console.log('━'.repeat(64) + '\n');
  process.exit(fail ? 1 : 0);
})();
