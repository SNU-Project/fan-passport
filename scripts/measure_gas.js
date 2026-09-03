/**
 * 제안서 6장(장단점)에 쓸 실측 가스비 수집.
 * 데모 시나리오의 핵심 동작들을 한 번씩 실행하고 gasUsed를 표로 뽑는다.
 */
const fs = require("fs");
const path = require("path");
const solc = require("solc");
const ganache = require("ganache");
const { ethers } = require("ethers");

const DIR = path.join(__dirname, "..", "contracts");
const read = (f) => fs.readFileSync(path.join(DIR, f), "utf8");
const input = {
  language: "Solidity",
  sources: { "FanPassport.sol": { content: read("FanPassport.sol") }, "TicketBox.sol": { content: read("TicketBox.sol") } },
  settings: { optimizer: { enabled: true, runs: 200 }, outputSelection: { "*": { "*": ["abi", "evm.bytecode.object"] } } },
};
const out = JSON.parse(solc.compile(JSON.stringify(input)));
const PASSPORT = out.contracts["FanPassport.sol"]["FanPassport"];
const TICKETBOX = out.contracts["TicketBox.sol"]["TicketBox"];

process.removeAllListeners("unhandledRejection");
process.on("unhandledRejection", () => {});

(async () => {
  const provider = new ethers.BrowserProvider(ganache.provider({ logging: { quiet: true }, wallet: { totalAccounts: 8 } }));
  const accts = await provider.listAccounts();
  const S = (i) => provider.getSigner(accts[i].address);
  const issuer = await S(0), promoter = await S(1), gate = await S(2), fan1 = await S(3), fan2 = await S(4);

  const rows = [];
  const record = async (label, txPromise) => {
    const tx = await txPromise;
    const r = await tx.wait();
    rows.push({ label, gas: Number(r.gasUsed) });
  };

  const pf = new ethers.ContractFactory(PASSPORT.abi, PASSPORT.evm.bytecode.object, issuer);
  let rcpt;
  const passportDeployTx = await pf.getDeployTransaction();
  const sentDeploy = await issuer.sendTransaction(passportDeployTx);
  rcpt = await sentDeploy.wait();
  rows.push({ label: "FanPassport 배포", gas: Number(rcpt.gasUsed) });
  const passport = new ethers.Contract(rcpt.contractAddress, PASSPORT.abi, issuer);

  const now = (await provider.getBlock("latest")).timestamp;
  const bf = new ethers.ContractFactory(TICKETBOX.abi, TICKETBOX.evm.bytecode.object, promoter);
  const deployTx = await bf.getDeployTransaction(await passport.getAddress(), 1, now + 30 * 86400, ethers.parseEther("0.1"), ethers.parseEther("0.02"), 4, gate.address);
  const sentBox = await promoter.sendTransaction(deployTx);
  rcpt = await sentBox.wait();
  rows.push({ label: "TicketBox 배포 (공연 1건)", gas: Number(rcpt.gasUsed) });
  const box = new ethers.Contract(rcpt.contractAddress, TICKETBOX.abi, promoter);

  await record("팬 여권 발급 (issue)", passport.connect(issuer).issue(fan1.address));
  await passport.connect(issuer).setAuthorized(await box.getAddress(), true).then(t => t.wait());
  await passport.connect(issuer).issue(fan2.address).then(t => t.wait());

  await record("배분 라운드 열기 (openRound)", box.connect(promoter).openRound(1, 1));
  await record("응모 (enterDraw, 점수 0)", box.connect(fan1).enterDraw(0));
  await record("응모 (enterDraw, 점수 100)", box.connect(fan2).enterDraw(0));
  await record("추첨 확정 (draw)", box.connect(promoter).draw());
  const w0 = await box.winners(0);
  const winner = w0 === fan1.address ? fan1 : fan2;
  await record("당첨자 발권 (claim)", box.connect(winner).claim({ value: ethers.parseEther("0.12") }));
  await record("현장 검표 (checkIn)", box.connect(gate).checkIn(1));

  const total = rows.reduce((a, r) => a + r.gas, 0);
  console.log("\n=== 실측 가스비 (ganache, optimizer runs=200) ===\n");
  for (const r of rows) console.log(`${r.label.padEnd(28, " ")} ${r.gas.toLocaleString()} gas`);
  console.log("-".repeat(45));
  console.log(`${"합계(1개 공연, 팬 1인당 1회전 기준)".padEnd(28," ")} ${total.toLocaleString()} gas`);

  // 참고용 KRW 환산 (이더리움 메인넷 계산 예시 — 실제 배포 시 L2 사용을 전제하므로 참고치일 뿐)
  const gwei = 20; // 가정: 20 gwei
  const ethKrw = 5_000_000; // 가정: 1 ETH = 500만원
  const krw = (gas) => Math.round((gas * gwei * 1e-9) * ethKrw);
  console.log("\n(참고) 가정: 20 gwei, 1 ETH = 500만원 기준 환산 — 실제로는 L2/사이드체인 사용을 전제해 훨씬 저렴함");
  for (const r of rows) console.log(`${r.label.padEnd(28," ")} ≈ ${krw(r.gas).toLocaleString()}원`);
  console.log(`${"합계".padEnd(28," ")} ≈ ${krw(total).toLocaleString()}원`);

  process.exit(0);
})();
