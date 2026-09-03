# 디지털 팬 여권 (Fan Passport)

[![실제 EVM 데모 검증](https://github.com/SNU-Project/fan-passport/actions/workflows/demo.yml/badge.svg)](https://github.com/SNU-Project/fan-passport/actions/workflows/demo.yml)

공연 티켓을 **돈이 아니라, 살 수 없는 참여 기록으로 배분**해서 암표를 사전 차단하는 스마트컨트랙트 프로토타입.

## 🖱️ [지금 바로 눌러보기 → snu-project.github.io/fan-passport](https://snu-project.github.io/fan-passport/)

버튼을 누르면 [`FanPassport.sol`](contracts/FanPassport.sol) · [`TicketBox.sol`](contracts/TicketBox.sol)과
동일한 규칙이 그 자리에서 실행되고, 결과가 실시간으로 바뀐다. 지갑도 설치도 필요 없다.

- 팬 여권을 팔아보기 → `Soulbound`로 거부
- 암표상이 티켓을 되팔기 → `NonTransferable`로 거부
- 응모 → 추첨 → 반납 → 재배정 → 체크인 → 노쇼 → 점수 감가까지 전 과정 체험

---

## 로컬에서 실제 EVM으로 돌려보기

```bash
git clone https://github.com/SNU-Project/fan-passport.git
cd fan-passport && npm install
npm run demo
```

인메모리 이더리움에 두 계약을 실제로 배포해 9단계 시나리오·22개 항목을 검증한다.
같은 검증을 커밋마다 GitHub Actions가 서버에서 재실행한다 (위 배지 클릭 → 전체 로그).

## 더 보기

- [제안서 (.docx)](docs/기말프로젝트_제안서.docx) — 왜 블록체인인가, 기존 시스템과의 비교
- [contracts/](contracts) — FanPassport.sol, TicketBox.sol
