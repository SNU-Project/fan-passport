// SPDX-License-Identifier: MIT
pragma solidity ^0.8.20;

import "./FanPassport.sol";

/**
 * @title TicketBox (공연 1건의 티켓 발권·배분 계약)
 * @notice 공연 하나당 계약 하나가 배포된다. 티켓은 NFT지만 "아무한테나 못 보내는" NFT다.
 *
 *  왜 이런 구조인가 — 제안서의 논증을 코드로:
 *
 *  (1) 초과수요는 없앨 수 없다. 기획사는 매진 자체가 상품 가치의 일부라서
 *      가격을 시장청산 수준까지 올리지 않는다(Becker 1991). 따라서 정가와
 *      실제 가치의 차이(=지대)는 구조적으로 남는다. 문제는 "그 지대를 누가 갖느냐"다.
 *      지금은 암표상이 가져가고, 그걸 노린 매크로 경쟁이 순수한 낭비를 만든다
 *      (Leslie & Sorensen 2014).
 *
 *  (2) 그래서 지불수단을 바꾼다. 좌석 일부를 원화가 아니라 "살 수 없는 화폐"인
 *      팬 점수로 배분한다. 입찰한 점수는 낙찰 여부와 무관하게 소모된다
 *      (올페이 옥션 — 경합·지대추구의 표준 모형, Baye/Kovenock/de Vries 1996).
 *      암표상은 자본이 아무리 많아도 입찰 자체가 불가능하다.
 *
 *  (3) 그러나 점수만으로 배분하면 신규 팬이 영원히 진입하지 못한다.
 *      좌석을 [우선권 트랙 = 점수 경매] + [일반 트랙 = 추첨] 두 갈래로 나눈다.
 *      점수 감가(FanPassport.DECAY_PERIOD)와 함께, 팬덤이 닫히지 않게 하는 장치다.
 *
 *  (4) 양도는 막되 재배분은 살린다. 못 가게 된 사람은 정가 그대로 반납하고,
 *      반납분은 다시 경매로 돌아간다. 티켓이 낮은 가치의 사람에서 높은 가치의
 *      사람으로 옮겨가는 재판매의 순기능은 유지하면서, 웃돈만 제거한다.
 *
 *  (5) 노쇼 보증금. 예매 시 소액을 예치하고 실제 입장하면 전액 돌려준다.
 *      반납에는 페널티가 없고 무단 노쇼만 몰수된다 — 빈 좌석을 만드느니
 *      반납하게 만드는 인센티브 설계다.
 */
contract TicketBox {
    /* ─────────────────────────── 공연 정보 ─────────────────────────── */

    FanPassport public immutable passport;
    address public immutable promoter;   // 기획사
    address public immutable gate;       // 현장 입장 검표 단말
    uint64  public immutable showId;
    uint64  public immutable showTime;

    uint256 public immutable faceValue;  // 정가 (반납 시 전액 환불)
    uint256 public immutable deposit;    // 노쇼 보증금 (입장 시 전액 환급)
    uint256 public immutable capacity;   // 총 좌석 수

    /// @notice 몰수된 보증금이 쌓이는 곳. 다음 회차 팬 리워드 재원으로 쓴다.
    uint256 public rewardPool;

    /* ─────────────────────────── 티켓 (제한된 NFT) ─────────────────────────── */

    enum TicketState {
        None,       // 존재하지 않음
        Held,       // 정상 보유 중
        Returned,   // 공식 반납됨 → 재배분 대기
        Used,       // 입장 완료
        Forfeited   // 무단 노쇼 (보증금 몰수)
    }

    mapping(uint256 => address) public ownerOf;
    mapping(address => uint256) public balanceOf;
    mapping(uint256 => TicketState) public stateOf;

    uint256 public minted;                 // 지금까지 발행된 티켓 수 (<= capacity)
    uint256[] public returnPool;           // 반납되어 재배분을 기다리는 티켓들

    event Transfer(address indexed from, address indexed to, uint256 indexed tokenId);
    event Returned(uint256 indexed tokenId, address indexed from, uint256 refund);
    event CheckedIn(uint256 indexed tokenId, address indexed fan);
    event NoShow(uint256 indexed tokenId, address indexed fan, uint256 forfeited);

    /* ─────────────────────────── 배분 라운드 ─────────────────────────── */

    struct Entry {
        address fan;
        uint256 pointsBid;   // 0이면 추첨 트랙만 응모
    }

    uint256 public round;                  // 1차 배정 = 1, 반납분 재배정 = 2, ...
    uint256 public prioritySeats;          // 이번 라운드의 점수 경매 좌석
    uint256 public lotterySeats;           // 이번 라운드의 추첨 좌석
    bool    public entryOpen;
    bool    public drawn;

    Entry[] public entries;
    mapping(uint256 => mapping(address => bool)) public entered;   // round => fan => 응모함
    mapping(uint256 => mapping(address => bool)) public isWinner;  // round => fan => 당첨
    address[] public winners;

    event RoundOpened(uint256 indexed round, uint256 prioritySeats, uint256 lotterySeats);
    event Entered(uint256 indexed round, address indexed fan, uint256 pointsBid);
    event Won(uint256 indexed round, address indexed fan, bool viaPriority);
    event Claimed(uint256 indexed round, address indexed fan, uint256 tokenId);

    /* ─────────────────────────── 오류 ─────────────────────────── */

    /// @dev 암표 차단의 핵심. 공식 반납 외의 모든 소유권 이전이 여기서 막힌다.
    error NonTransferable();
    error NotPromoter();
    error NotGate();
    error NotOwner();
    error NoPassport();
    error EntryClosed();
    error AlreadyEntered();
    error AlreadyHasTicket();
    error NotDrawnYet();
    error AlreadyDrawn();
    error NotWinner();
    error WrongPayment(uint256 sent, uint256 need);
    error SoldOut();
    error TooLate();
    error TooEarly();
    error BadState();

    modifier onlyPromoter() {
        if (msg.sender != promoter) revert NotPromoter();
        _;
    }

    constructor(
        FanPassport _passport,
        uint64  _showId,
        uint64  _showTime,
        uint256 _faceValue,
        uint256 _deposit,
        uint256 _capacity,
        address _gate
    ) {
        passport  = _passport;
        promoter  = msg.sender;
        gate      = _gate;
        showId    = _showId;
        showTime  = _showTime;
        faceValue = _faceValue;
        deposit   = _deposit;
        capacity  = _capacity;
    }

    /* ═══════════════════════ 1. 배분 라운드 ═══════════════════════ */

    /// @notice 좌석을 두 트랙으로 나눠 응모를 연다.
    function openRound(uint256 _prioritySeats, uint256 _lotterySeats) external onlyPromoter {
        if (entryOpen) revert AlreadyDrawn();
        if (minted + _prioritySeats + _lotterySeats > capacity + returnPool.length) revert SoldOut();

        round += 1;
        prioritySeats = _prioritySeats;
        lotterySeats  = _lotterySeats;
        entryOpen = true;
        drawn = false;
        delete entries;
        delete winners;

        emit RoundOpened(round, _prioritySeats, _lotterySeats);
    }

    /// @notice 응모한다. pointsBid > 0 이면 우선권 트랙에도 함께 참여한다.
    /// @dev    올페이 — 입찰 점수는 이 시점에 소모된다. 떨어져도 돌아오지 않는다.
    ///         "간절함"에 비용을 붙여야 무의미한 전원 최대입찰을 막을 수 있다.
    function enterDraw(uint256 pointsBid) external {
        if (!entryOpen) revert EntryClosed();
        if (!passport.hasPassport(msg.sender)) revert NoPassport();
        if (entered[round][msg.sender]) revert AlreadyEntered();
        if (balanceOf[msg.sender] > 0) revert AlreadyHasTicket();

        entered[round][msg.sender] = true;

        if (pointsBid > 0) {
            passport.spend(msg.sender, pointsBid);   // 점수 부족하면 여기서 revert
        }

        entries.push(Entry({fan: msg.sender, pointsBid: pointsBid}));
        emit Entered(round, msg.sender, pointsBid);
    }

    /// @notice 응모를 마감하고 배분을 확정한다.
    function draw() external onlyPromoter {
        if (!entryOpen) revert EntryClosed();
        if (drawn) revert AlreadyDrawn();

        entryOpen = false;
        drawn = true;

        uint256 n = entries.length;
        if (n == 0) return;

        bool[] memory taken = new bool[](n);

        // ── (a) 우선권 트랙: 점수 상위 순. 점수 0인 응모자는 여기서 뽑히지 않는다.
        for (uint256 s = 0; s < prioritySeats; s++) {
            uint256 best = type(uint256).max;
            uint256 bestPts = 0;
            for (uint256 i = 0; i < n; i++) {
                if (taken[i]) continue;
                if (entries[i].pointsBid > bestPts) {
                    bestPts = entries[i].pointsBid;
                    best = i;
                }
            }
            if (best == type(uint256).max) break;   // 남은 유효 입찰 없음
            taken[best] = true;
            _addWinner(entries[best].fan, true);
        }

        // ── (b) 일반 트랙: 남은 응모자 전원 대상 무작위 추첨.
        //        데모용 난수다. 실서비스는 체인링크 VRF나 커밋-리빌이 필요하다 — README 참조.
        for (uint256 s = 0; s < lotterySeats; s++) {
            uint256 remaining = 0;
            for (uint256 i = 0; i < n; i++) {
                if (!taken[i]) remaining++;
            }
            if (remaining == 0) break;

            uint256 r = uint256(
                keccak256(abi.encodePacked(blockhash(block.number - 1), round, s, remaining))
            ) % remaining;

            for (uint256 i = 0; i < n; i++) {
                if (taken[i]) continue;
                if (r == 0) {
                    taken[i] = true;
                    _addWinner(entries[i].fan, false);
                    break;
                }
                r--;
            }
        }
    }

    function _addWinner(address fan, bool viaPriority) private {
        isWinner[round][fan] = true;
        winners.push(fan);
        emit Won(round, fan, viaPriority);
    }

    /// @notice 당첨자가 정가 + 보증금을 내고 실제 티켓을 받아간다.
    function claim() external payable returns (uint256 tokenId) {
        if (!drawn) revert NotDrawnYet();
        if (!isWinner[round][msg.sender]) revert NotWinner();
        if (msg.value != faceValue + deposit) revert WrongPayment(msg.value, faceValue + deposit);

        isWinner[round][msg.sender] = false;   // 1회만
        tokenId = _issue(msg.sender);
        emit Claimed(round, msg.sender, tokenId);
    }

    /// @dev 반납된 티켓이 있으면 그것부터 재사용하고, 없으면 새로 발행한다.
    function _issue(address to) private returns (uint256 tokenId) {
        if (returnPool.length > 0) {
            tokenId = returnPool[returnPool.length - 1];
            returnPool.pop();
        } else {
            if (minted >= capacity) revert SoldOut();
            minted += 1;
            tokenId = minted;
        }

        ownerOf[tokenId] = to;
        balanceOf[to] += 1;
        stateOf[tokenId] = TicketState.Held;
        emit Transfer(address(0), to, tokenId);
    }

    /* ═══════════════════════ 2. 양도 차단 ═══════════════════════ */

    /// @notice ERC-721 형태를 갖추되, 공식 반납(to == 이 계약) 외에는 전부 거부한다.
    /// @dev    이것이 암표 차단의 전부다. 웃돈을 현금으로 아무리 주고받아도
    ///         소유권이 넘어가지 않으므로, 구매자는 입장할 수 없다(checkIn 참조).
    ///         "적발해서 취소"가 아니라 "애초에 불가능"으로 바뀐다.
    function transferFrom(address from, address to, uint256 tokenId) public {
        if (to != address(this)) revert NonTransferable();
        if (msg.sender != from || ownerOf[tokenId] != from) revert NotOwner();
        _returnTicket(tokenId);
    }

    function safeTransferFrom(address from, address to, uint256 tokenId) external {
        transferFrom(from, to, tokenId);
    }

    function approve(address, uint256) external pure {
        revert NonTransferable();
    }

    function setApprovalForAll(address, bool) external pure {
        revert NonTransferable();
    }

    /* ═══════════════════════ 3. 반납 ═══════════════════════ */

    /// @notice 못 가게 됐을 때 정가와 보증금을 전액 돌려받고 좌석을 반환한다.
    function returnTicket(uint256 tokenId) external {
        if (ownerOf[tokenId] != msg.sender) revert NotOwner();
        _returnTicket(tokenId);
    }

    function _returnTicket(uint256 tokenId) private {
        if (stateOf[tokenId] != TicketState.Held) revert BadState();
        if (block.timestamp >= showTime) revert TooLate();

        address holder = ownerOf[tokenId];

        // 상태를 먼저 바꾸고 송금한다 (재진입 방지).
        stateOf[tokenId] = TicketState.Returned;
        ownerOf[tokenId] = address(0);
        balanceOf[holder] -= 1;
        returnPool.push(tokenId);

        emit Transfer(holder, address(0), tokenId);

        uint256 refund = faceValue + deposit;
        (bool ok, ) = holder.call{value: refund}("");
        require(ok, "refund failed");

        emit Returned(tokenId, holder, refund);
    }

    /* ═══════════════════════ 4. 입장과 보증금 ═══════════════════════ */

    /// @notice 현장 검표. 이 순간에만 팬 여권에 스탬프가 찍힌다.
    /// @dev    예매가 아니라 입장이 기록의 기준이다. 사재기로는 점수가 쌓이지 않는다.
    function checkIn(uint256 tokenId) external {
        if (msg.sender != gate) revert NotGate();
        if (stateOf[tokenId] != TicketState.Held) revert BadState();

        address holder = ownerOf[tokenId];
        stateOf[tokenId] = TicketState.Used;

        passport.stamp(holder, showId);

        (bool ok, ) = holder.call{value: deposit}("");
        require(ok, "deposit refund failed");

        emit CheckedIn(tokenId, holder);
    }

    /// @notice 공연이 끝난 뒤, 반납도 입장도 하지 않은 좌석의 보증금을 몰수한다.
    function closeNoShow(uint256 tokenId) external onlyPromoter {
        if (block.timestamp < showTime) revert TooEarly();
        if (stateOf[tokenId] != TicketState.Held) revert BadState();

        address holder = ownerOf[tokenId];
        stateOf[tokenId] = TicketState.Forfeited;
        rewardPool += deposit;

        emit NoShow(tokenId, holder, deposit);
    }

    /// @notice 정가 대금을 기획사에 정산한다.
    function settleRevenue() external onlyPromoter {
        uint256 locked = _lockedDeposits();
        uint256 amount = address(this).balance - locked - rewardPool;
        (bool ok, ) = promoter.call{value: amount}("");
        require(ok, "settle failed");
    }

    function _lockedDeposits() private view returns (uint256 total) {
        for (uint256 i = 1; i <= minted; i++) {
            if (stateOf[i] == TicketState.Held) total += deposit;
        }
    }

    /* ═══════════════════════ 조회용 ═══════════════════════ */

    function entriesCount()   external view returns (uint256) { return entries.length; }
    function winnersCount()   external view returns (uint256) { return winners.length; }
    function returnPoolSize() external view returns (uint256) { return returnPool.length; }
}
