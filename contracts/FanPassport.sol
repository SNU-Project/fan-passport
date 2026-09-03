// SPDX-License-Identifier: MIT
pragma solidity ^0.8.20;

/**
 * @title FanPassport (디지털 팬 여권)
 * @notice 팬이 "실제로 공연에 갔다"는 사실을 본인 지갑에 쌓아두는 비양도성 증명서.
 *
 *  설계 원칙 3가지 — 제안서의 주장을 코드로 옮긴 것:
 *
 *  (1) 소울바운드(soulbound). 여권과 그 안의 출석 기록은 어떤 방법으로도 남에게
 *      넘길 수 없다. transferFrom 계열 함수는 존재하되 항상 revert 한다.
 *      "팔 수 없는 자산"이어야만 돈으로 살 수 없고, 돈으로 살 수 없어야만
 *      암표상이 흉내낼 수 없다. (Weyl, Ohlhaver, Buterin 2022)
 *
 *  (2) 감가(decay). 출석 점수는 시간이 지나면 선형으로 줄어들어 DECAY_PERIOD 후
 *      0이 된다. 20년 전 한 번 간 사람이 영원히 1순위가 되는 것을 막고,
 *      신규 팬이 따라잡을 수 있는 여지를 남긴다.
 *
 *  (3) 지갑 분실 대응. 소울바운드는 "영원히 못 옮긴다"가 아니라
 *      "본인 의사로도 시장에 팔 수 없다"는 뜻이다. 발급기관은 본인확인을 거쳐
 *      기록 전체를 새 주소로 이관할 수 있다(ERC-5484의 burn/재발행 모델).
 *      이관은 기록을 쪼개거나 합칠 수 없으므로 거래 수단이 되지 못한다.
 */
contract FanPassport {
    string public constant name = "Fan Passport";
    string public constant symbol = "FANP";

    /// @notice 발급기관(예: 공연예술통합전산망·티켓 발권사 컨소시엄). 오프체인 본인확인 담당.
    address public issuer;

    /* ─────────────────────────── 출석 기록 ─────────────────────────── */

    struct Stamp {
        uint64 showId;   // 어떤 공연이었는지
        uint64 time;     // 언제 입장했는지 (감가 계산의 기준점)
    }

    /// @notice 우선권 경매에서 소모한 점수. 적립분과 "동일한 속도로" 감가되어야 한다.
    /// @dev    적립만 감가시키고 소모는 명목값으로 남겨두면, 시간이 지날수록
    ///         차감액이 상대적으로 커져 모든 팬의 점수가 0으로 수렴한다.
    ///         적립과 소모를 같은 시간축 위에 두는 것이 이 설계의 전제다.
    struct Debit {
        uint128 amount;
        uint64  time;
    }

    mapping(address => Stamp[]) private _stamps;
    mapping(address => Debit[]) private _debits;
    mapping(address => bool) public hasPassport;

    /// @notice 스탬프를 찍고 점수를 차감할 수 있는 주체(= 각 공연의 TicketBox 계약).
    mapping(address => bool) public authorized;

    /// @notice 공연 1회 참석당 부여되는 점수.
    uint256 public constant POINTS_PER_ATTENDANCE = 100;

    /// @notice 이 기간에 걸쳐 점수가 100% → 0% 로 선형 감소한다.
    uint256 public constant DECAY_PERIOD = 730 days; // 2년

    event PassportIssued(address indexed fan);
    event Stamped(address indexed fan, uint64 indexed showId, uint64 time);
    event PointsSpent(address indexed fan, uint256 amount);
    event PassportMigrated(address indexed from, address indexed to, uint256 stampCount);

    error Soulbound();
    error NotIssuer();
    error NotAuthorized();
    error AlreadyHasPassport();
    error NoPassport();
    error InsufficientPoints(uint256 has, uint256 need);

    modifier onlyIssuer() {
        if (msg.sender != issuer) revert NotIssuer();
        _;
    }

    constructor() {
        issuer = msg.sender;
    }

    /* ─────────────────────────── 발급 ─────────────────────────── */

    /// @notice 오프체인 실명확인을 마친 사람에게 여권을 발급한다. 1인 1여권.
    /// @dev    "1인 1여권"의 진짜 근거는 체인이 아니라 본인확인이다.
    ///         체인은 그 결과를 위조 불가능하게 기록할 뿐이라는 점을 제안서에 명시할 것.
    function issue(address fan) external onlyIssuer {
        if (hasPassport[fan]) revert AlreadyHasPassport();
        hasPassport[fan] = true;
        emit PassportIssued(fan);
    }

    /// @notice 지갑을 분실했을 때 기록 전체를 새 주소로 옮긴다. 쪼개거나 합칠 수 없다.
    function migrate(address from, address to) external onlyIssuer {
        if (!hasPassport[from]) revert NoPassport();
        if (hasPassport[to]) revert AlreadyHasPassport();

        Stamp[] storage src = _stamps[from];
        uint256 n = src.length;
        for (uint256 i = 0; i < n; i++) {
            _stamps[to].push(src[i]);
        }
        Debit[] storage d = _debits[from];
        for (uint256 i = 0; i < d.length; i++) {
            _debits[to].push(d[i]);
        }

        delete _stamps[from];
        delete _debits[from];
        hasPassport[from] = false;
        hasPassport[to] = true;

        emit PassportMigrated(from, to, n);
    }

    /* ─────────────────────────── 권한 ─────────────────────────── */

    function setAuthorized(address box, bool ok) external onlyIssuer {
        authorized[box] = ok;
    }

    /* ─────────────────────────── 기록·차감 ─────────────────────────── */

    /// @notice 실제 입장이 확인된 순간 호출된다. 예매가 아니라 "입장"만 점수가 된다.
    function stamp(address fan, uint64 showId) external {
        if (!authorized[msg.sender]) revert NotAuthorized();
        if (!hasPassport[fan]) revert NoPassport();

        _stamps[fan].push(Stamp({showId: showId, time: uint64(block.timestamp)}));
        emit Stamped(fan, showId, uint64(block.timestamp));
    }

    /// @notice 우선권 경매 입찰 시 점수를 소모시킨다(올페이: 낙찰 여부와 무관하게 소모).
    function spend(address fan, uint256 amount) external {
        if (!authorized[msg.sender]) revert NotAuthorized();
        uint256 p = points(fan);
        if (p < amount) revert InsufficientPoints(p, amount);

        _debits[fan].push(Debit({amount: uint128(amount), time: uint64(block.timestamp)}));
        emit PointsSpent(fan, amount);
    }

    /* ─────────────────────────── 조회 ─────────────────────────── */

    /// @notice 감가를 반영한 현재 가용 점수. (적립 - 소모, 둘 다 같은 비율로 감가)
    function points(address fan) public view returns (uint256) {
        Stamp[] storage s = _stamps[fan];
        uint256 gross = 0;
        for (uint256 i = 0; i < s.length; i++) {
            gross += _decayed(POINTS_PER_ATTENDANCE, s[i].time);
        }

        Debit[] storage d = _debits[fan];
        uint256 used = 0;
        for (uint256 i = 0; i < d.length; i++) {
            used += _decayed(d[i].amount, d[i].time);
        }

        return gross > used ? gross - used : 0;
    }

    /// @dev DECAY_PERIOD에 걸쳐 선형으로 0이 되는 값.
    function _decayed(uint256 value, uint64 t) private view returns (uint256) {
        uint256 elapsed = block.timestamp - t;
        if (elapsed >= DECAY_PERIOD) return 0;
        return (value * (DECAY_PERIOD - elapsed)) / DECAY_PERIOD;
    }

    /// @notice 소모 이력 조회 (감가 미반영 명목값).
    function debitCount(address fan) external view returns (uint256) {
        return _debits[fan].length;
    }

    /// @notice 감가 이전의 총 참석 횟수 (굿즈 선구매 등 등급 판정용).
    function attendanceCount(address fan) external view returns (uint256) {
        return _stamps[fan].length;
    }

    function stampAt(address fan, uint256 i) external view returns (uint64 showId, uint64 time) {
        Stamp storage s = _stamps[fan][i];
        return (s.showId, s.time);
    }

    /* ─────────────────── 소울바운드: 양도 경로의 부재 ─────────────────── */

    /// @dev ERC-721 인터페이스를 흉내내되 전송은 항상 실패한다.
    ///      "구현하지 않음"이 아니라 "명시적으로 거부함"이어야
    ///      지갑·마켓플레이스가 이유를 사용자에게 보여줄 수 있다.
    function transferFrom(address, address, uint256) external pure {
        revert Soulbound();
    }

    function safeTransferFrom(address, address, uint256) external pure {
        revert Soulbound();
    }

    function approve(address, uint256) external pure {
        revert Soulbound();
    }

    function setApprovalForAll(address, bool) external pure {
        revert Soulbound();
    }
}
