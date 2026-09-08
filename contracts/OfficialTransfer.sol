// SPDX-License-Identifier: MIT
pragma solidity ^0.8.20;

import "./FanPassport.sol";
import "./TicketBox.sol";

/**
 * @title OfficialTransfer (공식 티켓 양도)
 * @notice 개인 간 직접 전송은 TicketBox가 전부 막는다. 그러면 정말 못 가게 된 사람이
 *         친구에게 넘길 방법도 사라지는데, 그 예외를 "규칙 안에서만" 열어 주는 계약이다.
 *
 *  통과해야 하는 조건 네 가지:
 *   (1) 가격 상한 — 정가를 넘는 금액으로는 애초에 등록이 되지 않는다.
 *   (2) 본인확인  — 양수인도 팬 여권을 가진 실명 확인 완료자여야 한다.
 *   (3) 중복 보유 — 이미 이 공연 표를 가진 계정은 받을 수 없다.
 *   (4) 수령 동의 — 양수인이 직접 accept를 호출해야 이전이 일어난다.
 *
 *  다만 이 계약이 통제할 수 있는 것은 온체인 결제액뿐이다. 계좌이체 같은 장외 웃돈까지
 *  막지는 못하며, 그 부분은 양수인 본인확인과 리워드 차등(TicketBox의 REWARD_ORIGINAL이
 *  직접 입장의 3분의 1)으로 유인을 줄이는 데 그친다. 제안서 5장에 한계로 명시했다.
 */
contract OfficialTransfer {
    TicketBox public immutable box;

    struct Offer {
        address seller;
        address buyer;   // 지정 양수인
        uint256 price;   // 정가 이하
        bool    open;
    }

    mapping(uint256 => Offer) public offers;

    event Offered(uint256 indexed tokenId, address indexed seller, address indexed buyer, uint256 price);
    event Cancelled(uint256 indexed tokenId);
    event Accepted(uint256 indexed tokenId, address indexed buyer, uint256 price);

    error NotOwner();
    error PriceAboveFaceValue(uint256 price, uint256 cap);
    error NoOffer();
    error NotDesignatedBuyer();
    error WrongPayment(uint256 sent, uint256 need);

    constructor(TicketBox _box) {
        box = _box;
    }

    /// @notice 보유자가 양도 조건을 등록한다. 정가를 넘으면 등록 자체가 실패한다.
    function offer(uint256 tokenId, address buyer, uint256 price) external {
        if (box.ownerOf(tokenId) != msg.sender) revert NotOwner();
        uint256 cap = box.faceValue();
        if (price > cap) revert PriceAboveFaceValue(price, cap);

        offers[tokenId] = Offer({seller: msg.sender, buyer: buyer, price: price, open: true});
        emit Offered(tokenId, msg.sender, buyer, price);
    }

    function cancel(uint256 tokenId) external {
        if (offers[tokenId].seller != msg.sender) revert NotOwner();
        delete offers[tokenId];
        emit Cancelled(tokenId);
    }

    /// @notice 양수인이 직접 호출해 수령에 동의하고 대금을 지불한다.
    /// @dev    여권 보유·중복 보유 검사는 TicketBox.officialTransferTo가 다시 확인한다.
    function accept(uint256 tokenId) external payable {
        Offer memory o = offers[tokenId];
        if (!o.open) revert NoOffer();
        if (o.buyer != msg.sender) revert NotDesignatedBuyer();
        if (msg.value != o.price) revert WrongPayment(msg.value, o.price);

        delete offers[tokenId];
        box.officialTransferTo(tokenId, msg.sender);

        if (o.price > 0) {
            (bool ok, ) = o.seller.call{value: o.price}("");
            require(ok, "seller payout failed");
        }
        emit Accepted(tokenId, msg.sender, o.price);
    }
}
