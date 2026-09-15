"""把客户路由记录转换为三个标准路由 Sheet 的业务记录。"""

from __future__ import annotations

import re
from collections import OrderedDict
from dataclasses import fields, replace
from decimal import Decimal, InvalidOperation
from typing import NoReturn, TypeVar

from converter_core.contracts import ConversionIssue
from converter_core.converters.routing.errors import RoutingConversionError
from converter_core.converters.routing.models import (
    CanPduRouteRecord,
    CustomerEndpoint,
    CustomerRoutingRow,
    EthPduRouteRecord,
    RoutingConversionData,
    SignalRouteRecord,
)
from converter_core.converters.routing.reader import ROUTING_SHEET

CAN_TYPES = {"CAN", "CANFD"}

_SOURCE_COLUMNS = {
    "can_type": "Source CAN Type",
    "tcp_udp": "Source TCP/UDP",
    "lin_delay": "Source LIN Delay(ms)",
    "lin_schedule_delay": "Source LIN Schedule Delay(ms)",
    "frame_id": "Source Frame ID(Hex)",
    "frame_length": "Source Frame Length(Byte)",
    "pdu_id": "Source PDU ID(Hex)",
    "pdu_length": "Source PDU Length(Byte)",
    "signal_start_bit": "Source Signal Start bit",
    "byte_order": "Source Byte Order",
    "source_ip": "Source Src IP Addr",
    "target_ip": "Source Tgt IP Addr",
    "source_port": "Source Src Port",
    "target_port": "Source Tgt Port",
}
_TARGET_COLUMNS = {
    "can_type": "Target CAN Type",
    "tcp_udp": "Target TCP/UDP",
    "lin_delay": "Target LIN Delay(ms)",
    "lin_schedule_delay": "Target LIN Schedule Delay(ms)",
    "frame_tx_mode": "Target Frame TxMode",
    "frame_period": "Target Frame Period/mRTI(ms)",
    "frame_id": "Target Frame ID(Hex)",
    "frame_length": "Target Frame Length(Byte)",
    "pdu_id": "Target PDU ID(Hex)",
    "pdu_triggering": "Target PDU Triggering",
    "pdu_period": "Target PDU Period/mRTI(ms)",
    "pdu_length": "Target PDU Length(Byte)",
    "signal_start_bit": "Destination Signal Start bit",
    "byte_order": "Target Byte Order",
    "source_ip": "Target Src IP Addr",
    "target_ip": "Target Tgt IP Addr",
    "source_port": "Target Src Port",
    "target_port": "Target Tgt Port",
}


def _text(value: object | None) -> str:
    return "" if value is None else str(value).strip()


def _cell_value(value: object | None) -> object | None:
    """清理字符串两端空白，但不把客户缺失值补成默认值。"""
    if isinstance(value, str):
        return value.strip() or None
    return value


def _fail(row: CustomerRoutingRow, column: str, message: str) -> NoReturn:
    raise RoutingConversionError(
        ConversionIssue(
            level="error",
            message=message,
            sheet=ROUTING_SHEET,
            row=row.source_row,
            column=column,
        )
    )


def _normalise_can_type(value: object | None) -> str | None:
    text = re.sub(r"[\s_-]+", "", _text(value)).upper()
    if not text:
        return None
    return {"CAN": "CAN", "CANFD": "CANFD"}.get(text)


def _protocol(
    row: CustomerRoutingRow,
    endpoint: CustomerEndpoint,
    columns: dict[str, str],
    side_name: str,
) -> str:
    """只使用各协议专用字段识别端点，避免根据网段名称猜测协议。"""
    protocols: list[str] = []
    can_text = _text(endpoint.can_type)
    if can_text:
        can_type = _normalise_can_type(endpoint.can_type)
        if can_type is None:
            _fail(row, columns["can_type"], f"{side_name} CAN 类型“{endpoint.can_type}”无法识别。")
        protocols.append(can_type)

    transport = _text(endpoint.tcp_udp)
    if transport:
        if transport.casefold() not in {"tcp", "udp"}:
            _fail(row, columns["tcp_udp"], f"{side_name} TCP/UDP 类型“{endpoint.tcp_udp}”无法识别。")
        protocols.append("ETH")

    if _text(endpoint.lin_delay) or _text(endpoint.lin_schedule_delay):
        protocols.append("LIN")

    protocols = list(dict.fromkeys(protocols))
    if not protocols:
        _fail(
            row,
            f"{columns['can_type']}／{columns['tcp_udp']}／{columns['lin_delay']}／{columns['lin_schedule_delay']}",
            f"无法根据{side_name}的协议专用字段识别网络类型。",
        )
    if len(protocols) > 1:
        _fail(
            row,
            f"{columns['can_type']}／{columns['tcp_udp']}／{columns['lin_delay']}／{columns['lin_schedule_delay']}",
            f"{side_name}同时出现多个协议的专用字段：{', '.join(protocols)}。",
        )
    return protocols[0]


def _integer(
    value: object | None,
    *,
    row: CustomerRoutingRow,
    column: str,
    field_name: str,
    allow_blank: bool,
) -> int | None:
    if not _text(value):
        if allow_blank:
            return None
        _fail(row, column, f"{field_name}不能为空。")
    try:
        if isinstance(value, bool):
            raise ValueError
        number = Decimal(_text(value))
    except (InvalidOperation, ValueError):
        _fail(row, column, f"{field_name}必须是整数，实际值为“{value}”。")
    if number != number.to_integral_value():
        _fail(row, column, f"{field_name}必须是整数，不能截断“{value}”。")
    return int(number)


def _normalise_id(
    value: object | None,
    protocol: str,
    *,
    row: CustomerRoutingRow,
    column: str,
) -> str | None:
    """ID 字段统一写成 0x 前缀；宽度不足时仅补零，不截断扩展 CAN ID。"""
    text = _text(value)
    if not text:
        return None
    try:
        if isinstance(value, bool):
            raise ValueError
        if isinstance(value, int):
            number = value
        elif isinstance(value, float) and value.is_integer():
            number = int(value)
        else:
            digits = text[2:] if text.casefold().startswith("0x") else text
            number = int(digits, 16)
        if number < 0:
            raise ValueError
    except (TypeError, ValueError):
        _fail(row, column, f"报文或 PDU ID 必须是十六进制整数，实际值为“{value}”。")
    width = 8 if protocol == "ETH" else 2 if protocol == "LIN" else 3
    return f"0x{number:0{width}X}"


def _frame_route_type(row: CustomerRoutingRow) -> str:
    value = _text(row.target.frame_tx_mode)
    key = re.sub(r"[\s_-]+", "", value).casefold()
    result = {
        "eventtriggered": "Event",
        "periodic": "Cycle",
        "ce": "Cycle",
    }.get(key)
    if result is None:
        _fail(
            row,
            _TARGET_COLUMNS["frame_tx_mode"],
            f"Target Frame TxMode 的值“{row.target.frame_tx_mode}”没有已确认的转换规则。",
        )
    return result


def _target_pdu_route_type(row: CustomerRoutingRow) -> str | None:
    value = _text(row.target.pdu_triggering)
    if not value:
        return None
    result = {"always": "Always", "never": "Never"}.get(value.casefold())
    if result is None:
        _fail(
            row,
            _TARGET_COLUMNS["pdu_triggering"],
            f"Target PDU Triggering 的值“{row.target.pdu_triggering}”没有已确认的转换规则。",
        )
    return result


def _positive_period(
    value: object | None,
    *,
    row: CustomerRoutingRow,
    column: str,
) -> int:
    """只提取 Period；CE 的 ``(Period, mRTI)`` 取第一个值，纯 mRTI 不冒充周期。"""
    text = _text(value)
    number_text: str | None = None
    period_match = re.search(r"\bperiod\s*:\s*([+-]?\d+(?:\.\d+)?)", text, re.IGNORECASE)
    if period_match:
        number_text = period_match.group(1)
    else:
        tuple_match = re.fullmatch(r"\(\s*([+-]?\d+(?:\.\d+)?)\s*,\s*[^)]*\)", text)
        plain_match = re.fullmatch(r"[+-]?\d+(?:\.\d+)?", text)
        if tuple_match:
            number_text = tuple_match.group(1)
        elif plain_match:
            number_text = plain_match.group(0)

    if number_text is None:
        _fail(row, column, f"Cycle 路由需要明确的 Period 正整数，实际值为“{value}”。")
    try:
        period = Decimal(number_text)
    except InvalidOperation:
        _fail(row, column, f"周期值“{value}”无法解析。")
    if period <= 0 or period != period.to_integral_value():
        _fail(row, column, f"周期必须是正整数毫秒，不能使用“{value}”。")
    return int(period)


def _target_cycle(row: CustomerRoutingRow, target_protocol: str, route_type: str | None) -> int | None:
    if route_type != "Cycle":
        return None
    if target_protocol in CAN_TYPES:
        return _positive_period(
            row.target.frame_period,
            row=row,
            column=_TARGET_COLUMNS["frame_period"],
        )
    if target_protocol == "ETH":
        return _positive_period(
            row.target.pdu_period,
            row=row,
            column=_TARGET_COLUMNS["pdu_period"],
        )
    return _positive_period(
        row.target.lin_schedule_delay,
        row=row,
        column=_TARGET_COLUMNS["lin_schedule_delay"],
    )


def _endpoint_id(
    row: CustomerRoutingRow,
    endpoint: CustomerEndpoint,
    protocol: str,
    columns: dict[str, str],
) -> str | None:
    field = "pdu_id" if protocol == "ETH" else "frame_id"
    return _normalise_id(getattr(endpoint, field), protocol, row=row, column=columns[field])


def _endpoint_length(endpoint: CustomerEndpoint, protocol: str) -> object | None:
    return _cell_value(endpoint.pdu_length if protocol == "ETH" else endpoint.frame_length)


def _signal_position(
    row: CustomerRoutingRow,
    endpoint: CustomerEndpoint,
    columns: dict[str, str],
) -> tuple[int | None, str | None]:
    order_text = _text(endpoint.byte_order)
    start_text = _text(endpoint.signal_start_bit)
    if not order_text:
        if start_text:
            _fail(row, columns["byte_order"], "信号起始位有值时必须提供字节序，才能转换为 LSB 定义。")
        return None, None

    order_key = re.sub(r"[\s_-]+", "", order_text).casefold()
    byte_order = {"msbfirst": "Motorola", "msblast": "Intel"}.get(order_key)
    if byte_order is None:
        _fail(row, columns["byte_order"], f"字节序“{endpoint.byte_order}”无法转换为 Intel 或 Motorola。")

    start_bit = _integer(
        endpoint.signal_start_bit,
        row=row,
        column=columns["signal_start_bit"],
        field_name="信号起始位",
        allow_blank=True,
    )
    if start_bit is None:
        return None, byte_order
    if start_bit < 0:
        _fail(row, columns["signal_start_bit"], "信号起始位不能是负数。")
    if byte_order == "Intel":
        return start_bit, byte_order

    signal_length = _integer(
        row.signal_length,
        row=row,
        column="Size of Signal",
        field_name="信号位长",
        allow_blank=False,
    )
    if signal_length is None or signal_length <= 0:
        _fail(row, "Size of Signal", "Motorola 信号需要正整数位长才能把 MSB 起始位换算为 LSB。")

    # 客户 MSBFirst 的起始位指向最高有效位。北汇模板需要最低有效位，
    # 因此按 Motorola sawtooth 位编号从 MSB 反向走 length - 1 位。
    lsb = start_bit
    for _ in range(signal_length - 1):
        lsb = lsb + 15 if lsb % 8 == 0 else lsb - 1
    if lsb < 0:
        _fail(row, columns["signal_start_bit"], "Motorola 起始位换算后超出报文起点，请核对起始位和位长。")
    return lsb, byte_order


def _validate_single_socket(row: CustomerRoutingRow, endpoint: CustomerEndpoint, columns: dict[str, str]) -> None:
    """多 IP 或多端口没有配对标识时不能取第一项，也不能生成笛卡尔积。"""
    for field in ("source_ip", "target_ip", "source_port", "target_port"):
        value = _text(getattr(endpoint, field))
        if value and re.search(r"[,;/、\n\r]", value):
            _fail(
                row,
                columns[field],
                f"{columns[field]} 含多个值“{value}”，客户表没有说明 IP 与端口的配对关系。",
            )


def _build_can_pdu(
    row: CustomerRoutingRow,
    source_protocol: str,
    target_protocol: str,
) -> CanPduRouteRecord:
    route_type = _frame_route_type(row)
    return CanPduRouteRecord(
        pdu_name=_cell_value(row.source.pdu_name),
        pdu_routing_type=route_type,
        src_network_name=_cell_value(row.source.network_name),
        src_network_type=source_protocol,
        src_pdu_header_id=_endpoint_id(row, row.source, source_protocol, _SOURCE_COLUMNS),
        # CAN-PDU 的长度来源由客户单独指定为 PDU Length，不能替换成 Frame Length。
        src_pdu_length=_cell_value(row.source.pdu_length),
        des_network_name=_cell_value(row.target.network_name),
        des_network_type=target_protocol,
        des_pdu_header_id=_endpoint_id(row, row.target, target_protocol, _TARGET_COLUMNS),
        des_cycle=_target_cycle(row, target_protocol, route_type),
        des_pdu_length=_cell_value(row.target.pdu_length),
        gateway_ecu=row.gateway_ecu,
        source_rows=(row.source_row,),
    )


def _build_signal(
    row: CustomerRoutingRow,
    source_protocol: str,
    target_protocol: str,
) -> SignalRouteRecord:
    route_type = _frame_route_type(row)
    source_start, source_order = _signal_position(row, row.source, _SOURCE_COLUMNS)
    target_start, target_order = _signal_position(row, row.target, _TARGET_COLUMNS)
    return SignalRouteRecord(
        sig_name=_cell_value(row.signal_name),
        sig_routing_type=route_type,
        sig_length=_cell_value(row.signal_length),
        src_network_name=_cell_value(row.source.network_name),
        src_network_type=source_protocol,
        src_pdu_header_id=_endpoint_id(row, row.source, source_protocol, _SOURCE_COLUMNS),
        src_pdu_length=_endpoint_length(row.source, source_protocol),
        src_startbit=source_start,
        src_startbit_type="LSB",
        src_byte_order=source_order,
        des_network_name=_cell_value(row.target.network_name),
        des_network_type=target_protocol,
        des_pdu_header_id=_endpoint_id(row, row.target, target_protocol, _TARGET_COLUMNS),
        des_cycle=_target_cycle(row, target_protocol, route_type),
        des_pdu_length=_endpoint_length(row.target, target_protocol),
        des_startbit=target_start,
        des_startbit_type="LSB",
        des_byte_order=target_order,
        gateway_ecu=row.gateway_ecu,
        source_rows=(row.source_row,),
    )


def _eth_values(endpoint: CustomerEndpoint, is_eth: bool) -> dict[str, object | None]:
    if not is_eth:
        return {
            "vlan_id": None,
            "client_port": None,
            "client_ip": None,
            "server_port": None,
            "server_ip": None,
        }
    return {
        "vlan_id": _cell_value(endpoint.vlan_id),
        # 客户在线说明把 Tgt 定义为 Client、Src 定义为 Server；IP 与端口必须保持同侧配对。
        "client_port": _cell_value(endpoint.target_port),
        "client_ip": _cell_value(endpoint.target_ip),
        "server_port": _cell_value(endpoint.source_port),
        "server_ip": _cell_value(endpoint.source_ip),
    }


def _build_eth_pdu(
    row: CustomerRoutingRow,
    source_protocol: str,
    target_protocol: str,
) -> EthPduRouteRecord:
    if source_protocol == "ETH":
        _validate_single_socket(row, row.source, _SOURCE_COLUMNS)
    if target_protocol == "ETH":
        _validate_single_socket(row, row.target, _TARGET_COLUMNS)

    route_type = _target_pdu_route_type(row) if target_protocol == "ETH" else _frame_route_type(row)
    source_eth = _eth_values(row.source, source_protocol == "ETH")
    target_eth = _eth_values(row.target, target_protocol == "ETH")
    return EthPduRouteRecord(
        pdu_name=_cell_value(row.source.pdu_name),
        pdu_routing_type=route_type,
        src_network_name=_cell_value(row.source.network_name),
        src_network_type=source_protocol,
        src_pdu_header_id=_endpoint_id(row, row.source, source_protocol, _SOURCE_COLUMNS),
        src_pdu_length=_endpoint_length(row.source, source_protocol),
        src_eth_vlan_id=source_eth["vlan_id"],
        src_eth_vlan_priority=None,
        src_eth_client_port=source_eth["client_port"],
        src_eth_client_mac=None,
        src_eth_client_ip=source_eth["client_ip"],
        src_eth_server_port=source_eth["server_port"],
        src_eth_server_mac=None,
        src_eth_server_ip=source_eth["server_ip"],
        des_network_name=_cell_value(row.target.network_name),
        des_network_type=target_protocol,
        des_pdu_header_id=_endpoint_id(row, row.target, target_protocol, _TARGET_COLUMNS),
        des_cycle=_target_cycle(row, target_protocol, route_type),
        des_pdu_length=_endpoint_length(row.target, target_protocol),
        des_eth_vlan_id=target_eth["vlan_id"],
        des_eth_vlan_priority=None,
        des_eth_client_port=target_eth["client_port"],
        des_eth_client_mac=None,
        des_eth_client_ip=target_eth["client_ip"],
        des_eth_server_port=target_eth["server_port"],
        des_eth_server_mac=None,
        des_eth_server_ip=target_eth["server_ip"],
        gateway_ecu=row.gateway_ecu,
        source_rows=(row.source_row,),
    )


RouteRecord = TypeVar("RouteRecord", CanPduRouteRecord, SignalRouteRecord, EthPduRouteRecord)


def _target_values(record: RouteRecord) -> tuple[object, ...]:
    return tuple(
        getattr(record, field.name)
        for field in fields(record)
        if field.name not in {"gateway_ecu", "source_rows"}
    )


def _can_identity(record: CanPduRouteRecord) -> tuple[object, ...]:
    return (
        record.gateway_ecu,
        record.pdu_name,
        record.src_network_name,
        record.src_network_type,
        record.src_pdu_header_id,
        record.des_network_name,
        record.des_network_type,
        record.des_pdu_header_id,
    )


def _signal_identity(record: SignalRouteRecord) -> tuple[object, ...]:
    return (
        record.gateway_ecu,
        record.sig_name,
        record.src_network_name,
        record.src_network_type,
        record.src_pdu_header_id,
        record.src_startbit,
        record.des_network_name,
        record.des_network_type,
        record.des_pdu_header_id,
        record.des_startbit,
    )


def _eth_identity(record: EthPduRouteRecord) -> tuple[object, ...]:
    return (
        record.gateway_ecu,
        record.pdu_name,
        record.src_network_name,
        record.src_network_type,
        record.src_pdu_header_id,
        record.src_eth_vlan_id,
        record.src_eth_client_port,
        record.src_eth_client_ip,
        record.src_eth_server_port,
        record.src_eth_server_ip,
        record.des_network_name,
        record.des_network_type,
        record.des_pdu_header_id,
        record.des_eth_vlan_id,
        record.des_eth_client_port,
        record.des_eth_client_ip,
        record.des_eth_server_port,
        record.des_eth_server_ip,
    )


def _add_record(
    record: RouteRecord,
    records: OrderedDict[tuple[object, ...], RouteRecord],
    identities: dict[tuple[object, ...], RouteRecord],
    identity: tuple[object, ...],
    conflict_column: str,
) -> None:
    """完整目标字段相同才合并；相同端点的重要属性不同时直接报告冲突。"""
    full_key = (record.gateway_ecu, *_target_values(record))
    existing = records.get(full_key)
    if existing is not None:
        records[full_key] = replace(existing, source_rows=(*existing.source_rows, *record.source_rows))
        return

    conflicting = identities.get(identity)
    if conflicting is not None:
        row = record.source_rows[0]
        gateway = record.gateway_ecu or "空"
        raise RoutingConversionError(
            ConversionIssue(
                level="error",
                message=(
                    f"Gateway Ecu“{gateway}”的相同路由端点在第 {conflicting.source_rows[0]} 行和"
                    f"第 {row} 行存在不同的长度、周期、路由方式或字节序。"
                ),
                sheet=ROUTING_SHEET,
                row=row,
                column=conflict_column,
            )
        )
    records[full_key] = record
    identities[identity] = record


def build_conversion_data(rows: tuple[CustomerRoutingRow, ...]) -> RoutingConversionData:
    """分流、转换、去重客户记录，并一次返回所有可定位的业务问题。"""
    can_records: OrderedDict[tuple[object, ...], CanPduRouteRecord] = OrderedDict()
    signal_records: OrderedDict[tuple[object, ...], SignalRouteRecord] = OrderedDict()
    eth_records: OrderedDict[tuple[object, ...], EthPduRouteRecord] = OrderedDict()
    can_identities: dict[tuple[object, ...], CanPduRouteRecord] = {}
    signal_identities: dict[tuple[object, ...], SignalRouteRecord] = {}
    eth_identities: dict[tuple[object, ...], EthPduRouteRecord] = {}
    issues: list[ConversionIssue] = []

    for row in rows:
        try:
            source_protocol = _protocol(row, row.source, _SOURCE_COLUMNS, "源端")
            target_protocol = _protocol(row, row.target, _TARGET_COLUMNS, "目标端")
            gateway_type = _text(row.gateway_type).casefold()
            if gateway_type == "pdu gateway":
                if source_protocol in CAN_TYPES and target_protocol in CAN_TYPES:
                    record = _build_can_pdu(row, source_protocol, target_protocol)
                    _add_record(
                        record,
                        can_records,
                        can_identities,
                        _can_identity(record),
                        "Gateway Type(PDU/Signal)／Target Frame TxMode／PDU Length",
                    )
                elif (source_protocol == "ETH" and target_protocol in CAN_TYPES) or (
                    source_protocol in CAN_TYPES and target_protocol == "ETH"
                ):
                    record = _build_eth_pdu(row, source_protocol, target_protocol)
                    _add_record(
                        record,
                        eth_records,
                        eth_identities,
                        _eth_identity(record),
                        "Gateway Type(PDU/Signal)／Target Frame TxMode／Target PDU Triggering／PDU Length",
                    )
                else:
                    _fail(
                        row,
                        "Gateway Type(PDU/Signal)",
                        f"PDU Gateway 暂不支持 {source_protocol} 到 {target_protocol} 的协议组合。",
                    )
            elif gateway_type == "signal gateway":
                record = _build_signal(row, source_protocol, target_protocol)
                _add_record(
                    record,
                    signal_records,
                    signal_identities,
                    _signal_identity(record),
                    "Gateway Type(PDU/Signal)／信号位置／字节序／周期",
                )
            else:
                _fail(
                    row,
                    "Gateway Type(PDU/Signal)",
                    f"Gateway Type 的值“{row.gateway_type}”无法识别。",
                )
        except RoutingConversionError as error:
            issues.extend(error.issues)

    if issues:
        raise RoutingConversionError(*issues)
    return RoutingConversionData(
        can_pdu_routes=tuple(can_records.values()),
        signal_routes=tuple(signal_records.values()),
        eth_pdu_routes=tuple(eth_records.values()),
    )
