"""客户路由表解析结果和标准模板业务记录。"""

from __future__ import annotations

from dataclasses import dataclass


CellValue = object | None


@dataclass(frozen=True)
class CustomerEndpoint:
    """保存客户表中同一侧端点的协议识别字段和业务字段。"""

    network_name: CellValue
    can_type: CellValue
    tcp_udp: CellValue
    vlan_id: CellValue
    source_ip: CellValue
    target_ip: CellValue
    source_port: CellValue
    target_port: CellValue
    frame_tx_mode: CellValue
    frame_period: CellValue
    frame_id: CellValue
    frame_length: CellValue
    lin_delay: CellValue
    lin_schedule_delay: CellValue
    pdu_name: CellValue
    pdu_id: CellValue
    pdu_triggering: CellValue
    pdu_period: CellValue
    pdu_length: CellValue
    signal_start_bit: CellValue
    byte_order: CellValue


@dataclass(frozen=True)
class CustomerRoutingRow:
    """表示客户 ``Routing Table`` 中一条保持原始单元格值的路由记录。"""

    source_row: int
    gateway_ecu: str | None
    gateway_type: str | None
    signal_name: CellValue
    signal_length: CellValue
    source: CustomerEndpoint
    target: CustomerEndpoint


@dataclass(frozen=True)
class CanPduRouteRecord:
    """表示标准模板 ``CAN-PDU路由`` 中的一条记录。"""

    pdu_name: CellValue
    pdu_routing_type: str | None
    src_network_name: CellValue
    src_network_type: str
    src_pdu_header_id: str | None
    src_pdu_length: CellValue
    des_network_name: CellValue
    des_network_type: str
    des_pdu_header_id: str | None
    des_cycle: int | None
    des_pdu_length: CellValue
    gateway_ecu: str | None
    source_rows: tuple[int, ...]


@dataclass(frozen=True)
class SignalRouteRecord:
    """表示标准模板 ``信号路由`` 中的一条记录。"""

    sig_name: CellValue
    sig_routing_type: str
    sig_length: CellValue
    src_network_name: CellValue
    src_network_type: str
    src_pdu_header_id: str | None
    src_pdu_length: CellValue
    src_startbit: int | None
    src_startbit_type: str
    src_byte_order: str | None
    des_network_name: CellValue
    des_network_type: str
    des_pdu_header_id: str | None
    des_cycle: int | None
    des_pdu_length: CellValue
    des_startbit: int | None
    des_startbit_type: str
    des_byte_order: str | None
    gateway_ecu: str | None
    source_rows: tuple[int, ...]


@dataclass(frozen=True)
class EthPduRouteRecord:
    """表示标准模板 ``ETH-PDU路由`` 中的一条记录。"""

    pdu_name: CellValue
    pdu_routing_type: str | None
    src_network_name: CellValue
    src_network_type: str
    src_pdu_header_id: str | None
    src_pdu_length: CellValue
    src_eth_vlan_id: CellValue
    src_eth_vlan_priority: None
    src_eth_client_port: CellValue
    src_eth_client_mac: None
    src_eth_client_ip: CellValue
    src_eth_server_port: CellValue
    src_eth_server_mac: None
    src_eth_server_ip: CellValue
    des_network_name: CellValue
    des_network_type: str
    des_pdu_header_id: str | None
    des_cycle: int | None
    des_pdu_length: CellValue
    des_eth_vlan_id: CellValue
    des_eth_vlan_priority: None
    des_eth_client_port: CellValue
    des_eth_client_mac: None
    des_eth_client_ip: CellValue
    des_eth_server_port: CellValue
    des_eth_server_mac: None
    des_eth_server_ip: CellValue
    gateway_ecu: str | None
    source_rows: tuple[int, ...]


@dataclass(frozen=True)
class RoutingConversionData:
    """保存写入三个路由 Sheet 的记录；网段映射按规则不生成。"""

    can_pdu_routes: tuple[CanPduRouteRecord, ...]
    signal_routes: tuple[SignalRouteRecord, ...]
    eth_pdu_routes: tuple[EthPduRouteRecord, ...]
