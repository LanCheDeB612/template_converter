"""SOME/IP 客户模板解析结果和标准模板业务记录。"""

from dataclasses import dataclass


@dataclass(frozen=True)
class ServiceDeployment:
    """表示 ``Service_Deployment`` 中一条服务部署记录。"""

    source_row: int
    server: str
    service_name: str
    service_id: int
    instance_id: int
    major_version: int
    minor_version: int
    server_port: int
    transport_protocol: str
    clients: tuple[str, ...]


@dataclass(frozen=True)
class ServiceDefinition:
    """表示 ``Service_Interface`` 中一个服务块的起点。"""

    source_row: int
    service_id: int
    service_name: str


@dataclass(frozen=True)
class EventgroupDefinition:
    """表示事件组及其明确标记的订阅客户端。"""

    source_row: int
    service_id: int
    service_name: str
    eventgroup_id: int
    protocol: str
    clients: tuple[str, ...]


@dataclass(frozen=True)
class InterfaceDefinition:
    """表示合并后的 Method 或单条 Event/Notification 定义。"""

    source_row: int
    service_id: int
    service_name: str
    element_name: str
    element_id: int
    rpc_type: str
    protocol: str
    eventgroup_id: int | None
    cycle_time: int | float | None


@dataclass(frozen=True)
class SomeipSourceData:
    """保存客户两个业务 Sheet 解析出的结构化数据。"""

    deployments: tuple[ServiceDeployment, ...]
    services: tuple[ServiceDefinition, ...]
    eventgroups: tuple[EventgroupDefinition, ...]
    interfaces: tuple[InterfaceDefinition, ...]


@dataclass(frozen=True)
class CommunicationBehaviorRecord:
    """表示标准模板 ``通信行为`` 中的一条服务关系。"""

    service_name: str
    service_id: str
    instance_id: str
    major_version: int
    minor_version: int
    transport_protocol: str
    server: str
    server_port_udp: int | None
    server_port_tcp: int | None
    eventgroup_id: str | None
    event_protocol: str | None
    client: str
    source_rows: tuple[int, ...]


@dataclass(frozen=True)
class InterfaceRecord:
    """表示标准模板 ``事件和方法`` 中的一条接口定义。"""

    service_name: str
    service_id: str
    element_name: str
    element_id: str
    rpc_type: str
    protocol: str
    eventgroup_id: str | None
    cycle_time: int | float | None
    source_rows: tuple[int, ...]


@dataclass(frozen=True)
class SomeipConversionData:
    """保存准备写入 SOME/IP 标准模板的节点及两类业务记录。"""

    node_names: tuple[str, ...]
    behavior_rows: tuple[CommunicationBehaviorRecord, ...]
    interface_rows: tuple[InterfaceRecord, ...]
