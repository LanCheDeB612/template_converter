"""DDS 客户模板解析后使用的业务数据。"""

from dataclasses import dataclass


@dataclass(frozen=True)
class ServiceDeployment:
    """描述一个服务在指定 Domain 中的提供方和参与客户端。"""

    source_row: int
    service_id: int
    domain_id: int
    server: str
    clients: tuple[str, ...]


@dataclass(frozen=True)
class TopicDefinition:
    """描述客户 Topic 行中需要写入通信矩阵的字段。"""

    source_row: int
    service_id: int
    topic_name: str
    direction: str
    reliability: str | None
    history: str | None
    history_depth: object | None
    resource_limits: str | None


@dataclass(frozen=True)
class NodeRecord:
    """表示标准模板节点配置中的一条唯一 Participant 记录。"""

    participant_name: str
    domain_id: int


@dataclass(frozen=True)
class MatrixRecord:
    """表示标准模板通信矩阵中的一条唯一通信关系。"""

    participant_name: str
    domain_id: int
    topic_name: str
    role: str
    reliability: str | None
    history: str | None
    history_depth: object | None
    resource_limits: str | None
    source_rows: tuple[int, ...]


@dataclass(frozen=True)
class DdsConversionData:
    """保存即将写入标准模板的节点和通信矩阵数据。"""

    nodes: tuple[NodeRecord, ...]
    matrix_rows: tuple[MatrixRecord, ...]
