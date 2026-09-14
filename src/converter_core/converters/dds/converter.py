"""把客户服务部署和 Topic 定义转换为标准 DDS 业务记录。"""

from __future__ import annotations

from collections import OrderedDict, defaultdict

from converter_core.contracts import ConversionIssue
from converter_core.converters.dds.errors import DdsConversionError
from converter_core.converters.dds.models import (
    DdsConversionData,
    MatrixRecord,
    NodeRecord,
    ServiceDeployment,
    TopicDefinition,
)


def build_conversion_data(
    deployments: tuple[ServiceDeployment, ...],
    topics: tuple[TopicDefinition, ...],
) -> DdsConversionData:
    """生成节点和通信关系，并合并目标字段完全相同的重复 Topic。"""
    deployments_by_service: dict[int, list[ServiceDeployment]] = defaultdict(list)
    for deployment in deployments:
        deployments_by_service[deployment.service_id].append(deployment)

    unmatched = sorted({topic.service_id for topic in topics if topic.service_id not in deployments_by_service})
    if unmatched:
        issues = tuple(
            ConversionIssue(
                level="error",
                message=f"Service ID 0x{service_id:04X} 在 ServiceDeployment 中没有对应服务。",
                sheet="ServicesAndTopicDefinition",
                column="Service ID",
            )
            for service_id in unmatched
        )
        raise DdsConversionError(*issues)

    nodes: OrderedDict[tuple[str, int], NodeRecord] = OrderedDict()
    matrix: OrderedDict[tuple[str, int, str, str], MatrixRecord] = OrderedDict()

    # 节点配置独立来源于 ServiceDeployment。即使某个服务暂时没有 Topic，
    # 它的服务端和已标记客户端仍需保留，供用户补充实际网络参数。
    for deployment in deployments:
        for participant in (deployment.server, *deployment.clients):
            node_key = (participant, deployment.domain_id)
            nodes.setdefault(
                node_key,
                NodeRecord(participant_name=participant, domain_id=deployment.domain_id),
            )

    for topic in topics:
        for deployment in deployments_by_service[topic.service_id]:
            participants = (deployment.server, *deployment.clients)
            for participant in participants:
                is_server = participant == deployment.server
                role = "Writer" if (topic.direction == "OUT") == is_server else "Reader"
                matrix_key = (participant, deployment.domain_id, topic.topic_name, role)
                existing = matrix.get(matrix_key)
                candidate_qos = (
                    topic.reliability,
                    topic.history,
                    topic.history_depth,
                    topic.resource_limits,
                )
                if existing is None:
                    matrix[matrix_key] = MatrixRecord(
                        participant_name=participant,
                        domain_id=deployment.domain_id,
                        topic_name=topic.topic_name,
                        role=role,
                        reliability=topic.reliability,
                        history=topic.history,
                        history_depth=topic.history_depth,
                        resource_limits=topic.resource_limits,
                        source_rows=(topic.source_row,),
                    )
                    continue

                existing_qos = (
                    existing.reliability,
                    existing.history,
                    existing.history_depth,
                    existing.resource_limits,
                )
                if existing_qos != candidate_qos:
                    raise DdsConversionError(
                        ConversionIssue(
                            level="error",
                            message=(
                                f"Topic“{topic.topic_name}”合并后的 Participant、Domain 和角色相同，"
                                f"但第 {existing.source_rows[0]} 行与第 {topic.source_row} 行的 QoS 不一致。"
                            ),
                            sheet="ServicesAndTopicDefinition",
                            row=topic.source_row,
                            column="TopicName／TopicQosConfig",
                        )
                    )

                # 同一 Topic 可由多个服务元素共用；目标通信字段和 QoS 相同时只保留一条，
                # 同时记录全部来源行，便于后续出现问题时定位客户表。
                matrix[matrix_key] = MatrixRecord(
                    participant_name=existing.participant_name,
                    domain_id=existing.domain_id,
                    topic_name=existing.topic_name,
                    role=existing.role,
                    reliability=existing.reliability,
                    history=existing.history,
                    history_depth=existing.history_depth,
                    resource_limits=existing.resource_limits,
                    source_rows=(*existing.source_rows, topic.source_row),
                )

    return DdsConversionData(nodes=tuple(nodes.values()), matrix_rows=tuple(matrix.values()))
