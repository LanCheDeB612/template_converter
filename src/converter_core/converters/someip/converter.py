"""把客户服务部署和接口定义转换为 SOME/IP 标准模板记录。"""

from __future__ import annotations

from collections import OrderedDict, defaultdict
from dataclasses import replace

from converter_core.contracts import ConversionIssue
from converter_core.converters.someip.errors import SomeipConversionError
from converter_core.converters.someip.models import (
    CommunicationBehaviorRecord,
    InterfaceRecord,
    SomeipConversionData,
    SomeipSourceData,
)


def _identifier(value: int) -> str:
    """按目标模板约定把各类 SOME/IP ID 写成至少四位十六进制。"""
    return f"0x{value:04X}"


def _merge_or_fail(
    records: OrderedDict[tuple[object, ...], object],
    key: tuple[object, ...],
    candidate: object,
    *,
    sheet: str,
    row: int,
) -> None:
    existing = records.get(key)
    if existing is None:
        records[key] = candidate
        return

    # 来源行不参与业务属性比较；属性完全相同的重复项可以合并，
    # 不同属性不能静默采用先出现的一行。
    if replace(existing, source_rows=()) != replace(candidate, source_rows=()):
        raise SomeipConversionError(
            ConversionIssue(
                level="error",
                message="相同目标键对应的服务、协议、版本、端口、周期或事件组属性不一致。",
                sheet=sheet,
                row=row,
            )
        )
    records[key] = replace(
        existing,
        source_rows=(*existing.source_rows, *candidate.source_rows),
    )


def _node_names(source: SomeipSourceData) -> tuple[str, ...]:
    """按客户模板中的首次出现顺序生成唯一节点名称。"""
    names: OrderedDict[str, None] = OrderedDict()

    # Server ECU 来自 Service_Deployment；先写服务端，再补入两个 Sheet 中用 x
    # 标记的客户端。未参与任何服务或事件组的 Clients 表头不是有效节点，不应输出。
    for deployment in source.deployments:
        names.setdefault(deployment.server, None)
    for deployment in source.deployments:
        for client in deployment.clients:
            names.setdefault(client, None)
    for eventgroup in source.eventgroups:
        for client in eventgroup.clients:
            names.setdefault(client, None)

    return tuple(names)


def build_conversion_data(source: SomeipSourceData) -> SomeipConversionData:
    """生成通信行为和事件/方法记录，并检查跨 Sheet 关联及重复冲突。"""
    deployments_by_service: dict[int, list[object]] = defaultdict(list)
    for deployment in source.deployments:
        deployments_by_service[deployment.service_id].append(deployment)

    service_names: dict[int, str] = {}
    service_rows: dict[int, int] = {}
    for service in source.services:
        existing = service_names.get(service.service_id)
        if existing is not None and existing != service.service_name:
            raise SomeipConversionError(
                ConversionIssue(
                    level="error",
                    message=(
                        f"Service ID {_identifier(service.service_id)} 在第 {service_rows[service.service_id]} 行"
                        f"和第 {service.source_row} 行使用了不同的 Service Name。"
                    ),
                    sheet="Service_Interface",
                    row=service.source_row,
                    column="Service Name",
                )
            )
        service_names[service.service_id] = service.service_name
        service_rows[service.service_id] = service.source_row

    issues: list[ConversionIssue] = []
    for deployment in source.deployments:
        interface_name = service_names.get(deployment.service_id)
        if interface_name is None:
            issues.append(
                ConversionIssue(
                    level="error",
                    message=(
                        f"Service ID {_identifier(deployment.service_id)} 在 Service_Interface 中没有服务定义。"
                    ),
                    sheet="Service_Deployment",
                    row=deployment.source_row,
                    column="Service ID",
                )
            )
        elif interface_name != deployment.service_name:
            issues.append(
                ConversionIssue(
                    level="error",
                    message=(
                        f"Service ID {_identifier(deployment.service_id)} 在两个 Sheet 中的 Service Name 不一致："
                        f"“{deployment.service_name}”与“{interface_name}”。"
                    ),
                    sheet="Service_Deployment",
                    row=deployment.source_row,
                    column="Service Name",
                )
            )
    for service in source.services:
        if service.service_id not in deployments_by_service:
            issues.append(
                ConversionIssue(
                    level="error",
                    message=(
                        f"Service ID {_identifier(service.service_id)} 在 Service_Deployment 中没有部署记录。"
                    ),
                    sheet="Service_Interface",
                    row=service.source_row,
                    column="Service ID",
                )
            )
    if issues:
        raise SomeipConversionError(*issues)

    eventgroups_by_service: dict[int, list[object]] = defaultdict(list)
    for eventgroup in source.eventgroups:
        eventgroups_by_service[eventgroup.service_id].append(eventgroup)

    behavior_rows: OrderedDict[tuple[object, ...], CommunicationBehaviorRecord] = OrderedDict()
    for deployment in source.deployments:
        subscribed_clients: set[str] = set()
        for eventgroup in eventgroups_by_service[deployment.service_id]:
            # Eventgroup 的 Clients 标记是订阅关系的唯一依据；即使部署表没有该节点，
            # 仍按接口表生成，不把两张表的用途混为一谈。
            for client in eventgroup.clients:
                subscribed_clients.add(client)
                candidate = CommunicationBehaviorRecord(
                    service_name=deployment.service_name,
                    service_id=_identifier(deployment.service_id),
                    instance_id=_identifier(deployment.instance_id),
                    major_version=deployment.major_version,
                    minor_version=deployment.minor_version,
                    transport_protocol=deployment.transport_protocol,
                    server=deployment.server,
                    server_port_udp=deployment.server_port,
                    server_port_tcp=None,
                    eventgroup_id=_identifier(eventgroup.eventgroup_id),
                    event_protocol=eventgroup.protocol,
                    client=client,
                    source_rows=(deployment.source_row, eventgroup.source_row),
                )
                key = (
                    deployment.server,
                    deployment.service_id,
                    deployment.instance_id,
                    client,
                    eventgroup.eventgroup_id,
                )
                _merge_or_fail(
                    behavior_rows,
                    key,
                    candidate,
                    sheet="Service_Interface",
                    row=eventgroup.source_row,
                )

        # 部署表只补充“参与服务但未订阅任何事件组”的关系；已经订阅的客户端
        # 不再生成一条无 EventgroupID 的重复服务关系。
        for client in deployment.clients:
            if client in subscribed_clients:
                continue
            candidate = CommunicationBehaviorRecord(
                service_name=deployment.service_name,
                service_id=_identifier(deployment.service_id),
                instance_id=_identifier(deployment.instance_id),
                major_version=deployment.major_version,
                minor_version=deployment.minor_version,
                transport_protocol=deployment.transport_protocol,
                server=deployment.server,
                server_port_udp=deployment.server_port,
                server_port_tcp=None,
                eventgroup_id=None,
                event_protocol=None,
                client=client,
                source_rows=(deployment.source_row,),
            )
            key = (
                deployment.server,
                deployment.service_id,
                deployment.instance_id,
                client,
                None,
            )
            _merge_or_fail(
                behavior_rows,
                key,
                candidate,
                sheet="Service_Deployment",
                row=deployment.source_row,
            )

    interface_rows: OrderedDict[tuple[int, int], InterfaceRecord] = OrderedDict()
    for interface in source.interfaces:
        candidate = InterfaceRecord(
            service_name=interface.service_name,
            service_id=_identifier(interface.service_id),
            element_name=interface.element_name,
            element_id=_identifier(interface.element_id),
            rpc_type=interface.rpc_type,
            protocol=interface.protocol,
            eventgroup_id=(
                _identifier(interface.eventgroup_id)
                if interface.eventgroup_id is not None
                else None
            ),
            cycle_time=interface.cycle_time,
            source_rows=(interface.source_row,),
        )
        _merge_or_fail(
            interface_rows,
            (interface.service_id, interface.element_id),
            candidate,
            sheet="Service_Interface",
            row=interface.source_row,
        )

    return SomeipConversionData(
        node_names=_node_names(source),
        behavior_rows=tuple(behavior_rows.values()),
        interface_rows=tuple(interface_rows.values()),
    )
