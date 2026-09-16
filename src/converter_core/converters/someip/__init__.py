"""客户 SOME/IP 服务矩阵到标准通信矩阵的业务转换。"""

from converter_core.converters.someip.converter import build_conversion_data
from converter_core.converters.someip.errors import SomeipConversionError
from converter_core.converters.someip.models import (
    CommunicationBehaviorRecord,
    EventgroupDefinition,
    InterfaceDefinition,
    InterfaceRecord,
    ServiceDefinition,
    ServiceDeployment,
    SomeipConversionData,
    SomeipSourceData,
)
from converter_core.converters.someip.reader import read_customer_workbook

__all__ = [
    "CommunicationBehaviorRecord",
    "EventgroupDefinition",
    "InterfaceDefinition",
    "InterfaceRecord",
    "ServiceDefinition",
    "ServiceDeployment",
    "SomeipConversionData",
    "SomeipConversionError",
    "SomeipSourceData",
    "build_conversion_data",
    "read_customer_workbook",
]
