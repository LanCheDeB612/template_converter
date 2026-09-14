"""转换任务统一入口。"""

from converter_core.contracts import ConversionRequest, ConversionResult


class ConversionService:
    """隔离界面和协议转换逻辑，后续由 DDS 转换器提供实际实现。"""

    def convert(self, request: ConversionRequest) -> ConversionResult:
        """执行一次转换；工程初始化阶段尚未接入 DDS 规则。"""
        raise NotImplementedError("DDS 转换功能尚未实现")
