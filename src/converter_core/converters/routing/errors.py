"""路由转换能够定位到客户工作簿的业务异常。"""

from converter_core.contracts import ConversionIssue


class RoutingConversionError(Exception):
    """携带一个或多个可直接展示给用户的路由转换问题。"""

    def __init__(self, *issues: ConversionIssue) -> None:
        self.issues = issues
        super().__init__(issues[0].message if issues else "路由转换失败")
