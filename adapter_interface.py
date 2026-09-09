"""格式适配器的抽象接口 —— 比对引擎只认归一化 Prop，不认具体格式。

docx/xlsx/pptx（office_adapter.py）与邮件 .eml（email_adapter.py）都实现同一接口，
引擎零改动接入。新增格式只需实现 FormatAdapter 子类并 AdapterRegistry.register()。
详见《校验器边界与规则需求清单.md》第 5 节。
"""
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import Iterable, Dict, List


@dataclass
class Prop:
    """归一化属性occurrence：比对引擎唯一依赖的结构。"""
    part: str          # 文件分区，如 word/document.xml / 邮件 header|body|html
    path: str          # 带位置索引的定位路径，如 p[1]/r[1]/rPr[1]/sz[1]
    element: str       # 元素名，如 sz / font-family / color
    attr: str          # 属性名，如 val / fill / w
    value: str         # 属性值（已解算为可读字符串）


@dataclass
class StructuralAnchor:
    """结构锚点：解决 A1/A2/A4 位置错位问题（按语义而非绝对位置对齐）。"""
    kind: str          # style | bookmark | placeholder | numbering
    name: str          # 如 "Heading1" / "anchor_x" / "{{name}}" / "list-3"
    path: str          # 该锚点在文件中的定位路径


class FormatAdapter(ABC):
    """所有格式适配器的基类。子类只需实现 extract / is_instance / anchors。"""

    #: 格式标识，引擎据此选择适配器（docx / xlsx / pptx / email）
    fmt: str = ""

    @abstractmethod
    def is_instance(self, path: str) -> bool:
        """该文件是否本适配器可处理。"""
        ...

    @abstractmethod
    def extract(self, path: str) -> List[Prop]:
        """抽全量属性（对称 diff 的左侧/右侧来源）。路径必须带位置索引。"""
        ...

    @abstractmethod
    def anchors(self, path: str) -> List[StructuralAnchor]:
        """返回结构锚点，供引擎在绝对位置失效时按语义对齐。"""
        ...

    # ---- 以下为引擎侧公共能力，子类一般无需重写 ----

    def to_map(self, props: List[Prop]) -> Dict[tuple, str]:
        """把 Prop 列表压成 (part, path, element, attr) -> value 的字典，
        作为对称 diff 的 key（位置索引保证同参数不同位置是不同 key）。"""
        return {(p.part, p.path, p.element, p.attr): p.value for p in props}


class AdapterRegistry:
    """适配器注册表：引擎从这里按文件类型取适配器，新增格式只需 register。"""

    _adapters: List[FormatAdapter] = []

    @classmethod
    def register(cls, adapter: FormatAdapter) -> None:
        cls._adapters.append(adapter)

    @classmethod
    def for_file(cls, path: str) -> FormatAdapter:
        for a in cls._adapters:
            if a.is_instance(path):
                return a
        raise ValueError(f"no adapter for: {path}")


# ===== 邮件适配器：已迁移至 email_adapter.py =====
# class EmailAdapter 的真实实现见 email_adapter.py（解析 .eml，产出与 docx 同构的
# 带位置索引 Prop 流：邮件头 + HTML 内联样式拆分 + 占位符 {{}} 捕获）。
# .msg（Outlook OLE）需 olefile，待后续接入。
