"""信息映射与办公文档批量处理工具。

对外暴露的主要能力：

- :func:`map_records` —— 按字段映射规则整理记录
- :func:`read_any` / :func:`read_folder` —— 多格式读取、目录批量读取
- :func:`write_output` —— 多格式输出
- :func:`convert_file` / :func:`convert_folder` —— 文档格式转换
- :func:`merge_any` / :func:`merge_tables` —— 大批量输入合并为一个输出文档
- :func:`organize` —— 文件自动归档
- :func:`upload_file` / :func:`upload_folder` —— 自动上传（占位符配置）
"""

from .classify import ClassifyConfig, ClassifyRule, MovePlan, organize, plan_moves
from .converters import ConvertResult, convert_file, convert_folder
from .core import map_records
from .formats import INPUT_SUFFIXES, OUTPUT_SUFFIXES, support_matrix
from .merger import MergeResult, merge_any, merge_tables
from .readers import read_any, read_folder
from .uploader import PingResult, UploadConfig, UploadResult, ping_server, upload_file, upload_folder
from .writers import write_output, write_text_document

__version__ = "0.2.0"

__all__ = [
    "__version__",
    "ClassifyConfig",
    "ClassifyRule",
    "ConvertResult",
    "INPUT_SUFFIXES",
    "MergeResult",
    "MovePlan",
    "OUTPUT_SUFFIXES",
    "PingResult",
    "UploadConfig",
    "UploadResult",
    "convert_file",
    "convert_folder",
    "map_records",
    "merge_any",
    "merge_tables",
    "organize",
    "ping_server",
    "plan_moves",
    "read_any",
    "read_folder",
    "support_matrix",
    "upload_file",
    "upload_folder",
    "write_output",
    "write_text_document",
]
