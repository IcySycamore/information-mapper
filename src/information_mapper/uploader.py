"""自动上传脚本（网址与 API Key 全部使用占位符）。

⚠️ 正式使用前必须替换下面两个占位符，或通过命令行参数 / 环境变量传入：

.. code-block:: python

    DEFAULT_UPLOAD_URL = "https://your-server.example.com/api/upload"   # 上传接口地址
    DEFAULT_API_TOKEN  = "YOUR_API_TOKEN_HERE"                          # 接口密钥

支持两种提交方式：

- ``multipart``（默认）：``multipart/form-data`` 表单文件上传
- ``json``：把文件内容 Base64 后放进 JSON 体，适合 API 网关 / 云函数

只依赖标准库，无需安装 requests。所有网络调用都带重试与超时保护，
``dry_run=True`` 时只组装请求、不真正发送，便于先核对地址与字段。
"""

from __future__ import annotations

import base64
import json
import mimetypes
import os
import time
import urllib.error
import urllib.request
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable

from .formats import INPUT_SUFFIXES
from .readers import iter_input_files

# ------------------------------------------------------------ 占位符配置
DEFAULT_UPLOAD_URL = "https://your-server.example.com/api/upload"  # TODO: 替换为真实上传地址
DEFAULT_API_TOKEN = "YOUR_API_TOKEN_HERE"  # TODO: 替换为真实 API Key

ENV_URL = "INFO_MAP_UPLOAD_URL"
ENV_TOKEN = "INFO_MAP_UPLOAD_TOKEN"

_PLACEHOLDER_MARKERS = ("example.com", "your-", "YOUR_", "REPLACE_ME", "CHANGE_ME")

SUPPORTED_METHODS = ("multipart", "json")


@dataclass
class UploadConfig:
    """上传配置。"""

    url: str = DEFAULT_UPLOAD_URL
    token: str = DEFAULT_API_TOKEN
    method: str = "multipart"
    field_name: str = "file"
    token_header: str = "Authorization"
    token_prefix: str = "Bearer "
    timeout: float = 30.0
    retries: int = 3
    retry_delay: float = 2.0
    verify_ssl: bool = True
    extra_fields: dict[str, str] = field(default_factory=dict)

    # -------------------------------------------------------- 构造
    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "UploadConfig":
        allowed = {field_name for field_name in cls.__dataclass_fields__}
        payload = {key: value for key, value in data.items() if key in allowed}
        payload["extra_fields"] = {
            str(key): str(value) for key, value in (data.get("extra_fields") or {}).items()
        }
        return cls(**payload)

    @classmethod
    def from_file(cls, path: str | Path) -> "UploadConfig":
        config_path = Path(path)
        return cls.from_dict(json.loads(config_path.read_text(encoding="utf-8")))

    def resolve(self) -> "UploadConfig":
        """用环境变量覆盖地址与密钥，便于在 CI / 服务器上安全注入。"""
        url = os.environ.get(ENV_URL, self.url)
        token = os.environ.get(ENV_TOKEN, self.token)
        return UploadConfig(
            url=url,
            token=token,
            method=self.method,
            field_name=self.field_name,
            token_header=self.token_header,
            token_prefix=self.token_prefix,
            timeout=self.timeout,
            retries=self.retries,
            retry_delay=self.retry_delay,
            verify_ssl=self.verify_ssl,
            extra_fields=dict(self.extra_fields),
        )

    # -------------------------------------------------------- 校验
    def placeholders(self) -> list[str]:
        """返回仍是占位符的配置项名称（为空表示已替换成真实值）。"""
        pending: list[str] = []
        if _looks_like_placeholder(self.url):
            pending.append("url")
        if _looks_like_placeholder(self.token):
            pending.append("token")
        return pending

    def validate(self) -> None:
        """确认配置可用于真实上传。"""
        if self.method not in SUPPORTED_METHODS:
            raise ValueError(f"不支持的上传方式：{self.method}，可选 {'/'.join(SUPPORTED_METHODS)}")
        pending = self.placeholders()
        if pending:
            raise ValueError(
                "上传地址/密钥仍是占位符（"
                + "、".join(pending)
                + "）。请修改 config/upload.example.json，"
                + f"或用命令行参数 --url/--token（也可设置环境变量 {ENV_URL} / {ENV_TOKEN}）。"
            )


@dataclass
class UploadResult:
    """单个文件的上传结果。"""

    path: Path
    ok: bool
    status: int | None = None
    message: str = ""
    attempts: int = 0
    bytes_sent: int = 0
    dry_run: bool = False

    def summary(self) -> str:
        if self.dry_run:
            return f"[预演] {self.path.name} → 已组装请求（{self.bytes_sent} 字节）"
        if self.ok:
            return f"[成功] {self.path.name}（HTTP {self.status}，尝试 {self.attempts} 次，{self.bytes_sent} 字节）"
        return f"[失败] {self.path.name}（尝试 {self.attempts} 次）：{self.message}"


# ---------------------------------------------------------------- 请求构造
def build_multipart_body(
    file_path: Path,
    *,
    field_name: str = "file",
    extra_fields: dict[str, str] | None = None,
) -> tuple[bytes, str]:
    """构造 ``multipart/form-data`` 请求体，返回 ``(body, content_type)``。"""
    boundary = f"----InfoMapBoundary{uuid.uuid4().hex}"
    content_type = mimetypes.guess_type(file_path.name)[0] or "application/octet-stream"
    chunks: list[bytes] = []

    for key, value in (extra_fields or {}).items():
        chunks.append(f"--{boundary}\r\n".encode())
        chunks.append(f'Content-Disposition: form-data; name="{key}"\r\n\r\n'.encode())
        chunks.append(f"{value}\r\n".encode())

    chunks.append(f"--{boundary}\r\n".encode())
    chunks.append(
        f'Content-Disposition: form-data; name="{field_name}"; filename="{file_path.name}"\r\n'.encode()
    )
    chunks.append(f"Content-Type: {content_type}\r\n\r\n".encode())
    chunks.append(file_path.read_bytes())
    chunks.append(f"\r\n--{boundary}--\r\n".encode())

    return b"".join(chunks), f"multipart/form-data; boundary={boundary}"


def build_json_body(file_path: Path, *, extra_fields: dict[str, str] | None = None) -> tuple[bytes, str]:
    """构造 JSON 请求体（文件内容 Base64），返回 ``(body, content_type)``。"""
    payload: dict[str, Any] = {
        "filename": file_path.name,
        "content_base64": base64.b64encode(file_path.read_bytes()).decode("ascii"),
        "content_type": mimetypes.guess_type(file_path.name)[0] or "application/octet-stream",
    }
    payload.update(extra_fields or {})
    return json.dumps(payload, ensure_ascii=False).encode("utf-8"), "application/json; charset=utf-8"


def build_request(file_path: str | Path, config: UploadConfig) -> urllib.request.Request:
    """根据配置组装 ``urllib`` 请求对象（不发网络请求）。"""
    path = Path(file_path)
    if not path.is_file():
        raise ValueError(f"找不到待上传文件：{path}")

    if config.method == "json":
        body, content_type = build_json_body(path, extra_fields=config.extra_fields)
    else:
        body, content_type = build_multipart_body(
            path, field_name=config.field_name, extra_fields=config.extra_fields
        )

    headers = {
        "Content-Type": content_type,
        "Content-Length": str(len(body)),
        "User-Agent": "information-mapper/0.2 (+uploader)",
    }
    if config.token:
        headers[config.token_header] = f"{config.token_prefix}{config.token}"

    return urllib.request.Request(config.url, data=body, headers=headers, method="POST")


# ---------------------------------------------------------------- 上传执行
def upload_file(
    file_path: str | Path,
    config: UploadConfig | None = None,
    *,
    dry_run: bool = False,
) -> UploadResult:
    """上传单个文件；失败会按配置重试。"""
    path = Path(file_path)
    settings = (config or UploadConfig()).resolve()

    try:
        if not dry_run:
            settings.validate()
        request = build_request(path, settings)
    except ValueError as error:
        return UploadResult(path=path, ok=False, message=str(error))

    body_size = len(request.data or b"")

    if dry_run:
        return UploadResult(
            path=path,
            ok=True,
            message=f"预演请求 {settings.url}",
            bytes_sent=body_size,
            dry_run=True,
        )

    context = None
    if not settings.verify_ssl:
        import ssl

        context = ssl._create_unverified_context()  # noqa: S323 - 仅供内网自签名证书场景

    attempts = 0
    last_message = ""
    for attempt in range(1, max(settings.retries, 1) + 1):
        attempts = attempt
        try:
            with urllib.request.urlopen(request, timeout=settings.timeout, context=context) as response:
                status = response.status
                text = response.read(2048).decode("utf-8", errors="ignore")
            return UploadResult(
                path=path,
                ok=200 <= status < 300,
                status=status,
                message=text[:200] or "上传成功",
                attempts=attempts,
                bytes_sent=body_size,
            )
        except urllib.error.HTTPError as error:
            last_message = f"HTTP {error.code}：{error.reason}"
            if error.code < 500:  # 4xx 属于请求本身的问题，重试无意义
                break
        except (urllib.error.URLError, TimeoutError, OSError) as error:
            last_message = f"网络错误：{error}"
        if attempt < settings.retries:
            time.sleep(settings.retry_delay)

    return UploadResult(
        path=path,
        ok=False,
        message=last_message or "上传失败",
        attempts=attempts,
        bytes_sent=body_size,
    )


@dataclass
class PingResult:
    """接口连通性测试结果。"""

    url: str
    ok: bool
    status: int | None = None
    elapsed_ms: float = 0.0
    message: str = ""

    def summary(self) -> str:
        if self.ok:
            code = f"HTTP {self.status}" if self.status else "已连接"
            return f"Ping 成功 · {code} · {self.elapsed_ms:.0f} 毫秒"
        return f"Ping 失败 · {self.message}"


def ping_server(
    config: UploadConfig | None = None,
    *,
    timeout: float = 5.0,
) -> PingResult:
    """测试上传接口是否可达（只发一个轻量 GET 请求，不推送任何文件）。

    只要服务器返回了 HTTP 响应（即使 404 / 405）就认为网络可达；
    地址仍是占位符、DNS 解析失败、连接超时则视为不可达。
    """
    settings = (config or UploadConfig()).resolve()
    if _looks_like_placeholder(settings.url):
        return PingResult(url=settings.url, ok=False, message="接口地址还没填（当前为空或占位符）")

    request = urllib.request.Request(settings.url, method="GET")
    request.add_header("User-Agent", "information-mapper/0.3 (+ping)")
    if settings.token:
        request.add_header(settings.token_header, f"{settings.token_prefix}{settings.token}")

    context = None
    if not settings.verify_ssl:
        import ssl

        context = ssl._create_unverified_context()  # noqa: S323 - 内网自签名证书场景

    started = time.perf_counter()
    try:
        with urllib.request.urlopen(request, timeout=timeout, context=context) as response:
            elapsed = (time.perf_counter() - started) * 1000
            return PingResult(
                url=settings.url,
                ok=True,
                status=response.status,
                elapsed_ms=elapsed,
                message="接口可达",
            )
    except urllib.error.HTTPError as error:
        elapsed = (time.perf_counter() - started) * 1000
        return PingResult(
            url=settings.url,
            ok=True,  # 有 HTTP 响应说明服务器是通的
            status=error.code,
            elapsed_ms=elapsed,
            message=f"接口可达（返回 HTTP {error.code}）",
        )
    except (urllib.error.URLError, TimeoutError, OSError) as error:
        elapsed = (time.perf_counter() - started) * 1000
        return PingResult(url=settings.url, ok=False, elapsed_ms=elapsed, message=f"无法连接：{error}")


def upload_paths(
    paths: Iterable[str | Path],
    config: UploadConfig | None = None,
    *,
    dry_run: bool = False,
) -> list[UploadResult]:
    """批量上传多个文件。"""
    settings = (config or UploadConfig()).resolve()
    return [upload_file(path, settings, dry_run=dry_run) for path in paths]


def upload_folder(
    directory: str | Path,
    config: UploadConfig | None = None,
    *,
    patterns: Iterable[str] | None = None,
    recursive: bool = False,
    dry_run: bool = False,
) -> list[UploadResult]:
    """上传目录下所有受支持文档。"""
    settings = (config or UploadConfig()).resolve()
    files = iter_input_files(directory, recursive=recursive, patterns=patterns)
    return upload_paths(files, settings, dry_run=dry_run)


def summarize(results: Iterable[UploadResult]) -> str:
    """汇总上传结果。"""
    items = list(results)
    ok = sum(1 for item in items if item.ok)
    head = f"上传预演：{len(items)} 个文件" if items and items[0].dry_run else f"上传完成：成功 {ok}/{len(items)}"
    lines = [head]
    lines.extend(f"  {item.summary()}" for item in items)
    return "\n".join(lines)


def _looks_like_placeholder(value: str) -> bool:
    """判断配置值是否还是占位符。"""
    if not value:
        return True
    text = str(value)
    return any(marker in text for marker in _PLACEHOLDER_MARKERS)


def supported_suffixes_for_upload() -> list[str]:
    """可上传的文件类型（与读取支持的格式一致）。"""
    return sorted(INPUT_SUFFIXES)
