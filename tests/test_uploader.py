"""自动上传测试（全部离线，不产生真实网络请求）。"""

from __future__ import annotations

import base64
import json
import urllib.error
from pathlib import Path

import pytest

from information_mapper.uploader import (
    DEFAULT_API_TOKEN,
    DEFAULT_UPLOAD_URL,
    UploadConfig,
    build_json_body,
    build_multipart_body,
    build_request,
    summarize,
    upload_file,
    upload_folder,
)

REAL_URL = "https://upload.gov.cn/api/files"  # 测试用的"非占位符"地址
REAL_TOKEN = "token-abc-123"


def test_default_config_is_placeholder() -> None:
    config = UploadConfig()
    assert config.url == DEFAULT_UPLOAD_URL
    assert config.token == DEFAULT_API_TOKEN
    assert set(config.placeholders()) == {"url", "token"}


def test_validate_rejects_placeholder() -> None:
    with pytest.raises(ValueError, match="占位符"):
        UploadConfig().validate()


def test_validate_accepts_real_values() -> None:
    UploadConfig(url=REAL_URL, token=REAL_TOKEN).validate()


def test_build_multipart_body_contains_file(xlsx_file: Path) -> None:
    body, content_type = build_multipart_body(
        xlsx_file, field_name="file", extra_fields={"team": "码上助基层"}
    )
    assert content_type.startswith("multipart/form-data; boundary=")
    assert xlsx_file.name.encode("utf-8") in body
    assert b'name="team"' in body
    assert xlsx_file.read_bytes()[:16] in body


def test_build_json_body_is_base64(xlsx_file: Path) -> None:
    body, content_type = build_json_body(xlsx_file, extra_fields={"team": "码上助基层"})
    payload = json.loads(body)
    assert payload["filename"] == xlsx_file.name
    assert payload["team"] == "码上助基层"
    assert base64.b64decode(payload["content_base64"]) == xlsx_file.read_bytes()
    assert "application/json" in content_type


def test_build_request_sets_headers(xlsx_file: Path) -> None:
    config = UploadConfig(url=REAL_URL, token=REAL_TOKEN, method="json")
    request = build_request(xlsx_file, config)
    assert request.method == "POST"
    assert request.get_header("Authorization") == f"Bearer {REAL_TOKEN}"
    assert request.get_header("Content-length") == str(len(request.data))


def test_build_request_rejects_missing_file(tmp_path: Path) -> None:
    with pytest.raises(ValueError, match="找不到待上传文件"):
        build_request(tmp_path / "不存在.xlsx", UploadConfig())


def test_upload_dry_run_does_not_send(xlsx_file: Path) -> None:
    result = upload_file(xlsx_file, UploadConfig(), dry_run=True)
    assert result.ok
    assert result.dry_run
    assert result.bytes_sent > 0
    assert "预演" in result.summary()


def test_upload_without_dry_run_reports_placeholder(xlsx_file: Path) -> None:
    result = upload_file(xlsx_file, UploadConfig())
    assert not result.ok
    assert "占位符" in result.message


def test_upload_retries_on_server_error(monkeypatch: pytest.MonkeyPatch, xlsx_file: Path) -> None:
    attempts = {"count": 0}

    def fake_urlopen(request, timeout=None, context=None):  # noqa: ANN001
        attempts["count"] += 1
        raise urllib.error.HTTPError(request.full_url, 500, "Server Error", {}, None)

    monkeypatch.setattr("information_mapper.uploader.urllib.request.urlopen", fake_urlopen)
    config = UploadConfig(url=REAL_URL, token=REAL_TOKEN, retries=2, retry_delay=0)

    result = upload_file(xlsx_file, config)

    assert not result.ok
    assert attempts["count"] == 2
    assert result.attempts == 2
    assert "HTTP 500" in result.message


def test_upload_stops_on_client_error(monkeypatch: pytest.MonkeyPatch, xlsx_file: Path) -> None:
    attempts = {"count": 0}

    def fake_urlopen(request, timeout=None, context=None):  # noqa: ANN001
        attempts["count"] += 1
        raise urllib.error.HTTPError(request.full_url, 401, "Unauthorized", {}, None)

    monkeypatch.setattr("information_mapper.uploader.urllib.request.urlopen", fake_urlopen)
    config = UploadConfig(url=REAL_URL, token=REAL_TOKEN, retries=3, retry_delay=0)

    result = upload_file(xlsx_file, config)

    assert not result.ok
    assert attempts["count"] == 1  # 4xx 不重试


def test_upload_success(monkeypatch: pytest.MonkeyPatch, xlsx_file: Path) -> None:
    class FakeResponse:
        status = 201

        def read(self, size: int = -1) -> bytes:
            return b'{"id": "abc"}'

        def __enter__(self) -> "FakeResponse":
            return self

        def __exit__(self, *args: object) -> bool:
            return False

    monkeypatch.setattr(
        "information_mapper.uploader.urllib.request.urlopen", lambda *a, **k: FakeResponse()
    )
    result = upload_file(xlsx_file, UploadConfig(url=REAL_URL, token=REAL_TOKEN))
    assert result.ok
    assert result.status == 201
    assert "abc" in result.message


def test_env_variables_override_config(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("INFO_MAP_UPLOAD_URL", REAL_URL)
    monkeypatch.setenv("INFO_MAP_UPLOAD_TOKEN", REAL_TOKEN)
    resolved = UploadConfig().resolve()
    assert resolved.url == REAL_URL
    assert resolved.placeholders() == []


def test_upload_folder_dry_run(inbox: Path) -> None:
    results = upload_folder(inbox, UploadConfig(), dry_run=True)
    assert results
    assert all(item.dry_run and item.ok for item in results)
    assert "预演" in summarize(results)


def test_config_from_file(tmp_path: Path) -> None:
    path = tmp_path / "upload.json"
    path.write_text(
        json.dumps({"url": REAL_URL, "token": REAL_TOKEN, "method": "json", "retries": 5}),
        encoding="utf-8",
    )
    config = UploadConfig.from_file(path)
    assert config.method == "json"
    assert config.retries == 5
    config.validate()
