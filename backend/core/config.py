"""
配置加载模块
从 cfg.yml 读取所有配置，提供类型安全的访问接口。
"""

import os
from pathlib import Path
from dataclasses import dataclass, field
from typing import Optional

import yaml


# ---------------------------------------------------------------------------
# 配置数据结构
# ---------------------------------------------------------------------------

@dataclass
class ComConfig:
    """WPS/Word COM 相关配置"""
    prog_id: str = "KWPS.Application"
    visible: bool = False
    display_alerts: bool = False

    @property
    def prog_ids(self) -> list[str]:
        """返回 ProgID 列表（支持字符串或列表配置）"""
        if isinstance(self.prog_id, str):
            return [self.prog_id]
        return self.prog_id


@dataclass
class LlmConfig:
    """LLM API 相关配置"""
    api_base: str = "https://your-company-api.example.com/v1"
    api_key: str = ""
    model: str = "your-model-name"
    timeout: int = 120
    temperature: float = 0.1
    max_tokens: int = 4096


@dataclass
class ReviewConfig:
    """审核相关配置"""
    enable_rules: bool = True
    enable_llm: bool = True
    strictness: str = "normal"  # strict / normal / loose


@dataclass
class ServerConfig:
    """服务相关配置"""
    host: str = "127.0.0.1"
    port: int = 8080
    max_upload_mb: int = 20
    output_dir: str = "./output"
    upload_dir: str = "./uploads"
    cleanup_days: int = 7


@dataclass
class AppConfig:
    """总配置"""
    com: ComConfig = field(default_factory=ComConfig)
    llm: LlmConfig = field(default_factory=LlmConfig)
    review: ReviewConfig = field(default_factory=ReviewConfig)
    server: ServerConfig = field(default_factory=ServerConfig)


# ---------------------------------------------------------------------------
# 配置加载
# ---------------------------------------------------------------------------

_config: Optional[AppConfig] = None


def _find_config_file() -> Path:
    """查找 cfg.yml，按以下顺序：
    1. 环境变量 DOCX_REVIEW_CONFIG
    2. 当前工作目录
    3. 本文件所在目录的上级（项目根目录）
    """
    env_path = os.environ.get("DOCX_REVIEW_CONFIG")
    if env_path and os.path.isfile(env_path):
        return Path(env_path)

    cwd_path = Path.cwd() / "cfg.yml"
    if cwd_path.exists():
        return cwd_path

    # 相对于本文件的上级目录
    this_dir = Path(__file__).resolve().parent.parent.parent
    project_path = this_dir / "cfg.yml"
    if project_path.exists():
        return project_path

    raise FileNotFoundError(
        "找不到 cfg.yml 配置文件。"
        "请确保 cfg.yml 位于工作目录或项目根目录，"
        "或设置环境变量 DOCX_REVIEW_CONFIG 指向配置文件。"
    )


def load_config(config_path: Optional[str | Path] = None) -> AppConfig:
    """加载并解析 cfg.yml 配置文件"""
    global _config

    if config_path is None:
        path = _find_config_file()
    else:
        path = Path(config_path)
        if not path.exists():
            raise FileNotFoundError(f"配置文件不存在: {path}")

    with open(path, "r", encoding="utf-8") as f:
        raw = yaml.safe_load(f)

    com_raw = raw.get("com", {})
    llm_raw = raw.get("llm", {})
    review_raw = raw.get("review", {})
    server_raw = raw.get("server", {})

    _config = AppConfig(
        com=ComConfig(
            prog_id=com_raw.get("prog_id", "WPS.Application"),
            visible=com_raw.get("visible", False),
            display_alerts=com_raw.get("display_alerts", False),
        ),
        llm=LlmConfig(
            api_base=llm_raw.get("api_base", ""),
            api_key=llm_raw.get("api_key", ""),
            model=llm_raw.get("model", ""),
            timeout=llm_raw.get("timeout", 120),
            temperature=llm_raw.get("temperature", 0.1),
            max_tokens=llm_raw.get("max_tokens", 4096),
        ),
        review=ReviewConfig(
            enable_rules=review_raw.get("enable_rules", True),
            enable_llm=review_raw.get("enable_llm", True),
            strictness=review_raw.get("strictness", "normal"),
        ),
        server=ServerConfig(
            host=server_raw.get("host", "127.0.0.1"),
            port=server_raw.get("port", 8080),
            max_upload_mb=server_raw.get("max_upload_mb", 20),
            output_dir=server_raw.get("output_dir", "./output"),
            upload_dir=server_raw.get("upload_dir", "./uploads"),
            cleanup_days=server_raw.get("cleanup_days", 7),
        ),
    )

    return _config


def get_config() -> AppConfig:
    """获取已加载的配置，未加载则自动加载"""
    global _config
    if _config is None:
        return load_config()
    return _config