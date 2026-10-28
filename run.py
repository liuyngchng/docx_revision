"""
党政机关公文格式审核与修订系统 — 启动入口

用法：
    python run.py                 # 使用默认 cfg.yml
    python run.py --config 自定义.yml   # 指定配置文件
    python run.py --port 9090     # 覆盖端口

启动后浏览器访问 http://127.0.0.1:8080
"""

import argparse
import sys
import webbrowser
from pathlib import Path

import uvicorn


def parse_args():
    parser = argparse.ArgumentParser(
        description="党政机关公文格式审核与修订系统"
    )
    parser.add_argument(
        "--config", "-c",
        default=None,
        help="配置文件路径（默认查找 cfg.yml）",
    )
    parser.add_argument(
        "--port", "-p",
        type=int,
        default=None,
        help="服务端口（覆盖 cfg.yml 中的配置）",
    )
    parser.add_argument(
        "--host",
        default=None,
        help="监听地址（默认 127.0.0.1）",
    )
    parser.add_argument(
        "--no-browser",
        action="store_true",
        help="启动时不自动打开浏览器",
    )
    return parser.parse_args()


def main():
    args = parse_args()

    # 加载配置以获取端口等信息
    from backend.core.config import load_config
    from backend.api.routes import create_app

    config = load_config(args.config)

    host = args.host or config.server.host
    port = args.port or config.server.port

    app = create_app(args.config)

    if not args.no_browser:
        url = f"http://{host}:{port}"
        # 延迟打开浏览器，等服务器起来
        import threading
        threading.Timer(1.5, lambda: webbrowser.open(url)).start()

    print("=" * 60)
    print("  党政机关公文格式审核与修订系统")
    print("  GB/T 9704 — 国家标准")
    print("=" * 60)
    print(f"  监听地址: http://{host}:{port}")
    print(f"  配置文件: {args.config or 'cfg.yml'}")
    print(f"  COM ProgID: {config.com.prog_id}")
    print(f"  LLM API: {config.llm.api_base}")
    print(f"  模型: {config.llm.model}")
    print("=" * 60)
    print("  按 Ctrl+C 停止服务")
    print()

    uvicorn.run(app, host=host, port=port, log_level="info")


if __name__ == "__main__":
    main()