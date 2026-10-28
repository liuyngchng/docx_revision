"""
FastAPI 路由：上传文档、触发审核、下载修订后的文件。
"""

import os
import shutil
import uuid
from pathlib import Path
from datetime import datetime

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from backend.core.config import get_config, load_config
from backend.core.models import ReviewResult, TaskStatus
from backend.review.processor import ReviewProcessor
from backend.revision.applier import apply_revisions as apply_revisions_com


# ---------------------------------------------------------------------------
# 应用初始化
# ---------------------------------------------------------------------------

def create_app(config_path: str | None = None) -> FastAPI:
    """创建并配置 FastAPI 应用"""
    config = load_config(config_path)

    app = FastAPI(
        title="党政机关公文格式审核与修订系统",
        description="基于 GB/T 9704 标准的公文审核与修订工具",
        version="0.1.0",
    )

    # 确保目录存在
    upload_dir = Path(config.server.upload_dir)
    output_dir = Path(config.server.output_dir)
    upload_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    # 前端静态文件
    frontend_dir = Path(__file__).resolve().parent.parent.parent / "frontend"
    static_dir = frontend_dir / "static"
    templates_dir = frontend_dir / "templates"

    if static_dir.exists():
        app.mount("/static", StaticFiles(directory=str(static_dir)), name="static")

    # 任务状态存储（内存，后续可换 Redis）
    task_store: dict[str, ReviewResult] = {}

    # ------------------------------------------------------------------
    # 路由
    # ------------------------------------------------------------------

    @app.get("/", response_class=HTMLResponse)
    async def index():
        """前端首页"""
        index_path = templates_dir / "index.html"
        if index_path.exists():
            return index_path.read_text(encoding="utf-8")
        return HTMLResponse("<h1>公文审核系统</h1><p>前端页面未找到</p>")

    @app.post("/api/upload")
    async def upload_document(file: UploadFile = File(...)):
        """上传 docx 文档"""
        if not file.filename or not file.filename.lower().endswith(".docx"):
            raise HTTPException(400, "仅支持 .docx 格式的文档")

        # 保存上传文件
        file_id = str(uuid.uuid4())[:12]
        safe_name = f"{file_id}_{Path(file.filename).name}"
        upload_path = Path(config.server.upload_dir) / safe_name
        upload_path.parent.mkdir(parents=True, exist_ok=True)

        content = await file.read()

        if len(content) > config.server.max_upload_mb * 1024 * 1024:
            raise HTTPException(
                400,
                f"文件大小超过限制（{config.server.max_upload_mb}MB）"
            )

        with open(upload_path, "wb") as f:
            f.write(content)

        # 创建审核任务
        result = ReviewResult(
            task_id=file_id,
            file_name=file.filename,
            status="uploaded",
        )
        task_store[file_id] = result

        return JSONResponse({
            "task_id": file_id,
            "file_name": file.filename,
            "status": "uploaded",
        })

    @app.post("/api/review/{task_id}")
    async def start_review(task_id: str, background_tasks: BackgroundTasks):
        """启动审核任务"""
        if task_id not in task_store:
            raise HTTPException(404, "任务不存在")

        result = task_store[task_id]

        # 找到上传文件
        upload_dir = Path(config.server.upload_dir)
        files = list(upload_dir.glob(f"{task_id}_*"))
        if not files:
            raise HTTPException(404, "上传文件未找到，请重新上传")
        doc_path = files[0]

        result.status = "processing"
        result.started_at = datetime.now()

        # 同步执行审核（后台任务）
        def run_review():
            try:
                processor = ReviewProcessor(config)
                review_result = processor.review(doc_path)
                result.issues = review_result.issues
                result.status = "reviewed"  # 审核完成，等待应用修订
            except Exception as e:
                result.status = "error"
                result.error_message = str(e)
            finally:
                result.finished_at = datetime.now()

        background_tasks.add_task(run_review)

        return JSONResponse({
            "task_id": task_id,
            "status": "processing",
            "message": "审核任务已启动",
        })

    @app.get("/api/status/{task_id}")
    async def get_status(task_id: str):
        """查询任务状态"""
        if task_id not in task_store:
            raise HTTPException(404, "任务不存在")

        result = task_store[task_id]

        # 检查是否有已生成的修订文件
        output_dir = Path(config.server.output_dir)
        rev_files = list(output_dir.glob(f"{task_id}_*"))
        download_url = None
        if rev_files:
            download_url = f"/api/download/{task_id}"

        return TaskStatus(
            task_id=result.task_id,
            status=result.status,
            file_name=result.file_name,
            total_issues=result.total_issues,
            error_count=result.error_count,
            warning_count=result.warning_count,
            info_count=result.info_count,
            issues=result.issues,
            download_url=download_url,
            error_message=result.error_message,
        )

    @app.post("/api/apply/{task_id}")
    async def apply_revisions(task_id: str, background_tasks: BackgroundTasks):
        """应用修订，生成修订后文档"""
        if task_id not in task_store:
            raise HTTPException(404, "任务不存在")

        result = task_store[task_id]
        if result.status not in ("reviewed", "pending"):
            raise HTTPException(400, f"当前状态 '{result.status}' 不允许应用修订")

        if not result.issues:
            raise HTTPException(400, "没有发现审核问题，无需应用修订")

        upload_dir = Path(config.server.upload_dir)
        files = list(upload_dir.glob(f"{task_id}_*"))
        if not files:
            raise HTTPException(404, "原始文件未找到")
        doc_path = files[0]

        result.status = "applying"

        output_dir = Path(config.server.output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / f"{task_id}_revised_{result.file_name}"

        def run_apply():
            try:
                apply_revisions_com(
                    doc_path, result.issues, output_path, config.com
                )
                result.status = "done"
            except Exception as e:
                result.status = "error"
                result.error_message = f"应用修订失败: {e}"
            finally:
                result.finished_at = datetime.now()

        background_tasks.add_task(run_apply)

        return JSONResponse({
            "task_id": task_id,
            "status": "applying",
            "message": "正在应用修订...",
        })

    @app.post("/api/review-and-apply/{task_id}")
    async def review_and_apply(task_id: str, background_tasks: BackgroundTasks):
        """一步完成：审核 + 应用修订"""
        if task_id not in task_store:
            raise HTTPException(404, "任务不存在")

        result = task_store[task_id]

        upload_dir = Path(config.server.upload_dir)
        files = list(upload_dir.glob(f"{task_id}_*"))
        if not files:
            raise HTTPException(404, "上传文件未找到，请重新上传")
        doc_path = files[0]

        result.status = "processing"
        result.started_at = datetime.now()

        output_dir = Path(config.server.output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / f"{task_id}_revised_{result.file_name}"

        def run_all():
            try:
                # 1. 审核
                processor = ReviewProcessor(config)
                review_result = processor.review(doc_path)
                result.issues = review_result.issues

                if not result.issues:
                    result.status = "done"
                    shutil.copy(doc_path, output_path)
                    return

                # 2. 通过 COM 以修订模式应用修改
                result.status = "applying"
                apply_revisions_com(
                    doc_path, result.issues, output_path, config.com
                )
                result.status = "done"
            except Exception as e:
                result.status = "error"
                result.error_message = str(e)
            finally:
                result.finished_at = datetime.now()

        background_tasks.add_task(run_all)

        return JSONResponse({
            "task_id": task_id,
            "status": "processing",
            "message": "审核和修订任务已启动",
        })

    @app.get("/api/download/{task_id}")
    async def download_revised(task_id: str):
        """下载修订后的文档"""
        if task_id not in task_store:
            raise HTTPException(404, "任务不存在")

        result = task_store[task_id]
        if result.status != "done":
            raise HTTPException(400, f"任务未完成，当前状态: {result.status}")

        output_dir = Path(config.server.output_dir)
        files = list(output_dir.glob(f"{task_id}_*"))
        if not files:
            raise HTTPException(404, "修订后的文件未找到")

        file_path = files[0]
        safe_name = f"修订_{result.file_name}"

        return FileResponse(
            path=str(file_path),
            filename=safe_name,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    @app.get("/api/health")
    async def health():
        """健康检查"""
        wps_available = False
        try:
            from backend.revision.applier import ComRevisionApplier
            applier = ComRevisionApplier(config.com)
            applier._ensure_com_available()
            wps_available = True
        except Exception:
            pass

        return {
            "status": "ok",
            "wps_com_available": wps_available,
        }

    return app