"""
FastAPI 路由：上传文档、触发审核、下载修订后的文件。
"""

import json
import shutil
import time
import uuid
from dataclasses import asdict
from pathlib import Path
from datetime import datetime

from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from backend.core.config import get_config, load_config
from backend.core.models import (
    Issue, IssueCategory, IssueSeverity, ReviewResult, TaskStatus,
)
from backend.review.processor import ReviewProcessor
from backend.revision.applier import apply_revisions as apply_revisions_com
from backend.revision.applier import get_com_queue


# ---------------------------------------------------------------------------
# 任务状态持久化（JSON 文件，重启不丢任务状态）
# ---------------------------------------------------------------------------

def _issue_to_dict(i: Issue) -> dict:
    d = asdict(i)
    d["category"] = i.category.value
    d["severity"] = i.severity.value
    return d


def _issue_from_dict(d: dict) -> Issue:
    d = dict(d)
    d["category"] = IssueCategory(d["category"])
    d["severity"] = IssueSeverity(d["severity"])
    return Issue(**d)


def _result_to_dict(r: ReviewResult) -> dict:
    return {
        "task_id": r.task_id,
        "file_name": r.file_name,
        "issues": [_issue_to_dict(i) for i in r.issues],
        "status": r.status,
        "error_message": r.error_message,
        "started_at": r.started_at.isoformat() if r.started_at else None,
        "finished_at": r.finished_at.isoformat() if r.finished_at else None,
    }


def _result_from_dict(d: dict) -> ReviewResult:
    return ReviewResult(
        task_id=d["task_id"],
        file_name=d["file_name"],
        issues=[_issue_from_dict(i) for i in d.get("issues", [])],
        status=d["status"],
        error_message=d.get("error_message"),
        started_at=datetime.fromisoformat(d["started_at"]) if d.get("started_at") else None,
        finished_at=datetime.fromisoformat(d["finished_at"]) if d.get("finished_at") else None,
    )


class TaskStore:
    """任务状态存储：内存 + JSON 文件持久化。"""

    def __init__(self, path: Path):
        self.path = path
        self.data: dict[str, ReviewResult] = {}
        self._load()

    def _load(self):
        if not self.path.exists():
            return
        try:
            raw = json.loads(self.path.read_text(encoding="utf-8"))
            self.data = {
                k: _result_from_dict(v) for k, v in raw.items()
            }
        except (OSError, json.JSONDecodeError, TypeError, ValueError):
            self.data = {}

    def save(self):
        try:
            self.path.write_text(
                json.dumps(
                    {k: _result_to_dict(v) for k, v in self.data.items()},
                    ensure_ascii=False, indent=2,
                ),
                encoding="utf-8",
            )
        except OSError:
            pass  # 持久化失败不影响主流程

    def __contains__(self, key):
        return key in self.data

    def __getitem__(self, key):
        return self.data[key]

    def __setitem__(self, key, value):
        self.data[key] = value
        self.save()

    def prune(self, days: int):
        """清理超过保留天数的已完成/出错任务。"""
        cutoff = time.time() - days * 86400
        stale = []
        for task_id, r in self.data.items():
            ts = r.finished_at or r.started_at
            if ts is not None and ts.timestamp() < cutoff and r.status in ("done", "error"):
                stale.append(task_id)
        for task_id in stale:
            del self.data[task_id]
        if stale:
            self.save()

    def reset_transient_states(self):
        """服务重启时，将瞬态状态（后台线程已消失）回退到稳定状态。"""
        transitions = {
            "queued": "reviewed",    # 审核已完成，COM 锁丢失，重试即可
            "applying": "reviewed",  # COM 操作中断，重试即可
            "processing": "uploaded", # 审核中断，需重新触发
        }
        fixed = []
        for task_id, r in self.data.items():
            if r.status in transitions:
                r.status = transitions[r.status]
                fixed.append(task_id)
        if fixed:
            self.save()


# ---------------------------------------------------------------------------
# 过期文件清理
# ---------------------------------------------------------------------------

def cleanup_old_files(config, days: int) -> int:
    """删除 uploads/output 中超过保留天数的文件，返回删除数量。"""
    cutoff = time.time() - days * 86400
    removed = 0
    for d in (Path(config.server.upload_dir), Path(config.server.output_dir)):
        if not d.exists():
            continue
        for f in d.iterdir():
            try:
                if f.is_file() and f.stat().st_mtime < cutoff:
                    f.unlink()
                    removed += 1
            except OSError:
                pass
    return removed


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

    # 任务状态存储（内存 + JSON 持久化，重启不丢）
    task_store = TaskStore(output_dir / "task_store.json")

    # 启动时清理过期文件与过期任务
    cleanup_old_files(config, config.server.cleanup_days)
    task_store.prune(config.server.cleanup_days)
    task_store.reset_transient_states()

    # 前端静态文件
    frontend_dir = Path(__file__).resolve().parent.parent.parent / "frontend"
    static_dir = frontend_dir / "static"
    templates_dir = frontend_dir / "templates"

    if static_dir.exists():
        app.mount("/static", StaticFiles(directory=str(static_dir)), name="static")

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
        task_store.save()

        # 后台执行审核
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
                task_store.save()

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

        output_dir = Path(config.server.output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / f"{task_id}_revised_{result.file_name}"

        # 入队前先持久化，避免前端误以为"未排队"
        result.status = "queued"
        result.finished_at = None
        result.error_message = None
        task_store.save()

        def com_work():
            """由 COM worker 线程执行（串行）。"""
            try:
                result.status = "applying"
                task_store.save()
                outcome = apply_revisions_com(
                    doc_path, result.issues, output_path, config.com
                )
                stats = outcome.get("stats", {})
                print(f"[apply] {task_id} stats={stats}")
                result.status = "done"
            except Exception as e:
                result.status = "error"
                result.error_message = f"应用修订失败: {e}"
            finally:
                result.finished_at = datetime.now()
                task_store.save()

        com_queue = get_com_queue()
        com_queue.enqueue(task_id, result.file_name, com_work)

        return JSONResponse({
            "task_id": task_id,
            "status": "queued",
            "message": "已加入修订队列，等待 WPS/Word COM 资源空闲后自动执行...",
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
        task_store.save()

        output_dir = Path(config.server.output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / f"{task_id}_revised_{result.file_name}"

        def run_all():
            try:
                # 1. 审核
                processor = ReviewProcessor(config)
                review_result = processor.review(doc_path)
                result.issues = review_result.issues
                task_store.save()

                if not result.issues:
                    result.status = "done"
                    shutil.copy(doc_path, output_path)
                    result.finished_at = datetime.now()
                    task_store.save()
                    return

                # 2. 审核完成，入队等待 COM worker 执行修订
                result.status = "queued"
                task_store.save()

                def com_work():
                    """COM 修订（queue worker 串行执行）"""
                    try:
                        result.status = "applying"
                        task_store.save()
                        outcome = apply_revisions_com(
                            doc_path, result.issues, output_path, config.com
                        )
                        stats = outcome.get("stats", {})
                        print(f"[apply] {task_id} stats={stats}")
                        result.status = "done"
                    except Exception as e:
                        result.status = "error"
                        result.error_message = str(e)
                    finally:
                        result.finished_at = datetime.now()
                        task_store.save()

                com_queue = get_com_queue()
                com_queue.enqueue(task_id, result.file_name, com_work)

            except Exception as e:
                result.status = "error"
                result.error_message = str(e)
                result.finished_at = datetime.now()
                task_store.save()

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

    @app.get("/api/queue")
    async def get_queue():
        """查询 COM 修订队列（当前执行中 + 排队中）"""
        return get_com_queue().get_state()

    @app.post("/api/cancel/{task_id}")
    async def cancel_queued_task(task_id: str):
        """取消排队中的修订任务（已在执行中的无法取消）"""
        if task_id not in task_store:
            raise HTTPException(404, "任务不存在")

        result = task_store[task_id]
        if result.status != "queued":
            raise HTTPException(400, f"当前状态 '{result.status}' 不可取消")

        # 尝试从队列移除
        if get_com_queue().cancel(task_id):
            result.status = "reviewed"  # 回退到"审核完成"，可重新申请修订
            result.error_message = "已由用户取消排队"
            task_store.save()
            return JSONResponse({
                "task_id": task_id,
                "status": "reviewed",
                "message": "已取消排队",
            })

        # 已不在排队队列里（可能正在执行或已执行完）
        return JSONResponse({
            "task_id": task_id,
            "status": result.status,
            "message": "任务已在执行中或已完成，无法取消",
        })

    @app.get("/api/health")
    async def health():
        """健康检查：实际探测 WPS/Word COM 是否可用（不只检查 pywin32）"""
        wps_available = False
        wps_detail = ""
        try:
            from backend.revision.applier import ComRevisionApplier
            applier = ComRevisionApplier(config.com)
            prog_id = applier.connect()
            wps_available = True
            wps_detail = prog_id
            applier.close()
        except Exception as e:
            wps_detail = str(e)

        return {
            "status": "ok",
            "wps_com_available": wps_available,
            "wps_prog_id": wps_detail if wps_available else None,
            "wps_error": None if wps_available else wps_detail,
        }

    return app
