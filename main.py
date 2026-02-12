"""
HWP 영어시험 문항 분석 서비스 - FastAPI 서버
업로드 → 파싱 → 문항 추출 → 엑셀 다운로드
"""

import os
import sys
import io
import uuid
import time
import shutil
import asyncio
import logging
import tempfile
from pathlib import Path
from typing import Dict, List, Optional
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor

from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware

# 모듈 임포트
sys.path.insert(0, os.path.dirname(__file__))
from hwp_parser import extract_text_from_hwp, HwpParseResult
from question_extractor import extract_questions, ExtractionResult, QuestionData
from excel_generator import generate_excel, generate_merged_excel

# ── 로깅 설정 ──
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
logger = logging.getLogger(__name__)

# ── 앱 설정 ──
TEMP_DIR = os.path.join(os.path.dirname(__file__), "temp_uploads")
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "temp_outputs")
MAX_FILE_SIZE = 50 * 1024 * 1024  # 50MB
ALLOWED_EXTENSIONS = {".hwp"}
MAX_CONCURRENT = 3
CLEANUP_MINUTES = 60

os.makedirs(TEMP_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ── FastAPI 앱 ──
app = FastAPI(
    title="HWP 영어시험 문항 분석기",
    description="HWP 기출문제 파일을 업로드하면 문항을 자동 분석하여 엑셀로 다운로드",
    version="1.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ── 상태 관리 ──
file_tasks: Dict[str, dict] = {}  # task_id → task info
executor = ThreadPoolExecutor(max_workers=MAX_CONCURRENT)
processing_semaphore = asyncio.Semaphore(MAX_CONCURRENT)


def _process_hwp_file(task_id: str, filepath: str, filename: str):
    """HWP 파일 처리 (스레드에서 실행)"""
    task = file_tasks[task_id]
    task["status"] = "processing"
    task["started_at"] = datetime.now().isoformat()

    try:
        # 1. HWP → 텍스트 추출
        logger.info(f"[{task_id}] 파싱 시작: {filename}")
        parse_result = extract_text_from_hwp(filepath)

        if not parse_result.success:
            task["status"] = "failed"
            task["error"] = f"HWP 파싱 실패: {parse_result.error}"
            logger.error(f"[{task_id}] 파싱 실패: {parse_result.error}")
            return

        task["parse_method"] = parse_result.method_used
        task["parse_time_ms"] = parse_result.parse_time_ms

        # 2. 텍스트 → 문항 추출
        logger.info(f"[{task_id}] 문항 추출 시작 (텍스트 {len(parse_result.full_text)}자)")
        extraction = extract_questions(parse_result.full_text, filename)

        task["total_questions"] = extraction.total_questions
        task["answer_section_found"] = extraction.answer_section_found
        task["warnings"] = extraction.warnings
        task["metadata"] = extraction.metadata

        # 3단계 검증 데이터
        task["expected_numbers"] = extraction.expected_numbers
        task["extracted_numbers"] = extraction.extracted_numbers
        task["missing_numbers"] = extraction.missing_numbers
        task["recovered_numbers"] = extraction.recovered_numbers
        task["verification_log"] = extraction.verification_log
        # v2.0 Validation / Repair 데이터
        task["detected_questions"] = extraction.detected_questions
        task["detected_answers"] = extraction.detected_answers
        task["missing_answer_numbers"] = extraction.missing_answer_numbers
        task["questions_without_answer"] = extraction.questions_without_answer
        task["suspicious_number_jumps"] = extraction.suspicious_number_jumps
        task["questions_without_question_text"] = extraction.questions_without_question_text
        task["repair_log"] = extraction.repair_log

        # 문항 데이터를 직렬화 가능한 형태로 저장 (9열 + v2.0 강화 필드)
        questions_data = []
        for q in extraction.questions:
            questions_data.append({
                "question_number": q.question_number,
                "question_text": q.question_text,
                "common_passage": q.common_passage,
                "question_passage": q.question_passage,
                "choices": q.choices,
                "answer": q.answer,
                "question_type": q.question_type,
                "confidence": q.confidence,
                "notes": q.notes,
                "passage_group_id": q.passage_group_id,
                "raw_block_text": q.raw_block_text[:500],
                "school": q.school,
                "grade": q.grade,
                # v2.0 강화 필드
                "seq_no": q.seq_no,
                "raw_no": q.raw_no,
                "answer_source": q.answer_source,
                "source_block_ids": q.source_block_ids,
                "item_warnings": q.item_warnings,
                "choices_list": q.choices_list,
            })

        task["questions"] = questions_data
        task["status"] = "completed"
        task["completed_at"] = datetime.now().isoformat()
        logger.info(f"[{task_id}] 완료: {extraction.total_questions}문항 추출")

    except Exception as e:
        task["status"] = "failed"
        task["error"] = str(e)
        logger.error(f"[{task_id}] 처리 오류: {e}", exc_info=True)
    finally:
        # 임시 파일 정리
        try:
            if os.path.exists(filepath):
                os.remove(filepath)
        except Exception:
            pass


# ── API 엔드포인트 ──

@app.post("/api/upload")
async def upload_files(files: List[UploadFile] = File(...)):
    """HWP 파일 다중 업로드"""
    results = []

    for file in files:
        # 확장자 검증
        ext = os.path.splitext(file.filename or "")[1].lower()
        if ext not in ALLOWED_EXTENSIONS:
            results.append({
                "filename": file.filename,
                "task_id": None,
                "status": "rejected",
                "error": f"지원하지 않는 파일 형식: {ext}. HWP 파일만 업로드 가능합니다.",
            })
            continue

        # 파일 크기 검증
        content = await file.read()
        if len(content) > MAX_FILE_SIZE:
            results.append({
                "filename": file.filename,
                "task_id": None,
                "status": "rejected",
                "error": f"파일 크기 초과: {len(content) / 1024 / 1024:.1f}MB (최대 50MB)",
            })
            continue

        # 임시 파일 저장
        task_id = str(uuid.uuid4())[:8]
        safe_filename = f"{task_id}_{file.filename}"
        filepath = os.path.join(TEMP_DIR, safe_filename)

        with open(filepath, "wb") as f:
            f.write(content)

        # 작업 등록
        file_tasks[task_id] = {
            "task_id": task_id,
            "filename": file.filename,
            "filepath": filepath,
            "status": "queued",
            "created_at": datetime.now().isoformat(),
            "total_questions": 0,
            "questions": [],
            "error": None,
            "warnings": [],
            "metadata": {},
        }

        # 비동기 처리 시작
        loop = asyncio.get_event_loop()
        loop.run_in_executor(executor, _process_hwp_file, task_id, filepath, file.filename)

        results.append({
            "filename": file.filename,
            "task_id": task_id,
            "status": "queued",
        })

    return {"files": results}


@app.get("/api/status/{task_id}")
async def get_task_status(task_id: str):
    """작업 상태 조회"""
    if task_id not in file_tasks:
        raise HTTPException(status_code=404, detail="작업을 찾을 수 없습니다.")

    task = file_tasks[task_id]

    # 민감 정보 제외
    return {
        "task_id": task["task_id"],
        "filename": task["filename"],
        "status": task["status"],
        "total_questions": task.get("total_questions", 0),
        "error": task.get("error"),
        "warnings": task.get("warnings", []),
        "metadata": task.get("metadata", {}),
        "parse_method": task.get("parse_method", ""),
        "parse_time_ms": task.get("parse_time_ms", 0),
        "created_at": task.get("created_at", ""),
        "completed_at": task.get("completed_at", ""),
    }


@app.get("/api/status")
async def get_all_status():
    """모든 작업 상태 조회"""
    statuses = []
    for task_id, task in file_tasks.items():
        statuses.append({
            "task_id": task["task_id"],
            "filename": task["filename"],
            "status": task["status"],
            "total_questions": task.get("total_questions", 0),
            "error": task.get("error"),
            "warnings": task.get("warnings", []),
            "created_at": task.get("created_at", ""),
            "completed_at": task.get("completed_at", ""),
        })
    return {"tasks": statuses}


@app.get("/api/questions/{task_id}")
async def get_questions(task_id: str):
    """문항 데이터 조회 (미리보기용)"""
    if task_id not in file_tasks:
        raise HTTPException(status_code=404, detail="작업을 찾을 수 없습니다.")

    task = file_tasks[task_id]

    if task["status"] != "completed":
        raise HTTPException(
            status_code=400,
            detail=f"아직 처리 중입니다. 현재 상태: {task['status']}"
        )

    return {
        "task_id": task_id,
        "filename": task["filename"],
        "total_questions": task["total_questions"],
        "metadata": task.get("metadata", {}),
        "questions": task["questions"],
        # 3단계 검증 데이터
        "expected_numbers": task.get("expected_numbers", []),
        "extracted_numbers": task.get("extracted_numbers", []),
        "missing_numbers": task.get("missing_numbers", []),
        "recovered_numbers": task.get("recovered_numbers", []),
        "verification_log": task.get("verification_log", ""),
        # v2.0 Validation / Repair 데이터
        "detected_questions": task.get("detected_questions", 0),
        "detected_answers": task.get("detected_answers", 0),
        "missing_answer_numbers": task.get("missing_answer_numbers", []),
        "questions_without_answer": task.get("questions_without_answer", []),
        "suspicious_number_jumps": task.get("suspicious_number_jumps", []),
        "questions_without_question_text": task.get("questions_without_question_text", []),
        "repair_log": task.get("repair_log", []),
    }


@app.post("/api/retry/{task_id}")
async def retry_task(task_id: str):
    """실패한 작업 재시도"""
    if task_id not in file_tasks:
        raise HTTPException(status_code=404, detail="작업을 찾을 수 없습니다.")

    task = file_tasks[task_id]

    if task["status"] not in ("failed",):
        raise HTTPException(status_code=400, detail="재시도 가능한 상태가 아닙니다.")

    filepath = task.get("filepath", "")
    if not os.path.exists(filepath):
        raise HTTPException(status_code=400, detail="원본 파일이 삭제되었습니다. 다시 업로드해주세요.")

    # 상태 초기화
    task["status"] = "queued"
    task["error"] = None
    task["warnings"] = []
    task["questions"] = []
    task["total_questions"] = 0

    loop = asyncio.get_event_loop()
    loop.run_in_executor(executor, _process_hwp_file, task_id, filepath, task["filename"])

    return {"task_id": task_id, "status": "queued", "message": "재시도가 시작되었습니다."}


@app.post("/api/export/excel")
async def export_excel(task_ids: Optional[List[str]] = None):
    """엑셀 다운로드 (선택된 작업들 또는 전체)"""
    # 완료된 작업만 필터
    if task_ids:
        target_tasks = [file_tasks[tid] for tid in task_ids if tid in file_tasks]
    else:
        target_tasks = list(file_tasks.values())

    completed_tasks = [t for t in target_tasks if t["status"] == "completed"]

    if not completed_tasks:
        raise HTTPException(status_code=400, detail="추출 완료된 파일이 없습니다.")

    # QuestionData 객체 복원 (9열 + v2.0 구조)
    all_questions = []
    for task in completed_tasks:
        for q_data in task.get("questions", []):
            q = QuestionData(
                question_number=q_data["question_number"],
                question_text=q_data["question_text"],
                common_passage=q_data.get("common_passage", ""),
                question_passage=q_data.get("question_passage", ""),
                choices=q_data["choices"],
                answer=q_data["answer"],
                question_type=q_data.get("question_type", ""),
                confidence=q_data.get("confidence", 0),
                notes=q_data.get("notes", ""),
                passage_group_id=q_data.get("passage_group_id"),
                school=q_data.get("school", ""),
                grade=q_data.get("grade", 0),
                # v2.0 강화 필드
                seq_no=q_data.get("seq_no", 0),
                raw_no=q_data.get("raw_no"),
                answer_source=q_data.get("answer_source", "missing"),
                source_block_ids=q_data.get("source_block_ids", []),
                item_warnings=q_data.get("item_warnings", []),
                choices_list=q_data.get("choices_list", []),
            )
            all_questions.append(q)

    if not all_questions:
        raise HTTPException(status_code=400, detail="추출된 문항이 없습니다.")

    # 엑셀 생성
    output_filename = f"exam_questions_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    output_path = os.path.join(OUTPUT_DIR, output_filename)

    try:
        generate_excel(all_questions, output_path)
    except Exception as e:
        logger.error(f"엑셀 생성 실패: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"엑셀 생성 실패: {str(e)}")

    return FileResponse(
        path=output_path,
        filename=output_filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.delete("/api/tasks/{task_id}")
async def delete_task(task_id: str):
    """작업 삭제"""
    if task_id not in file_tasks:
        raise HTTPException(status_code=404, detail="작업을 찾을 수 없습니다.")

    task = file_tasks.pop(task_id)

    # 관련 파일 정리
    filepath = task.get("filepath", "")
    if filepath and os.path.exists(filepath):
        try:
            os.remove(filepath)
        except:
            pass

    return {"message": f"작업 {task_id}이 삭제되었습니다."}


@app.delete("/api/tasks")
async def clear_all_tasks():
    """모든 작업 삭제"""
    for task_id in list(file_tasks.keys()):
        task = file_tasks.pop(task_id, {})
        filepath = task.get("filepath", "")
        if filepath and os.path.exists(filepath):
            try:
                os.remove(filepath)
            except:
                pass
    return {"message": "모든 작업이 삭제되었습니다."}


# ── 주기적 정리 ──
async def cleanup_old_files():
    """만료된 임시 파일 정리"""
    while True:
        await asyncio.sleep(300)  # 5분마다
        now = datetime.now()
        expired_ids = []

        for task_id, task in file_tasks.items():
            created = task.get("created_at", "")
            if created:
                try:
                    created_dt = datetime.fromisoformat(created)
                    if now - created_dt > timedelta(minutes=CLEANUP_MINUTES):
                        expired_ids.append(task_id)
                except:
                    pass

        for task_id in expired_ids:
            task = file_tasks.pop(task_id, {})
            filepath = task.get("filepath", "")
            if filepath and os.path.exists(filepath):
                try:
                    os.remove(filepath)
                except:
                    pass
            logger.info(f"만료 작업 정리: {task_id}")

    # 출력 디렉토리 정리
    for f in os.listdir(OUTPUT_DIR):
        fpath = os.path.join(OUTPUT_DIR, f)
        try:
            if os.path.isfile(fpath):
                age = time.time() - os.path.getmtime(fpath)
                if age > CLEANUP_MINUTES * 60:
                    os.remove(fpath)
        except:
            pass


@app.on_event("startup")
async def startup_event():
    asyncio.create_task(cleanup_old_files())
    logger.info("HWP 분석 서버 시작")
    logger.info(f"임시 디렉토리: {TEMP_DIR}")
    logger.info(f"출력 디렉토리: {OUTPUT_DIR}")


@app.get("/api/health")
async def health_check():
    """서버 상태 확인"""
    # COM 사용 가능 여부 확인
    com_available = False
    try:
        import win32com.client
        com_available = True
    except ImportError:
        pass

    return {
        "status": "ok",
        "com_available": com_available,
        "active_tasks": len(file_tasks),
        "temp_dir": TEMP_DIR,
    }


# ── 서버 실행 ──
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(
        "main:app",
        host="0.0.0.0",
        port=8000,
        reload=True,
        log_level="info",
    )

