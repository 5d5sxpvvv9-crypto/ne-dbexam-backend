"""HWP 파싱 테스트 스크립트"""
import os
import sys
import glob
import logging

sys.path.insert(0, os.path.dirname(__file__))
logging.basicConfig(level=logging.DEBUG, handlers=[
    logging.StreamHandler(sys.stdout)
])

# stdout UTF-8 설정
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from hwp_parser import extract_text_from_hwp
from question_extractor import extract_questions

# 상위 디렉토리에서 HWP 파일 찾기
parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
hwp_files = glob.glob(os.path.join(parent_dir, "*.hwp"))

if not hwp_files:
    print("HWP 파일을 찾을 수 없습니다.")
    sys.exit(1)

for hwp_file in hwp_files:
    print(f"\n{'='*60}")
    print(f"파일: {os.path.basename(hwp_file)}")
    print(f"{'='*60}")

    result = extract_text_from_hwp(hwp_file)
    print(f"성공: {result.success}")
    print(f"방법: {result.method_used}")
    print(f"시간: {result.parse_time_ms}ms")
    print(f"블록 수: {len(result.text_blocks)}")

    if not result.success:
        print(f"오류: {result.error}")
        continue

    print(f"\n--- 텍스트 미리보기 (처음 2000자) ---")
    print(result.full_text[:2000])

    print(f"\n--- 문항 추출 ---")
    extraction = extract_questions(result.full_text, os.path.basename(hwp_file))
    print(f"총 문항: {extraction.total_questions}")
    print(f"정답 섹션 발견: {extraction.answer_section_found}")

    for q in extraction.questions:
        print(f"\n  문항 {q.question_number}:")
        print(f"    질문: {q.question_text[:80]}...")
        print(f"    지문: {q.reading_passage[:60]}..." if q.reading_passage else "    지문: (없음)")
        print(f"    보기: {q.choices[:60]}..." if q.choices else "    보기: (없음)")
        print(f"    정답: {q.answer}")
        print(f"    유형: {q.question_type}/{q.type_detail} (신뢰도: {q.confidence:.1f})")

