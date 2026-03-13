"""
영어시험 문항 분리/분류 모듈 (v5 — 강화 지시서 v2.0 반영)

3단계 프로세스 + Validation + Repair 로 문항 누락 0건 보장

[1단계] 전체 문항 번호 스캔 — 시험지의 모든 문항 번호를 사전 탐지
[2단계] 상세 파싱 — 각 문항의 출제문항/지문/보기/정답 추출
[3단계] 검증 및 복구 — 누락 문항 확인 → 복구 시도 → 검증 리포트 생성
[Validation] 품질 검증 — 번호 연속성, 정답 매핑, 최소 품질 기준
[Repair] 강제 복구 — 정답키 기반 슬롯 생성, 선지 기반 추가 탐지

최우선 원칙: 문항 누락 절대 금지
  - 공통지문 공유 문항도 각각 별도 행(row)
  - 서술형 문항 반드시 포함 (보기만 비움)
  - "윗글~" 문항도 별도 행
  - 번호 점프 허용 (경고만 출력)
  - 과분할보다 과포함 허용 (누락 방지 우선)

출력 열 (고정 9열):
  문제번호 | 학교 | 학년 | 출제문항 | 공통지문 | 문제지문 | 보기/조건 | 정답 | 문제유형 | 문항형태
"""

import re
import os
import json
import yaml
import logging
from typing import List, Dict, Optional, Tuple, Set
from dataclasses import dataclass, field

logger = logging.getLogger(__name__)

# ── 규칙 설정 로드 ──
RULES_PATH = os.path.join(os.path.dirname(__file__), "rules", "config.yaml")


def load_rules() -> dict:
    try:
        with open(RULES_PATH, 'r', encoding='utf-8') as f:
            return yaml.safe_load(f)
    except Exception as e:
        logger.warning(f"규칙 파일 로드 실패: {e}, 기본값 사용")
        return {}


RULES = load_rules()


@dataclass
class QuestionData:
    """문항 구조화 데이터 (9열 대응 + v2.0 강화 필드)"""
    question_number: int = 0
    question_text: str = ""           # 출제문항
    common_passage: str = ""           # 공통지문 (여러 문항 공유)
    question_passage: str = ""         # 문제지문 (해당 문항만의 지문)
    choices: str = ""                  # 보기/조건 (서술형이면 비움)
    answer: str = ""                   # 정답
    question_type: str = ""            # 문제유형 (4가지 고정)
    question_format: str = ""          # 문항형태 (객관식 / 서술형)
    # ── 내부 메타데이터 ──
    confidence: float = 0.0
    notes: str = ""
    passage_group_id: Optional[int] = None
    raw_block_text: str = ""
    school: str = ""
    grade: int = 0
    # ── v2.0 강화 필드 ──
    seq_no: int = 0                                            # 문서 순서 기반 번호
    raw_no: Optional[int] = None                               # 텍스트 원본 번호 (raw)
    answer_source: str = "missing"                             # answer_key | inline | missing
    source_block_ids: List[int] = field(default_factory=list)  # 소스 라인 인덱스
    item_warnings: List[str] = field(default_factory=list)     # 문항별 경고
    choices_list: List[str] = field(default_factory=list)      # 선지 배열 (JSON 호환)


@dataclass
class ExtractionResult:
    """문항 추출 결과 (3단계 검증 + Validation/Repair 포함)"""
    questions: List[QuestionData] = field(default_factory=list)
    total_questions: int = 0
    answer_section_found: bool = False
    answer_section_raw: str = ""
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    metadata: Dict = field(default_factory=dict)
    # ── 3단계 검증 필드 ──
    expected_numbers: List[int] = field(default_factory=list)
    extracted_numbers: List[int] = field(default_factory=list)
    missing_numbers: List[int] = field(default_factory=list)
    recovered_numbers: List[int] = field(default_factory=list)
    verification_log: str = ""
    # ── v2.0 Validation / Repair 필드 ──
    detected_questions: int = 0
    detected_answers: int = 0
    missing_answer_numbers: List[int] = field(default_factory=list)
    questions_without_answer: List[int] = field(default_factory=list)
    suspicious_number_jumps: List[str] = field(default_factory=list)
    questions_without_question_text: List[int] = field(default_factory=list)
    repair_log: List[str] = field(default_factory=list)


# ────────────────────────────────────────────────
# v2.0 문항 시작 탐지 패턴
# ────────────────────────────────────────────────

PATTERN_Q_START = [
    re.compile(r'^\s*(\d{1,2})\s*\.\s+'),     # "N. "
    re.compile(r'^\s*(\d{1,2})\s*\)\s+'),      # "N) "
    re.compile(r'^\s*\((\d{1,2})\)\s+'),       # "(N) "
    re.compile(r'^\s*\[(\d{1,2})\]\s+'),       # "[N] "
]

# v2.0 선지 탐지 패턴
PATTERN_CHOICE = [
    re.compile(r'[①②③④⑤]'),
    re.compile(r'^\s*[1-5][.)]\s+'),
    re.compile(r'^\s*[A-Ea-e][.)]\s+'),
]


# ────────────────────────────────────────────────
# 종결 패턴 (하드코딩 + config.yaml 병합)
# ────────────────────────────────────────────────

_OBJECTIVE_STEMS = [
    '것은', '것', '문장은', '개수는', '질문은', '단어는',
    '곳은', '고르면', '인가', '않은가', '무엇인가',
    '알맞은가', '무엇', '누구', '어디', '올바른가',
    '적절한가', '부적절한가', '맞는가', '틀린가',
]

_SUBJECTIVE_STEMS = [
    '쓰시오', '완성하시오', '영작하시오', '답하시오',
    '채우시오', '서술하시오', '설명하시오', '구하시오',
    '고르시오', '나타내시오', '적으시오', '바꾸시오',
    '고치시오', '표현하시오', '변형하시오', '배열하시오',
    '작성하시오', '만드시오',
]


def _build_ending_patterns() -> Tuple[List[str], List[str]]:
    obj_pats = [re.escape(s) + r'\?' for s in _OBJECTIVE_STEMS]
    subj_pats = [re.escape(s) + r'\.' for s in _SUBJECTIVE_STEMS]
    for pat in RULES.get('question_text', {}).get('endings', []):
        if pat not in obj_pats:
            obj_pats.append(pat)
    for kw in RULES.get('question_text', {}).get('subjective_endings', []):
        p = re.escape(kw) if '\\' in kw else kw
        if p not in subj_pats:
            subj_pats.append(p)
    return obj_pats, subj_pats


OBJ_ENDING_RES, SUBJ_ENDING_RES = _build_ending_patterns()


# ────────────────────────────────────────────────
# 라인 전처리 유틸리티
# ────────────────────────────────────────────────

def _clean_line_suffixes(line: str) -> str:
    clean = re.sub(r'\s*\[[^\]]{1,20}\]\s*$', '', line).strip()
    clean = re.sub(r'\s*\(단[^)]*\)\s*$', '', clean).strip()
    clean = re.sub(r'\s*\(정답\s*\d+\s*개\)\s*$', '', clean).strip()
    clean = re.sub(r'\s*\([^)]*(?:하시오|[가-힣]\s*것)\)\s*$', '', clean).strip()
    if re.search(r'시오\.\s*\(', clean):
        clean = re.sub(r'\s*\([^)]+\)\s*$', '', clean).strip()
    clean = re.sub(r'(?<=[가-힣a-zA-Z?.!])\s+\d{1,2}\s*$', '', clean).strip()
    return clean


def _strip_leading_qnum(line: str) -> Tuple[Optional[int], str]:
    """문항 번호 제거: PATTERN_Q_START 전체 탐색 (v2.0 강화)"""
    for pat in PATTERN_Q_START:
        m = pat.match(line)
        if m:
            num = int(m.group(1))
            rest = line[m.end():].strip()
            return num, rest
    return None, line


# ────────────────────────────────────────────────
# 라인 판별 함수들
# ────────────────────────────────────────────────

def _is_passage_intro(line: str) -> bool:
    clean = _clean_line_suffixes(line)
    return ('물음에 답하시오' in clean) and ('다음' in clean or '읽고' in clean)


def _is_question_text(line: str) -> bool:
    if not line or not line.strip():
        return False
    stripped = line.strip()
    if _is_passage_intro(stripped):
        return False
    if _is_metadata_line(stripped):
        return False
    if re.match(r'^[•→\-·▶▷]', stripped):
        return False
    if re.match(r'^[①②③④⑤ⓐⓑⓒⓓⓔ]', stripped):
        return False
    if re.match(r'^\(\d+\)\s', stripped):
        return False
    if len(stripped) <= 5:
        return False
    clean = _clean_line_suffixes(stripped)
    _, clean = _strip_leading_qnum(clean)
    if not clean:
        return False
    for pat in OBJ_ENDING_RES:
        if re.search(pat + r'\s*$', clean):
            return True
    for pat in SUBJ_ENDING_RES:
        if re.search(pat + r'\s*$', clean):
            return True
    if re.search(r'문장의\s*기호를', clean):
        return True
    if clean.endswith('?'):
        if _korean_ratio(clean) > 0.25:
            return True
        # English-only exam questions: "Which of the following...", "What does..." etc.
        if (len(clean) > 25
                and re.search(r'(?i)\b(which|what|who|whom|where|when|how|why)\b', clean)
                and not re.match(r'^.{1,20}:\s', clean)):
            return True
    return False


def _is_metadata_line(line: str) -> bool:
    skip_patterns = [
        r'^\d{4}년\s+중?\d',
        r'^동아\(', r'^YBM', r'^천재\(', r'^비상\(', r'^금성\(',
        r'^능률\(', r'^지학사\(', r'^미래엔\(',
        r'^타\s*사이트',
        r'^<출제\s*범위>',
        r'^교과서:', r'^모의고사:', r'^부교재:',
        r'^\d{4}년\s+중?\d\s+\d학기',
        r'^정답\s*$',
        r'^무단.*게시', r'^무단.*복제',
    ]
    for pat in skip_patterns:
        if re.search(pat, line):
            return True
    return False


def _is_answer_line(line: str) -> bool:
    s = line.strip()
    if not s:
        return False
    if re.match(r'^[①②③④⑤]$', s):
        return True
    if re.match(r'^[①②③④⑤](?:\s*,\s*[①②③④⑤])+$', s):
        return True
    if re.match(r'^[ⓐⓑⓒⓓⓔ](?:\s*,\s*[ⓐⓑⓒⓓⓔ])*$', s):
        return True
    if re.match(r'^[1-5]$', s):
        return True
    return False


def _is_choice_line(line: str) -> bool:
    """선지 라인 판별 (v2.0 강화: 다중 패턴)"""
    for pat in PATTERN_CHOICE:
        if pat.search(line):
            return True
    return False


def _is_writing_area(line: str) -> bool:
    return bool(re.match(r'^\s*→\s*[_\s]{3,}', line))


def _is_end_of_content(line: str) -> bool:
    """정답 섹션 헤더 감지 (v2.0: 확장 키워드)"""
    return bool(re.match(
        r'^\s*(정답|정답\s*및\s*해설|정답\s*/\s*해설|Answer\s*Key?|모범\s*답안)\s*$',
        line, re.IGNORECASE
    ))


def _korean_ratio(text: str) -> float:
    if not text:
        return 0.0
    korean = len(re.findall(r'[가-힣]', text))
    total = len(text.replace(' ', ''))
    return korean / max(total, 1)


def _is_subjective_question(question_text: str) -> bool:
    first_line = question_text.split('\n')[0]
    clean = _clean_line_suffixes(first_line)
    for pat in SUBJ_ENDING_RES:
        if re.search(pat + r'\s*$', clean):
            return True
    return False


def _references_previous_passage(text: str) -> bool:
    """윗글/위 대화 패턴 감지 (강화 버전)

    감지 패턴:
      윗글, 위 글, 위의 글, 위 대화, 위의 대화
      + 위 형태의 조사 결합 (윗글의, 위 대화의, 위 글의, 위의 대화의 …)
    """
    ref_pats = RULES.get('common_passage', {}).get(
        'reference_patterns',
        [r'^윗글', r'^위\s*글', r'^위의\s*글',
         r'^위\s*대화', r'^위의\s*대화',
         r'^위\s*독백', r'^위의\s*독백'],
    )
    for pat in ref_pats:
        if re.match(pat, text):
            return True
    return False


# ────────────────────────────────────────────────
# v2.0 질문 후보 점수화
# ────────────────────────────────────────────────

def _score_question_candidate(line: str) -> int:
    """질문 후보 점수화 (높을수록 질문 가능성 높음).

    +2 : '?' 포함
    +2 : '시오.' 포함
    +1 : '다음', '윗글', '위 글', '대화', '표', '읽고' 포함
    +1 : 한글 10자 이상
    -3 : 선지 패턴 ①②③④⑤ 으로 시작
    """
    score = 0
    if '?' in line:
        score += 2
    if '시오.' in line:
        score += 2
    for kw in ('다음', '윗글', '위 글', '대화', '표', '읽고'):
        if kw in line:
            score += 1
            break
    korean_chars = len(re.findall(r'[가-힣]', line))
    if korean_chars >= 10:
        score += 1
    if re.match(r'^\s*[①②③④⑤]', line):
        score -= 3
    return score


_MULTI_PART_ANSWER_RE = re.compile(
    r'^\s*(?:'
    r'\([A-Za-z]\)\s*'          # (A), (B), (a), (b)
    r'|\(\d\)\s*'               # (1), (2)
    r'|[ⓐⓑⓒⓓⓔ]\s*:\s*'      # ⓐ:, ⓑ:
    r'|-\s+'                    # - bullet list
    r')'
)


def _is_multi_part_answer_line(line: str) -> bool:
    return bool(_MULTI_PART_ANSWER_RE.match(line))


def _extract_multi_part_label(line: str) -> Optional[int]:
    """Extract numeric label from multi-part answer line for reset detection."""
    m = re.match(r'^\s*\((\d)\)', line)
    if m:
        return int(m.group(1))
    m = re.match(r'^\s*\(([A-Za-z])\)', line)
    if m:
        return ord(m.group(1).upper()) - ord('A') + 1
    m = re.match(r'^\s*([ⓐⓑⓒⓓⓔ])', line)
    if m:
        return 'ⓐⓑⓒⓓⓔ'.index(m.group(1)) + 1
    return None


def _split_choices(text: str) -> List[str]:
    """선지 텍스트를 개별 선지 리스트로 분리 (v2.0).

    ①~⑤ 이 한 줄에 여러 개 있으면 분리.
    """
    parts = re.split(r'(?=[①②③④⑤])', text)
    result = [p.strip() for p in parts if p.strip()]
    return result


# ────────────────────────────────────────────────
# [1단계] 전체 문항 번호 스캔
# ────────────────────────────────────────────────

def _prescan_question_numbers(lines: List[str]) -> Tuple[List[int], Dict[int, int]]:
    """시험지 전체를 스캔하여 모든 문항 번호와 위치를 추출 (v2.0 강화).

    3단 탐지:
      Pass 1: PATTERN_Q_START 패턴 전체를 사용한 명시적 문항 번호 탐색
      Pass 2: 점수 기반 질문 후보 보완 탐지
      Pass 3: 번호 없는 경우 → 질문 패턴으로 순차 번호 부여

    Returns:
        (sorted_numbers, {question_number: line_index})
    """
    # ── Pass 1: PATTERN_Q_START 로 명시적 번호 탐색 ──
    numbered_positions: Dict[int, int] = {}

    for i, line in enumerate(lines):
        stripped = line.strip()
        if not stripped or _is_metadata_line(stripped) or _is_end_of_content(stripped):
            continue

        detected_num = None
        rest = ""
        for pat in PATTERN_Q_START:
            m = pat.match(stripped)
            if m:
                detected_num = int(m.group(1))
                rest = stripped[m.end():].strip()
                break

        if detected_num is None:
            continue

        if detected_num < 1 or detected_num > 50:
            continue

        # 실제 선지(①②③…)인지 확인 — 선지면 스킵
        if re.match(r'^[①②③④⑤]', stripped):
            continue

        clean_rest = _clean_line_suffixes(rest)
        has_korean = _korean_ratio(rest) > 0.10
        ends_with_q = clean_rest.rstrip().endswith('?')
        ends_with_subj = any(
            re.search(pat + r'\s*$', clean_rest) for pat in SUBJ_ENDING_RES
        )
        is_question = _is_question_text(stripped)
        q_score = _score_question_candidate(rest)

        if has_korean or ends_with_q or ends_with_subj or is_question or q_score >= 2:
            if detected_num not in numbered_positions:
                numbered_positions[detected_num] = i

    if numbered_positions:
        sorted_nums = sorted(numbered_positions.keys())
        return sorted_nums, numbered_positions

    # ── Pass 2: 번호만 있는 단독 블록 → 다음 줄과 병합하여 재검사 ──
    for i, line in enumerate(lines):
        stripped = line.strip()
        if not stripped or _is_metadata_line(stripped):
            continue
        m = re.match(r'^(\d{1,2})\s*$', stripped)
        if m and i + 1 < len(lines):
            num = int(m.group(1))
            next_line = lines[i + 1].strip()
            merged = f"{num}. {next_line}"
            if 1 <= num <= 50 and (_is_question_text(merged) or _score_question_candidate(next_line) >= 2):
                if num not in numbered_positions:
                    numbered_positions[num] = i

    if numbered_positions:
        sorted_nums = sorted(numbered_positions.keys())
        return sorted_nums, numbered_positions

    # ── Pass 3: 번호 없음 → 질문 패턴으로 순차 탐지 ──
    pattern_positions: Dict[int, int] = {}
    q_num = 0

    for i, line in enumerate(lines):
        stripped = line.strip()
        if not stripped or _is_metadata_line(stripped) or _is_end_of_content(stripped):
            continue
        if _is_question_text(stripped):
            la = i + 1
            while la < len(lines) and not lines[la].strip():
                la += 1
            if la < len(lines):
                nxt = lines[la].strip()
                if (_is_choice_line(nxt) and len(nxt) > 5) or _is_writing_area(nxt):
                    continue
            q_num += 1
            pattern_positions[q_num] = i

    sorted_nums = sorted(pattern_positions.keys())
    return sorted_nums, pattern_positions


# ────────────────────────────────────────────────
# 정답 섹션 파싱
# ────────────────────────────────────────────────

def _find_and_parse_answer_section(lines: List[str]) -> Dict[int, str]:
    """시험지 마지막 '정답' 섹션에서 문항별 정답 추출 (v2.0 강화).

    정답 헤더: 정답, 정답 및 해설, 정답/해설, Answer, Answer Key, 모범답안
    {문항번호: 정답} 반환. 정답 섹션이 없거나 비어있으면 빈 dict.
    """
    answers: Dict[int, str] = {}

    # ── 정답 섹션 시작 위치 탐색 (하단부터 역방향) ──
    answer_start = None
    for i in range(len(lines) - 1, -1, -1):
        if _is_end_of_content(lines[i].strip()):
            answer_start = i + 1
            break

    if answer_start is None:
        return answers

    answer_lines: List[str] = []
    for i in range(answer_start, len(lines)):
        line = lines[i].strip()
        if line:
            answer_lines.append(line)

    if not answer_lines:
        return answers

    answer_text = ' '.join(answer_lines)

    # ── 패턴 1: "N ①" / "N. ①" / "N) ①" 등 (객관식) ──
    obj_matches = re.findall(
        r'(\d+)\s*[.:\-→)]*\s*([①②③④⑤])', answer_text
    )
    for m in obj_matches:
        answers[int(m[0])] = m[1]

    # ── 패턴 2: "N 숫자" (1~5) ──
    num_matches = re.findall(
        r'(\d{1,2})\s*[.:\-→)]+\s*([1-5])\b', answer_text
    )
    for m in num_matches:
        q_num = int(m[0])
        if q_num not in answers:
            # 숫자 → 원문자 정규화
            answers[q_num] = _normalize_answer(m[1])

    # ── 패턴 3: ⓐ-ⓔ 조합 (서술형) ──
    combo_matches = re.findall(
        r'(\d+)\s*[.:\-→)]*\s*([ⓐⓑⓒⓓⓔ](?:\s*,\s*[ⓐⓑⓒⓓⓔ])*)', answer_text
    )
    for m in combo_matches:
        q_num = int(m[0])
        if q_num not in answers:
            answers[q_num] = m[1]

    # ── 패턴 4: "N. 텍스트" (서술형 모범답안, 다중 줄 지원) ──
    for idx, line in enumerate(answer_lines):
        subj_match = re.match(r'^(\d+)\s*[.:\-→)]+\s+(.{5,})$', line)
        if subj_match:
            q_num = int(subj_match.group(1))
            if q_num not in answers:
                answer_text_parts = [subj_match.group(2).strip()]
                for j in range(idx + 1, len(answer_lines)):
                    cont = answer_lines[j].strip()
                    if re.match(r'^\d+\s*[.:\-→)]', cont):
                        break
                    answer_text_parts.append(cont)
                answers[q_num] = '\n'.join(answer_text_parts)

    return answers


def _normalize_answer(ans: str) -> str:
    """정답 정규화: 1~5 숫자를 ①~⑤ 원문자로 변환.
    이미 원문자이면 그대로 반환."""
    mapping = {'1': '①', '2': '②', '3': '③', '4': '④', '5': '⑤'}
    ans = ans.strip()
    if ans in mapping:
        return mapping[ans]
    return ans


# ────────────────────────────────────────────────
# 문제유형 분류 (4가지 고정)
# ────────────────────────────────────────────────

_GRAMMAR_KEYWORDS = [
    '문법', '어법', '문법적으로', '문법상',
    '영작', '올바르게 표현', '문장을 완성',
]
_VOCABULARY_KEYWORDS = [
    '의미로', '뜻으로', '다른 뜻', '의미가', '의미는',
]
_LISTENING_KEYWORDS = [
    '듣기', '들으시오', '들려주는', '방송',
]


def _classify_question_type(q: QuestionData) -> Tuple[str, float, str]:
    """문제유형 분류 → (type, confidence, notes)
    Grammar > Vocabulary > Listening > Reading/Comprehension(기본)"""
    text = q.question_text
    notes_parts: List[str] = []

    for kw in _GRAMMAR_KEYWORDS:
        if kw in text:
            notes_parts.append(f"'{kw}' → Grammar")
            return "Grammar", 0.9, "; ".join(notes_parts)

    for kw in _VOCABULARY_KEYWORDS:
        if kw in text:
            notes_parts.append(f"'{kw}' → Vocabulary")
            return "Vocabulary", 0.9, "; ".join(notes_parts)

    for kw in _LISTENING_KEYWORDS:
        if kw in text:
            notes_parts.append(f"'{kw}' → Listening")
            return "Listening", 0.9, "; ".join(notes_parts)

    notes_parts.append("기본값 → Reading/Comprehension")
    return "Reading/Comprehension", 0.7, "; ".join(notes_parts)


# ────────────────────────────────────────────────
# 규칙 2 후처리: "윗글/위 대화" 공통지문 전파
# ────────────────────────────────────────────────

def _postprocess_common_passages(questions: List[QuestionData]) -> List[str]:
    """규칙 2 (강화 버전): "윗글~", "위 대화~" 패턴 후처리

    ※ questions 는 번호순 정렬된 상태여야 합니다.

    처리 절차 (문항 N이 "윗글/위 대화" 패턴일 때):
      1️⃣ 이전 문항(N-1) 확인
      2️⃣ N-1에 공통지문이 있으면 → N.공통지문 = N-1.공통지문 (복사)
         N-1에 문제지문만 있으면 → N.공통지문 = N-1.문제지문 (복사)
         N.문제지문 = 비움
      3️⃣ N-1이 문제지문만 가지고 있었다면:
         N-1.공통지문 = N-1.문제지문 (이동)
         N-1.문제지문 = 비움
      4️⃣ 검증: 모든 "윗글/위 대화" 문항의 공통지문 ≠ ""

    Returns:
        경고 메시지 리스트
    """
    warnings: List[str] = []

    for idx in range(len(questions)):
        q = questions[idx]
        if not _references_previous_passage(q.question_text):
            continue

        # ── 이미 메인 루프에서 공통지문이 정상 할당된 경우 → 스킵 ──
        if q.common_passage:
            logger.debug(
                f"[후처리] 문항 {q.question_number}: "
                f"이미 공통지문 있음 (메인 루프에서 할당됨) → 스킵"
            )
            continue

        # ── 이전 문항 찾기 ──
        if idx == 0:
            warnings.append(
                f"문항 {q.question_number}: '윗글/위 대화' 패턴이지만 이전 문항 없음"
            )
            continue

        prev_q = questions[idx - 1]

        # ── Case 1: 이전 문항에 공통지문이 있음 → 그대로 복사 ──
        if prev_q.common_passage:
            q.common_passage = prev_q.common_passage
            q.passage_group_id = prev_q.passage_group_id
            q.question_passage = ""
            logger.debug(
                f"[후처리] 문항 {q.question_number}: "
                f"이전 문항 {prev_q.question_number}의 공통지문 복사"
            )

        # ── Case 2: 이전 문항에 문제지문만 있음 → 공통지문으로 이동 ──
        elif prev_q.question_passage:
            shared_passage = prev_q.question_passage
            # 이전 문항 수정: 문제지문 → 공통지문
            prev_q.common_passage = shared_passage
            prev_q.question_passage = ""
            # 현재 문항 설정
            q.common_passage = shared_passage
            q.passage_group_id = prev_q.passage_group_id
            q.question_passage = ""
            logger.debug(
                f"[후처리] 문항 {q.question_number}: "
                f"이전 문항 {prev_q.question_number}의 문제지문→공통지문 이동"
            )

        # ── Case 3: 이전 문항에 지문 자체가 없음 ──
        else:
            warnings.append(
                f"문항 {q.question_number}: '윗글/위 대화' 패턴이지만 "
                f"이전 문항 {prev_q.question_number}에 지문 없음"
            )

    # ── 최종 검증: 모든 "윗글/위 대화" 문항의 공통지문 확인 ──
    for q in questions:
        if _references_previous_passage(q.question_text) and not q.common_passage:
            warnings.append(
                f"⚠️ 문항 {q.question_number}: '윗글/위 대화' 패턴인데 공통지문 = None"
            )

    return warnings


# ────────────────────────────────────────────────
# [3단계] 누락 문항 복구
# ────────────────────────────────────────────────

def _recover_missing_question(
    lines: List[str],
    q_num: int,
    line_idx: int,
    next_boundary: int,
    active_common_passage: str,
    active_passage_group_id: Optional[int],
    meta: dict,
) -> Optional[QuestionData]:
    """2단계에서 누락된 문항을 해당 위치에서 복구 시도.

    복구 전략:
      1. 해당 번호 줄에서 출제문항 추출
      2. 다음 문항 경계까지 인라인정답/지문/보기 수집
      3. 공통지문 참조 여부 자동 판별
    """
    if line_idx >= len(lines):
        return None

    q = QuestionData()
    q.question_number = q_num
    q.school = meta.get("school", "")
    q.grade = meta.get("grade", 0)

    first_line = lines[line_idx].strip()
    _, qtext = _strip_leading_qnum(first_line)
    question_text = _clean_line_suffixes(qtext)

    # 최소 유효성: 3자 이상
    if len(question_text) < 3:
        return None

    # ── 인라인 정답 ──
    j = line_idx + 1
    while j < min(next_boundary, len(lines)) and not lines[j].strip():
        j += 1

    inline_answer = ""
    if j < min(next_boundary, len(lines)):
        pa = lines[j].strip()
        if _is_answer_line(pa):
            inline_answer = pa
            j += 1
        elif (not _is_question_text(pa)
              and not _is_passage_intro(pa)
              and not _is_choice_line(pa)
              and not pa.startswith('<')
              and not _is_writing_area(pa)
              and not _is_metadata_line(pa)
              and len(pa) < 120):
            inline_answer = pa
            j += 1

    # ── 지문 + 보기 ──
    passage_lines: List[str] = []
    choice_lines: List[str] = []
    found_choices = False

    while j < min(next_boundary, len(lines)):
        nl = lines[j].strip()
        j += 1
        if not nl or _is_metadata_line(nl) or _is_end_of_content(nl):
            continue
        if _is_question_text(nl) or _is_passage_intro(nl):
            break
        if _is_writing_area(nl):
            continue
        if _is_choice_line(nl) and not found_choices:
            found_choices = True
        if found_choices:
            choice_lines.append(nl)
        else:
            passage_lines.append(nl)

    # ── 공통지문 vs 문제지문 ──
    refs_passage = _references_previous_passage(question_text)
    if refs_passage and active_common_passage:
        q.common_passage = active_common_passage
        q.passage_group_id = active_passage_group_id
        q.question_passage = ""
        if passage_lines:
            question_text += '\n' + '\n'.join(passage_lines)
    else:
        q.common_passage = ""
        q.question_passage = '\n'.join(passage_lines)

    q.question_text = question_text
    q.answer = inline_answer

    if _is_subjective_question(question_text) and not choice_lines:
        q.choices = ""
    else:
        q.choices = '\n'.join(choice_lines)

    q.raw_block_text = first_line
    q.question_type, q.confidence, q.notes = _classify_question_type(q)
    q.question_format = "서술형" if (_is_subjective_question(question_text) and not choice_lines) else "객관식"
    q.notes += "; [3단계 복구]"
    q.source_block_ids = list(range(line_idx, min(next_boundary, len(lines))))
    q.raw_no = q_num
    q.item_warnings.append("RECOVERED_STAGE3")

    return q


# ────────────────────────────────────────────────
# v2.0 Validation 단계
# ────────────────────────────────────────────────

def _validate_extraction(
    questions: List[QuestionData],
    answer_map: Dict[int, str],
    expected_numbers: List[int],
) -> Dict:
    """Validation: 품질 검증 및 이상 탐지.

    검증 항목:
      1. detected_questions / detected_answers 비교
      2. 번호 연속성 검사 (점프 > 2)
      3. 품질 최소 조건 (question_text >= 5 OR choices >= 3 OR passage >= 50)
    """
    result: Dict = {
        "detected_questions": len(questions),
        "detected_answers": len(answer_map),
        "missing_answer_numbers": [],
        "questions_without_answer": [],
        "suspicious_number_jumps": [],
        "questions_without_question_text": [],
    }

    q_numbers = sorted(q.question_number for q in questions)
    answer_numbers = sorted(answer_map.keys())

    # 정답은 있는데 문항이 없는 번호
    q_set = set(q_numbers)
    result["missing_answer_numbers"] = [n for n in answer_numbers if n not in q_set]

    # 문항은 있는데 정답이 없는 번호
    result["questions_without_answer"] = [
        q.question_number for q in questions if not q.answer
    ]

    # 번호 연속성 검사
    if q_numbers:
        for idx in range(1, len(q_numbers)):
            gap = q_numbers[idx] - q_numbers[idx - 1]
            if gap > 2:
                result["suspicious_number_jumps"].append(
                    f"{q_numbers[idx-1]}→{q_numbers[idx]} (gap={gap})"
                )

    # 품질 최소 조건
    for q in questions:
        qtext_ok = len(q.question_text) >= 5
        choices_ok = len(q.choices_list) >= 3
        passage_ok = len(q.question_passage) >= 50 or len(q.common_passage) >= 50

        if not (qtext_ok or choices_ok or passage_ok):
            result["questions_without_question_text"].append(q.question_number)
            q.item_warnings.append("품질 최소 조건 미충족 (qtext<5, choices<3, passage<50)")

    return result


# ────────────────────────────────────────────────
# v2.0 Repair 단계 (누락 방지 핵심)
# ────────────────────────────────────────────────

def _repair_missing_questions(
    questions: List[QuestionData],
    answer_map: Dict[int, str],
    lines: List[str],
    meta: Dict,
    expected_numbers: List[int],
    validation: Dict,
) -> Tuple[List[QuestionData], List[str]]:
    """Repair: Validation에서 발견된 누락을 강제 복구.

    Case 1: 정답키에 존재하나 문항 없는 번호 → 강제 슬롯 생성
    Case 2: 선지(①~⑤) 반복 등장하나 문항번호 없음 → 새 문항 생성
    Case 3: 번호 점프 발생 → 점프 구간 주변 블록 재스캔
    """
    repair_log: List[str] = []
    q_map = {q.question_number: q for q in questions}
    new_questions: List[QuestionData] = []
    max_seq = max((q.seq_no for q in questions), default=0)

    # ── Case 1: 정답키 기반 강제 슬롯 생성 ──
    for num in validation.get("missing_answer_numbers", []):
        if num in q_map:
            continue
        max_seq += 1
        forced = QuestionData(
            question_number=num,
            seq_no=max_seq,
            raw_no=num,
            question_text=f"⚠️ {num}번 문항 — 정답키에 존재하나 본문에서 미탐지",
            answer=answer_map.get(num, ""),
            answer_source="answer_key" if num in answer_map else "missing",
            school=meta.get("school", ""),
            grade=meta.get("grade", 0),
            item_warnings=["FORCED_SLOT_FROM_ANSWER_KEY"],
            notes="[Repair] 정답키 기반 강제 생성",
        )
        new_questions.append(forced)
        repair_log.append(f"Case1: {num}번 강제 슬롯 생성 (FORCED_SLOT_FROM_ANSWER_KEY)")
        logger.warning(f"[Repair] Case1: {num}번 강제 슬롯 생성")

    # ── Case 2: 선지 ①~⑤ 반복 등장하나 문항 없는 영역 탐지 ──
    # 이미 문항에 포함된 라인 인덱스 수집
    covered_lines: Set[int] = set()
    for q in questions:
        covered_lines.update(q.source_block_ids)

    orphan_choice_runs: List[List[int]] = []
    current_run: List[int] = []

    for idx, line in enumerate(lines):
        if idx in covered_lines:
            if current_run:
                orphan_choice_runs.append(current_run)
                current_run = []
            continue
        stripped = line.strip()
        if stripped and _is_choice_line(stripped):
            current_run.append(idx)
        else:
            if current_run:
                orphan_choice_runs.append(current_run)
                current_run = []
    if current_run:
        orphan_choice_runs.append(current_run)

    for run in orphan_choice_runs:
        if len(run) < 3:
            continue
        # 선지 3개 이상 → 객관식 문항 가능
        all_q_nums = set(q.question_number for q in questions) | set(nq.question_number for nq in new_questions)
        # 앞 라인에서 문항 번호 추출 시도
        candidate_num = None
        for back in range(run[0] - 1, max(run[0] - 5, -1), -1):
            if back < 0:
                break
            bl = lines[back].strip()
            det, _ = _strip_leading_qnum(bl)
            if det and det not in all_q_nums:
                candidate_num = det
                break

        if candidate_num is None:
            continue

        max_seq += 1
        choice_text = '\n'.join(lines[j].strip() for j in run if lines[j].strip())
        choice_parts = []
        for j in run:
            choice_parts.extend(_split_choices(lines[j].strip()))

        forced = QuestionData(
            question_number=candidate_num,
            seq_no=max_seq,
            raw_no=candidate_num,
            question_text=f"⚠️ {candidate_num}번 문항 — 선지 기반 추론 생성",
            choices=choice_text,
            choices_list=choice_parts,
            answer=answer_map.get(candidate_num, ""),
            answer_source="answer_key" if candidate_num in answer_map else "missing",
            school=meta.get("school", ""),
            grade=meta.get("grade", 0),
            source_block_ids=run,
            item_warnings=["FORCED_SLOT_FROM_ORPHAN_CHOICES"],
            notes="[Repair] 고아 선지 기반 생성",
        )
        new_questions.append(forced)
        repair_log.append(f"Case2: {candidate_num}번 선지 기반 슬롯 생성")
        logger.warning(f"[Repair] Case2: {candidate_num}번 선지 기반 생성")

    # ── Case 3: 번호 점프 구간 재스캔 ──
    for jump_info in validation.get("suspicious_number_jumps", []):
        # "N→M (gap=G)" 형태 파싱
        m = re.match(r'(\d+)→(\d+)', jump_info)
        if not m:
            continue
        start_n, end_n = int(m.group(1)), int(m.group(2))
        for gap_num in range(start_n + 1, end_n):
            all_q_nums = set(q.question_number for q in questions) | set(nq.question_number for nq in new_questions)
            if gap_num in all_q_nums:
                continue

            # gap_num의 위치를 lines에서 다시 찾기
            for idx, line in enumerate(lines):
                stripped = line.strip()
                det, _ = _strip_leading_qnum(stripped)
                if det == gap_num and idx not in covered_lines:
                    # 다음 문항 경계까지
                    next_boundary = len(lines)
                    for k_idx in range(idx + 1, len(lines)):
                        k_det, _ = _strip_leading_qnum(lines[k_idx].strip())
                        if k_det and k_det > gap_num:
                            next_boundary = k_idx
                            break
                        if _is_end_of_content(lines[k_idx].strip()):
                            next_boundary = k_idx
                            break

                    max_seq += 1
                    recovered = _recover_missing_question(
                        lines, gap_num, idx, next_boundary,
                        "", None, meta,
                    )
                    if recovered:
                        recovered.seq_no = max_seq
                        recovered.raw_no = gap_num
                        recovered.answer = answer_map.get(gap_num, recovered.answer)
                        if gap_num in answer_map:
                            recovered.answer_source = "answer_key"
                        recovered.item_warnings.append("RECOVERED_FROM_NUMBER_JUMP")
                        new_questions.append(recovered)
                        repair_log.append(f"Case3: {gap_num}번 번호 점프 재스캔 복구")
                        logger.warning(f"[Repair] Case3: {gap_num}번 복구")
                    break

    return new_questions, repair_log


# ────────────────────────────────────────────────
# 검증 로그 생성
# ────────────────────────────────────────────────

def _generate_verification_log(
    expected: List[int],
    extracted: List[int],
    missing: List[int],
    recovered: List[int],
    questions: Optional[List[QuestionData]] = None,
    validation: Optional[Dict] = None,
    repair_log_entries: Optional[List[str]] = None,
) -> str:
    """파싱 검증 결과 로그 (v2.0 Validation/Repair 정보 포함)"""
    log = [
        "===== 파싱 검증 결과 =====",
        f"시험지 총 문항 수: {len(expected)}개",
        f"출력된 행 수: {len(extracted)}개",
        f"문항 번호 리스트: {extracted}",
    ]
    if recovered:
        log.append(f"복구된 문항: {[f'{n}번' for n in recovered]}")
    if missing:
        log.append(f"누락된 문항: {[f'{n}번' for n in missing]}")
    else:
        log.append("누락된 문항: [없음]")

    if len(extracted) == len(expected) and not missing:
        log.append("✅ 검증 통과: 모든 문항 정상 추출")
    else:
        log.append(f"⚠️ 불일치: 예상 {len(expected)}개 vs 추출 {len(extracted)}개")

    # ── v2.0 Validation 필드 ──
    if validation:
        log.append("")
        log.append("── Validation 결과 ──")
        log.append(f"  detected_questions: {validation.get('detected_questions', 0)}")
        log.append(f"  detected_answers: {validation.get('detected_answers', 0)}")
        man = validation.get("missing_answer_numbers", [])
        if man:
            log.append(f"  missing_answer_numbers: {man}")
        qwa = validation.get("questions_without_answer", [])
        if qwa:
            log.append(f"  questions_without_answer: {qwa}")
        snj = validation.get("suspicious_number_jumps", [])
        if snj:
            log.append(f"  suspicious_number_jumps: {snj}")
        qwqt = validation.get("questions_without_question_text", [])
        if qwqt:
            log.append(f"  questions_without_question_text: {qwqt}")

    # ── v2.0 Repair 로그 ──
    if repair_log_entries:
        log.append("")
        log.append("── Repair 로그 ──")
        for entry in repair_log_entries:
            log.append(f"  {entry}")

    # ── "윗글/위 대화" 패턴 검증 ──
    if questions:
        ref_questions = [
            q for q in questions
            if _references_previous_passage(q.question_text)
        ]
        if ref_questions:
            log.append("")
            log.append("── 윗글/위 대화 패턴 문항 검증 ──")
            log.append(f"해당 문항 수: {len(ref_questions)}개")
            ok_list = [q.question_number for q in ref_questions if q.common_passage]
            ng_list = [q.question_number for q in ref_questions if not q.common_passage]
            if ok_list:
                log.append(f"  공통지문 정상: {[f'{n}번' for n in ok_list]}")
            if ng_list:
                log.append(f"  ⚠️ 공통지문 없음: {[f'{n}번' for n in ng_list]}")
            else:
                log.append("  ✅ 모든 윗글/위 대화 문항의 공통지문 확인됨")

    log.append("=" * 26)
    return '\n'.join(log)


# ────────────────────────────────────────────────
# 메인 추출 함수 (3단계 프로세스)
# ────────────────────────────────────────────────

def extract_questions(full_text: str, filename: str = "") -> ExtractionResult:
    """3단계 프로세스로 문항 추출 (누락 절대 금지)

    [1단계] 전체 스캔 → 문항 번호 리스트 (사전 탐지)
    [2단계] 상세 파싱 → 문항 데이터 추출 (기존 패턴 기반)
    [3단계] 검증 → 누락 문항 복구 → 정답 적용 → 검증 리포트
    """
    result = ExtractionResult()
    meta = _parse_filename_metadata(filename)
    result.metadata = meta

    lines = full_text.split('\n')

    # ═══════════════════════════════════════════
    # [1단계] 전체 문항 번호 스캔
    # ═══════════════════════════════════════════
    expected_numbers, number_positions = _prescan_question_numbers(lines)
    result.expected_numbers = expected_numbers
    logger.info(f"[1단계] 문항 번호 스캔 완료: {expected_numbers}")

    # 번호 연속성 검사 (점프 경고)
    if expected_numbers:
        for idx in range(1, len(expected_numbers)):
            gap = expected_numbers[idx] - expected_numbers[idx - 1]
            if gap > 1:
                for skip in range(expected_numbers[idx - 1] + 1, expected_numbers[idx]):
                    result.warnings.append(
                        f"⚠️ {skip}번 문항이 시험지에 없음 (번호 점프: "
                        f"{expected_numbers[idx - 1]}→{expected_numbers[idx]})"
                    )

    # ── 정답 섹션 파싱 ──
    answer_map = _find_and_parse_answer_section(lines)
    if answer_map:
        result.answer_section_found = True
        logger.info(f"정답 섹션 발견: {len(answer_map)}개 정답")

    # ═══════════════════════════════════════════
    # [2단계] 상세 파싱
    # ═══════════════════════════════════════════
    questions: List[QuestionData] = []

    q_num = 0
    seq_counter = 0  # v2.0 문서 순서 카운터
    current_common_passage = ""
    current_passage_group_id = 0
    active_passage_group: Optional[int] = None

    i = 0

    # 헤더/메타 건너뛰기
    while i < len(lines):
        line = lines[i].strip()
        if line and (_is_question_text(line) or _is_passage_intro(line)):
            break
        i += 1

    # 메인 파싱 루프
    while i < len(lines):
        line = lines[i].strip()

        if not line:
            i += 1
            continue

        if _is_end_of_content(line):
            break

        if _is_metadata_line(line):
            i += 1
            continue

        # ── 공통지문 도입부 ──
        if _is_passage_intro(line):
            current_passage_group_id += 1
            active_passage_group = current_passage_group_id
            passage_lines: List[str] = []
            i += 1

            while i < len(lines):
                nl = lines[i].strip()
                if not nl:
                    i += 1
                    continue
                if _is_question_text(nl) or _is_passage_intro(nl) or _is_end_of_content(nl):
                    break
                if _is_metadata_line(nl):
                    i += 1
                    continue
                passage_lines.append(nl)
                i += 1

            current_common_passage = '\n'.join(passage_lines)
            logger.info(f"공통지문 그룹 {current_passage_group_id}: {len(passage_lines)}줄")
            continue

        # ── 문항 질문문 ──
        if _is_question_text(line):
            q_num += 1
            seq_counter += 1
            q = QuestionData()
            q.school = meta.get("school", "")
            q.grade = meta.get("grade", 0)
            q.seq_no = seq_counter               # v2.0
            q.source_block_ids = [i]              # v2.0 시작 라인

            detected_num, qtext_body = _strip_leading_qnum(line)
            if detected_num is not None:
                q.question_number = detected_num
                q.raw_no = detected_num           # v2.0
                q_num = detected_num
            else:
                q.question_number = q_num
                q.raw_no = None                   # v2.0 번호 없음

            question_text = _clean_line_suffixes(
                qtext_body if detected_num is not None else line
            )
            raw_block_lines = [line]

            # 인라인 정답 수집
            i += 1
            while i < len(lines) and not lines[i].strip():
                i += 1

            inline_answer = ""
            if i < len(lines):
                pa = lines[i].strip()
                if _is_answer_line(pa):
                    inline_answer = pa
                    raw_block_lines.append(pa)
                    q.source_block_ids.append(i)
                    i += 1
                elif (not _is_question_text(pa)
                      and not _is_passage_intro(pa)
                      and not _is_choice_line(pa)
                      and not pa.startswith('<')
                      and not _is_writing_area(pa)
                      and not _is_metadata_line(pa)
                      and (len(pa) < 120 or ' 또는 ' in pa)):
                    inline_answer = pa
                    raw_block_lines.append(pa)
                    q.source_block_ids.append(i)
                    i += 1

            # 다중 줄 인라인 정답 수집: (A)/(B), (1)/(2), ⓐ:/ⓑ: 패턴
            if inline_answer and _is_multi_part_answer_line(inline_answer):
                answer_parts = [inline_answer]
                prev_label = _extract_multi_part_label(inline_answer)
                while i < len(lines):
                    next_l = lines[i].strip()
                    if not next_l:
                        i += 1
                        continue
                    curr_label = _extract_multi_part_label(next_l)
                    if curr_label is not None and prev_label is not None and curr_label <= prev_label:
                        break
                    if (_is_multi_part_answer_line(next_l)
                            and not _is_question_text(next_l)
                            and not _is_passage_intro(next_l)
                            and not _is_choice_line(next_l)
                            and not next_l.startswith('<')
                            and not _is_writing_area(next_l)):
                        answer_parts.append(next_l)
                        raw_block_lines.append(next_l)
                        q.source_block_ids.append(i)
                        if curr_label is not None:
                            prev_label = curr_label
                        i += 1
                    else:
                        break
                inline_answer = '\n'.join(answer_parts)

            # 지문 + 보기 수집
            passage_lines_q: List[str] = []
            choice_lines: List[str] = []
            found_choices = False

            while i < len(lines):
                nl = lines[i].strip()
                if not nl:
                    i += 1
                    continue
                if _is_question_text(nl) or _is_passage_intro(nl) or _is_end_of_content(nl):
                    if _is_question_text(nl) and not _is_passage_intro(nl) and not _is_end_of_content(nl):
                        la = i + 1
                        while la < len(lines) and not lines[la].strip():
                            la += 1
                        if la < len(lines):
                            nxt = lines[la].strip()
                            nxt_is_real_choice = _is_choice_line(nxt) and len(nxt) > 5
                            if nxt_is_real_choice or (inline_answer and _is_writing_area(nxt)):
                                passage_lines_q.append(nl)
                                raw_block_lines.append(nl)
                                q.source_block_ids.append(i)
                                i += 1
                                continue
                    break
                if _is_metadata_line(nl):
                    i += 1
                    continue

                raw_block_lines.append(nl)
                q.source_block_ids.append(i)      # v2.0

                if _is_writing_area(nl):
                    i += 1
                    continue

                if _is_choice_line(nl) and not found_choices:
                    found_choices = True

                if found_choices:
                    choice_lines.append(nl)
                else:
                    passage_lines_q.append(nl)

                i += 1

            # 공통지문 vs 문제지문
            refs_passage = _references_previous_passage(question_text)

            if refs_passage and active_passage_group is not None:
                # "윗글/위 대화" + 활성 공통지문 그룹 존재 → 공통지문 복사
                q.passage_group_id = active_passage_group
                q.common_passage = current_common_passage
                # 부가 지문이 있으면 문제지문(question_passage)에 저장
                q.question_passage = '\n'.join(passage_lines_q) if passage_lines_q else ""
            elif refs_passage:
                # "윗글/위 대화" 패턴이지만 활성 공통지문 없음
                # → 후처리(_postprocess_common_passages)에서 이전 문항 지문 복사 예정
                q.common_passage = ""
                q.question_passage = '\n'.join(passage_lines_q)
            else:
                q.common_passage = ""
                q.question_passage = '\n'.join(passage_lines_q)
                # "다음~"으로 시작하면 새 맥락이므로 공통지문 그룹 리셋
                # 영어 전용 문항 등 "다음"으로 시작하지 않으면 그룹 유지
                if re.match(r'^다음', question_text):
                    active_passage_group = None
                    current_common_passage = ""
                elif active_passage_group is not None:
                    q.passage_group_id = active_passage_group
                    q.common_passage = current_common_passage

            q.question_text = question_text

            # v2.0 answer_source 추적
            q.answer = inline_answer
            if inline_answer:
                q.answer_source = "inline"

            # v2.0 choices_list 생성
            if _is_subjective_question(question_text) and not choice_lines:
                q.choices = ""
                q.choices_list = []
            else:
                q.choices = '\n'.join(choice_lines)
                all_choice_parts: List[str] = []
                for cl in choice_lines:
                    all_choice_parts.extend(_split_choices(cl))
                q.choices_list = all_choice_parts

            q.raw_block_text = '\n'.join(raw_block_lines[:30])
            q.question_type, q.confidence, q.notes = _classify_question_type(q)
            q.question_format = "서술형" if (_is_subjective_question(question_text) and not choice_lines) else "객관식"

            # v2.0 품질 경고
            if len(question_text) < 5:
                q.item_warnings.append("question_text 길이 < 5")

            questions.append(q)
            logger.debug(
                f"[2단계] 문항 {q.question_number} (seq={seq_counter}): "
                f"{question_text[:40]}… → 정답={q.answer}, 유형={q.question_type}"
            )
            continue

        i += 1

    logger.info(f"[2단계] 파싱 완료: {len(questions)}문항")

    # ═══════════════════════════════════════════
    # [3단계] 검증 및 복구
    # ═══════════════════════════════════════════
    extracted_set = set(q.question_number for q in questions)
    missing_in_stage2 = sorted(set(expected_numbers) - extracted_set)
    recovered_numbers: List[int] = []

    if missing_in_stage2:
        logger.warning(f"[3단계] 2단계 누락 문항 감지: {missing_in_stage2}")

        for num in missing_in_stage2:
            if num not in number_positions:
                continue

            line_idx = number_positions[num]

            # 다음 문항 경계 찾기
            next_nums = [n for n in sorted(number_positions.keys()) if n > num]
            if next_nums:
                next_boundary = number_positions[next_nums[0]]
            else:
                next_boundary = len(lines)
                for k in range(line_idx + 1, len(lines)):
                    if _is_end_of_content(lines[k].strip()):
                        next_boundary = k
                        break

            # 복구 시점의 활성 공통지문 추정 (가장 가까운 이전 공통지문)
            active_cp = ""
            active_pg = None
            for q in sorted(questions, key=lambda x: x.question_number):
                if q.question_number < num and q.common_passage:
                    active_cp = q.common_passage
                    active_pg = q.passage_group_id
                if q.question_number >= num:
                    break

            recovered = _recover_missing_question(
                lines, num, line_idx, next_boundary,
                active_cp, active_pg, meta,
            )
            if recovered:
                questions.append(recovered)
                recovered_numbers.append(num)
                result.warnings.append(f"문항 {num}: 3단계에서 복구됨")
                logger.info(f"[3단계] 문항 {num} 복구 성공")
            else:
                result.warnings.append(f"문항 {num}: 복구 실패 (내용 불충분)")
                logger.warning(f"[3단계] 문항 {num} 복구 실패")
    else:
        logger.info("[3단계] 누락 문항 없음 — 검증 통과")

    # ── 번호순 정렬 ──
    questions.sort(key=lambda q: q.question_number)

    # ═══════════════════════════════════════════
    # 규칙 2 후처리: "윗글/위 대화" 공통지문 전파
    # ═══════════════════════════════════════════
    passage_warnings = _postprocess_common_passages(questions)
    result.warnings.extend(passage_warnings)
    logger.info(
        f"[후처리] 윗글/위 대화 공통지문 전파 완료 "
        f"(경고 {len(passage_warnings)}건)"
    )

    # ── 정답 적용 (정답 섹션 우선 > 인라인 폴백) ──
    for q in questions:
        if q.question_number in answer_map:
            q.answer = answer_map[q.question_number]
            q.answer_source = "answer_key"
        elif q.answer and q.answer_source == "inline":
            pass  # 인라인 정답 유지
        # answer_source 가 여전히 "missing" 인 경우 그대로 유지

    # ═══════════════════════════════════════════
    # [Validation] 품질 검증
    # ═══════════════════════════════════════════
    validation = _validate_extraction(questions, answer_map, expected_numbers)
    result.detected_questions = validation["detected_questions"]
    result.detected_answers = validation["detected_answers"]
    result.missing_answer_numbers = validation["missing_answer_numbers"]
    result.questions_without_answer = validation["questions_without_answer"]
    result.suspicious_number_jumps = validation["suspicious_number_jumps"]
    result.questions_without_question_text = validation["questions_without_question_text"]

    logger.info(
        f"[Validation] questions={result.detected_questions}, "
        f"answers={result.detected_answers}, "
        f"missing_ans_nums={result.missing_answer_numbers}, "
        f"jumps={result.suspicious_number_jumps}"
    )

    # ═══════════════════════════════════════════
    # [Repair] 강제 복구 (누락 방지 핵심)
    # ═══════════════════════════════════════════
    need_repair = (
        validation["missing_answer_numbers"]
        or validation["suspicious_number_jumps"]
    )

    repair_log_entries: List[str] = []

    if need_repair:
        logger.info("[Repair] 누락 가능성 감지 — Repair 단계 실행")
        repaired_qs, repair_log_entries = _repair_missing_questions(
            questions, answer_map, lines, meta, expected_numbers, validation,
        )
        if repaired_qs:
            questions.extend(repaired_qs)
            questions.sort(key=lambda q: q.question_number)
            for rq in repaired_qs:
                recovered_numbers.append(rq.question_number)
            logger.info(f"[Repair] {len(repaired_qs)}건 추가 복구")
    else:
        logger.info("[Repair] 복구 불필요 — 스킵")

    result.repair_log = repair_log_entries

    # ── 최종 결과 ──
    result.questions = questions
    result.total_questions = len(questions)
    result.extracted_numbers = sorted(q.question_number for q in questions)
    result.missing_numbers = sorted(set(expected_numbers) - set(result.extracted_numbers))
    result.recovered_numbers = recovered_numbers

    for q in questions:
        if not q.answer:
            result.warnings.append(f"문항 {q.question_number}: 정답 미발견")

    if result.total_questions == 0:
        result.errors.append("문항을 찾을 수 없습니다. 텍스트 형식을 확인해주세요.")

    result.verification_log = _generate_verification_log(
        expected_numbers,
        result.extracted_numbers,
        result.missing_numbers,
        recovered_numbers,
        questions,
        validation,
        repair_log_entries,
    )

    logger.info(f"[최종] {result.total_questions}문항 추출 완료")
    return result


# ────────────────────────────────────────────────
# 파일명 메타 추출
# ────────────────────────────────────────────────

def _parse_filename_metadata(filename: str) -> dict:
    meta = {
        "year": "", "grade": 0, "semester": "", "exam_type": "",
        "school": "", "school_short": "", "region": "",
        "publisher": "", "raw_filename": filename,
    }
    if not filename:
        return meta

    name = os.path.splitext(filename)[0]

    year_match = re.search(r'(\d{4})년', name)
    if year_match:
        meta["year"] = year_match.group(1)

    grade_match = re.search(r'중(\d)', name)
    if grade_match:
        meta["grade"] = int(grade_match.group(1))
    else:
        grade_match = re.search(r'고(\d)', name)
        if grade_match:
            meta["grade"] = int(grade_match.group(1))

    sem_match = re.search(r'(\d)학기\s*(중간|기말)', name)
    if sem_match:
        meta["semester"] = f"{sem_match.group(1)}학기"
        meta["exam_type"] = sem_match.group(2)

    school_match = re.search(r'_([\w]+(?:중학교|고등학교|중|고))_', name)
    if school_match:
        meta["school"] = school_match.group(1)
        meta["school_short"] = meta["school"].replace("학교", "")

    parts = name.split('_')
    for part in parts:
        if any(region in part for region in ['시', '도', '군', '구']):
            if '학교' not in part and '학기' not in part:
                meta["region"] = part.strip()
                break

    pub_match = re.search(r'_([^_]+)\(', name)
    if pub_match:
        meta["publisher"] = pub_match.group(1)

    return meta


# ────────────────────────────────────────────────
# 독립 실행 테스트
# ────────────────────────────────────────────────

if __name__ == "__main__":
    import sys
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

    test_file = os.path.join(os.path.dirname(__file__), "extracted_text.txt")
    if os.path.exists(test_file):
        with open(test_file, 'r', encoding='utf-8') as f:
            text = f.read()

        result = extract_questions(
            text,
            "2025년_중3_1학기 기말_봉황중학교_충청남도 공주시_동아(윤정미).hwp"
        )

        # ── 문항별 출력 ──
        for q in result.questions:
            print(f"\n{'='*60}")
            print(f"[{q.question_number}] seq={q.seq_no} raw_no={q.raw_no} "
                  f"| {q.question_type} | answer_source={q.answer_source}")
            print(f"  출제문항: {q.question_text[:80]}")
            print(f"  공통지문: {q.common_passage[:50]}…"
                  if q.common_passage else "  공통지문: (없음)")
            print(f"  문제지문: {q.question_passage[:50]}…"
                  if q.question_passage else "  문제지문: (없음)")
            print(f"  보기: {q.choices[:50]}…"
                  if q.choices else "  보기: (없음)")
            print(f"  보기(list): {q.choices_list[:3]}…"
                  if q.choices_list else "  보기(list): []")
            print(f"  정답: {q.answer}")
            print(f"  source_blocks: {q.source_block_ids[:5]}…"
                  if len(q.source_block_ids) > 5 else f"  source_blocks: {q.source_block_ids}")
            if q.item_warnings:
                print(f"  ⚠️ warnings: {q.item_warnings}")

        # ── 검증 리포트 출력 (필수) ──
        print(f"\n{result.verification_log}")

        # ── 필수 로그 출력 (v2.0) ──
        print(f"\n── 필수 로그 (v2.0) ──")
        print(f"detected_questions: {result.detected_questions}")
        print(f"detected_answers: {result.detected_answers}")
        print(f"missing_answer_numbers: {result.missing_answer_numbers}")
        print(f"questions_without_answer: {result.questions_without_answer}")
        print(f"suspicious_number_jumps: {result.suspicious_number_jumps}")
        print(f"questions_without_question_text: {result.questions_without_question_text}")
        if result.repair_log:
            print(f"repair_log: {result.repair_log}")

        # ── 경고 출력 ──
        if result.warnings:
            print("\n⚠️ 경고:")
            for w in result.warnings:
                print(f"  - {w}")
    else:
        print("extracted_text.txt를 찾을 수 없습니다.")
