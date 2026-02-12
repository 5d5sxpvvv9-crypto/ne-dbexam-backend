"""HWP PrvText 스트림 및 올바른 파싱 테스트"""
import os, sys, struct, zlib, io, glob
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import olefile

parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
hwp_files = glob.glob(os.path.join(parent_dir, "*.hwp"))
filepath = hwp_files[0]

ole = olefile.OleFileIO(filepath)

# 1) PrvText 읽기
print("=== PrvText ===")
prvtext_data = ole.openstream("PrvText").read()
# PrvText는 UTF-16LE 인코딩
try:
    prvtext = prvtext_data.decode('utf-16-le')
    print(prvtext[:3000])
except:
    print("UTF-16LE 디코딩 실패")
    # CP949 시도
    try:
        prvtext = prvtext_data.decode('cp949')
        print(prvtext[:3000])
    except:
        print("CP949도 실패")

# 2) 수정된 PARA_TEXT 파싱 (16바이트 제어문자)
print("\n\n=== 수정된 파싱 결과 ===")

header = ole.openstream("FileHeader").read()
props = struct.unpack_from('<I', header, 36)[0]
is_compressed = bool(props & 0x01)

section_data = ole.openstream("BodyText/Section0").read()
if is_compressed:
    try:
        section_data = zlib.decompress(section_data, -15)
    except:
        section_data = zlib.decompress(section_data)

HWPTAG_BEGIN = 0x010
HWPTAG_PARA_TEXT = HWPTAG_BEGIN + 51

def extract_para_text_v2(data):
    """수정된 텍스트 추출 - 제어문자 16바이트"""
    chars = []
    i = 0
    while i < len(data):
        if i + 1 >= len(data):
            break
        code = data[i] | (data[i+1] << 8)
        
        if code == 10:  # 줄바꿈
            chars.append('\n')
            i += 2
        elif code == 13:  # 문단 끝
            i += 2
        elif code < 32:  # 모든 제어 문자 16바이트
            i += 16
        else:
            chars.append(chr(code))
            i += 2
    return ''.join(chars)

offset = 0
paragraphs = []
while offset < len(section_data):
    if offset + 4 > len(section_data):
        break
    header_val = struct.unpack_from('<I', section_data, offset)[0]
    tag_id = header_val & 0x3FF
    level = (header_val >> 10) & 0x3FF
    size = (header_val >> 20) & 0xFFF
    offset += 4
    if size == 0xFFF:
        if offset + 4 > len(section_data):
            break
        size = struct.unpack_from('<I', section_data, offset)[0]
        offset += 4
    if offset + size > len(section_data):
        break
    record_data = section_data[offset:offset+size]
    offset += size
    
    if tag_id == HWPTAG_PARA_TEXT:
        text = extract_para_text_v2(record_data)
        if text.strip():
            paragraphs.append(text.strip())

full_text = '\n'.join(paragraphs)
print(full_text[:5000])

print(f"\n\n=== 총 문단 수: {len(paragraphs)} ===")
print(f"=== 전체 텍스트 길이: {len(full_text)} 자 ===")

ole.close()

