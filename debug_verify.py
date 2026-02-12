"""한글 인코딩 검증 - UTF-8 파일로 저장"""
import os, sys, struct, zlib, io, glob
import olefile

parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
hwp_files = glob.glob(os.path.join(parent_dir, "*.hwp"))
filepath = hwp_files[0]

ole = olefile.OleFileIO(filepath)
header = ole.openstream("FileHeader").read()
props = struct.unpack_from('<I', header, 36)[0]
is_compressed = bool(props & 0x01)

section_data = ole.openstream("BodyText/Section0").read()
if is_compressed:
    try:
        section_data = zlib.decompress(section_data, -15)
    except:
        section_data = zlib.decompress(section_data)

HWPTAG_PARA_TEXT = 0x010 + 51

def extract_text(data):
    chars = []
    i = 0
    while i < len(data):
        if i + 1 >= len(data):
            break
        code = data[i] | (data[i+1] << 8)
        if code == 10:
            chars.append('\n')
            i += 2
        elif code == 13:
            i += 2
        elif code < 32:
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
    h = struct.unpack_from('<I', section_data, offset)[0]
    tag_id = h & 0x3FF
    size = (h >> 20) & 0xFFF
    offset += 4
    if size == 0xFFF:
        if offset + 4 > len(section_data):
            break
        size = struct.unpack_from('<I', section_data, offset)[0]
        offset += 4
    if offset + size > len(section_data):
        break
    if tag_id == HWPTAG_PARA_TEXT:
        text = extract_text(section_data[offset:offset+size])
        if text.strip():
            paragraphs.append(text.strip())
    offset += size

full_text = '\n'.join(paragraphs)

# UTF-8 파일로 저장
with open('extracted_text.txt', 'w', encoding='utf-8') as f:
    f.write(full_text)

print(f"Total paragraphs: {len(paragraphs)}")
print(f"Total chars: {len(full_text)}")
print("Saved to extracted_text.txt")

ole.close()

