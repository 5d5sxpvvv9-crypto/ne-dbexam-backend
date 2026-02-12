"""HWP 바이너리 구조 디버깅"""
import os, sys, struct, zlib, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import olefile

parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
import glob
hwp_files = glob.glob(os.path.join(parent_dir, "*.hwp"))
filepath = hwp_files[0]
print(f"파일: {filepath}")

ole = olefile.OleFileIO(filepath)

# Streams
print("\n=== OLE Streams ===")
for entry in ole.listdir():
    stream_name = "/".join(entry)
    size = ole.get_size(stream_name)
    print(f"  {stream_name} ({size} bytes)")

# FileHeader
header = ole.openstream("FileHeader").read()
props = struct.unpack_from('<I', header, 36)[0]
is_compressed = bool(props & 0x01)
print(f"\nCompressed: {is_compressed}")

# BodyText/Section0
section_data = ole.openstream("BodyText/Section0").read()
if is_compressed:
    try:
        section_data = zlib.decompress(section_data, -15)
    except:
        section_data = zlib.decompress(section_data)
print(f"Section0 decompressed: {len(section_data)} bytes")

# Parse records and show PARA_TEXT records
HWPTAG_BEGIN = 0x010
HWPTAG_PARA_HEADER = HWPTAG_BEGIN + 50  # 66
HWPTAG_PARA_TEXT = HWPTAG_BEGIN + 51  # 67

offset = 0
para_count = 0
while offset < len(section_data) and para_count < 30:
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
        para_count += 1
        # Try to extract text - show raw bytes and decoded text
        print(f"\n--- PARA_TEXT #{para_count} (size={size}, level={level}) ---")
        
        # Show first 100 bytes hex
        hex_preview = ' '.join(f'{b:02x}' for b in record_data[:min(100, len(record_data))])
        print(f"  Hex: {hex_preview}")
        
        # Try UTF-16LE decode with control char filtering
        chars = []
        i = 0
        while i < len(record_data):
            if i + 1 >= len(record_data):
                break
            code = record_data[i] | (record_data[i+1] << 8)
            if code <= 9:
                chars.append(f'[CTRL{code}]')
                i += 8  # try 8 bytes for extended controls
            elif code == 10:
                chars.append('\n')
                i += 2
            elif code == 13:
                chars.append('[CR]')
                i += 2
            elif code < 32:
                chars.append(f'[C{code}]')
                i += 2
            else:
                chars.append(chr(code))
                i += 2
        text = ''.join(chars)
        print(f"  Text: {text[:200]}")
        
    elif tag_id == HWPTAG_PARA_HEADER:
        # Show paragraph header for context
        pass

ole.close()

