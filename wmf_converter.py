import os
import sys
import ctypes
from ctypes import wintypes

# S? d?ng tr?c ti?p GDI+ có s?n trên m?i b?n Windows (không c?n cài thêm gì)
class GdiplusStartupInput(ctypes.Structure):
    _fields_ = [
        ('GdiplusVersion', wintypes.UINT),
        ('DebugEventCallback', ctypes.c_void_p),
        ('SuppressBackgroundThread', wintypes.BOOL),
        ('SuppressExternalCodecs', wintypes.BOOL)
    ]

class GUID(ctypes.Structure):
    _fields_ = [
        ('Data1', wintypes.DWORD),
        ('Data2', wintypes.WORD),
        ('Data3', wintypes.WORD),
        ('Data4', ctypes.c_byte * 8)
    ]

# Mã GUID c?a d?nh d?ng ?nh PNG chu?n: {557CF406-1A04-11D3-9A73-0000F81EF32E}
PNG_CLSID = GUID(
    0x557CF406, 0x1A04, 0x11D3,
    (ctypes.c_byte * 8)(0x9A, 0x73, 0x00, 0x00, 0xF8, 0x1E, 0xF3, 0x2E)
)

def convert_wmf_folder(folder_path):
    if not os.path.exists(folder_path):
        return

    gdiplus = ctypes.windll.gdiplus
    token = ctypes.c_ulong()
    startup_in = GdiplusStartupInput(1, None, False, False)
    gdiplus.GdiplusStartup(ctypes.byref(token), ctypes.byref(startup_in), None)

    try:
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                ext = os.path.splitext(file)[1].lower()
                if ext in ['.wmf', '.emf']:
                    src_file = os.path.join(root, file)
                    dst_file = os.path.splitext(src_file)[0] + '.png'

                    image = ctypes.c_void_p()
                    # M? ?nh WMF b?ng GDI+
                    status = gdiplus.GdipLoadImageFromFile(ctypes.c_wchar_p(src_file), ctypes.byref(image))
                    if status == 0:
                        # Luu thành PNG ch?t lu?ng cao
                        gdiplus.GdipSaveImageToFile(image, ctypes.c_wchar_p(dst_file), ctypes.byref(PNG_CLSID), None)
                        gdiplus.GdipDisposeImage(image)
    finally:
        gdiplus.GdiplusShutdown(token)

if __name__ == '__main__':
    if len(sys.argv) > 1:
        convert_wmf_folder(sys.argv[1])