"""视觉常量；新增主题只需向 THEMES 添加字典，不改布局。"""
from copy import deepcopy
import platform

PALETTE = {
    "paper": "F4F3ED", "paper_alt": "E9EAE1", "ink": "222B26",
    "ink_soft": "59635B", "accent": "4B6047", "accent_mid": "98A48D",
    "emphasis": "A63D35", "dark_bg": "1F2A25", "white": "FFFFFF",
    "offwhite": "F0EFE6",
}
FONTS = {
    "cjk": "Microsoft YaHei" if platform.system() == "Windows" else "Noto Sans CJK SC",
    "head": "Microsoft YaHei" if platform.system() == "Windows" else "Noto Sans CJK SC",
    "mono": "Consolas" if platform.system() == "Windows" else "DejaVu Sans Mono",
}
GEOMETRY = {"slide_w": 13.333333, "slide_h": 7.5, "margin": 0.62,
            "content_top": 1.9, "content_bottom": 6.78, "max_pages": 99}
TYPE_SCALE = {"h1": 28, "h2": 20, "body": 18, "small": 11,
              "caption": 10, "min_body": 12, "cover": 48}
THEMES = {"military_light": {"palette": PALETTE, "fonts": FONTS,
                            "geometry": GEOMETRY, "type_scale": TYPE_SCALE}}


def get_theme(name):
    return deepcopy(THEMES[name])
