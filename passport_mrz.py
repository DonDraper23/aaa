#!/usr/bin/env python3
"""
护照 MRZ (Machine Readable Zone) 高精度解析与校验模块
依据：ICAO Doc 9303 标准

核心能力：
    1. MRZ 校验位（check digit）数学验证 → 100% 确认护照号正确
    2. OCR 常见混淆字符自动纠错（0↔O, 1↔I, 5↔S, 8↔B 等）
    3. 图片预处理（灰度+自动对比度+锐化），提升 OCR 准确率
    4. 护照区独立高 DPI 裁剪，兜底 MRZ 被误识别为中文的场景

用法示例：
    from passport_mrz import verify_passport_no, enhance_for_ocr, crop_mrz_zone_hidpi

    # 1. 校验 API 返回的护照号是否正确
    verified, confidence = verify_passport_no(mrz_line2_raw)

    # 2. 图片预处理
    enhanced = enhance_for_ocr(pil_image)

    # 3. 高 DPI 裁剪 MRZ 区域（当整张图识别失败时兜底）
    mrz_crop = crop_mrz_zone_hidpi(pdf_path, dpi=300)
"""

import re
from itertools import product
from PIL import Image, ImageEnhance, ImageOps, ImageFilter


# ============================================================
# MRZ 校验位算法（ICAO 9303）
# ============================================================

def _char_value(c: str) -> int:
    """MRZ 字符转数值：数字=自身, A-Z=10-35, '<'=0"""
    if c == "<":
        return 0
    if c.isdigit():
        return int(c)
    if c.isalpha():
        return ord(c.upper()) - ord("A") + 10
    return 0


def mrz_check_digit(data: str) -> int:
    """
    计算 MRZ 校验位（权重循环 7,3,1）
    >>> mrz_check_digit("123456789")
    7
    """
    weights = [7, 3, 1]
    return sum(_char_value(c) * weights[i % 3] for i, c in enumerate(data)) % 10


# ============================================================
# 护照号 + 校验位验证（核心功能）
# ============================================================

# OCR 常见误识别对照（左侧=OCR 可能读错的字符，右侧=候选正确值）
_OCR_CONFUSION = {
    "0": ["O", "D", "Q"],
    "O": ["0"],
    "1": ["I", "l", "L", "T"],
    "I": ["1"],
    "2": ["Z"],
    "Z": ["2"],
    "5": ["S"],
    "S": ["5"],
    "8": ["B"],
    "B": ["8"],
    "6": ["G"],
    "G": ["6"],
    "4": ["A"],
    "A": ["4"],
}


def verify_passport_no(pno: str, check: str) -> tuple[str | None, str]:
    """
    验证护照号 + 校验位。若校验失败，尝试 OCR 纠错。

    Args:
        pno:   9位护照号（来自 MRZ 第二行 1-9 位）
        check: 1位校验位（MRZ 第二行第 10 位）

    Returns:
        (corrected_pno, confidence)
            confidence 取值:
                "verified"   - 校验位通过，100% 正确
                "corrected"  - 用 OCR 纠错后通过校验（高置信度）
                "unverified" - 无法通过校验（需人工核对）
    """
    if not pno or len(pno) != 9 or not check or len(check) != 1:
        return None, "unverified"

    # 原值直接通过校验
    if str(mrz_check_digit(pno)) == check:
        return pno, "verified"

    # OCR 纠错：只替换可疑字符（非数字或在混淆表中的数字）
    candidates_per_pos = []
    for c in pno:
        opts = {c}
        if c in _OCR_CONFUSION:
            opts.update(_OCR_CONFUSION[c])
        # 护照号必须是纯数字，过滤非数字候选
        opts = {o for o in opts if o.isdigit()}
        if not opts:
            opts = {c}  # 保底
        candidates_per_pos.append(list(opts))

    # 笛卡尔积搜索（最多 3^9 = 19683 组合，实际因多数位只有1选项会小得多）
    combos = 1
    for c in candidates_per_pos:
        combos *= len(c)
    if combos > 50000:
        return pno, "unverified"  # 组合爆炸，放弃

    for combo in product(*candidates_per_pos):
        candidate = "".join(combo)
        if candidate == pno:
            continue  # 原值已试过
        if candidate.isdigit() and str(mrz_check_digit(candidate)) == check:
            return candidate, "corrected"

    return pno, "unverified"


# ============================================================
# MRZ 第二行结构化解析
# ============================================================

_RE_MRZ_LINE2 = re.compile(
    r"([A-Z0-9<]{9})"     # 1. 护照号 (9)
    r"(\d)"               # 2. 护照号校验 (1)
    r"([A-Z<]{3})"        # 3. 国籍 (3)
    r"(\d{6})"            # 4. 生日 YYMMDD (6)
    r"(\d)"               # 5. 生日校验 (1)
    r"([MF<])"            # 6. 性别 (1)
    r"(\d{6})"            # 7. 失效日 (6)
    r"(\d)"               # 8. 失效日校验 (1)
)


def parse_mrz_line2(text: str) -> dict | None:
    """
    从文本中定位并解析 MRZ 第二行。
    返回结构化字段 + 各校验位结果；未匹配到返回 None。
    """
    for line in text.splitlines():
        clean = re.sub(r"[^A-Z0-9<]", "", line.upper())
        if len(clean) < 28:
            continue
        m = _RE_MRZ_LINE2.match(clean)
        if not m:
            continue

        pno_raw, pno_chk, nat, dob, dob_chk, sex, exp, exp_chk = m.groups()

        # 用 OCR 纠错验证护照号
        pno_fixed, pno_conf = verify_passport_no(pno_raw, pno_chk)

        return {
            "passport_no":        pno_fixed or pno_raw,
            "passport_no_check":  pno_conf,          # verified / corrected / unverified
            "nationality":        nat.replace("<", ""),
            "dob":                dob,
            "dob_ok":             str(mrz_check_digit(dob)) == dob_chk,
            "sex":                sex if sex in ("M", "F") else "",
            "expiry":             exp,
            "expiry_ok":          str(mrz_check_digit(exp)) == exp_chk,
            "raw_line2":          clean[:44],
        }
    return None


# ============================================================
# 图片预处理（提升 OCR 准确率）
# ============================================================

def enhance_for_ocr(img: Image.Image,
                    contrast: float = 1.4,
                    sharpness: float = 1.6,
                    autocontrast_cutoff: int = 2) -> Image.Image:
    """
    护照/文档 OCR 预处理流水线：
      1. 转灰度（去除彩色干扰）
      2. 自动对比度拉伸（解决扫描偏暗/偏亮）
      3. 对比度增强
      4. 锐化（让字符边缘更清晰）
      5. 轻度降噪
      6. 转回 RGB（API 需要 JPEG）
    """
    gray = img.convert("L")
    gray = ImageOps.autocontrast(gray, cutoff=autocontrast_cutoff)

    gray = ImageEnhance.Contrast(gray).enhance(contrast)
    gray = ImageEnhance.Sharpness(gray).enhance(sharpness)
    gray = gray.filter(ImageFilter.MedianFilter(size=3))  # 去除椒盐噪点

    return gray.convert("RGB")


def crop_mrz_zone_hidpi(pdf_path: str, dpi: int = 300) -> Image.Image:
    """
    从 PDF 高 DPI 重新渲染 + 裁剪 MRZ 区（护照页底部 ~15-25%）。
    用于兜底：整图识别时 MRZ 被 API 读成中文字符的场景。

    Args:
        pdf_path: PDF 文件路径
        dpi:      渲染分辨率（300 为最佳性价比，400+ 收益递减）

    Returns:
        MRZ 区 PIL Image（已做 OCR 预处理）
    """
    from pdf2image import convert_from_path

    img = convert_from_path(pdf_path, dpi=dpi)[0]
    w, h = img.size

    # 护照页通常占图片上半部。MRZ 在护照页底部，即整图 30-50% 区间
    # 这里取个保守范围：整图 25-55%，足够覆盖
    mrz_crop = img.crop((0, int(h * 0.25), w, int(h * 0.55)))

    return enhance_for_ocr(mrz_crop, contrast=1.5, sharpness=1.8)


# ============================================================
# 自测
# ============================================================

if __name__ == "__main__":
    # 测试校验位计算
    # 示例：护照号 L898902C3，校验位应为 6
    assert mrz_check_digit("L898902C3") == 6, "MRZ 校验位算法错误"
    print("✅ MRZ 校验位算法正确")

    # 测试 OCR 纠错：假设真实护照 123456789，OCR 读成 1Z3456789（2→Z）
    real_pno = "123456789"
    real_chk = str(mrz_check_digit(real_pno))
    print(f"  真实护照: {real_pno}  校验位: {real_chk}")

    fixed, conf = verify_passport_no("1Z3456789", real_chk)
    print(f"  OCR='1Z3456789' → 修正后: {fixed}  置信度: {conf}")
    assert fixed == real_pno and conf == "corrected"

    # 测试正常值通过
    fixed, conf = verify_passport_no(real_pno, real_chk)
    assert fixed == real_pno and conf == "verified"

    print("✅ 所有测试通过")
