import sys
import colorsys
import numpy as np
from PIL import Image, ImageFilter
from sklearn.cluster import KMeans
from pptx import Presentation
from pptx.dml.color import RGBColor
from lxml import etree
import json
import zipfile
import tempfile
import shutil
import os
import re

VALID_MODES = {'analogous', 'complementary', 'square', 'triad'}

# 引数受け取り
# 1: 入力画像パス, 2: テンプレートパス, 3: 出力パス, 4: 配色モード, 5(任意): palette(HEX6を6個カンマ区切り)
if len(sys.argv) < 5:
    print("Error: Missing arguments")
    sys.exit(1)

input_path = sys.argv[1]
template_path = sys.argv[2]
output_path = sys.argv[3]
mode = sys.argv[4] # 'analogous', 'complementary', 'square', 'triad'
palette_arg = sys.argv[5] if len(sys.argv) >= 6 else None

if mode not in VALID_MODES:
    print(f"Error: Invalid mode '{mode}'", file=sys.stderr)
    sys.exit(1)

ANALYZE_ONLY = output_path == '-'

# --- 1. 色抽出ロジック (グラデーション・ノイズ対策) ---
def extract_dominant_color(image_path):
    img = Image.open(image_path).convert('RGB')
    img = img.resize((150, 150)) # 高速化
    
    # 輪郭線やノイズを潰す（グラデーション対策）
    img = img.filter(ImageFilter.GaussianBlur(radius=3))
    
    # 減色処理（ポスタリゼーション）
    img = img.quantize(colors=16).convert('RGB')

    img_np = np.array(img)
    pixels = img_np.reshape(-1, 3)

    # K-meansで「最も面積の広い色」を1つ抽出
    kmeans = KMeans(n_clusters=1, n_init=10)
    kmeans.fit(pixels)
    return tuple(kmeans.cluster_centers_[0].astype(int))

# --- 2. 配色生成ロジック (Adobe Color風) ---
def generate_palette(base_rgb, mode):
    r, g, b = base_rgb
    h, s, v = colorsys.rgb_to_hsv(r/255, g/255, b/255)
    
    # ベースカラーの彩度・明度を少し調整（暗すぎ/鮮やかすぎを防ぐ）
    s = max(0.3, min(0.9, s))
    v = max(0.3, min(0.95, v))

    def clamp(x, lo, hi):
        return max(lo, min(hi, x))

    def to_rgb(hh, ss, vv):
        rr, gg, bb = colorsys.hsv_to_rgb(hh % 1.0, clamp(ss, 0.0, 1.0), clamp(vv, 0.0, 1.0))
        return (int(rr * 255), int(gg * 255), int(bb * 255))

    def rgb_dist(a, b):
        return ((a[0]-b[0])**2 + (a[1]-b[1])**2 + (a[2]-b[2])**2) ** 0.5

    def push_unique(colors, hh, ss, vv, min_dist=55.0):
        # 近すぎる色が出たら、色相を少しずつ回して離す
        for step in range(18):
            candidate = to_rgb(hh + step * 0.055, ss, vv)
            if all(rgb_dist(candidate, c) >= min_dist for c in colors):
                colors.append(candidate)
                return
        # 最後の手段：距離条件を諦めて追加（ただし極力重複は避けたはず）
        colors.append(to_rgb(hh, ss, vv))

    palette = []

    # 6色すべて異なるアクセント色を作る（PowerPointのaccent1..6に対応）
    if mode == 'complementary':
        # base 2色 + それぞれの明暗バリエーション
        push_unique(palette, h,       s,        v)
        push_unique(palette, h + 0.5, s,        v)
        push_unique(palette, h,       s * 0.85, v * 1.15)
        push_unique(palette, h + 0.5, s * 0.85, v * 1.15)
        push_unique(palette, h,       s * 1.05, v * 0.78)
        push_unique(palette, h + 0.5, s * 1.05, v * 0.78)
    elif mode == 'triad':
        push_unique(palette, h,          s,        v)
        push_unique(palette, h + 1/3,    s,        v)
        push_unique(palette, h + 2/3,    s,        v)
        push_unique(palette, h,          s * 0.80, v * 1.18)
        push_unique(palette, h + 1/3,    s * 1.05, v * 0.80)
        push_unique(palette, h + 2/3,    s * 0.90, v * 1.05)
    elif mode == 'square':
        push_unique(palette, h,          s,        v)
        push_unique(palette, h + 0.25,   s,        v)
        push_unique(palette, h + 0.50,   s,        v)
        push_unique(palette, h + 0.75,   s,        v)
        push_unique(palette, h,          s * 0.80, v * 1.18)
        push_unique(palette, h + 0.50,   s * 1.05, v * 0.78)
    else:
        # analogous
        shifts = [0.0, -0.06, 0.06, -0.12, 0.12, 0.18]
        variations = [
            (1.00, 1.00),
            (0.90, 1.10),
            (1.05, 0.82),
            (0.85, 1.15),
            (1.00, 0.75),
            (0.95, 1.05),
        ]
        for (shift, (s_mul, v_mul)) in zip(shifts, variations):
            push_unique(palette, h + shift, s * s_mul, v * v_mul)

    return palette


def parse_palette_arg(palette_str):
    if not palette_str:
        return None
    parts = [p.strip() for p in palette_str.split(',') if p.strip()]
    if len(parts) != 6:
        raise ValueError('palette must have 6 hex colors')
    rgb = []
    for p in parts:
        if len(p) != 6 or not re.fullmatch(r'[0-9A-Fa-f]{6}', p):
            raise ValueError('palette color must be 6 hex chars')
        r = int(p[0:2], 16)
        g = int(p[2:4], 16)
        b = int(p[4:6], 16)
        rgb.append((r, g, b))
    return rgb

# --- 3. PPT適用 ---
def apply_theme(tpl, out, palette):
    prs = Presentation(tpl)

    # python-pptxはSlideMaster.themeのような高レベルAPIを提供しないため、
    # themeパーツ内の a:clrScheme (accent1..6) をXMLで直接書き換える。
    theme_part = None
    for rel in prs.part.rels.values():
        if rel.reltype.endswith('/theme'):
            theme_part = rel._target
            break

    if theme_part is None:
        raise RuntimeError('Theme part not found in template PPTX')

    root = etree.fromstring(theme_part.blob)
    ns = {
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    }

    clr_scheme = root.find('.//a:themeElements/a:clrScheme', namespaces=ns)
    if clr_scheme is None:
        raise RuntimeError('a:clrScheme not found in theme XML')

    def rgb_to_hex(rgb):
        r, g, b = rgb
        return f"{r:02X}{g:02X}{b:02X}"

    # 6色→accent1..6へ割り当て（重複させない）
    if len(palette) < 6:
        raise RuntimeError('Palette must contain 6 colors')

    accents = {
        'accent1': rgb_to_hex(palette[0]),
        'accent2': rgb_to_hex(palette[1]),
        'accent3': rgb_to_hex(palette[2]),
        'accent4': rgb_to_hex(palette[3]),
        'accent5': rgb_to_hex(palette[4]),
        'accent6': rgb_to_hex(palette[5]),
    }

    for accent_name, hexval in accents.items():
        accent_el = clr_scheme.find(f'a:{accent_name}', namespaces=ns)
        if accent_el is None:
            # まれに欠けている場合は作る
            accent_el = etree.SubElement(clr_scheme, f"{{{ns['a']}}}{accent_name}")

        srgb = accent_el.find('a:srgbClr', namespaces=ns)
        if srgb is None:
            # 既存がsysClr等の場合は置き換え
            for child in list(accent_el):
                accent_el.remove(child)
            srgb = etree.SubElement(accent_el, f"{{{ns['a']}}}srgbClr")
        srgb.set('val', hexval)

    theme_part._blob = etree.tostring(
        root,
        xml_declaration=True,
        encoding='UTF-8',
        standalone=True,
    )

    # 一時ファイルに.pptxとして保存
    with tempfile.NamedTemporaryFile(suffix='.pptx', delete=False) as tmp:
        tmp_pptx = tmp.name
    prs.save(tmp_pptx)

    # .potxとして出力する場合、Content-Typeを変更する必要がある
    if out.lower().endswith('.potx'):
        convert_pptx_to_potx(tmp_pptx, out)
        os.unlink(tmp_pptx)
    else:
        shutil.move(tmp_pptx, out)


def convert_pptx_to_potx(pptx_path, potx_path):
    """
    .pptxを.potxに変換する。
    [Content_Types].xml内のContentTypeを変更する。
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        # pptxをZIPとして展開
        with zipfile.ZipFile(pptx_path, 'r') as zf:
            zf.extractall(tmpdir)
        
        # [Content_Types].xmlを編集
        content_types_path = os.path.join(tmpdir, '[Content_Types].xml')
        with open(content_types_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # pptx -> potx に変更
        content = content.replace(
            'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
            'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml'
        )
        
        with open(content_types_path, 'w', encoding='utf-8') as f:
            f.write(content)
        
        # 新しいZIPファイルとして保存
        with zipfile.ZipFile(potx_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for root_dir, dirs, files in os.walk(tmpdir):
                for file in files:
                    file_path = os.path.join(root_dir, file)
                    arcname = os.path.relpath(file_path, tmpdir)
                    zf.write(file_path, arcname)

# 実行
try:
    base_color = None
    palette = parse_palette_arg(palette_arg)
    if palette is None:
        base_color = extract_dominant_color(input_path)
        palette = generate_palette(base_color, mode)

    if not ANALYZE_ONLY:
        apply_theme(template_path, output_path, palette)

    result = {
        'status': 'ok',
        'mode': mode,
        'palette': [f"{r:02X}{g:02X}{b:02X}" for (r, g, b) in palette],
        'base_color': (f"{base_color[0]:02X}{base_color[1]:02X}{base_color[2]:02X}" if base_color else None),
    }
    print(json.dumps(result, ensure_ascii=False))
except Exception as e:
    print(f"Error: {e}", file=sys.stderr)
    sys.exit(1)