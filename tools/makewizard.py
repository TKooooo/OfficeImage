# -*- coding: utf-8 -*-
"""
生成安装向导的侧边图 / 右上角小图（Inno modern 样式）。
    wizard-large.bmp  164x314  左侧竖条
    wizard-small.bmp   55x55   内页右上角

复用 makeicon.py 里的 SDF 渲染（同一套蓝 + 白色放大镜），保证安装程序和
插件按钮、exe 图标是同一套视觉语言。
"""
import os
import struct
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import makeicon as mi


def lerp(a, b, t):
    return a + (b - a) * t


def draw_tile(px, py, cx, cy, half, radius):
    """蓝色圆角方块，返回 (r,g,b) 或 None。坐标是像素坐标。"""
    if mi.sdf_round_rect(px, py, cx, cy, half, half, radius) > 0:
        return None
    # 对角渐变
    t = ((px - (cx - half)) + (py - (cy - half))) / (4.0 * half)
    t = 0.0 if t < 0 else (1.0 if t > 1 else t)
    return (int(lerp(mi.C_TOP[0], mi.C_BOT[0], t)),
            int(lerp(mi.C_TOP[1], mi.C_BOT[1], t)),
            int(lerp(mi.C_TOP[2], mi.C_BOT[2], t)))


def magnifier_hit(px, py, cx, cy, scale):
    """白色放大镜（圆环 + 手柄），scale 为整体大小（像素）。"""
    ring_c = (cx + (mi.RING_C[0] - 0.5) * scale, cy + (mi.RING_C[1] - 0.5) * scale)
    r = mi.RING_R * scale
    w = mi.RING_W * scale
    if mi.sdf_ring(px, py, ring_c[0], ring_c[1], r, w) <= 0:
        return True
    ax = cx + (mi.HANDLE_A[0] - 0.5) * scale
    ay = cy + (mi.HANDLE_A[1] - 0.5) * scale
    bx = cx + (mi.HANDLE_B[0] - 0.5) * scale
    by = cy + (mi.HANDLE_B[1] - 0.5) * scale
    return mi.sdf_segment(px, py, ax, ay, bx, by) <= mi.HANDLE_W * scale * 0.5


def render_large(w, h):
    """浅色竖条 + 居中的品牌方块 + 底部淡淡的水印放大镜。"""
    SS = 4
    rows = []
    bg_top = (247, 250, 255)
    bg_bot = (228, 238, 251)
    tile_c = (w * 0.5, h * 0.34)
    tile_half = 48.0
    ghost_c = (w * 0.52, h * 0.80)
    ghost_scale = 230.0
    for y in range(h):
        row = []
        for x in range(w):
            # 背景渐变
            t = y / float(h - 1)
            br, bg, bb = (lerp(bg_top[i], bg_bot[i], t) for i in range(3))

            # 水印：淡淡的放大镜（圆环 + 手柄），压在底部会被裁掉一截，像装饰
            if magnifier_hit(x + 0.5, y + 0.5, ghost_c[0], ghost_c[1], ghost_scale):
                br, bg, bb = (br * 0.86 + 46 * 0.14, bg * 0.86 + 108 * 0.14, bb * 0.86 + 212 * 0.14)

            # 右边界一条细线
            if x >= w - 2:
                br, bg, bb = (216, 226, 240)

            acc = [0.0, 0.0, 0.0]
            hits = 0
            tile_hits = 0
            for sy in range(SS):
                py = y + (sy + 0.5) / SS
                for sx in range(SS):
                    px = x + (sx + 0.5) / SS
                    c = draw_tile(px, py, tile_c[0], tile_c[1], tile_half, tile_half * 0.45)
                    if c is None:
                        continue
                    tile_hits += 1
                    if magnifier_hit(px, py, tile_c[0], tile_c[1], tile_half * 1.5):
                        acc[0] += 255; acc[1] += 255; acc[2] += 255
                        hits += 1
                    else:
                        acc[0] += c[0]; acc[1] += c[1]; acc[2] += c[2]
            if tile_hits:
                k = tile_hits / float(SS * SS)
                ink = hits / float(tile_hits)
                tr = acc[0] / tile_hits * (1 - ink) + 255 * ink
                tg = acc[1] / tile_hits * (1 - ink) + 255 * ink
                tb = acc[2] / tile_hits * (1 - ink) + 255 * ink
                br = br * (1 - k) + tr * k
                bg = bg * (1 - k) + tg * k
                bb = bb * (1 - k) + tb * k
            row.append((int(br), int(bg), int(bb)))
        rows.append(row)
    return rows


def render_small(size):
    SS = 4
    rows = []
    for y in range(size):
        row = []
        for x in range(size):
            r = g = b = 255
            tile_hits = 0
            white_hits = 0
            for sy in range(SS):
                py = y + (sy + 0.5) / SS
                for sx in range(SS):
                    px = x + (sx + 0.5) / SS
                    c = draw_tile(px, py, size / 2.0, size / 2.0, size * 0.47, size * 0.21)
                    if c is None:
                        continue
                    tile_hits += 1
                    if magnifier_hit(px, py, size / 2.0, size / 2.0, size * 0.72):
                        white_hits += 1
                    else:
                        r_, g_, b_ = c
                        r += (r_ - r) / float(tile_hits)
                        g += (g_ - g) / float(tile_hits)
                        b += (b_ - b) / float(tile_hits)
            if white_hits:
                w = white_hits / float(tile_hits)
                r = r * (1 - w) + 255 * w
                g = g * (1 - w) + 255 * w
                b = b * (1 - w) + 255 * w
            row.append((int(r), int(g), int(b)))
        rows.append(row)
    return rows


def save_bmp(path, rows):
    h = len(rows)
    w = len(rows[0])
    row_bytes = ((w * 3 + 3) // 4) * 4
    data = bytearray()
    for y in range(h - 1, -1, -1):
        line = bytearray()
        for (r, g, b) in rows[y]:
            line += bytes((b, g, r))
        line += b"\x00" * (row_bytes - w * 3)
        data += line
    hdr = struct.pack("<IiiHHIIiiII", 40, w, h, 1, 24, 0, len(data), 2835, 2835, 0, 0)
    fh = struct.pack("<HIHHI", 0x4D42, 14 + 40 + len(data), 0, 0, 14 + 40)
    with open(path, "wb") as f:
        f.write(fh + hdr + bytes(data))
    return len(fh) + len(hdr) + len(data)


def main():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    out = os.path.join(root, "assets")
    big = save_bmp(os.path.join(out, "wizard-large.bmp"), render_large(164, 314))
    small = save_bmp(os.path.join(out, "wizard-small.bmp"), render_small(55))
    print("wizard-large.bmp  %d bytes (164x314)" % big)
    print("wizard-small.bmp  %d bytes (55x55)" % small)


if __name__ == "__main__":
    sys.exit(main())
