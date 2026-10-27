# -*- coding: utf-8 -*-
"""
生成 officeimage.ico（多尺寸 32 位 ICO，无第三方依赖）。

图形：蓝色圆角方块 + 白色放大镜，和插件浮动按钮的视觉语言一致。
做法：每个像素 4x4 超采样，用有符号距离场（SDF）算圆角矩形 / 圆环 / 胶囊，
      覆盖率直接当 alpha，缩到 16px 也不会糊成一团。

用法：
    python tools/makeicon.py            ->  assets/officeimage.ico
"""
import math
import os
import struct
import sys
import zlib

SIZES = [16, 20, 24, 32, 40, 48, 64, 128, 256]
SS = 4  # 每个像素轴向上的超采样倍数

# 配色（和插件按钮同一套蓝）
C_TOP = (0x3B, 0x7B, 0xEA)
C_BOT = (0x1B, 0x47, 0xA8)
C_INK = (0xFF, 0xFF, 0xFF)

# 几何（单位坐标系 0..1）
MARGIN = 0.020
RADIUS = 0.215
RING_C = (0.440, 0.430)
RING_R = 0.220
RING_W = 0.108
HANDLE_A = (0.578, 0.568)
HANDLE_B = (0.782, 0.772)
HANDLE_W = 0.120


def sdf_round_rect(px, py, cx, cy, hw, hh, r):
    dx = abs(px - cx) - (hw - r)
    dy = abs(py - cy) - (hh - r)
    ax = dx if dx > 0 else 0.0
    ay = dy if dy > 0 else 0.0
    return math.hypot(ax, ay) + (min(max(dx, dy), 0.0)) - r


def sdf_ring(px, py, cx, cy, r, w):
    return abs(math.hypot(px - cx, py - cy) - r) - w * 0.5


def sdf_segment(px, py, ax, ay, bx, by):
    vx, vy = bx - ax, by - ay
    wx, wy = px - ax, py - ay
    t = (wx * vx + wy * vy) / (vx * vx + vy * vy)
    t = 0.0 if t < 0 else (1.0 if t > 1 else t)
    return math.hypot(px - (ax + t * vx), py - (ay + t * vy))


def inside_tile(px, py):
    return sdf_round_rect(px, py, 0.5, 0.5, 0.5 - MARGIN, 0.5 - MARGIN, RADIUS) <= 0.0


def inside_ink(px, py):
    if sdf_ring(px, py, RING_C[0], RING_C[1], RING_R, RING_W) <= 0.0:
        return True
    return sdf_segment(px, py, HANDLE_A[0], HANDLE_A[1], HANDLE_B[0], HANDLE_B[1]) <= HANDLE_W * 0.5


def render(size):
    """返回 rows[y][x] = (r, g, b, a)，y=0 是顶部。"""
    step = 1.0 / (size * SS)          # 单个子采样在单位坐标里的宽度
    rows = []
    for y in range(size):
        row = []
        for x in range(size):
            tile_hits = 0
            ink_hits = 0
            for sy in range(SS):
                py = (y * SS + sy + 0.5) * step
                for sx in range(SS):
                    px = (x * SS + sx + 0.5) * step
                    if inside_tile(px, py):
                        tile_hits += 1
                        if inside_ink(px, py):
                            ink_hits += 1
            total = SS * SS
            if tile_hits == 0:
                row.append((0, 0, 0, 0))
                continue
            # 对角渐变
            g = (x + y) / (2.0 * (size - 1)) if size > 1 else 0.0
            br = C_TOP[0] + (C_BOT[0] - C_TOP[0]) * g
            bg = C_TOP[1] + (C_BOT[1] - C_TOP[1]) * g
            bb = C_TOP[2] + (C_BOT[2] - C_TOP[2]) * g
            k = ink_hits / float(tile_hits)  # 只在方块内部统计白墨覆盖率
            r = br + (C_INK[0] - br) * k
            gg = bg + (C_INK[1] - bg) * k
            b = bb + (C_INK[2] - bb) * k
            a = int(round(255.0 * tile_hits / total))
            row.append((int(round(r)), int(round(gg)), int(round(b)), a))
        rows.append(row)
    return rows


def bmp_entry(size, rows):
    """32 位 BMP（DIB）+ AND 掩码，ICO 里最通用的存法。"""
    hdr = struct.pack('<IiiHHIIiiII', 40, size, size * 2, 1, 32, 0,
                      size * size * 4, 0, 0, 0, 0)
    xor = bytearray()
    for y in range(size - 1, -1, -1):          # 自下而上
        for (r, g, b, a) in rows[y]:
            # 预乘无关，BMP 存直通 alpha
            xor += bytes((b, g, r, a))
    row_bytes = ((size + 31) // 32) * 4
    mask = bytes(row_bytes * size)
    return hdr + bytes(xor) + mask


def png_bytes(size, rows):
    raw = bytearray()
    for y in range(size):
        raw.append(0)
        for (r, g, b, a) in rows[y]:
            raw += bytes((r, g, b, a))

    def chunk(tag, data):
        c = struct.pack('>I', len(data)) + tag + data
        return c + struct.pack('>I', zlib.crc32(tag + data) & 0xFFFFFFFF)

    ihdr = struct.pack('>IIBBBBB', size, size, 8, 6, 0, 0, 0)
    return (b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', ihdr) +
            chunk(b'IDAT', zlib.compress(bytes(raw), 9)) + chunk(b'IEND', b''))


def build_ico(path):
    images = []
    for s in SIZES:
        images.append((s, bmp_entry(s, render(s))))
    out = bytearray(struct.pack('<HHH', 0, 1, len(images)))
    offset = 6 + 16 * len(images)
    for (s, data) in images:
        out += struct.pack('<BBBBHHII',
                           s if s < 256 else 0,
                           s if s < 256 else 0,
                           0, 0, 1, 32, len(data), offset)
        offset += len(data)
    for (_, data) in images:
        out += data
    with open(path, 'wb') as f:
        f.write(bytes(out))
    return len(out)


def main():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    outdir = os.path.join(root, 'assets')
    if not os.path.isdir(outdir):
        os.makedirs(outdir)
    ico = os.path.join(outdir, 'officeimage.ico')
    n = build_ico(ico)
    print('ICO  %s  %d bytes' % (ico, n))

    # 预览图：256 原尺寸 + 32px / 16px 各放大到 256 宽（看小尺寸是否还认得出）
    prev = os.path.join(outdir, 'icon-preview.png')
    side = 256
    big = render(256)
    slots = [(big, 1)]                      # (像素数据, 放大倍数)
    for s in (32, 16):
        slots.append((render(s), side // s))
    gap = 8
    width = sum(side + gap for _ in slots)
    canvas = [[(255, 255, 255)] * width for _ in range(side)]
    ox = 0
    for (pixels, scale) in slots:
        n = len(pixels)
        for y in range(side):
            for x in range(side):
                px = pixels[y // scale][x // scale]
                canvas[y][ox + x] = (px[0], px[1], px[2])
        ox += side + gap

    raw = bytearray()
    for y in range(side):
        raw.append(0)
        for (r, g, b) in canvas[y]:
            raw += bytes((r, g, b, 255))

    def chunk(tag, data):
        return struct.pack('>I', len(data)) + tag + data + struct.pack('>I', zlib.crc32(tag + data) & 0xFFFFFFFF)

    png = (b'\x89PNG\r\n\x1a\n' +
           chunk(b'IHDR', struct.pack('>IIBBBBB', width, side, 8, 6, 0, 0, 0)) +
           chunk(b'IDAT', zlib.compress(bytes(raw), 9)) + chunk(b'IEND', b''))
    with open(prev, 'wb') as f:
        f.write(png)
    print('PNG  %s' % prev)


if __name__ == '__main__':
    sys.exit(main())
