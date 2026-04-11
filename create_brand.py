"""
ARTILEGENZ Brand Asset Generator
Generates 5 distinct logo concepts + banners in PNG and SVG formats
Run with WinPython: python create_brand.py

Requirements: pip install Pillow
"""

import os
import math
import random
from PIL import Image, ImageDraw, ImageFont, ImageFilter

OUT = os.path.join(os.path.expanduser("~"), "Downloads", "Artilegenz_Brand")
os.makedirs(OUT, exist_ok=True)

# ─── TAGLINES ────────────────────────────────────────────────────────────────
TAGLINES = {
    1: "Break Fixed. Think Smart. Move Fast.",
    2: "AI That Gets It Done.",
    3: "Your Data. Our AI. Zero Drama.",
    4: "Smart Systems. Zero Downtime. All Yours.",
    5: "Intelligence Built In. Drama Left Out.",
}

# ─── PALETTES ─────────────────────────────────────────────────────────────────
PALETTES = {
    # Concept 1: Neural Pulse — electric lime + hot magenta on deep space
    "neural_pulse": {
        "bg":      (8, 6, 20),
        "primary": (0, 255, 128),    # electric green
        "accent":  (255, 0, 200),    # hot magenta
        "mid":     (0, 200, 255),    # cyan
        "text":    (255, 255, 255),
    },
    # Concept 2: Fuzzy Logic — fiery orange-gold spectrum
    "fuzzy_logic": {
        "bg":      (10, 5, 30),
        "primary": (255, 160, 0),    # amber
        "accent":  (255, 50, 120),   # coral-red
        "mid":     (255, 220, 0),    # golden yellow
        "text":    (255, 255, 255),
    },
    # Concept 3: LLM Brain — vivid purple-violet with teal sparks
    "llm_brain": {
        "bg":      (5, 5, 25),
        "primary": (160, 0, 255),    # vivid violet
        "accent":  (0, 240, 200),    # teal
        "mid":     (220, 80, 255),   # bright purple
        "text":    (255, 255, 255),
    },
    # Concept 4: Data Vortex — plasma blue + electric orange
    "data_vortex": {
        "bg":      (0, 8, 30),
        "primary": (0, 140, 255),    # bright blue
        "accent":  (255, 120, 0),    # electric orange
        "mid":     (0, 220, 255),    # sky blue
        "text":    (255, 255, 255),
    },
    # Concept 5: Quantum Grid — chromatic red + ice white
    "quantum_grid": {
        "bg":      (4, 0, 18),
        "primary": (255, 30, 80),    # vivid red
        "accent":  (200, 220, 255),  # ice blue
        "mid":     (255, 100, 0),    # orange-red
        "text":    (255, 255, 255),
    },
}


# ─── HELPERS ──────────────────────────────────────────────────────────────────
def lerp_color(c1, c2, t):
    return tuple(int(c1[i] + (c2[i] - c1[i]) * t) for i in range(3))


def alpha_blend(base, color, alpha):
    """Blend color over base with given alpha (0–1)."""
    return tuple(int(base[i] * (1 - alpha) + color[i] * alpha) for i in range(3))


def draw_glowing_circle(draw, cx, cy, r, color, steps=6):
    """Draw a soft glowing dot with concentric alpha circles."""
    for i in range(steps, 0, -1):
        frac = i / steps
        radius = int(r * (1 + frac * 0.8))
        alpha = int(80 * (1 - frac))
        draw.ellipse(
            [cx - radius, cy - radius, cx + radius, cy + radius],
            fill=(*color, alpha)
        )
    draw.ellipse([cx - r, cy - r, cx + r, cy + r], fill=(*color, 255))


def draw_glowing_line(draw, x1, y1, x2, y2, color, width=2):
    """Draw a line with a soft glow halo."""
    draw.line([(x1, y1), (x2, y2)], fill=(*color, 40), width=width + 6)
    draw.line([(x1, y1), (x2, y2)], fill=(*color, 90), width=width + 2)
    draw.line([(x1, y1), (x2, y2)], fill=(*color, 220), width=width)


def best_font(sizes):
    """Try to load a good font, fall back gracefully."""
    candidates = [
        "arialbd.ttf", "arial.ttf", "DejaVuSans-Bold.ttf",
        "DejaVuSans.ttf", "LiberationSans-Bold.ttf", "FreeSansBold.ttf",
    ]
    for size in sizes:
        for name in candidates:
            try:
                return ImageFont.truetype(name, size)
            except Exception:
                pass
        try:
            return ImageFont.load_default(size=size)
        except Exception:
            return ImageFont.load_default()
    return ImageFont.load_default()


def text_center(draw, x, y, text, font, color, anchor="mm"):
    draw.text((x, y), text, font=font, fill=color, anchor=anchor)


def save_png(img, name):
    path = os.path.join(OUT, name)
    img.save(path)
    print(f"  Saved: {path}")
    return path


# ─── CONCEPT 1 : NEURAL PULSE ─────────────────────────────────────────────────
def concept1_neural_pulse(W=800, H=900):
    """
    Bright neon neural network forming the letter A.
    Nodes pulse outward from the apex in electric green + hot magenta.
    """
    p = PALETTES["neural_pulse"]
    img = Image.new("RGBA", (W, H), (*p["bg"], 255))
    draw = ImageDraw.Draw(img, "RGBA")

    cx, cy = W // 2, H // 2 - 60

    # Seed random for reproducibility
    random.seed(42)

    # ── background nebula ──
    for _ in range(300):
        nx = random.randint(0, W)
        ny = random.randint(0, H)
        nr = random.randint(1, 2)
        draw.ellipse([nx - nr, ny - nr, nx + nr, ny + nr],
                     fill=(255, 255, 255, random.randint(20, 60)))

    # ── neural network nodes (concentric rings) ──
    layers = [
        [(cx, cy - 160)],                          # apex
        [(cx - 90, cy - 60), (cx + 90, cy - 60)],  # level 1
        [(cx - 170, cy + 40), (cx, cy + 60), (cx + 170, cy + 40)],  # level 2
        [(cx - 240, cy + 150), (cx - 80, cy + 160), (cx + 80, cy + 160), (cx + 240, cy + 150)],  # level 3
    ]

    node_colors = [p["primary"], p["mid"], p["accent"], p["mid"]]

    # Draw edges first
    for li in range(len(layers) - 1):
        for (ax, ay) in layers[li]:
            for (bx, by) in layers[li + 1]:
                draw_glowing_line(draw, ax, ay, bx, by, p["primary"], width=1)

    # A-shape backbone lines
    apex = layers[0][0]
    bot_l = layers[3][0]
    bot_r = layers[3][3]
    cross_l = layers[2][0]
    cross_r = layers[2][2]

    draw_glowing_line(draw, apex[0], apex[1], bot_l[0], bot_l[1], p["primary"], width=3)
    draw_glowing_line(draw, apex[0], apex[1], bot_r[0], bot_r[1], p["primary"], width=3)
    draw_glowing_line(draw, cross_l[0], cross_l[1], cross_r[0], cross_r[1], p["accent"], width=3)

    # Draw nodes
    radii = [14, 10, 8, 7]
    for li, layer in enumerate(layers):
        for (nx, ny) in layer:
            draw_glowing_circle(draw, nx, ny, radii[li], node_colors[li])

    # ── text ──
    font_big = best_font([90, 72, 60])
    font_tag = best_font([22, 18])
    font_sub = best_font([18, 14])

    text_center(draw, cx, cy + 240, "ARTILEGENZ", font_big, p["text"])
    text_center(draw, cx, cy + 340, TAGLINES[1], font_tag, p["primary"])
    text_center(draw, cx, cy + 380, "AI-Powered Enterprise Integration", font_sub, p["mid"])

    return img


# ─── CONCEPT 2 : FUZZY LOGIC ──────────────────────────────────────────────────
def concept2_fuzzy_logic(W=800, H=900):
    """
    Overlapping probability bubbles in amber-gold-coral showing fuzzy reasoning.
    The overlap zones glow bright to show emergent intelligence.
    """
    p = PALETTES["fuzzy_logic"]
    img = Image.new("RGBA", (W, H), (*p["bg"], 255))
    draw = ImageDraw.Draw(img, "RGBA")

    cx, cy = W // 2, H // 2 - 40

    # ── fuzzy circles (overlapping) ──
    circles = [
        (cx,      cy - 90,  120, p["primary"],  0.55),
        (cx - 95, cy + 50,  110, p["accent"],   0.50),
        (cx + 95, cy + 50,  110, p["mid"],      0.50),
        (cx,      cy + 10,  80,  p["primary"],  0.70),  # center overlap glow
    ]

    # Soft glow layers — outermost first
    for (bx, by, br, col, alph) in circles:
        for rr in range(5, 0, -1):
            frac = rr / 5
            r_now = int(br + br * frac * 0.6)
            a_now = int(40 * (1 - frac) * alph)
            draw.ellipse([bx - r_now, by - r_now, bx + r_now, by + r_now],
                         fill=(*col, a_now))
        draw.ellipse([bx - br, by - br, bx + br, by + br],
                     fill=(*col, int(255 * alph)))

    # Central convergence point — bright white-gold
    draw_glowing_circle(draw, cx, cy + 10, 20, (255, 240, 100))

    # ── "A" letter silhouette using dots on the circles ──
    random.seed(7)
    for angle_deg in range(0, 360, 12):
        for radius_fac in [0.6, 0.85, 1.0]:
            ang = math.radians(angle_deg)
            for (bx, by, br, col, _) in circles[:3]:
                r = br * radius_fac
                px = int(bx + r * math.cos(ang))
                py = int(by + r * math.sin(ang))
                dot_r = 3
                draw.ellipse([px - dot_r, py - dot_r, px + dot_r, py + dot_r],
                             fill=(255, 255, 255, 80))

    # Connecting arcs between circles (dashed)
    for i in range(0, 360, 20):
        ang = math.radians(i)
        r = 160
        px = int(cx + r * math.cos(ang))
        py = int(cy + r * math.sin(ang))
        draw.ellipse([px - 2, py - 2, px + 2, py + 2], fill=(*p["mid"], 120))

    # ── text ──
    font_big = best_font([90, 72, 60])
    font_tag = best_font([22, 18])
    font_sub = best_font([18, 14])

    text_center(draw, cx, cy + 230, "ARTILEGENZ", font_big, p["text"])
    text_center(draw, cx, cy + 330, TAGLINES[2], font_tag, p["primary"])
    text_center(draw, cx, cy + 370, "Fuzzy Smart. Precisely Right.", font_sub, p["mid"])

    return img


# ─── CONCEPT 3 : LLM BRAIN ────────────────────────────────────────────────────
def concept3_llm_brain(W=800, H=900):
    """
    Abstract brain silhouette made of glowing token-like text blocks and
    neural layers, rendered in vivid violet + teal.
    """
    p = PALETTES["llm_brain"]
    img = Image.new("RGBA", (W, H), (*p["bg"], 255))
    draw = ImageDraw.Draw(img, "RGBA")

    cx, cy = W // 2, H // 2 - 60

    # ── brain layers (horizontal token strips) ──
    layer_y = [cy - 160, cy - 110, cy - 60, cy - 10, cy + 40, cy + 90, cy + 130]
    layer_w = [120, 180, 240, 260, 240, 180, 120]

    token_colors = [p["primary"], p["mid"], p["accent"]]

    for li, (ly, lw) in enumerate(zip(layer_y, layer_w)):
        n_tokens = max(2, lw // 38)
        x_start = cx - lw // 2
        tok_w = lw // n_tokens - 4
        for ti in range(n_tokens):
            tx = x_start + ti * (tok_w + 4)
            col = token_colors[(li + ti) % len(token_colors)]
            alpha = random.randint(120, 220)
            glow_r = 6
            # glow
            draw.rectangle([tx - glow_r, ly - 10 - glow_r,
                             tx + tok_w + glow_r, ly + 22 + glow_r],
                            fill=(*col, 30))
            # token body
            draw.rectangle([tx, ly - 10, tx + tok_w, ly + 22],
                            fill=(*col, alpha))
            # token highlight
            draw.rectangle([tx + 2, ly - 8, tx + tok_w - 2, ly - 4],
                            fill=(255, 255, 255, 80))

    # ── vertical attention connections ──
    random.seed(99)
    for _ in range(30):
        li1 = random.randint(0, len(layer_y) - 2)
        li2 = li1 + 1
        x1 = cx + random.randint(-layer_w[li1] // 2, layer_w[li1] // 2)
        x2 = cx + random.randint(-layer_w[li2] // 2, layer_w[li2] // 2)
        draw_glowing_line(draw, x1, layer_y[li1] + 6, x2, layer_y[li2] - 10,
                          p["accent"], width=1)

    # ── outer brain halo ──
    for r in [170, 185, 200]:
        for deg in range(0, 360, 4):
            ang = math.radians(deg)
            px = int(cx + r * 1.1 * math.cos(ang))
            py = int(cy - 30 + r * math.sin(ang))
            alpha = int(60 * abs(math.sin(math.radians(deg * 2))))
            draw.ellipse([px - 2, py - 2, px + 2, py + 2],
                         fill=(*p["mid"], alpha))

    # ── text ──
    font_big = best_font([90, 72, 60])
    font_tag = best_font([22, 18])
    font_sub = best_font([18, 14])

    text_center(draw, cx, cy + 235, "ARTILEGENZ", font_big, p["text"])
    text_center(draw, cx, cy + 335, TAGLINES[3], font_tag, p["primary"])
    text_center(draw, cx, cy + 375, "Language. Logic. Lightning Fast.", font_sub, p["mid"])

    return img


# ─── CONCEPT 4 : DATA VORTEX ──────────────────────────────────────────────────
def concept4_data_vortex(W=800, H=900):
    """
    Spiraling data streams converging into a bright plasma core.
    Streams alternate plasma blue and electric orange.
    """
    p = PALETTES["data_vortex"]
    img = Image.new("RGBA", (W, H), (*p["bg"], 255))
    draw = ImageDraw.Draw(img, "RGBA")

    cx, cy = W // 2, H // 2 - 60

    # ── vortex spiral arms ──
    n_arms = 6
    for arm in range(n_arms):
        arm_angle = (2 * math.pi / n_arms) * arm
        col = p["primary"] if arm % 2 == 0 else p["accent"]
        points = []
        for step in range(120):
            t = step / 119.0
            r = 30 + 170 * t
            angle = arm_angle + t * 3.5 * math.pi
            x = cx + r * math.cos(angle)
            y = cy + r * math.sin(angle)
            points.append((x, y))

        # Draw with tapering alpha
        for i in range(len(points) - 1):
            t = i / len(points)
            alpha = int(80 + 160 * t)
            draw.line([points[i], points[i + 1]],
                      fill=(*col, alpha), width=2)

    # ── secondary data particles along arms ──
    random.seed(55)
    for _ in range(80):
        arm = random.randint(0, n_arms - 1)
        arm_angle = (2 * math.pi / n_arms) * arm
        t = random.random()
        r = 30 + 170 * t
        angle = arm_angle + t * 3.5 * math.pi
        px = int(cx + r * math.cos(angle))
        py = int(cy + r * math.sin(angle))
        col = p["primary"] if arm % 2 == 0 else p["accent"]
        dot_r = random.randint(2, 5)
        draw.ellipse([px - dot_r, py - dot_r, px + dot_r, py + dot_r],
                     fill=(*col, 200))

    # ── central plasma core ──
    for r in [40, 28, 18, 10]:
        alpha = int(180 - r * 2)
        draw.ellipse([cx - r, cy - r, cx + r, cy + r],
                     fill=(*p["mid"], alpha))
    draw.ellipse([cx - 8, cy - 8, cx + 8, cy + 8],
                 fill=(255, 255, 255, 255))

    # ── outer ring ──
    for deg in range(0, 360, 3):
        ang = math.radians(deg)
        r = 200
        px = int(cx + r * math.cos(ang))
        py = int(cy + r * math.sin(ang))
        col = p["primary"] if (deg // 30) % 2 == 0 else p["accent"]
        draw.ellipse([px - 2, py - 2, px + 2, py + 2],
                     fill=(*col, 120))

    # ── text ──
    font_big = best_font([90, 72, 60])
    font_tag = best_font([22, 18])
    font_sub = best_font([18, 14])

    text_center(draw, cx, cy + 235, "ARTILEGENZ", font_big, p["text"])
    text_center(draw, cx, cy + 335, TAGLINES[4], font_tag, p["primary"])
    text_center(draw, cx, cy + 375, "Where Data Meets Destiny.", font_sub, p["mid"])

    return img


# ─── CONCEPT 5 : QUANTUM GRID ─────────────────────────────────────────────────
def concept5_quantum_grid(W=800, H=900):
    """
    Diamond-lattice geometric structure representing quantum-state AI.
    Vivid red-crimson nodes on ice-blue grid, with color burst at center.
    """
    p = PALETTES["quantum_grid"]
    img = Image.new("RGBA", (W, H), (*p["bg"], 255))
    draw = ImageDraw.Draw(img, "RGBA")

    cx, cy = W // 2, H // 2 - 60

    # ── isometric diamond grid ──
    grid_size = 45
    cols = range(-4, 5)
    rows = range(-4, 5)

    nodes = {}
    for row in rows:
        for col in cols:
            # Diamond offset layout
            gx = cx + col * grid_size + (row % 2) * (grid_size // 2)
            gy = cy + row * (grid_size * 0.55)
            dist = math.sqrt((gx - cx) ** 2 + (gy - cy) ** 2)
            if dist <= 200:
                nodes[(row, col)] = (int(gx), int(gy))

    # Draw edges
    directions = [(0, 1), (1, 0), (1, 1), (-1, 1)]
    for (row, col), (nx, ny) in nodes.items():
        for dr, dc in directions:
            neighbor = (row + dr, col + dc)
            if neighbor in nodes:
                mx, my = nodes[neighbor]
                dist = math.sqrt((nx - cx) ** 2 + (ny - cy) ** 2)
                t = 1 - min(dist / 200, 1)
                col_edge = lerp_color(p["accent"], p["primary"], t)
                alpha = int(60 + 120 * t)
                draw.line([(nx, ny), (mx, my)],
                          fill=(*col_edge, alpha), width=1)

    # Draw nodes
    for (row, col_idx), (nx, ny) in nodes.items():
        dist = math.sqrt((nx - cx) ** 2 + (ny - cy) ** 2)
        t = 1 - min(dist / 200, 1)
        node_color = lerp_color(p["accent"], p["primary"], t)
        node_r = int(3 + 7 * t)
        # glow
        draw.ellipse([nx - node_r - 4, ny - node_r - 4,
                      nx + node_r + 4, ny + node_r + 4],
                     fill=(*node_color, 40))
        draw.ellipse([nx - node_r, ny - node_r,
                      nx + node_r, ny + node_r],
                     fill=(*node_color, 220))

    # Central quantum burst
    burst_colors = [p["primary"], p["mid"], (255, 200, 50), (255, 255, 255)]
    for i, bc in enumerate(burst_colors):
        r = 28 - i * 5
        if r > 0:
            draw.ellipse([cx - r, cy - r, cx + r, cy + r],
                         fill=(*bc, 200 - i * 40))

    # ── 4-pointed star at center ──
    star_pts = []
    for i in range(8):
        ang = math.radians(i * 45 - 22.5)
        r = 35 if i % 2 == 0 else 15
        star_pts.append((cx + r * math.cos(ang), cy + r * math.sin(ang)))
    draw.polygon(star_pts, fill=(*p["primary"], 220))

    # ── text ──
    font_big = best_font([90, 72, 60])
    font_tag = best_font([22, 18])
    font_sub = best_font([18, 14])

    text_center(draw, cx, cy + 235, "ARTILEGENZ", font_big, p["text"])
    text_center(draw, cx, cy + 335, TAGLINES[5], font_tag, p["primary"])
    text_center(draw, cx, cy + 375, "Quantum-Ready. Enterprise-Proven.", font_sub, p["mid"])

    return img


# ─── BANNERS (1920x600) ───────────────────────────────────────────────────────
def draw_banner(concept_num, W=1920, H=600):
    concept_map = {
        1: ("neural_pulse",  TAGLINES[1], "AI-Powered Enterprise Integration"),
        2: ("fuzzy_logic",   TAGLINES[2], "Fuzzy Smart. Precisely Right."),
        3: ("llm_brain",     TAGLINES[3], "Language. Logic. Lightning Fast."),
        4: ("data_vortex",   TAGLINES[4], "Where Data Meets Destiny."),
        5: ("quantum_grid",  TAGLINES[5], "Quantum-Ready. Enterprise-Proven."),
    }
    palette_key, tagline, sub = concept_map[concept_num]
    p = PALETTES[palette_key]

    img = Image.new("RGBA", (W, H), (*p["bg"], 255))
    draw = ImageDraw.Draw(img, "RGBA")

    # ── gradient stripe sweep ──
    for x in range(W):
        t = x / W
        col = lerp_color(p["bg"], lerp_color(p["primary"], p["bg"], 0.8), abs(math.sin(t * math.pi)))
        draw.line([(x, 0), (x, H)], fill=(*col, 30))

    # ── animated particle field ──
    random.seed(concept_num * 13)
    for _ in range(200):
        px = random.randint(0, W)
        py = random.randint(0, H)
        pr = random.randint(1, 4)
        pcol = random.choice([p["primary"], p["accent"], p["mid"]])
        draw.ellipse([px - pr, py - pr, px + pr, py + pr],
                     fill=(*pcol, random.randint(30, 100)))

    # ── concept-specific icon on the left ──
    icon_cx, icon_cy = W // 6, H // 2
    icon_r = 120

    if concept_num == 1:  # neural nodes
        for angle_deg in range(0, 360, 45):
            ang = math.radians(angle_deg)
            nx = int(icon_cx + icon_r * math.cos(ang))
            ny = int(icon_cy + icon_r * math.sin(ang))
            draw_glowing_line(draw, icon_cx, icon_cy, nx, ny, p["primary"], 2)
            draw_glowing_circle(draw, nx, ny, 10, p["accent"])
        draw_glowing_circle(draw, icon_cx, icon_cy, 18, p["primary"])

    elif concept_num == 2:  # fuzzy circles
        for offset, col, alph in [(-50, p["accent"], 160), (50, p["mid"], 160), (0, p["primary"], 180)]:
            r = 70
            ox = icon_cx + offset
            draw.ellipse([ox - r, icon_cy - r, ox + r, icon_cy + r],
                         fill=(*col, alph))

    elif concept_num == 3:  # token strips
        for li in range(5):
            ly = icon_cy - 60 + li * 28
            lw = [80, 120, 150, 120, 80][li]
            x0 = icon_cx - lw // 2
            col = [p["primary"], p["mid"], p["accent"], p["mid"], p["primary"]][li]
            draw.rectangle([x0, ly, x0 + lw, ly + 20], fill=(*col, 180))

    elif concept_num == 4:  # mini vortex
        for arm in range(4):
            arm_angle = (2 * math.pi / 4) * arm
            for step in range(40):
                t = step / 39
                r = 20 + 90 * t
                angle = arm_angle + t * 2.5 * math.pi
                px_ = int(icon_cx + r * math.cos(angle))
                py_ = int(icon_cy + r * math.sin(angle))
                col = p["primary"] if arm % 2 == 0 else p["accent"]
                draw.ellipse([px_ - 2, py_ - 2, px_ + 2, py_ + 2],
                             fill=(*col, int(80 + 150 * t)))
        draw_glowing_circle(draw, icon_cx, icon_cy, 14, (255, 255, 255))

    elif concept_num == 5:  # diamond lattice
        for row in range(-3, 4):
            for col_i in range(-3, 4):
                gx = icon_cx + col_i * 28 + (row % 2) * 14
                gy = icon_cy + row * 16
                dist = math.sqrt((gx - icon_cx) ** 2 + (gy - icon_cy) ** 2)
                if dist <= 90:
                    t = 1 - dist / 90
                    nc = lerp_color(p["accent"], p["primary"], t)
                    nr = int(2 + 5 * t)
                    draw.ellipse([gx - nr, gy - nr, gx + nr, gy + nr],
                                 fill=(*nc, 200))

    # ── text block on right 2/3 ──
    tx = W // 3 + 40
    font_logo = best_font([120, 96, 80])
    font_tag = best_font([36, 28])
    font_sub = best_font([24, 20])

    draw.text((tx, H // 2 - 80), "ARTILEGENZ", font=font_logo, fill=p["text"])
    draw.text((tx, H // 2 + 60), tagline, font=font_tag, fill=p["primary"])
    draw.text((tx, H // 2 + 108), sub, font=font_sub, fill=p["mid"])

    return img


# ─── SVG GENERATORS ───────────────────────────────────────────────────────────
def concept1_svg():
    return '''<?xml version="1.0" encoding="UTF-8"?>
<svg width="400" height="450" viewBox="0 0 400 450" xmlns="http://www.w3.org/2000/svg">
  <defs>
    <radialGradient id="glow1" cx="50%" cy="50%" r="50%">
      <stop offset="0%" stop-color="#00FF80" stop-opacity="0.9"/>
      <stop offset="100%" stop-color="#00FF80" stop-opacity="0"/>
    </radialGradient>
    <filter id="blur"><feGaussianBlur stdDeviation="3"/></filter>
    <filter id="glow"><feGaussianBlur stdDeviation="5" result="b"/>
      <feMerge><feMergeNode in="b"/><feMergeNode in="SourceGraphic"/></feMerge>
    </filter>
  </defs>
  <rect width="400" height="450" fill="#08061A"/>
  <!-- Spine lines -->
  <line x1="200" y1="60" x2="70"  y2="230" stroke="#00FF80" stroke-width="3" filter="url(#glow)"/>
  <line x1="200" y1="60" x2="330" y2="230" stroke="#00FF80" stroke-width="3" filter="url(#glow)"/>
  <line x1="100" y1="190" x2="300" y2="190" stroke="#FF00C8" stroke-width="3" filter="url(#glow)"/>
  <!-- Network connections -->
  <line x1="200" y1="60" x2="155" y2="130" stroke="#00C8FF" stroke-width="1" opacity="0.5"/>
  <line x1="200" y1="60" x2="245" y2="130" stroke="#00C8FF" stroke-width="1" opacity="0.5"/>
  <line x1="155" y1="130" x2="120" y2="200" stroke="#00C8FF" stroke-width="1" opacity="0.5"/>
  <line x1="155" y1="130" x2="200" y2="210" stroke="#00C8FF" stroke-width="1" opacity="0.5"/>
  <line x1="245" y1="130" x2="200" y2="210" stroke="#00C8FF" stroke-width="1" opacity="0.5"/>
  <line x1="245" y1="130" x2="280" y2="200" stroke="#00C8FF" stroke-width="1" opacity="0.5"/>
  <!-- Nodes -->
  <circle cx="200" cy="60"  r="12" fill="#00FF80" filter="url(#glow)"/>
  <circle cx="155" cy="130" r="9"  fill="#00C8FF" filter="url(#glow)"/>
  <circle cx="245" cy="130" r="9"  fill="#00C8FF" filter="url(#glow)"/>
  <circle cx="120" cy="200" r="7"  fill="#FF00C8" filter="url(#glow)"/>
  <circle cx="200" cy="210" r="7"  fill="#FF00C8" filter="url(#glow)"/>
  <circle cx="280" cy="200" r="7"  fill="#FF00C8" filter="url(#glow)"/>
  <circle cx="70"  cy="230" r="6"  fill="#00C8FF" filter="url(#glow)"/>
  <circle cx="330" cy="230" r="6"  fill="#00C8FF" filter="url(#glow)"/>
  <!-- Text -->
  <text x="200" y="310" font-family="Arial,sans-serif" font-size="52" font-weight="bold"
        text-anchor="middle" fill="white">ARTILEGENZ</text>
  <text x="200" y="360" font-family="Arial,sans-serif" font-size="14"
        text-anchor="middle" fill="#00FF80">Break Fixed. Think Smart. Move Fast.</text>
</svg>'''


def concept5_svg():
    """Quantum Grid SVG."""
    nodes_svg = ""
    edges_svg = ""
    cx, cy = 200, 200
    grid_size = 36
    for row in range(-4, 5):
        for col in range(-4, 5):
            gx = cx + col * grid_size + (row % 2) * (grid_size // 2)
            gy = cy + row * int(grid_size * 0.55)
            dist = math.sqrt((gx - cx) ** 2 + (gy - cy) ** 2)
            if dist <= 155:
                t = 1 - dist / 155
                r = int(2 + 6 * t)
                opacity = 0.4 + 0.6 * t
                color = f"rgb({int(200 * t + 200 * (1 - t))},{int(220 * (1 - t))},{int(255 * (1 - t) + 80 * t)})"
                nodes_svg += f'<circle cx="{int(gx)}" cy="{int(gy)}" r="{r}" fill="{color}" opacity="{opacity:.2f}"/>\n'
    return f'''<?xml version="1.0" encoding="UTF-8"?>
<svg width="400" height="430" viewBox="0 0 400 430" xmlns="http://www.w3.org/2000/svg">
  <defs>
    <filter id="glow"><feGaussianBlur stdDeviation="4" result="b"/>
      <feMerge><feMergeNode in="b"/><feMergeNode in="SourceGraphic"/></feMerge>
    </filter>
    <radialGradient id="burst" cx="50%" cy="50%" r="40%">
      <stop offset="0%" stop-color="white" stop-opacity="1"/>
      <stop offset="40%" stop-color="#FF1E50" stop-opacity="0.8"/>
      <stop offset="100%" stop-color="#FF1E50" stop-opacity="0"/>
    </radialGradient>
  </defs>
  <rect width="400" height="430" fill="#040012"/>
  <g filter="url(#glow)">
{nodes_svg}  </g>
  <circle cx="{cx}" cy="{cy}" r="45" fill="url(#burst)" filter="url(#glow)"/>
  <circle cx="{cx}" cy="{cy}" r="12" fill="white"/>
  <text x="200" y="310" font-family="Arial,sans-serif" font-size="52" font-weight="bold"
        text-anchor="middle" fill="white">ARTILEGENZ</text>
  <text x="200" y="358" font-family="Arial,sans-serif" font-size="13"
        text-anchor="middle" fill="#FF1E50">Intelligence Built In. Drama Left Out.</text>
</svg>'''


# ─── TRANSPARENT VERSIONS ─────────────────────────────────────────────────────
def make_transparent(img, bg_color):
    """Replace near-background pixels with full transparency."""
    img = img.convert("RGBA")
    data = img.getdata()
    new_data = []
    for pixel in data:
        r, g, b, a = pixel
        dist = math.sqrt(sum((pixel[i] - bg_color[i]) ** 2 for i in range(3)))
        if dist < 30:
            new_data.append((r, g, b, 0))
        else:
            new_data.append(pixel)
    img.putdata(new_data)
    return img


# ─── MAIN ─────────────────────────────────────────────────────────────────────
def main():
    print("=" * 60)
    print("  ARTILEGENZ Brand Generator — 5 Concepts")
    print("=" * 60)

    generators = [
        (1, "NeuralPulse",  concept1_neural_pulse,  PALETTES["neural_pulse"]["bg"]),
        (2, "FuzzyLogic",   concept2_fuzzy_logic,   PALETTES["fuzzy_logic"]["bg"]),
        (3, "LLMBrain",     concept3_llm_brain,     PALETTES["llm_brain"]["bg"]),
        (4, "DataVortex",   concept4_data_vortex,   PALETTES["data_vortex"]["bg"]),
        (5, "QuantumGrid",  concept5_quantum_grid,  PALETTES["quantum_grid"]["bg"]),
    ]

    for num, name, gen_fn, bg in generators:
        print(f"\n[Concept {num}] {name}")

        # Dark version
        img = gen_fn()
        save_png(img, f"C{num}_{name}_dark.png")

        # Light version (white background)
        light = Image.new("RGBA", img.size, (245, 245, 245, 255))
        light.paste(img, mask=img)
        save_png(light, f"C{num}_{name}_light.png")

        # Transparent version
        transp = make_transparent(img, bg)
        save_png(transp, f"C{num}_{name}_transparent.png")

    # ── Banners ──
    print("\n[Banners] 1920x600")
    for num in range(1, 6):
        banner = draw_banner(num)
        save_png(banner, f"Banner_C{num}_1920x600.png")

    # ── Square social icons 512x512 ──
    print("\n[Social Icons] 512x512")
    for num, name, gen_fn, bg in generators:
        full = gen_fn(W=512, H=512)
        # Crop to just the logo area (top 512 wide x 512 tall)
        icon = full.crop((0, 0, 512, 512))
        save_png(icon, f"Icon_C{num}_{name}_512x512.png")

    # ── SVG files ──
    print("\n[SVG Files]")
    for svg_data, fname in [
        (concept1_svg(), "C1_NeuralPulse.svg"),
        (concept5_svg(), "C5_QuantumGrid.svg"),
    ]:
        path = os.path.join(OUT, fname)
        with open(path, "w", encoding="utf-8") as f:
            f.write(svg_data)
        print(f"  Saved: {path}")

    # ── Summary sheet ──
    print("\n[Summary Sheet]")
    W, H = 2400, 1800
    sheet = Image.new("RGBA", (W, H), (15, 15, 30, 255))
    draw = ImageDraw.Draw(sheet, "RGBA")
    font_title = best_font([40, 32])
    font_label = best_font([28, 22])

    headers = ["Concept 1\nNeural Pulse", "Concept 2\nFuzzy Logic",
               "Concept 3\nLLM Brain", "Concept 4\nData Vortex",
               "Concept 5\nQuantum Grid"]

    for i, (num, name, gen_fn, bg) in enumerate(generators):
        col = i % 3
        row = i // 3
        ox = col * 800 + 40
        oy = row * 900 + 40
        thumb = gen_fn(W=720, H=820)
        sheet.paste(thumb, (ox, oy))
        draw.text((ox + 360, oy + 835), headers[i],
                  font=font_label, fill=(200, 200, 200), anchor="mt")

    save_png(sheet, "ARTILEGENZ_All_Concepts_Sheet.png")

    print("\n" + "=" * 60)
    print(f"  All files saved to: {OUT}")
    print("=" * 60)
    print("\nTaglines included:")
    for k, v in TAGLINES.items():
        print(f"  Concept {k}: {v}")


if __name__ == "__main__":
    main()
