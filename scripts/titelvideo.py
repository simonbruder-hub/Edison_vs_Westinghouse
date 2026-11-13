"""Kurzes, stummes Video für die Frontfolie: In New York gehen die Lampen an.

Grundlage ist der Stich von William Allen Rogers für «Harper's Weekly» (1889). Das Video beginnt und endet mit dem
unveränderten Bild; dazwischen wird es abgedunkelt, dann zünden die Lampen nacheinander (eigene Bearbeitung).

Eingabe:  latex/handout/bilder/new_york_nacht_1889.jpg
Ausgabe:  assets/titelvideo.mp4 (H.264, 1280 × 664, 25 Bilder/s, ohne Ton) und assets/titelvideo_poster.jpg
Aufruf:   python3 scripts/titelvideo.py     (braucht numpy, Pillow und ffmpeg)
"""
import shutil
import subprocess
import tempfile
from pathlib import Path

import numpy as np
from PIL import Image

ROOT = Path(__file__).resolve().parent.parent
BILD = ROOT / 'latex' / 'handout' / 'bilder' / 'new_york_nacht_1889.jpg'
VIDEO = ROOT / 'assets' / 'titelvideo.mp4'
POSTER = ROOT / 'assets' / 'titelvideo_poster.jpg'
W, H, FPS = 1280, 664, 25

# Ablauf in Sekunden
T_HALTEN, T_DUNKEL, T_NACHT, T_HELL, T_ENDE = 1.0, 2.4, 3.2, 8.6, 11.0
DUNKEL = 0.12                      # Resthelligkeit der Nacht ohne Lampen
NACHTFARBE = np.array([0.60, 0.75, 1.0], dtype=np.float32)

# Lampen im zugeschnittenen Bild: (x, y, Radius in Pixeln, Zündzeitpunkt in Sekunden)
LAMPEN = [
    (132, 160, 200, 3.3),                                          # grosse Bogenlampe links
    (130, 290, 120, 4.2), (250, 308, 60, 4.5), (350, 268, 70, 4.8), (395, 322, 60, 5.0),   # Läden, Bahnhof, Tram
    (677, 317, 115, 5.5), (745, 316, 115, 5.8), (822, 312, 120, 6.1), (891, 312, 120, 6.4),
    (979, 309, 125, 6.7), (1078, 305, 130, 7.0), (1178, 302, 135, 7.3), (1234, 299, 135, 7.6),  # Arkade, von hinten nach vorn
]


def zuschnitt():
    src = Image.open(BILD).convert('RGB')
    im = src.resize((W, round(src.height * W / src.width)), Image.LANCZOS)
    oben = round((im.height - H) * 0.33)
    return im.crop((0, oben, W, oben + H))


def weich(x):
    x = np.clip(x, 0.0, 1.0)
    return x * x * (3 - 2 * x)


def zuenden(tau):
    """Helligkeit einer Lampe tau Sekunden nach dem Zünden: kurzes Flackern, dann ruhig."""
    if tau < 0:
        return 0.0
    for bis, wert in ((0.06, 0.9), (0.12, 0.15), (0.20, 1.0), (0.26, 0.5)):
        if tau < bis:
            return wert
    return 1.0 - 0.3 * float(np.exp(-(tau - 0.26) / 0.15))


def main():
    bild = zuschnitt()
    POSTER.parent.mkdir(exist_ok=True)
    bild.save(POSTER, quality=93)
    a = np.asarray(bild).astype(np.float32) / 255
    yy, xx = np.mgrid[0:H, 0:W].astype(np.float32)
    scheine = []
    for x, y, r, _ in LAMPEN:
        dy = yy - y
        ry = np.where(dy > 0, r * 2.2, r)          # nach unten weiter: Spiegelung auf der nassen Strasse
        scheine.append(np.exp(-(((xx - x) / r) ** 2 + (dy / ry) ** 2) / 2).astype(np.float32))

    tmp = Path(tempfile.mkdtemp(prefix='titelvideo_'))
    n = round(T_ENDE * FPS)
    for k in range(n + 1):
        t = k / FPS
        if t <= T_HALTEN or t >= T_HELL + 1.4:
            frame = a
        else:
            if t < T_NACHT:
                grund = 1 - (1 - DUNKEL) * weich((t - T_HALTEN) / (T_DUNKEL - T_HALTEN))
            else:
                grund = DUNKEL + (1 - DUNKEL) * weich((t - T_HELL) / 1.4)
            hell = np.full((H, W), grund, dtype=np.float32)
            for schein, (_, _, _, t0) in zip(scheine, LAMPEN):
                hell += zuenden(t - t0) * schein
            hell = np.clip(hell, 0, 1)[..., None]
            frame = a * hell * (1 - (1 - hell) * (1 - NACHTFARBE))
        Image.fromarray((np.clip(frame, 0, 1) * 255 + 0.5).astype(np.uint8)).save(tmp / f'f_{k:04d}.png', compress_level=1)

    subprocess.run(['ffmpeg', '-y', '-loglevel', 'error', '-framerate', str(FPS), '-i', str(tmp / 'f_%04d.png'),
                    '-c:v', 'libx264', '-profile:v', 'high', '-pix_fmt', 'yuv420p', '-crf', '19', '-preset', 'slow',
                    '-movflags', '+faststart', '-an', str(VIDEO)], check=True)
    shutil.rmtree(tmp)
    print('OK', VIDEO, f'{VIDEO.stat().st_size / 1e6:.1f} MB,', n + 1, 'Bilder;', POSTER)


if __name__ == '__main__':
    main()
