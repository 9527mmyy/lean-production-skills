#!/usr/bin/env python3
"""
视频抽帧工具（IE 动作分析 Step 1）

从生产现场作业视频中按时间顺序导出帧图，生成带时间戳的缩略图拼图（contact sheet）
和清单 manifest.json，供后续逐帧动作拆分使用。

后端按 ffmpeg -> OpenCV -> imageio 自动选择；三者均不可用时给出明确安装提示。

用法:
  python extract_frames.py --video VIDEO --out OUT_DIR [--fps 2] [--cols 4]
                           [--start 0] [--end 60] [--thumb-width 320] [--jpeg-quality 88]
                           [--backend auto|ffmpeg|cv2|imageio] [--every-n-frames 1]
"""

import argparse
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

try:
    from PIL import Image, ImageDraw, ImageFont, ImageOps
except ImportError:
    sys.stderr.write("[错误] 需要 Pillow：pip install Pillow\n")
    sys.exit(2)


def fmt_ts(sec: float) -> str:
    h = int(sec // 3600)
    m = int((sec % 3600) // 60)
    s = sec - h * 3600 - m * 60
    return f"{h:02d}:{m:02d}:{s:06.3f}"


def load_font(size: int):
    candidates = [
        "C:/Windows/Fonts/consola.ttf", "C:/Windows/Fonts/cour.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf",
        "/System/Library/Fonts/Menlo.ttc",
    ]
    for c in candidates:
        if os.path.exists(c):
            try:
                return ImageFont.truetype(c, size)
            except Exception:
                pass
    return ImageFont.load_default()


class VideoSource:
    """统一视频读取接口：返回 (索引, 时间秒, PIL.Image)。"""

    def __init__(self, video: str, backend: str):
        self.video = video
        self.backend = backend
        self.fps = None
        self.duration = None
        self.width = None
        self.height = None
        self.frame_count = None
        self._impl = None
        self._open()

    @staticmethod
    def _find_ffmpeg():
        """按 PATH -> imageio-ffmpeg 内置二进制 的顺序查找 ffmpeg 可执行文件。"""
        p = shutil.which("ffmpeg")
        if p:
            return p
        try:
            import imageio_ffmpeg
            exe = imageio_ffmpeg.get_ffmpeg_exe()
            if exe and os.path.exists(exe):
                return exe
        except Exception:
            pass
        return None

    def _open(self):
        self._ffmpeg = self._find_ffmpeg()
        if self.backend in ("auto", "ffmpeg") and self._ffmpeg:
            self._open_ffmpeg()
            return
        if self.backend in ("auto", "cv2"):
            try:
                import cv2
                self._impl = "cv2"
                cap = cv2.VideoCapture(self.video)
                if not cap.isOpened():
                    raise RuntimeError("OpenCV 无法打开视频")
                self.fps = cap.get(cv2.CAP_PROP_FPS) or 25.0
                self.frame_count = int(cap.get(cv2.CAP_PROP_FRAME_COUNT) or 0)
                self.width = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH) or 0)
                self.height = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT) or 0)
                self.duration = self.frame_count / self.fps if self.fps else None
                self._cap = cap
                return
            except Exception as e:
                if self.backend == "cv2":
                    raise
                sys.stderr.write(f"[提示] OpenCV 后端不可用：{e}\n")
        if self.backend in ("auto", "imageio"):
            try:
                import imageio.v2 as imageio
                self._impl = "imageio"
                reader = imageio.get_reader(self.video)
                meta = reader.get_meta_data()
                self.fps = meta.get("fps") or 25.0
                self.duration = meta.get("duration")
                size = meta.get("size") or (0, 0)
                self.width, self.height = size[0], size[1]
                self.frame_count = int(self.duration * self.fps) if self.duration else None
                self._reader = reader
                return
            except Exception as e:
                if self.backend == "imageio":
                    raise
                sys.stderr.write(f"[提示] imageio 后端不可用：{e}\n")
        raise RuntimeError(
            "未找到可用的视频解码后端。请安装其一：\n"
            "  pip install opencv-python-headless\n"
            "  或 pip install imageio-ffmpeg\n"
            "  或将 ffmpeg 加入 PATH"
        )

    def _open_ffmpeg(self):
        self._impl = "ffmpeg"
        # 用 ffprobe 获取元信息；没有 ffprobe 时解析 ffmpeg -i 输出
        duration = fps = width = height = None
        if shutil.which("ffprobe"):
            cmd = ["ffprobe", "-v", "error", "-select_streams", "v:0",
                   "-show_entries", "stream=width,height,r_frame_rate,duration:format=duration",
                   "-of", "json", self.video]
            out = subprocess.run(cmd, capture_output=True, text=True).stdout
            try:
                info = json.loads(out)
                st = info["streams"][0]
                width, height = st.get("width"), st.get("height")
                num, den = (st.get("r_frame_rate") or "25/1").split("/")
                fps = float(num) / float(den)
                duration = float(st.get("duration") or info.get("format", {}).get("duration") or 0)
            except Exception:
                pass
        if duration is None:
            out = subprocess.run([self._ffmpeg, "-i", self.video], capture_output=True, text=True).stderr
            import re
            dm = re.search(r"Duration:\s*(\d+):(\d+):(\d+\.\d+)", out)
            if dm:
                duration = int(dm[1]) * 3600 + int(dm[2]) * 60 + float(dm[3])
            fm = re.search(r"(\d+(?:\.\d+)?)\s*fps", out)
            if fm:
                fps = float(fm[1])
            sm = re.search(r"(\d{3,4})x(\d{3,4})", out)
            if sm:
                width, height = int(sm[1]), int(sm[2])
        self.fps = fps or 25.0
        self.duration = duration
        self.width, self.height = width, height

    def iter_frames(self, every_n: int = 1):
        if self._impl == "cv2":
            idx = 0
            while True:
                ok, frame = self._cap.read()
                if not ok:
                    break
                if idx % every_n == 0:
                    import cv2
                    img = Image.fromarray(cv2.cvtColor(frame, cv2.COLOR_BGR2RGB))
                    yield idx, idx / self.fps, img
                idx += 1
            self._cap.release()
        elif self._impl == "imageio":
            for idx, frame in enumerate(self._reader):
                if idx % every_n == 0:
                    yield idx, idx / self.fps, Image.fromarray(frame)
            self._reader.close()
        elif self._impl == "ffmpeg":
            # 直接让 ffmpeg 逐帧输出 PNG 到临时目录，保证"每帧"能力
            tmp = Path(self.video).parent / f".ie_frames_tmp_{os.getpid()}"
            tmp.mkdir(exist_ok=True)
            vf = f"select='not(mod(n\\,{every_n}))'" if every_n > 1 else "null"
            cmd = [self._ffmpeg, "-y", "-loglevel", "error", "-i", self.video,
                   "-vf", vf, "-fps_mode", "passthrough", str(tmp / "f_%08d.png")]
            r = subprocess.run(cmd, capture_output=True, text=True)
            if r.returncode != 0:
                # 兼容老版本 ffmpeg（无 -fps_mode）
                cmd = [self._ffmpeg, "-y", "-loglevel", "error", "-i", self.video,
                       "-vf", vf, "-vsync", "0", str(tmp / "f_%08d.png")]
                subprocess.run(cmd, check=True)
            for i, p in enumerate(sorted(tmp.glob("f_*.png"))):
                yield i * every_n, (i * every_n) / self.fps, Image.open(p).convert("RGB")
                p.unlink()
            try:
                tmp.rmdir()
            except OSError:
                shutil.rmtree(tmp, ignore_errors=True)

    def close(self):
        pass


def build_contact_sheets(frames_dir: Path, out_dir: Path, cols: int, thumb_w: int, quality: int):
    """frames_dir 内为 frame_*.jpg；生成 contact_sheets/sheet_*.jpg。"""
    files = sorted(frames_dir.glob("frame_*.jpg"))
    if not files:
        return []
    sheet_dir = out_dir / "contact_sheets"
    sheet_dir.mkdir(parents=True, exist_ok=True)
    rows = 4
    per_sheet = cols * rows
    font = load_font(max(14, thumb_w // 18))
    label_h = font.size + 10
    produced = []
    for s in range(0, len(files), per_sheet):
        chunk = files[s:s + per_sheet]
        thumbs = []
        for fp in chunk:
            im = Image.open(fp).convert("RGB")
            w = thumb_w
            h = int(im.height * w / im.width)
            im = im.resize((w, h), Image.LANCZOS)
            canvas = Image.new("RGB", (w, h + label_h), "white")
            canvas.paste(im, (0, 0))
            d = ImageDraw.Draw(canvas)
            d.rectangle([0, h, w, h + label_h], fill=(30, 30, 30))
            d.text((5, h + 4), fp.stem, fill="white", font=font)
            thumbs.append(canvas)
        th = thumbs[0].height
        sheet = Image.new("RGB", (cols * thumb_w, rows * th), (235, 235, 235))
        d = ImageDraw.Draw(sheet)
        for i, t in enumerate(thumbs):
            x = (i % cols) * thumb_w
            y = (i // cols) * th
            sheet.paste(t, (x, y))
            d.rectangle([x, y, x + thumb_w - 1, y + th - 1], outline=(120, 120, 120), width=1)
        out = sheet_dir / f"sheet_{s // per_sheet + 1:03d}.jpg"
        sheet.save(out, quality=quality)
        produced.append(str(out))
    return produced


def main():
    ap = argparse.ArgumentParser(description="视频抽帧与时间戳拼图生成")
    ap.add_argument("--video", required=True, help="视频文件路径")
    ap.add_argument("--out", required=True, help="输出目录")
    ap.add_argument("--fps", type=float, default=2.0, help="每秒抽取帧数（默认 2）")
    ap.add_argument("--start", type=float, default=0.0, help="起始秒")
    ap.add_argument("--end", type=float, default=None, help="结束秒")
    ap.add_argument("--cols", type=int, default=4, help="拼图列数")
    ap.add_argument("--thumb-width", type=int, default=320, help="拼图缩略图宽度")
    ap.add_argument("--jpeg-quality", type=int, default=88)
    ap.add_argument("--backend", choices=["auto", "ffmpeg", "cv2", "imageio"], default="auto")
    ap.add_argument("--every-n-frames", type=int, default=1,
                    help="每 N 原始帧取 1 帧；设为 1 且 --fps 高时可近似逐帧")
    args = ap.parse_args()

    out_dir = Path(args.out)
    frames_dir = out_dir / "frames"
    frames_dir.mkdir(parents=True, exist_ok=True)

    src = VideoSource(args.video, args.backend)
    step = 1.0 / args.fps if args.fps > 0 else 1.0
    end = args.end if args.end is not None else (src.duration or 1e9)

    manifest = {
        "video": str(Path(args.video).resolve()),
        "backend": src._impl,
        "source_fps": src.fps,
        "source_duration_sec": src.duration,
        "resolution": [src.width, src.height],
        "sample_fps": args.fps,
        "start": args.start,
        "end": end,
        "frames": [],
    }

    next_t = args.start
    saved = 0
    for idx, t, img in src.iter_frames(args.every_n_frames):
        if t < args.start:
            continue
        if t > end:
            break
        if t + 1e-6 < next_t:
            continue
        next_t += step
        name = f"frame_{saved + 1:06d}.jpg"
        img.save(frames_dir / name, quality=args.jpeg_quality)
        manifest["frames"].append({
            "file": f"frames/{name}",
            "frame_index": idx,
            "time_sec": round(t, 3),
            "timestamp": fmt_ts(t),
        })
        saved += 1
    src.close()

    sheets = build_contact_sheets(frames_dir, out_dir, args.cols, args.thumb_width, args.jpeg_quality)
    manifest["contact_sheets"] = sheets
    manifest["total_sampled_frames"] = saved
    manifest_path = out_dir / "manifest.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    print(json.dumps({
        "manifest": str(manifest_path),
        "backend": src._impl,
        "sampled_frames": saved,
        "contact_sheets": len(sheets),
        "source_duration_sec": src.duration,
        "source_fps": src.fps,
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
