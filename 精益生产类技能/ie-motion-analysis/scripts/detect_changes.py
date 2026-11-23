#!/usr/bin/env python3
"""
动作变化点检测（IE 动作分析 Step 1 辅助）

基于抽帧生成的帧图，计算相邻帧的像素差异，输出候选动作变化点
（画面明显变化的时间点），帮助确定动作拆分边界。纯 Pillow 实现，无数值库依赖。

用法:
  python detect_changes.py --manifest OUT_DIR/manifest.json [--top 40] [--threshold 0.0]
                           [--out changes.json]

输出 JSON：{"changes": [{"time_sec", "timestamp", "frame_file", "score", "level"}, ...], ...}
level: high/medium/low，按得分相对阈值划分；得分 0-1（归一化平均像素差）。
"""

import argparse
import json
import sys
from pathlib import Path

try:
    from PIL import Image, ImageChops, ImageStat
except ImportError:
    sys.stderr.write("[错误] 需要 Pillow：pip install Pillow\n")
    sys.exit(2)


def fmt_ts(sec: float) -> str:
    h = int(sec // 3600)
    m = int((sec % 3600) // 60)
    s = sec - h * 3600 - m * 60
    return f"{h:02d}:{m:02d}:{s:06.3f}"


def diff_score(a: Image.Image, b: Image.Image) -> float:
    """归一化平均像素差 0-1。"""
    size = (160, 120)
    a = a.convert("RGB").resize(size)
    b = b.convert("RGB").resize(size)
    diff = ImageChops.difference(a, b)
    stat = ImageStat.Stat(diff)
    mean = sum(stat.mean) / 3.0
    return min(1.0, mean / 255.0)


def main():
    ap = argparse.ArgumentParser(description="帧间差异变化点检测")
    ap.add_argument("--manifest", required=True, help="extract_frames.py 生成的 manifest.json")
    ap.add_argument("--top", type=int, default=40, help="返回前 N 个变化点")
    ap.add_argument("--threshold", type=float, default=0.0, help="最低得分阈值（0-1）")
    ap.add_argument("--out", default=None, help="输出 JSON 路径；默认写在 manifest 同目录")
    args = ap.parse_args()

    manifest_path = Path(args.manifest)
    mf = json.loads(manifest_path.read_text(encoding="utf-8"))
    frames = mf.get("frames", [])
    if len(frames) < 2:
        print(json.dumps({"error": "帧数不足，无法检测变化点", "frames": len(frames)}, ensure_ascii=False))
        return

    base = manifest_path.parent
    scores = []
    prev_img = None
    prev_meta = None
    for fr in frames:
        p = base / fr["file"]
        if not p.exists():
            continue
        img = Image.open(p).convert("RGB")
        if prev_img is not None:
            sc = diff_score(prev_img, img)
            t = fr["time_sec"]
            scores.append({
                "time_sec": t,
                "timestamp": fr.get("timestamp", fmt_ts(t)),
                "frame_file": fr["file"],
                "score": round(sc, 4),
            })
        prev_img = img
        prev_meta = fr

    if not scores:
        print(json.dumps({"error": "未计算出帧间差异"}, ensure_ascii=False))
        return

    selected = [s for s in scores if s["score"] >= args.threshold]
    selected.sort(key=lambda x: x["score"], reverse=True)
    selected = selected[:args.top]
    max_score = max(s["score"] for s in selected) or 1.0
    for s in selected:
        ratio = s["score"] / max_score
        s["level"] = "high" if ratio >= 0.6 else ("medium" if ratio >= 0.3 else "low")
    selected.sort(key=lambda x: x["time_sec"])

    result = {
        "manifest": str(manifest_path),
        "frame_pairs_compared": len(scores),
        "max_score": round(max_score, 4),
        "threshold": args.threshold,
        "changes": selected,
        "note": "变化点仅供定位动作边界参考，最终以实际观看画面为准",
    }
    out_path = Path(args.out) if args.out else manifest_path.parent / "changes.json"
    out_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"output": str(out_path), "changes": len(selected), "max_score": round(max_score, 4)},
                     ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
