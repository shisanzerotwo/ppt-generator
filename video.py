"""阶段二 M3：截图序列 → MP4（ffmpeg 配方层，配方参考 claude-ffmpeg-skill 的 slideshow 工作流）。

配方：单页 zoompan（Ken Burns 缓推 1→1.08）→ xfade 链式交叉淡入 → libx264/yuv420p，
画幅强制 1280×720（16:9）。

命令构造与执行分离：build_page_args / build_final_args 是纯函数（可单测命令正确性，
不必真跑视频）；synthesize 负责真实执行（逐页片段 → 拼合 → 清理临时文件）。
"""

import glob
import os
import shutil
import subprocess

W, H = 1280, 720  # 16:9，与 shot.py 截图尺寸一致
FPS = 25
PRESET = "veryfast"  # 速度优先：单机合成，画质差异可忽略


def ffmpeg_path() -> str | None:
    """定位 ffmpeg：PATH → WinGet Links → WinGet Packages（winget 装的不一定进当前 PATH）。"""
    p = shutil.which("ffmpeg")
    if p:
        return p
    links = os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\WinGet\Links\ffmpeg.exe")
    if os.path.isfile(links):
        return links
    hits = glob.glob(os.path.expandvars(
        r"%LOCALAPPDATA%\Microsoft\WinGet\Packages\Gyan.FFmpeg*\ffmpeg-*\bin\ffmpeg.exe"))
    return hits[0] if hits else None


def ffprobe_path() -> str | None:
    """ffprobe 与 ffmpeg 同 bin 目录；找不到返回 None（验证步骤可跳过）。"""
    ff = ffmpeg_path()
    if ff:
        probe = os.path.join(os.path.dirname(ff), "ffprobe.exe")
        if os.path.isfile(probe):
            return probe
    return shutil.which("ffprobe")


def build_page_args(img: str, out: str, seconds: float = 4.0, fps: int = FPS,
                    size: tuple = (W, H)) -> list[str]:
    """单页 Ken Burns 片段：zoom 缓推 1→1.08（居中），输出定长片段。"""
    d = int(round(seconds * fps))
    w, h = size
    step = (1.08 - 1) / d
    vf = (f"zoompan=z='min(zoom+{step:.6f},1.08)':d={d}:"
          f"x='iw/2-(iw/zoom/2)':y='ih/2-(ih/zoom/2)':s={w}x{h},"
          f"scale={w}:{h},format=yuv420p")
    return ["ffmpeg", "-y", "-loop", "1", "-framerate", str(fps), "-i", img,
            "-t", f"{seconds:.3f}", "-vf", vf, "-r", str(fps),
            "-c:v", "libx264", "-preset", PRESET, "-pix_fmt", "yuv420p", out]


def build_final_args(page_clips: list[str], out: str, seconds: float = 4.0,
                     fade: float = 0.8, fps: int = FPS) -> list[str]:
    """xfade 链式拼合：第 k 个转场起点 = k×(每页时长−转场时长)。"""
    n = len(page_clips)
    if n < 2:
        raise ValueError("至少需要两个片段才能转场拼合")
    inputs: list[str] = []
    for c in page_clips:
        inputs += ["-i", c]
    parts: list[str] = []
    prev = "0:v"
    for k in range(1, n):
        offset = k * (seconds - fade)
        out_label = f"v{k}"
        parts.append(f"[{prev}][{k}]xfade=transition=fade:duration={fade}:offset={offset:.3f}[{out_label}]")
        prev = out_label
    return (["ffmpeg", "-y"] + inputs +
            ["-filter_complex", ";".join(parts), "-map", f"[{prev}]",
             "-r", str(fps), "-c:v", "libx264", "-preset", PRESET,
             "-pix_fmt", "yuv420p", out])


def _run(args: list[str], err: str, timeout: float = 300):
    """带超时执行：ffmpeg 挂死（磁盘满/杀毒拦截）时不至于永久占用 _video_jobs。"""
    try:
        p = subprocess.run(args, capture_output=True, text=True, timeout=timeout)
    except subprocess.TimeoutExpired as e:
        raise RuntimeError(f"{err}：执行超时（>{timeout:.0f}s），已终止") from e
    if p.returncode != 0:
        raise RuntimeError(f"{err}：{p.stderr[-400:]}")


def synthesize(shots: list[str], out_path: str, seconds: float = 4.0,
               fade: float = 0.8, progress_cb=None) -> str:
    """执行合成：逐页片段 → xfade 拼合 → 清理临时片段，返回 out_path。

    progress_cb(done, total) 用于日志进度。
    """
    if not 0 < fade < seconds:
        raise ValueError(f"转场时长需在 (0, 每页时长) 之间，当前 fade={fade}, seconds={seconds}"
                         "（fade≥seconds 会让 ffmpeg 静默产出丢页的坏视频，审计 L1）")
    ff = ffmpeg_path()
    if not ff:
        raise RuntimeError("未找到 ffmpeg，请先安装：winget install ffmpeg")
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    tmp_dir = out_path + "_parts"
    os.makedirs(tmp_dir, exist_ok=True)
    try:
        clips = []
        for i, img in enumerate(shots, 1):
            clip = os.path.join(tmp_dir, f"page_{i:03d}.mp4")
            _run(build_page_args(img, clip, seconds), f"第 {i} 页片段合成失败",
                 timeout=max(180, seconds * 60))
            clips.append(clip)
            if progress_cb:
                progress_cb(i, len(shots))
        if len(clips) == 1:
            shutil.copyfile(clips[0], out_path)
        else:
            _run(build_final_args(clips, out_path, seconds, fade), "视频拼合失败",
                 timeout=max(300, seconds * len(clips) * 30))
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)
    return out_path
