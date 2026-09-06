"""阶段二视频配方层：命令构造纯函数 + 真 ffmpeg 合成冒烟（无 ffmpeg 则跳过）。"""

import os
import subprocess

import pytest

import video


def test_build_page_args_ken_burns_16_9():
    args = video.build_page_args("in.png", "out.mp4", seconds=4.0)
    joined = " ".join(args)
    assert "zoompan" in joined and "s=1280x720" in joined     # Ken Burns + 16:9 画幅
    assert "yuv420p" in joined and "libx264" in joined        # 兼容性编码
    assert "-t 4.000" in joined
    assert args[0] == "ffmpeg" and args[-1] == "out.mp4"


def test_build_final_args_xfade_offsets():
    clips = [f"p{i}.mp4" for i in range(3)]  # 3 片段、每页 4s、转场 0.8s
    args = video.build_final_args(clips, "out.mp4", seconds=4.0, fade=0.8)
    joined = " ".join(args)
    assert joined.count("-i") == 3
    # 转场起点：第 k 个 = k×(4−0.8) → 3.2 / 6.4
    assert "offset=3.200" in joined and "offset=6.400" in joined
    assert "-map [v2]" in joined
    assert args[-1] == "out.mp4"


def test_build_final_args_rejects_single_clip():
    with pytest.raises(ValueError):
        video.build_final_args(["only.mp4"], "out.mp4")


def _make_png(path):
    from PIL import Image
    Image.new("RGB", (1280, 720), (30, 60, 120)).save(path)


@pytest.mark.skipif(not video.ffmpeg_path(), reason="未安装 ffmpeg")
def test_synthesize_real_ffmpeg(tmp_path):
    imgs = []
    for i in (1, 2):
        p = tmp_path / f"slide_{i}.png"
        _make_png(p)
        imgs.append(str(p))
    out = str(tmp_path / "out.mp4")
    done = []
    video.synthesize(imgs, out, seconds=1.0, fade=0.4,
                     progress_cb=lambda d, t: done.append((d, t)))
    assert os.path.isfile(out) and os.path.getsize(out) > 0
    assert done == [(1, 2), (2, 2)]
    assert not os.path.exists(out + "_parts")  # 临时片段已清理

    probe = video.ffprobe_path()
    if probe:
        r = subprocess.run(
            [probe, "-v", "error", "-select_streams", "v:0",
             "-show_entries", "stream=width,height,codec_name:format=duration",
             "-of", "json", out],
            capture_output=True, text=True)
        info = __import__("json").loads(r.stdout)
        stream = info["streams"][0]
        assert (stream["width"], stream["height"]) == (1280, 720)
        assert stream["codec_name"] == "h264"
        # 2 页 × 1s − 1 个 0.4s 转场 = 1.6s
        assert abs(float(info["format"]["duration"]) - 1.6) < 0.5
