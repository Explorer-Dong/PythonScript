# @Time   : 2024-03-01 20:36
# @File   : Mp4ToGif.py
# @Author : Mr_Dwj

import moviepy.editor as mpy
from typing import Tuple


def Mp4ToGif(inputPath: str, outputPath: str,
             t_start: Tuple[float, float], t_end: Tuple[float, float],
             speed: float = 2) -> None:
	clip = mpy.VideoFileClip(inputPath).subclip(t_start, t_end)
	clip = clip.fx(mpy.vfx.speedx, speed)
	clip.write_gif(outputPath, fps=15)
	clip.close()


""" tutorial
import Mp4ToGif

def _run():
	inputPath = "D:\Huawei Share\Screenshot\hashQtShow.mp4"
	outputPath = "D:\desktop\hash.gif"
	t_start = (0, 3)
	t_end = (0, 18)
	Mp4ToGif.Mp4ToGif(inputPath, outputPath, t_start, t_end)


if __name__ == '__main__':
	_run()
"""
