import os

from PIL import Image, ImageSequence
from PIL.Image import Resampling, Palette, Dither

from utils.logger_utils import get_logger

ROOT_PATH = os.path.dirname(__file__)
logger = get_logger(os.path.join(ROOT_PATH, "logs"), "compress.log")


def compress_gif(input_path, output_path, scale_factor=None, width=None, colors=256, palette=Palette.ADAPTIVE,
                 dither=Dither.FLOYDSTEINBERG):
    with Image.open(input_path) as img:
        frames = [frame.copy() for frame in ImageSequence.Iterator(img)]
        img_info_dict = {
            "frame": len(frames),
            "width": img.width,
            "height": img.height,
            "mode": img.mode,
            "format": img.format,
            "loop": img.info['loop'],
            "duration": img.info['duration']
        }
        logger.info(f"input_path: {input_path},  img_info_dict: {img_info_dict}")
        compressed_frames = []
        for frame in frames:
            res_width = frame.width
            res_height = frame.height
            if width:
                res_width = int(width)
                res_height = int(width / (frame.width / frame.height))
            if scale_factor:
                res_width = int(frame.width * scale_factor)
                res_height = int(frame.height * scale_factor)
            frame = frame.resize(
                (res_width, res_height),
                Resampling.LANCZOS
            )
            frame = frame.convert('P', palette=palette, colors=colors, dither=dither)
            compressed_frames.append(frame)
        compressed_frames[0].save(
            output_path,
            save_all=True,
            append_images=compressed_frames[1:],
            optimize=True,
            duration=img.info['duration'],
            loop=img.info['loop']
        )
        return img_info_dict


if __name__ == '__main__':
    input_gif_path = '/Users/grantit/Desktop/750c8f0c1f22ba524e3d7cbbc96013e0.gif'
    output_gif_path = '/Users/grantit/Desktop/750c8f0c1f22ba524e3d7cbbc96013e0_cp2.gif'
    compress_gif(input_gif_path, output_gif_path, width=201, colors=32,
                 palette=Palette.ADAPTIVE,
                 dither=Dither.FLOYDSTEINBERG)
