"""Mechanical image preparation for new handouts; never writes source files."""
from io import BytesIO
from math import ceil
from PIL import Image


def padded_picture(path, crop, content_scale):
    """Expand a white canvas; retain selected image pixels without resampling.

    The caller validates scale and crop. Pixel crop boundaries must be integral
    for padding: this avoids silently rounding a teacher's selected action.
    Returns PNG bytes and the resulting pixel dimensions.
    """
    with Image.open(path) as original:
        if getattr(original, 'n_frames', 1) != 1:
            raise ValueError('补白只接受静态图片；请先明确选用哪一帧')
        selected = original.crop(tuple(int(v) for v in crop)) if crop else original.copy()
        rgba = selected.convert('RGBA')
        width, height = rgba.size
        canvas_height = ceil(height / content_scale)
        if width * canvas_height > Image.MAX_IMAGE_PIXELS:
            raise ValueError('补白画布过大，请提高content_scale或使用较小的选图')
        canvas = Image.new('RGB', (width, canvas_height), 'white')
        top = (canvas_height - height) // 2
        canvas.paste(rgba, (0, top), rgba)
        result = BytesIO()
        canvas.save(result, format='PNG', dpi=(96, 96))
        return result.getvalue(), (width, canvas_height)
