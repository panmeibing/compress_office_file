from PIL import Image


def resize_image(original_image: Image.Image, max_size):
    if original_image.width > max_size or original_image.height > max_size:
        scale = min(max_size / original_image.width, max_size / original_image.height)
        new_width = int(original_image.width * scale)
        new_height = int(original_image.height * scale)
        resized_image = original_image.resize((new_width, new_height))
        return resized_image
    return original_image


def compress_local_image(image_path, max_size=None, quality=90):
    print(f"start compress_local_image 0000000000: {image_path}")
    img = Image.open(image_path)
    if max_size:
        img = resize_image(img, max_size)
    img.save(image_path, optimize=True, quality=quality, format=img.format)


if __name__ == '__main__':
    compress_local_image('/Users/grantit/Desktop/人工智能/ppt/media/image10.jpeg')
