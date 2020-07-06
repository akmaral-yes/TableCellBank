import cv2
import os
from skimage.measure import compare_ssim
import imutils
import numpy as np
from operator import itemgetter


def pixelwisecomp(image_name, images_fuchsia_path, images_aqua_path,
                  table_folder):
    """
    Given two images where only the outside table borders are of different
    color, detect the tables positions
    """
    # Read a page image with potential tables of fuchsia color border,
    # convert it to the gray scale image
    image_fuchsia_path = os.path.join(images_fuchsia_path, image_name)
    img_fuchsia = cv2.imread(image_fuchsia_path)
    gray_fuchsia = cv2.cvtColor(img_fuchsia, cv2.COLOR_BGR2GRAY)

    # Read a page image with potential tables of aqua color border,
    # convert it to the gray scale image
    image_aqua_path = os.path.join(images_aqua_path, image_name)
    img_aqua = cv2.imread(image_aqua_path)
    gray_aqua = cv2.cvtColor(img_aqua, cv2.COLOR_BGR2GRAY)

    # Compare the above images: score 1 - the images are the same,
    # score 0  - the images do not have any same pixel
    (score, diff) = compare_ssim(gray_fuchsia, gray_aqua, full=True)

    # No table in the images
    if score == 1:
        return

    # Convert the difference image to binary
    diff = (diff * 255).astype("uint8")
    thresh = cv2.threshold(
        diff, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)[1]

    # Find countours
    cnts = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    cnts = imutils.grab_contours(cnts)

    rects = []
    for c in cnts:
        # Find the bounding rectangle over countour c
        (x, y, w, h) = cv2.boundingRect(c)

        # Skip if the table image is less than 10 in width or height
        if w <= 10 or h <= 10:
            continue

        # Sort out small wrongly cropped images
        if w < 500 and h < 500:
            continue

        rects.append((x, y, w, h))

    # Sort tables by top y-coord
    rects.sort(key=itemgetter(1))
    image_dict = {}

    color_rgb = (255, 0, 255)  # FUCHSIA
    # Iterate over all tables in the image
    for idx, (x, y, w, h) in enumerate(rects):
        table_image = img_fuchsia[y:(y + h), x:(x + w)]

        # Remove 5 pixels from every side which usually contains table border
        table_wo_borders = img_fuchsia[(
            y + 5):(y + h - 5), (x + 5):(x + w - 5)]

        # If the table_wo_borders has FUCHSIA inside: it means
        # the table has nested tables or non-rectangular shape
        if not detect_color_presence(table_wo_borders, color_rgb):
            table_name = image_name[:-4] + "_" + str(idx) + ".png"
            table_path = os.path.join(table_folder, table_name)
            cv2.imwrite(table_path, table_image)
            image_dict[table_name] = (x, y, w, h)

    return image_dict


def detect_color_presence(img, color_rgb):
    """
    Check if the given color_rgb is present in the img
    """
    lower_color = np.array(color_rgb)
    upper_color = np.array(color_rgb)
    # Find a mask of the given color - binary image:
    # all black except given color, which is white
    mask = cv2.inRange(img, lower_color, upper_color)
    # Find all contours
    cnts = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    cnts = imutils.grab_contours(cnts)
    # No countours = the color is absent on the image
    if len(cnts):
        return True
    return False


def crop_tables(table_name, images_path, tables_path, gt_tables_dict):
    """
    Given tables positions in gt_tables_dict crop the tables from page images
    """
    image_name = table_name.split(
        "_")[0] + "_" + table_name.split("_")[1] + ".png"
    image_path = os.path.join(images_path, image_name)
    image = cv2.imread(image_path)
    (x, y, w, h) = gt_tables_dict[table_name].loc
    table_image = image[y:y + h, x:x + w]
    tables_path = os.path.join(tables_path, table_name)
    cv2.imwrite(tables_path, table_image)
