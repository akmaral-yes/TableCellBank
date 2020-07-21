import cv2
import os
import random
from utils.preproc_utils import binarization
from utils.box_utils import bounding_rects_comp

def draw_cell_borders(table_name, cells_list, tables_path, gt_cells_path):
    """
    Draw cell borders with green color and save the images
    """
    table_path = os.path.join(tables_path, table_name)
    if len(cells_list) != 0:
        img = cv2.imread(table_path)
        img = draw_rectangles_xywh(img, cells_list, color=(0, 255, 0))
        cv2.imwrite(os.path.join(gt_cells_path, table_name), img)


def draw_rectangles_xywh(img, rect_list, color, thickness=3):
    """
    Given a list of rectangles draw them with a given color
    """
    for x, y, w, h in rect_list:
        cv2.rectangle(img, (x, y), (x + w, y + h), color, 3)
    return img


def draw_lines(table_name, tables_path, gt_rows_cols_path, gt_dict_lines):
    """
    Draw horizontal and vertical lines
    """
    table_path = os.path.join(tables_path, table_name)
    img = cv2.imread(table_path)
    vertical_lines = gt_dict_lines[table_name].vertical_lines
    horizontal_lines = gt_dict_lines[table_name].horizontal_lines
    for l in vertical_lines:
        cv2.line(img, (l[0], l[1]), (l[2], l[3]), (0, 0, 255), 4, cv2.LINE_AA)
    for l in horizontal_lines:
        cv2.line(img, (l[0], l[1]), (l[2], l[3]), (0, 0, 255), 4, cv2.LINE_AA)
    cv2.imwrite(os.path.join(gt_rows_cols_path, table_name), img)
