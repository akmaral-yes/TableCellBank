import cv2
import numpy as np
import imutils


def cell_borders_detection(table_path, colors, number_of_cells):
    """
    Given a table_image where every cell has different background color
    from colors, try to find these colors and corresponding
    bounding boxes

    If table spans across more than one page, use only table on the first
    page and drop the rest

    Args:
        table_path: a path with the table image
        colors: a list of colors of cell backgrounds
        number_of_cells: the maximum number of cells in all tables in the
                        document

    Returns:
        a list of cells that the table contains[(x,y,w,h),...]
    """
    cells_list = []
    not_found_thresh = 50

    # Count how many times in a row the color was not found
    not_found = 0

    # Read the image of the colored table
    table_image = cv2.imread(table_path)

    # Iterate over colors
    for (i, color) in enumerate(colors[:number_of_cells]):
        color_rgb = color[1:]
        lower_color = np.array(color_rgb)
        upper_color = np.array(color_rgb)

        # Find a mask of the given color - binary image:
        # all black except given color, which is white
        mask = cv2.inRange(table_image, lower_color, upper_color)
        # Find all contours
        cnts = cv2.findContours(
            mask,
            cv2.RETR_EXTERNAL,
            cv2.CHAIN_APPROX_SIMPLE)
        cnts = imutils.grab_contours(cnts)

        # No countour => no presence of the color
        if len(cnts) == 0:
            not_found += 1
            if not_found == 2 and i == 1:
                # It is not a part of the table on the first page,
                # if the color of the first cell is missing 2 times in a row
                return cells_list
            # For documents with too many max number of cells(ex.10000),
            # then stop after not_found_thresh
            # Some colors might be missing because of spanning cells, etc.
            if not_found == not_found_thresh:
                break
        # Iterate over the countours and find the bounding rectangle
        for idx, c in enumerate(cnts):
            (x, y, w, h) = cv2.boundingRect(c)
            cells_list.append((x, y, w, h))

    # If one countour is nested in another keep only the most outer
    # for ex: letters like o,p,q, have countour inside
    return _drop_nested_cells(cells_list)


def _drop_nested_cells(cells_list):
    """
    Delete from the cells_list the cells that are nested in a bigger cell
    For example the countour inside "O" should be dropped
    """
    cells_list_shortened = []
    broken = False
    for cell1 in cells_list:
        for cell2 in cells_list:
            if _box_in_box_xywh(cell1, cell2) and cell1 != cell2:
                broken = True
                break
        if broken:
            # cell1 is nested in the other cell
            broken = False
            continue
        else:
            # cell1 is not nested in any of the cells
            cells_list_shortened.append(cell1)
    return cells_list_shortened


def _box_in_box_xywh(box1, box2):
    """
    Check if box1 is inside box2
    """
    x1, y1, w1, h1 = box1
    x2, y2, w2, h2 = box2
    return x1 >= x2 and x1 + w1 <= x2 + w2 and y1 >= y2 and y1 + h1 <= y2 + h2
