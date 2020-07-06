def build_coord(cells_list):
    """
    Return: a list of unique x and y coordinates given cells_list([x,y,w,h],..)
    """
    x_list = []
    y_list = []
    for x, y, w, h in cells_list:
        y_list.append(y)
        y_list.append(y + h)
        x_list.append(x)
        x_list.append(x + w)
    return list(set(x_list)), list(set(y_list))


def build_lines(cells_list):
    """
    Convert a list of table cells to lists of horizontal and vertical lines
    """
    if len(cells_list) == 0:
        return [], []

    # Build lists of x, y values
    x_list, y_list = build_coord(cells_list)

    # Max and min in x- and y-coord
    x_min = min(x_list)
    x_max = max(x_list)
    y_min = min(y_list)
    y_max = max(y_list)

    horizontal_lines = []
    vertical_lines = []

    # Iterate over all y coordinates and draw horizontal lines
    # from leftmost position to the rightmost position
    y_list.sort()
    for idx_y, y in enumerate(y_list):
        horizontal_lines.append((x_min, y, x_max, y))

    # Iterate over all y coordinates and draw vertical lines
    # from topmost position to the bottommost position
    x_list.sort()
    for idx_x, x in enumerate(x_list):
        vertical_lines.append((x, y_min, x, y_max))

    # If there is small distance between 2 lines, keep only one
    if len(horizontal_lines) != 0:
        x1_prev, y1_prev, x2_prev, y2_prev = horizontal_lines[-1]
        for x1, y1, x2, y2 in reversed(horizontal_lines[:-1]):
            if y1_prev - y1 < 25:  # chosen based on experiments
                horizontal_lines.remove((x1, y1, x2, y2))
            y1_prev = y1

    if len(vertical_lines) != 0:
        x1_prev, y1_prev, x2_prev, y2_prev = vertical_lines[-1]
        for x1, y1, x2, y2 in reversed(vertical_lines[:-1]):
            if x1_prev - x1 < 20:  # chosen based on experiments
                vertical_lines.remove((x1, y1, x2, y2))
            x1_prev = x1

    return horizontal_lines, vertical_lines
