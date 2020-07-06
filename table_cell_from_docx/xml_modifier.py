from lxml import etree
import os


class XMLModifier():

    def __init__(self, folder_name, unzipped_path):
        self.doc_xml_path = os.path.join(
            unzipped_path, folder_name, "word", "document.xml"
        )
        self.styles_xml_path = os.path.join(
            unzipped_path, folder_name, "word", "styles.xml"
        )
        self.doc_tree = None
        self.styles_tree = None
        try:
            # Some files made from .doc are not readable
            self.doc_tree = etree.parse(self.doc_xml_path)
        except BaseException:
            pass
        try:
            self.styles_tree = etree.parse(self.styles_xml_path)
        except BaseException:
            pass

    def xml_draw_border(self, color_code):
        """
        Draw table borders with the given color in document.xml and styles.xml
        Note: if the border is defined in styles.xml, it overrids the border
        defined in document.xml
        """
        self.color = color_code

        if self.doc_tree is not None:
            self.root = self.doc_tree.getroot()
            self.nsmap = self.root.nsmap
            self._doc_table_border()
            self.doc_tree.write(self.doc_xml_path, pretty_print=True)

        if self.styles_tree is not None:
            self.root = self.styles_tree.getroot()
            self.nsmap = self.root.nsmap
            self._styles_table_border()
            self.styles_tree.write(self.styles_xml_path, pretty_print=True)

    def cell_background_colorful(self, color_code, colors):
        """
        1. Change background color of every cell in document.xml in the order
        of colors
        2. Delete all cell overridings in styles.xml
        Returns:
            the maximum number of cells among the tables of this document
        """
        self.color = color_code
        self.colors = colors

        # Change background color of every cell in document.xml
        self.root = self.doc_tree.getroot()
        self.nsmap = self.root.nsmap
        num_of_cells = self._cell_background_colorful_doc()
        # Save the changed document.xml back to the same place
        self.doc_tree.write(self.doc_xml_path, pretty_print=True)

        # Delete all cell overridings in styles.xml
        self.root = self.styles_tree.getroot()
        self.nsmap = self.root.nsmap
        self._delete_cell_background_styles()
        # Save the changed styles.xml back to the same place
        self.styles_tree.write(self.styles_xml_path, pretty_print=True)
        return num_of_cells

    def _doc_table_border(self):
        """
        Draw outside border of tables with the given color in 2 steps
        1. do it in tblBorders
        2. iterate over edge cells and draw their edge borders, because
        cells border overrids tblBorders
        """
        # Find all tables
        tables = self.doc_tree.findall('.//w:tbl', self.nsmap)
        for table in tables:
            # Check if the table has more than one cell, does not have nested
            # tables and it is rectangular
            if not self._table_checks(table):
                continue

            # If there is no tblBorders add it
            properties = table.find('.//w:tblPr', self.nsmap)
            table_borders = properties.find('.//w:tblBorders', self.nsmap)
            if table_borders is None:
                table_borders = etree.SubElement(
                    properties, "{" + self.nsmap['w'] + "}" + "tblBorders"
                )
            # Change the color of the outside borders of the tables to the
            # given color
            self._color_tblBorders(table_borders)

            # If table_spacing non-zero do not do cell borders
            table_spacing = properties.find('.//w:tblCellSpacing', self.nsmap)
            if table_spacing is not None:
                table_space_w = table_spacing.get(
                    "{" + self.nsmap['w'] + "}" + "w"
                )
                # The size of table_spacing bigger than 0
                if int(table_space_w) > 0:
                    continue

            # If there is a cell border it would overwrite table border =>
            # Draw edge cell borders
            nrows = table.findall('.//w:tr', self.nsmap)
            self._color_edge_cells(nrows)

    def _color_edge_cells(self, nrows):
        """
        Color the sides of edge cells that overlaps with the table border
        """
        # Iteratate over rows and then cells inside the rows
        for idx_row, row in enumerate(nrows):
            ncells = row.findall('.//w:tc', self.nsmap)
            for idx_cell, cell in enumerate(ncells):
                cell_props = cell.find('.//w:tcPr', self.nsmap)
                cell_borders = cell_props.find(
                    './/w:tcBorders', self.nsmap
                )
                if cell_borders is None:
                    cell_borders = etree.SubElement(
                        cell_props, "{" + self.nsmap['w'] + "}" + "tcBorders")

                # First row in the table
                if idx_row == 0:
                    self._set_border(cell_borders, "top")
                    # top left cell
                    if idx_cell == 0:
                        self._set_border(cell_borders, "left")
                    # top right cell
                    if idx_cell == len(ncells) - 1:
                        self._set_border(cell_borders, "right")

                # Last row in the table
                if idx_row == len(nrows) - 1:
                    self._set_border(cell_borders, "bottom")
                    # bottom left cell
                    if idx_cell == 0:
                        self._set_border(cell_borders, "left")
                    # bottom right cell
                    if idx_cell == len(ncells) - 1:
                        self._set_border(cell_borders, "right")

                # Neither the first, nor the last row
                if idx_row != 0 and idx_row != len(nrows) - 1:
                    if idx_cell == 0:
                        self._set_border(cell_borders, "left")
                    if idx_cell == len(ncells) - 1:
                        self._set_border(cell_borders, "right")

    def _set_border(self, cell_borders, side):
        """
        Given a cell draw build the given border with the given color
        """
        cell_side = cell_borders.find('.//w:' + side, self.nsmap)
        if cell_side is None:
            cell_side = etree.SubElement(
                cell_borders, "{" + self.nsmap['w'] + "}" + side)
        self._border_params(cell_side)
        try:
            del cell_side.attrib["{" + self.nsmap['w'] + "}" + "themeColor"]
        except BaseException:
            pass

    def _color_tblBorders(self, table_borders):
        """
        Change the color of the outside borders of the tables to the
        given color
        """
        for side_name in ["top", "bottom", "left", "right"]:
            side = table_borders.find('.//w:' + side_name, self.nsmap)
            if side is None:
                side = etree.SubElement(
                    table_borders, "{" + self.nsmap['w'] + "}" + side_name
                )
            self._border_params(side)

    def _border_params(self, side):
        """
        Set the border parameters
        """
        side.set("{" + self.nsmap['w'] + "}" + "val", 'single')
        side.set("{" + self.nsmap['w'] + "}" + "sz", '3')
        side.set("{" + self.nsmap['w'] + "}" + "space", '0')
        side.set("{" + self.nsmap['w'] + "}" + "color", self.color)

    def _table_rectangular(self, table):
        """
        Check if the table is rectangular: compare the number of columns in
        gridCol with the sum of cells in the row, taking into account the
        number of cells covered by the spanning cells
        """
        nrows = table.findall('.//w:tr', self.nsmap)
        tblGrid = table.find('.//w:tblGrid', self.nsmap)
        ncols = tblGrid.findall('.//w:gridCol', self.nsmap)
        number_of_cols = len(ncols)
        for row in nrows:
            ncells = row.findall('.//w:tc', self.nsmap)
            if len(ncells) == number_of_cols:
                continue
            # Number of cells in the row is not equal to the number of columns
            number_of_cells = 0
            for cell in ncells:
                cell_props = cell.find('.//w:tcPr', self.nsmap)
                cell_grid_span = cell_props.find('.//w:gridSpan', self.nsmap)
                val = 1
                if cell_grid_span is not None:
                    val = cell_grid_span.get(
                        "{" + self.nsmap['w'] + "}" + "val"
                    )
                number_of_cells += int(val)
            if number_of_cells < number_of_cols:
                return False
        return True

    def _one_cell_check(self, table, ):
        """
        Check if the table consists of only one cell
        """
        nrows = table.findall('.//w:tr', self.nsmap)
        if len(nrows) > 1:
            return True
        ncells = nrows[0].findall('.//w:tc', self.nsmap)
        if len(ncells) == 1:
            return False
        return True

    def _table_checks(self, table):
        # If a table consists of only one cell => not a table
        if not self._one_cell_check(table):
            return False
        # If a table has other tables inside => skip this table
        if table.find('.//w:tbl', self.nsmap):
            return False
        # If a table does not have rectangular shape => skip this table
        if not self._table_rectangular(table):
            return False
        return True

    def _styles_table_border(self):
        """
        Change background color in styles.xml
        """
        # Find all styles
        styles = self.styles_tree.findall('.//w:style', self.nsmap)
        for style in styles:
            style_type = style.get("{" + self.nsmap['w'] + "}" + "type")
            if style_type != "table":
                continue
            properties = style.find('.//w:tblPr', self.nsmap)
            if properties is None:
                continue
            # Change border color of the table
            table_borders = properties.find('.//w:tblBorders', self.nsmap)
            if table_borders is not None:
                self._color_tblBorders(table_borders)

    def _cell_background_colorful_doc(self):
        """
        Change background color of every cell in document.xml in the order of
        self.colors
        """
        tables = self.doc_tree.findall('.//w:tbl', self.nsmap)
        max_num_cells = 0
        for table in tables:
            k = 0  # Count the number of cells
            # Check if table has more than one cell, does not have nested
            # tables and it is rectangular
            if not self._table_checks(table):
                continue
            # Change background color of cells
            nrows = table.findall('.//w:tr', self.nsmap)
            for row in nrows:
                ncells = row.findall('.//w:tc', self.nsmap)
                for cell in ncells:
                    cell_props = cell.find('.//w:tcPr', self.nsmap)
                    if cell_props is None:
                        cell_props = etree.SubElement(
                            cell, "{" + self.nsmap['w'] + "}" + "shd")
                    cell_background = cell_props.find('.//w:shd', self.nsmap)
                    if cell_background is None:
                        cell_background = etree.SubElement(
                            cell_props, "{" + self.nsmap['w'] + "}" + "shd")
                    cell_background.set(
                        "{" + self.nsmap['w'] + "}" + "fill",
                        self.colors[k][0])
                    cell_background.set(
                        "{" + self.nsmap['w'] + "}" + "color",
                        self.colors[k][0])
                    cell_background.set(
                        "{" + self.nsmap['w'] + "}" + "val", "clear")
                    k += 1
                    # Delete if there is any color in "themeFill"
                    try:
                        del cell_background.attrib[
                            "{" + self.nsmap['w'] + "}" + "themeFill"]
                    except BaseException:
                        pass
                    # Remove the paragraph background if any in document.xml
                    self._remove_par_background(cell)
            if k > max_num_cells:
                max_num_cells = k
        return max_num_cells

    def _remove_par_background(self, cell):
        """
        The paragraph background overrrids the cell background
        Remove the paragraph background if any in document.xml
        """
        for par in cell.findall('.//w:p', self.nsmap):
            par_props = par.find('.//w:pPr', self.nsmap)
            if par_props is None:
                continue
            par_bkgd = par_props.find('.//w:shd', self.nsmap)
            try:
                del par_bkgd.attrib["{" + self.nsmap['w'] + "}" + "fill"]
            except BaseException:
                pass
            try:
                del par_bkgd.attrib["{" + self.nsmap['w'] + "}" + "themeFill"]
            except BaseException:
                pass
            self._remove_char_background(par)

    def _remove_char_background(self, paragraph):
        """
        Character background overrids the paragraph background
        Remove character background if any in document.xml
        """
        for par_row in paragraph.findall('.//w:r', self.nsmap):
            par_row_prop = par_row.find('.//w:rPr', self.nsmap)
            if par_row_prop is None:
                continue
            par_row_background = par_row_prop.find('.//w:shd', self.nsmap)
            if par_row_background is not None:
                par_row_prop.remove(par_row_background)
            par_row_highlight = par_row_prop.find('.//w:highlight', self.nsmap)
            if par_row_highlight is not None:
                par_row_prop.remove(par_row_highlight)

    def _delete_cell_background_styles(self):
        """
        Remove the paragraph background if any in styles.xml
        """
        styles = self.styles_tree.findall('.//w:style', self.nsmap)
        for style in styles:
            style_type = style.get("{" + self.nsmap['w'] + "}" + "type")
            if style_type != "paragraph":
                continue
            properties = style.find('.//w:pPr', self.nsmap)
            if properties is not None:
                shd = properties.find('.//w:shd', self.nsmap)
                if shd is not None:
                    properties.remove(shd)
