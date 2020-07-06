# https://imagemagick.org/script/download.php#windows
# https://www.ghostscript.com/download/gsdnld.html
import argparse
import os
import pandas as pd
import shutil
import multiprocessing
from line_builder import build_lines
from xml_modifier import XMLModifier
from cell_detector import cell_borders_detection
from converter import save_docx, unpack_zip, docx_to_pdf, pdf_to_image
from table_detector import pixelwisecomp, crop_tables
from utils.file_utils import save_dict, append_to_file
from utils.draw_utils import draw_lines, draw_cell_borders


def create_docs(docx_path, docx_names, output_path, multiproc, debug):
    """
    For a set of .docx documents find tables, crop them, build ground truth for
    cell, separating horizontal and vertical line positions
    """
    dirs = Directories(output_path)
    colors = Colors()
    dirs.create_folders()
    wrapper = DocProcessorWrapper(docx_path, colors, dirs, debug)
    if multiproc:
        # Parallel run
        processes_number = multiprocessing.cpu_count()
        p = multiprocessing.Pool(processes_number)
        p.map(wrapper, docx_names)
    else:
        # Sequential run
        for docx_name in docx_names:
            wrapper(docx_name)
    dirs.delete_folders()


class DocProcessorWrapper():

    def __init__(self, docx_path, colors, dirs, debug):
        self.docx_path = docx_path
        self.colors = colors
        self.dirs = dirs
        self.debug = debug

    def __call__(self, docx_name):
        doc = DocProcessor(docx_name, self.docx_path, self.colors,
                           self.dirs, self.debug)
        doc.retrieve_tables_structure()


class Directories():

    def __init__(self, output_path):
        # Absolute path is necessary for converting .docx for .pdf
        self.output_path = os.path.abspath(output_path)
        self.unzipped_path = os.path.join(self.output_path, "unzipped")
        self.fuchsia_docx_path = os.path.join(self.output_path, "docx_fuchsia")
        self.fuchsia_pdf_path = os.path.join(self.output_path, "pdf_fuchsia")
        self.aqua_docx_path = os.path.join(self.output_path, "docx_aqua")
        self.aqua_pdf_path = os.path.join(self.output_path, "pdf_aqua")
        self.fuchsia_images_path = os.path.join(
            self.output_path, "images_fuchsia"
        )
        self.aqua_images_path = os.path.join(self.output_path, "images_aqua")
        self.tables_path = os.path.join(self.output_path, "table_fuchsia")
        self.gt_tables_dict_path = os.path.join(
            self.output_path, "gt_tables_dict")
        self.color_docx_path = os.path.join(self.output_path, "docx_color")
        self.color_pdf_path = os.path.join(self.output_path, "pdf_color")
        self.color_images_path = os.path.join(self.output_path, "images_color")
        self.color_tables_path = os.path.join(self.output_path, "table_color")
        self.gt_cells_path = os.path.join(self.output_path, "gt_cells")
        self.gt_rows_cols_path = os.path.join(self.output_path, "gt_rows_cols")
        self.temp_path = os.path.join(output_path, "_temp")

    def create_folders(self):
        os.makedirs(self.output_path, exist_ok=True)
        os.makedirs(self.unzipped_path, exist_ok=True)
        os.makedirs(self.fuchsia_docx_path, exist_ok=True)
        os.makedirs(self.fuchsia_pdf_path, exist_ok=True)
        os.makedirs(self.aqua_docx_path, exist_ok=True)
        os.makedirs(self.aqua_pdf_path, exist_ok=True)
        os.makedirs(self.fuchsia_images_path, exist_ok=True)
        os.makedirs(self.aqua_images_path, exist_ok=True)
        os.makedirs(self.tables_path, exist_ok=True)
        os.makedirs(self.gt_tables_dict_path, exist_ok=True)
        os.makedirs(self.color_docx_path, exist_ok=True)
        os.makedirs(self.color_pdf_path, exist_ok=True)
        os.makedirs(self.color_images_path, exist_ok=True)
        os.makedirs(self.color_tables_path, exist_ok=True)
        os.makedirs(self.gt_cells_path, exist_ok=True)
        os.makedirs(self.gt_rows_cols_path, exist_ok=True)
        os.makedirs(self.temp_path, exist_ok=True)

    def delete_folders(self):
        self.folders_to_del = [
            self.unzipped_path,
            self.fuchsia_docx_path,
            self.fuchsia_pdf_path,
            self.aqua_docx_path,
            self.aqua_pdf_path,
            self.fuchsia_images_path,
            self.aqua_images_path,
            self.color_docx_path,
            self.color_pdf_path,
            self.color_images_path,
            self.color_tables_path,
            self.temp_path]
        for folder in self.folders_to_del:
            if os.path.exists(folder):
                os.rmdir(folder)


class Colors():
    def __init__(self):
        # Load generated data frame of random colors
        color_code_df = pd.read_csv('../dictionaries/random_colors_100000.csv')
        # Create list HEX+BGR
        self.colors = []
        for row in color_code_df.index:
            rgb = color_code_df.loc[row, "RGB"]
            r = int(rgb.split("-")[0])
            g = int(rgb.split("-")[1])
            b = int(rgb.split("-")[2])
            self.colors.append([color_code_df.loc[row, "HEX"]] + [b, g, r])


class DocProcessor():
    """
    For .docx find tables, crop them, build ground truth for cell, separating
    horizontal and vertical line positions
    """

    def __init__(self, docx_name, docx_path, colors, dirs, debug):
        """
        Args:
            docx_name: a word document name incl. .docx
            docx_path: a path with the word document
            colors: an object that contain a list of different colors: hex+rgb
            output_path: a path to save output tables and ground truth
        """
        print("name: ", docx_name)
        self.colors = colors
        self.dirs = dirs
        self.debug = debug
        self.docx_path = docx_path
        self.docx_name = docx_name
        self.docx_file_path = os.path.join(self.docx_path, self.docx_name)
        self.fuchsia = '#FF00FF'  # fuchsia
        self.aqua = '#00ffff'  # aqua
        self.name = self.docx_name.split(".")[0]
        self.pdf_name = self.name + ".pdf"
        self.json_name = self.name + ".json"

    def unzipped_to_images(
            self,
            file_docx_path,
            file_pdf_path,
            file_images_path):
        save_docx(self.name, self.dirs.unzipped_path, file_docx_path)
        if not docx_to_pdf(self.docx_name, file_docx_path, file_pdf_path,
                           self.dirs.output_path):
            return False
        pdf_to_image(file_pdf_path, self.pdf_name, file_images_path,
                     self.dirs.output_path, self.dirs.temp_path)
        return True

    def retrieve_tables_structure(self):
        """
        Crop tables from .docx files and build ground truth of cell postions,
        separating horizontal and vertical line positions
        """
        try:
            image_names = []
            table_names = []

            # Step 1: data_unzipped
            if not unpack_zip(self.docx_name, self.docx_file_path,
                              self.dirs.unzipped_path, self.dirs.output_path):
                return
            print("Step 1 done")

            # Step 2: Draw table border with FUCHSIA color in document.xml
            # and styles.xml
            xml_modifier = XMLModifier(self.name, self.dirs.unzipped_path)
            xml_modifier.xml_draw_border(self.fuchsia)
            # unzipped_folder=>.docx=>.pdf=>.png
            if not self.unzipped_to_images(
                    self.dirs.fuchsia_docx_path,
                    self.dirs.fuchsia_pdf_path,
                    self.dirs.fuchsia_images_path):
                return
            print("Step 2 done")

            # Step 3: Draw table border with AQUA color in document.xml
            # and styles.xml
            xml_modifier.xml_draw_border(self.aqua)
            # unzipped_folder=>.docx=>.pdf=>.png
            if not self.unzipped_to_images(
                    self.dirs.aqua_docx_path,
                    self.dirs.fuchsia_pdf_path,
                    self.dirs.aqua_images_path):
                return
            print("Step 3 done")

            # Step 4: From comparing images_fuchsia vs. images_aqua get
            # tables positions
            all_image_names = os.listdir(self.dirs.fuchsia_images_path)
            gt_tables_dict = {}
            for image_name in all_image_names:
                # Only images that correspond to the current document
                if image_name.split("_")[0] == self.name:
                    image_names.append(image_name)
                    image_dict = pixelwisecomp(
                        image_name,
                        self.dirs.fuchsia_images_path,
                        self.dirs.aqua_images_path,
                        self.dirs.tables_path)
                    if image_dict is not None:
                        for table_name, loc in image_dict.items():
                            gt_tables_dict[table_name] = Table(loc)

            if not len(gt_tables_dict.keys()):
                # No tables
                append_to_file(self.dirs.output_path,
                               'no_tables.csv', self.docx_name)
                print("No tables in the document")
                return

            # Step 4: Change cells' background to different colors and count
            # the maximum number of cells in tables in this document
            num_of_cells = xml_modifier.cell_background_colorful(
                self.aqua, self.colors.colors
            )
            if num_of_cells == 0:
                append_to_file(self.dirs.output_path,
                               'no_tables.csv', self.docx_name)
                print("No tables in the document")
                return
            print("Step 4 done")

            # Step 5: # unzipped_folder=>.docx=>.pdf=>.png
            if not self.unzipped_to_images(
                    self.dirs.color_docx_path,
                    self.dirs.color_pdf_path,
                    self.dirs.color_images_path):
                return
            print("Step 5 done")

            # Step 6: Crop colored tables based on gt_tables_dict
            table_names = list(gt_tables_dict.keys())
            for table_name in table_names:
                crop_tables(table_name,
                            self.dirs.color_images_path,
                            self.dirs.color_tables_path,
                            gt_tables_dict)
            print("Step 6 done")

            # Step 7: Find cell positions
            for table_name in table_names:
                color_table_path = os.path.join(
                    self.dirs.color_tables_path, table_name)
                cells_list = cell_borders_detection(
                    color_table_path, self.colors.colors, num_of_cells)
                gt_tables_dict[table_name].cells = cells_list
            print("Step 7 done")

            # Step 8: Draw retrieved cells borders
            if self.debug:
                for table_name in table_names:
                    draw_cell_borders(
                        table_name,
                        gt_tables_dict[table_name].cells,
                        self.dirs.tables_path,
                        self.dirs.gt_cells_path)
                print("Step 8 done")

            # Step 9: Find separating horizontal and vertical lines
            for table_name in table_names:
                horizontal_lines, vertical_lines = build_lines(
                    gt_tables_dict[table_name].cells)
                # At least two rows and at least two columns
                if len(horizontal_lines) <= 2 or len(vertical_lines) <= 2:
                    del gt_tables_dict[table_name]
                    table_path = os.path.join(
                        self.dirs.tables_path, table_name)
                    os.remove(table_path)
                    continue
                gt_tables_dict[table_name].horizontal_lines = horizontal_lines
                gt_tables_dict[table_name].vertical_lines = vertical_lines
            print("Step 9 done")

            # Step 10: Draw horizontal and vertical lines
            if self.debug:
                for table_name in gt_tables_dict.keys():
                    draw_lines(
                        table_name,
                        self.dirs.tables_path,
                        self.dirs.gt_rows_cols_path,
                        gt_tables_dict)
            print("Step 10 done")

            # Step 11: Save gt_tables_dict,
            # write down to the list of processed files
            save_dict(self.dirs.gt_tables_dict_path,
                      self.json_name, gt_tables_dict)
            append_to_file(self.dirs.output_path,
                           'processed.csv', self.docx_name)
            print("Step 11 done")

        finally:
            # Delete all intermediate files
            self.clean_up(image_names, table_names)

    def clean_up(self, image_names=[], table_names=[]):
        """
        Delete all intermediate files if they exists:
        unzipped, docx, pdf, images, table_color
        Except:
            table images, gt_tables_dict
        """
        # Delete unzipped_folder
        file_unzipped_path = os.path.join(self.dirs.unzipped_path, self.name)
        if os.path.exists(file_unzipped_path):
            shutil.rmtree(file_unzipped_path)

        # Delete .docx
        folder_names = ["docx_aqua", "docx_color", "docx_fuchsia"]
        for folder in folder_names:
            file_path = os.path.join(
                self.dirs.output_path, folder, self.docx_name)
            if os.path.exists(file_path):
                os.remove(file_path)

        # Delete .pdf
        folder_names = ["pdf_aqua", "pdf_color", "pdf_fuchsia"]
        for folder in folder_names:
            file_path = os.path.join(
                self.dirs.output_path, folder, self.pdf_name)
            if os.path.exists(file_path):
                os.remove(file_path)

        # Delete images
        folder_names = ["images_aqua", "images_color", "images_fuchsia"]
        for folder in folder_names:
            for image_name in image_names:
                file_path = os.path.join(
                    self.dirs.output_path, folder, image_name)
                if os.path.exists(file_path):
                    os.remove(file_path)

        # Delete in table_color
        for table_name in table_names:
            file_path = os.path.join(
                self.dirs.output_path, "table_color", table_name
            )
            if os.path.exists(file_path):
                os.remove(file_path)


class Table():
    def __init__(self, loc):
        self.loc = loc
        self.cells = []
        self.horizontal_lines = []
        self.vertical_lines = []


def find_index(output_path, uuids, url_df):
    """
    Find the index from which continue to process documents after interuption
    """
    # All files that were processed incl. successfully and unsuccessfully
    csvs = [f for f in os.listdir(output_path) if f[-3:] == "csv"]
    files_done = []
    for csv_doc in csvs:
        csv_path = os.path.join(output_path, csv_doc)
        csv_df = pd.read_csv(csv_path, header=None)
        files = csv_df.iloc[:, 0].to_list()
        files = [f.split(".")[0] for f in files]  # file names wo extension
        files_done.extend(files)

    # Indices that expected to still be processed
    files_to_be_done = list(set(uuids) - set(files_done))
    file_name = files_to_be_done[0]
    index_min = url_df.index[url_df['uuid'] == file_name][0]
    indices_to_be_done = []
    for file_name in files_to_be_done[1:]:
        index = url_df.index[url_df['uuid'] == file_name][0]
        indices_to_be_done.append(index)
        if index_min > index:
            index_min = index

    return indices_to_be_done, index_min


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--start_idx', default='0',
                        help="Choose start index")
    parser.add_argument('--end_idx', default='0',
                        help="Choose end index(incl. end_idx)")
    parser.add_argument('--do', default='run',
                        help="Process documents(run) or \
                        find index to start (index)")
    parser.add_argument('--multiproc', action='store_true',
                        help="Use multiprocessing: True/False")
    parser.add_argument('--debug', action='store_true',
                        help="Debug: True/False")

    args = parser.parse_args()
    start_idx = int(args.start_idx)
    end_idx = int(args.end_idx)

    # Build a list of file-names to be processed
    url_df_path = "../url_docx/url_table_structure_recognition_uuid_final.csv"
    url_df = pd.read_csv(url_df_path)
    uuids = url_df.loc[start_idx:end_idx, "uuid"].to_list()
    docx_names = [name + ".docx" for name in uuids]
    docx_path = "../data_raw/data_docx_recognition"
    output_path = "../output"

    if args.do == "run":
        create_docs(docx_path, docx_names, output_path,
                    args.multiproc, args.debug)
    elif args.do == "index":
        indices_to_be_done, idx_min = find_index(output_path, uuids, url_df)
        print("idx_found: ", idx_min)
        print("indices_to_be_done: ", indices_to_be_done)


if __name__ == "__main__":
    main()
    """
    Filter tables:
    step draw borders:
        - do not draw outside border if the table has only one cell or has
        other tables inside
    step compare images with fuchsia and aqua borders:
        - drop tables with fuchsia inside
    step detect_cells:
        - if a table spans across more than one page, use only a part of the
        table on the first page and drop the rest
    step draw vertical and horizontal lines:
        - at least two columns and at least two rows
    """
