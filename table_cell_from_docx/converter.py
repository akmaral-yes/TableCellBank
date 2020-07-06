import os
from PyPDF2 import PdfFileWriter, PdfFileReader
from wand.image import Image as WANDImage
from wand.color import Color
from win32com import client
import zipfile
import shutil
from utils.file_utils import append_to_file


def unpack_zip(file_name, zip_file_path, unzipped_path, output_path):
    """
    Unpack the .zip if possible
    """
    directory_to_extract_to = os.path.join(unzipped_path, file_name[:-5])
    try:
        with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
            zip_ref.extractall(directory_to_extract_to)
        return True
    except BaseException:
        append_to_file(output_path, 'unpack_failed.csv', file_name)
        return False


def save_docx(folder_name, unzipped_path, docx_path):
    """
    Convert the unzipped folder to .docx in docx_path
    """
    docx_file_path = os.path.join(docx_path, folder_name + ".docx")
    if os.path.exists(docx_file_path):
        os.remove(docx_file_path)
    folders_to_be_archived_path = os.path.join(unzipped_path, folder_name)
    zipped_folder_path = os.path.join(docx_path, folder_name)
    shutil.make_archive(zipped_folder_path, 'zip', folders_to_be_archived_path)
    os.rename(os.path.join(docx_path, folder_name + ".zip"), docx_file_path)


def docx_to_pdf(file_name, docx_path, pdf_path, output_path):
    """
    Convert .docx to .pdf
    """
    wdFormatPDF = 17  # PDF format
    wdExportDocumentContent = 0
    word = client.DispatchEx("Word.Application")

    in_file = os.path.join(docx_path, file_name)
    out_file = os.path.join(pdf_path, file_name[:-4] + "pdf")
    if os.path.exists(out_file):
        os.remove(out_file)

    # Indicator of whether the pdf was successfully saved
    done = False

    # Create a new Word Object
    word = client.DispatchEx("Word.Application")

    # Open the word file for read-only
    try:
        worddoc = word.Documents.Open(in_file, ReadOnly=1)
    except BaseException:
        # Save file names that couldn't open
        append_to_file(output_path, 'docx_failed.csv', file_name)

    # Save as .pdf
    try:
        worddoc.ExportAsFixedFormat(
            OutputFileName=out_file, Item=wdExportDocumentContent,
            ExportFormat=wdFormatPDF
        )
        done = True
    except BaseException:
        # Save file names that couldn't convert to .pdf
        append_to_file(output_path, 'pdf_failed.csv', file_name)

    finally:
        # Close file without saving changes
        try:
            worddoc.Close(SaveChanges=0)
            word.Quit()
        except BaseException:
            pass
        return done


def pdf_to_image(input_dir, file_name, output_dir, output_path, temp_path):
    """
    Convert given pdf to images

    Args:
        input_dir: a path to a folder with the pdf
        file_name: the pdf file name
        output_dir: a path to save the images from pdf
    """
    # Create temp_folder to save there separate pages and then delete
    name = file_name.split(".")[0]
    temp_folder = os.path.join(temp_path, name)
    os.makedirs(temp_folder, exist_ok=True)

    # Open .pdf
    file_path = os.path.join(input_dir, file_name)
    try:
        f = open(file_path, "rb")
        inputpdf = PdfFileReader(f)
    except BaseException:
        append_to_file(output_path, 'pdf_broken.csv', file_name)

    number_of_pages = inputpdf.numPages
    RESOLUTION_COMPRESSION_FACTOR = 300

    # Iterate over all pages
    for i in range(number_of_pages):

        # Split into single page pdf
        output = PdfFileWriter()
        output.addPage(inputpdf.getPage(i))
        try:
            with open(
                os.path.join(temp_folder, "document-page%i.pdf") % i, "wb"
            ) as outputStream:
                output.write(outputStream)
        except BaseException:
            f.close()
            append_to_file(output_path, 'pdf_stream.csv', file_name)
            return False

        # Open the single page pdf as image, compress and save
        with WANDImage(
                filename=os.path.join(temp_folder, "document-page%i.pdf") % i,
                resolution=300) as img:
            img.background_color = Color("white")
            img.alpha_channel = 'remove'
            img.compression_quality = RESOLUTION_COMPRESSION_FACTOR
            img.save(filename=os.path.join(
                output_dir, file_name[:-4] + "_%i.png") % i
            )

        # Remove the single page pdf
        os.remove(os.path.join(temp_folder, "document-page%i.pdf") % i)

    f.close()
    os.rmdir(temp_folder)
