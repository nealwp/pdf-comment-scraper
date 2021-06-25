import csv
import PyPDF2
import os
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename
from tkinter import messagebox

headers = ['date', 'note', 'provider', 'exhibit', 'page']


def bookmark_dict(reader_obj, bookmark_list):
    result = {}
    for item in bookmark_list:
        if isinstance(item, list):
            # recursive call
            result.update(bookmark_dict(reader_obj, item))
        else:
            try:
                result[reader_obj.getDestinationPageNumber(item)+1] = item.title
            except Exception:
                pass
    return result


def comment_list(reader_obj, bm_dict):

    cmt_list = []
    comment_delim = ";"
    n_pages = reader_obj.getNumPages()

    for i in range(n_pages):
        page0 = reader_obj.getPage(i)
        page_num = i + 1
        text = page0.extractText().split()
        try:
            for annotation in page0['/Annots']:
                annotation_dict = annotation.getObject()
                comment = annotation_dict.get('/Contents', None)
                if comment is not None:
                    comment = comment.split(comment_delim)
                    cmt_date = comment[0]
                    cmt_text = comment[1]
                    try:
                        provider = comment[3]
                    except IndexError:
                        provider = ""
                    if not bm_dict.get(page_num)[0] == "(":
                        citation = bm_dict.get(page_num)
                    else:
                        citation = text[0]
                    line = [cmt_date, cmt_text, provider, citation, page_num]
                    cmt_list.append(line)
        except Exception:
            pass
    return cmt_list


def write_csv(file_name, heads, rows):

    if os.path.exists(file_name):
        os.remove(file_name)

    with open(file_name, 'w', newline='') as cf:
        cw = csv.writer(cf, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
        cw.writerow(heads)
        for row in rows:
            cw.writerow(row)


def get_file_path():

    Tk().withdraw()
    file_path = askopenfilename(filetypes=[("PDF files", "*.pdf")])
    return file_path


def get_save_path():

    Tk().withdraw()
    save_path = asksaveasfilename(filetypes=[("CSV Files", "*.csv")], defaultextension=[("CSV Files", "*.csv")])
    return save_path


if __name__ == '__main__':

    src = get_file_path()

    if os.path.exists(src):

        reader = PyPDF2.PdfFileReader(src)
        bd = bookmark_dict(reader, reader.getOutlines())
        comments = comment_list(reader, bd)
        if len(comments) > 0:
            dest = get_save_path()
            if dest != '':
                write_csv(dest, headers, comments)
                messagebox.showinfo("Information", "File saved to " + dest)
            else:
                messagebox.showerror("Error", "File not saved!")
        else:
            messagebox.showwarning("Warning", "No comments found in " + src)
    else:
        messagebox.showerror("Error", "Source path does not exist!")
