from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LAParams, LTFigure, LTTextBox
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import PDFPageInterpreter, PDFResourceManager
from pdfminer.pdfpage import PDFPage, PDFTextExtractionNotAllowed
from pdfminer.pdfparser import PDFParser
from pathlib import Path
import xlsxwriter

encoding = 'ISO-8859-1'


def get_data_from_pdfs(number):
    path_pdfs = "Pdf_/" + str(number)

    pdf_text_paragraphs_list = []
    pdf_figures_stack = [""]
    paper_count = 0
    row = 0

    workbook = xlsxwriter.Workbook('dataset_' + str(number) + '.xlsx')

    worksheet = workbook.add_worksheet()

    for path in Path(path_pdfs).glob("*.pdf"):

        with path.open("rb") as f:
            parser = PDFParser(f)
            doc = PDFDocument(parser)
            page = list(PDFPage.create_pages(doc))[0]
            rsrcmgr = PDFResourceManager()
            device = PDFPageAggregator(rsrcmgr, laparams=LAParams(line_margin=1.50, word_margin=0.1))
            interpreter = PDFPageInterpreter(rsrcmgr, device)
            interpreter.process_page(page)
            layout = device.get_result()

            for obj in layout:

                print("PAPER: ")
                if isinstance(obj, LTTextBox):
                    # get each paragraph and append to a list  with lowercase
                    pdf_text_paragraphs_list.append(obj.get_text().lower())

                    print(obj.get_text().lower())
                    # text += obj.get_text()
                    # text += "\n"

                elif isinstance(obj, LTFigure):
                    pdf_figures_stack += list(obj)

            word_exist_abstract = "a b s t r a c t"

            # check the paragraph if it is abstract
            for paragraph in pdf_text_paragraphs_list:

                if word_exist_abstract in paragraph:
                    print("Found!")
                    abstract_paragraph = paragraph.replace(word_exist_abstract, "")
                    abstract_paragraph = abstract_paragraph.replace("\n", "")

                    # abstract_index = pdf_text_paragraphs_list.index(word)
                    # pdf_text_paragraphs_list[abstract_index] = word.replace(word_exist_abstract,"")
                    # word.replace(word_exist_abstract,"")
                    paper_count += 1
                # else:
                # print("Not found!")
            # print(pdf_text_paragraphs_list[abstract_index])
            ##abstract_data = pdf_text_paragraphs_list[abstract_index]
            # print("dic info :")
            # print(doc.info[0])
            abstract_data = abstract_paragraph

            dict_doc_info = doc.info[0]
            # retrive title of the paper
            title_data = dict_doc_info["Title"].decode(encoding)
            # print(dict_doc_info["Title"].decode(encoding))

            # retrive keywords from the paper
            keywords_data = dict_doc_info["Keywords"].decode(encoding)
            # print(dict_doc_info["Keywords"].decode(encoding))

            print("paper : " + str(paper_count))
            print("Tittle: " + title_data)
            print("keywords: " + keywords_data)
            print("ABSTRACT: " + abstract_data)
            # listq = [title_data, keywords_data, abstract_data]
            # print(listq)
            # print(pdf_text_paragraphs_list)
            pdf_text_paragraphs_list.clear()

            # save on the excel file dataset_papers

            titles_excel = ["Title", "Keywords", "Abstract", "Year", "PDF Directory"]
            content = [title_data, keywords_data, abstract_data, path_pdfs, str(path)]

            column = 0

            if row == 0:
                for item in titles_excel:
                    # write operation perform
                    worksheet.write(row, column, item)
                    column += 1
            else:
                # iterating through content list
                for item in content:
                    worksheet.write(row, column, item)
                    # incrementing the value of row by one
                    # with each iterations.
                    column += 1
            row += 1

    workbook.close()

def begin_pdf_path_extract():
    for number in range(10):
        get_data_from_pdfs(number)


begin_pdf_path_extract().init