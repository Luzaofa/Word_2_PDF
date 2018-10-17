import win32com.client


class Word_2_PDF(object):

    def __init__(self, filepath, Debug=False):
        """
        :param filepath:
        :param Debug: 控制过程是否可视化
        """
        self.wordApp = win32com.client.Dispatch('word.Application')
        self.wordApp.Visible = Debug
        self.myDoc = self.wordApp.Documents.Open(filepath)

    def export_pdf(self, output_file_path):
        """
        将Word文档转化为PDF文件
        :param output_file_path:
        :return:
        """
        self.myDoc.ExportAsFixedFormat(output_file_path, 17, Item=7, CreateBookmarks=0)

if __name__ == '__main__':

    rootpath = 'C:\\word_2_PDF\\'       # 文件夹根目录

    Word_2_PDF = Word_2_PDF(rootpath + 'Docfile.docx', True)

    Word_2_PDF.export_pdf(rootpath + 'PDFfile.pdf')
    
