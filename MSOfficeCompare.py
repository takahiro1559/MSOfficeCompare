import win32com.client
from pathlib import Path
import pathlib
import cv2
import numpy as np
import pdf2image
import os
from matplotlib import pyplot as plt

wd = win32com.client.Dispatch('Word.Application')
xl = win32com.client.Dispatch('Excel.Application')
wd.Visible = -1

doc = wd.Documents.Open(r'C:\work\python\PdfCompare\source_img\foo.docx')
xls = xl.Workbooks.Open(r'C:\work\python\PdfCompare\source_img\foo.xlsx')

wdFormatPDF = 17
doc.SaveAs2(FileName=r'C:\work\python\PdfCompare\source_img\foo2.pdf', FileFormat=wdFormatPDF)
xls.WorkSheets('Sheet1').Select()
xls.ActiveSheet.ExportAsFixedFormat(0,r'C:\work\python\PdfCompare\source_img\foo3.pdf')

doc = wd.Documents.Open(r'C:\work\python\PdfCompare\target_img\foo2.docx')
xls = xl.Workbooks.Open(r'C:\work\python\PdfCompare\target_img\foo2.xlsx')

wdFormatPDF = 17
doc.SaveAs2(FileName=r'C:\work\python\PdfCompare\target_img\foo2.pdf', FileFormat=wdFormatPDF)
xls.WorkSheets('Sheet1').Select()
xls.ActiveSheet.ExportAsFixedFormat(0,r'C:\work\python\PdfCompare\target_img\foo3.pdf')

xls.Close()
wd.Quit()
xl.Quit()

poppler_dir = Path(__file__).parent.absolute() / "poppler/bin"
os.environ["PATH"] += os.pathsep + str(poppler_dir)

source_dir = pathlib.Path('source_img')
source_files = source_dir.glob('*.*')
for file in source_files:
    if '.pdf' in file.name:
        images = pdf2image.convert_from_path(file, grayscale=True, size=1000)
        fname = os.path.splitext(file)[0]
        for index, image in enumerate(images):
            image.save(fname +"-"+ str(index+1) + '.png')

target_dir = pathlib.Path('target_img')
target_files = target_dir.glob('*.*')
for file in target_files:
    if '.pdf' in file.name:
        images = pdf2image.convert_from_path(file, grayscale=True, size=1000)
        fname = os.path.splitext(file)[0]
        for index, image in enumerate(images):
            image.save(fname +"-"+ str(index+1) + '.png')

source_dir = pathlib.Path('source_img')
source_files = source_dir.glob('*.*')
target_dir = pathlib.Path('target_img')

for source_file in source_files:
    print(source_file.name)
    if '.png' in source_file.name:
        source_img = cv2.imread(str(source_file))
        target_file = target_dir / source_file.name
        target_img = cv2.imread(str(target_file))
        if target_img is None:
            fs.write(target_file + '...skipped.\n')
            continue
        source_gray_img = cv2.cvtColor(source_img, cv2.COLOR_BGR2GRAY)
        target_gray_img = cv2.cvtColor(target_img, cv2.COLOR_BGR2GRAY)

        max_hight = source_img.shape[0]
        max_width = source_img.shape[1]

        result_window = np.zeros((max_hight, max_width), dtype=source_img.dtype)
        for start_y in range(0, max_hight-100, 50):
            for start_x in range(0, max_width-100, 50):
                window = source_gray_img[start_y:start_y+100, start_x:start_x+100]
                match = cv2.matchTemplate(target_gray_img, window, cv2.TM_CCOEFF_NORMED)
                _, _, _, max_loc = cv2.minMaxLoc(match)
                matched_window = target_gray_img[max_loc[1]:max_loc[1]+100, max_loc[0]:max_loc[0]+100]
                result = cv2.absdiff(window, matched_window)
                result_window[start_y:start_y+100, start_x:start_x+100] = result

        _, result_window_bin = cv2.threshold(result_window, 127, 255, cv2.THRESH_BINARY)
        contours, _ = cv2.findContours(result_window_bin, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        imgC = source_img.copy()
        for contour in contours:
            min = np.nanmin(contour, 0)
            max = np.nanmax(contour, 0)
            loc1 = (min[0][0], min[0][1])
            loc2 = (max[0][0], max[0][1])
            cv2.rectangle(imgC, loc1, loc2, 255, 2)

        plt.subplot(1, 3, 1), plt.imshow(cv2.cvtColor(source_img, cv2.COLOR_BGR2RGB)), plt.title('A'), plt.xticks([]), plt.yticks([])
        plt.subplot(1, 3, 2), plt.imshow(cv2.cvtColor(target_img, cv2.COLOR_BGR2RGB)), plt.title('B'), plt.xticks([]), plt.yticks([])
        plt.subplot(1, 3, 3), plt.imshow(cv2.cvtColor(imgC, cv2.COLOR_BGR2RGB)), plt.title('Answer'), plt.xticks([]), plt.yticks([])
        plt.show()

