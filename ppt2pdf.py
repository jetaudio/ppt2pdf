import sys
import os
from glob import glob
from shutil import rmtree
from os.path import splitext, isfile, isdir, basename

try:
        from comtypes import client
except:
        print('Cài comtypes sử dụng: pip install --upgrade comtypes')
        input('Nhấn Enter để thoát...')
        sys.exit(-1)

try:
        from fpdf import FPDF
except:
        print('Cài fpdf sử dụng: pip install --upgrade fpdf')
        input('Nhấn Enter để thoát...')
        sys.exit(-1)
        
def ppt2pdf(f):
        if not os.path.exists(f):
                print('Lỗi! Không có file ' + f)
                input('Nhấn Enter để thoát...')
                sys.exit(-1)

        powerpoint = client.CreateObject('Powerpoint.Application')
        powerpoint.Presentations.Open(f)
        f = splitext(f)[0].rstrip(' ') + splitext(f)[1]
        powerpoint.ActivePresentation.Export(splitext(f)[0], 'JPG')
        powerpoint.ActivePresentation.Close()
        powerpoint.Quit()
        pdf = FPDF('L', 'pt', [1440, 1920])
        pdf.set_margins(0, 0, 0)        
        for image in sorted(glob(splitext(f)[0] + "\\*.JPG"), key = lambda f: int(''.join(filter(str.isdigit, f)))):
                pdf.add_page()
                pdf.image(image,0,0,1280,960)

        pdf.output(splitext(f)[0] + '.PDF', 'F')
        pdf.close()
        rmtree(splitext(f)[0])
        
if __name__ == '__main__':
        if len(sys.argv) != 2:
        	  print('Yêu cầu máy có cài sẵn Office 2007 trở lên để sử dụng.')
        	  print('Kéo và thả file hoặc folder vào biểu tượng của chương trình để convert')
        	  input('Nhấn Enter để thoát...')
        	  sys.exit(-1)
        if isfile(sys.argv[1]):
                path = sys.argv[1]
                print('Đang chuyển đổi: ' + path)
                ppt2pdf(path)
        elif isdir(sys.argv[1]):
                path = sys.argv[1]
                for file in glob(path + '\\**\\*.ppt', recursive = True):
                        print('Đang chuyển đổi: ' + file)
                        ppt2pdf(file)
                        
                for file in glob(path + '\\**\\*.pptx', recursive = True):
                        print('Đang chuyển đổi: ' + file)
                        ppt2pdf(file)

                        
        input("Đã chuyển đổi xong. Nhấn Enter để thoát...")
