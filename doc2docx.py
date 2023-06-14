
import os
import win32com.client as win32
import argparse

def find_files(dir_path, file_ext=["docx"]):
    file_list = []
    for root, dirs, files in os.walk(dir_path):
        for filename in files:
            if any(filename.endswith(ext) for ext in file_ext):
                if not os.path.basename(filename).startswith('~$'):
                    file_list.append(os.path.join(root, filename))
    return file_list

def parse():
    parser = argparse.ArgumentParser()
    parser.add_argument('folder', default='',help='')
    args = parser.parse_args()
    return args

def run():
    args = parse()
    docs = find_files(args.folder)
    for i,doc in enumerate(docs):
        try:
            file_name = os.path.basename(doc)
            if not file_name.startswith('~$'):
                docx = doc.replace(".doc",".docx")
                if not os.path.isfile(docx):
                    print(f"[{i+1}/{len(docs)}]:{doc}")
                    application = win32.DispatchEx('Word.Application')
                    word = application.Documents.Open(doc)
                    word.SaveAs2(docx, FileFormat=16)
                    application.Quit()
        except Exception as e:
            print(e)
    
if __name__=='__main__':
    run()
