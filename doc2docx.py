
import os
import win32com.client as win32
from tqdm import tqdm
from utils.utils import *
import pythoncom
import concurrent.futures
import pebble
import multiprocessing

def find_files(dir_path, file_ext=[".doc"]):
    file_list = []
    for root, dirs, files in os.walk(dir_path):
        for filename in files:
            if any(filename.endswith(ext) for ext in file_ext) and not filename.startswith('~$'):
                    file_list.append(os.path.join(root, filename))
    return file_list

def parse():
    parser = argparse.ArgumentParser()
    parser.add_argument('folder', default='',help='')
    parser.add_argument('--mp', type=bool, default=True,
                        help='use multiprocess or not ')
    args = parser.parse_args()
    return args
    
def worker(doc):
    try:
        file_name = os.path.basename(doc)
        if not file_name.startswith('~$'):
            docx = doc.replace(".doc",".docx")
            if not os.path.isfile(docx):
                pythoncom.CoInitialize()
                word = win32.DispatchEx('Word.Application')
                doc_object = word.Documents.Open(doc)
                doc_object.SaveAs2(docx, FileFormat=16)
                word.Quit()
                pythoncom.CoUninitialize()
    except Exception as e:
        #pass
        doc=None
        print(f"{str(e)}")
    return doc
    
def run():
    args = parse()
    docs = find_files(args.folder)
    if args.mp:
        cpu_count = multiprocessing.cpu_count()
        print(f"amount of cpu:{cpu_count}")
        processeser = int(cpu_count*3/4)
        print(f"use {processeser} cpu")
       
        pool = pebble.ProcessPool(max_workers=processeser)
        with pool:
            for doc in tqdm(docs):
                future = pool.schedule(worker, args=(doc,), timeout=120,)
                try:
                    future.result()
                except Exception as e:
                    print(e)
    else:
        for _file in tqdm(docs):
            worker(_file)
    
if __name__=='__main__':
    run()
