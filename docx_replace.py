import zipfile
import tempfile
import os
import shutil
import random

def docx_replace(docx_filename, db_list, placeholder, output):
   """replace placeholder in word document with db_list elements"""
    
   tmp_dir = tempfile.mkdtemp()
   
   #open template docx by unzipping
   with open(docx_filename,'r+b') as f:
      zip = zipfile.ZipFile(f)
      #extract zip tree to our temp directory for later use
      zip.extractall(tmp_dir)
      filenames = zip.namelist()
      #extract xml from docx file
      xml_content = zip.read('word/document.xml')

   #replace placeholders within document's xml string
   for address in db_list:
      print(address)
      #one occurance of the given place holder per item 
      xml_content = xml_content.replace(placeholder,address,1)
   #create new file with xml contents within tmp directory
   with open(os.path.join(tmp_dir,'word/document.xml'), 'w') as myzip:
      myzip.write(xml_content)
      
   #create docx file by zipping contents in output file
   with zipfile.ZipFile(output, "w") as docx:
      for filename in filenames:
         docx.write(os.path.join(tmp_dir,filename), filename)
         
   #delete tmp directory
   shutil.rmtree(tmp_dir)

def RandomAddress():
   address = str(random.randint(0,999))
   address += " Fake St <w:br/>Gold Coast<w:br/>QLD<w:br/>4000"
   return address
   
if __name__ == "__main__":
   dblist = []
   for x in range(0,24):
       dblist.append(RandomAddress())
   placeholder = "HOLDER"
   docx_replace('template.docx', dblist, placeholder, 'mod_template.docx')

   import tempfile
import win32api
import win32print

filename = tempfile.mktemp (".txt")
open (filename, "w").write ("This is a test")
win32api.ShellExecute (
  0,
  "print",
  filename,
  #
  # If this is None, the default printer will
  # be used anyway.
  #
  '/d:"%s"' % win32print.GetDefaultPrinter (),
  ".",
  0
)

