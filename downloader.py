import os
import pandas as pd
import time

from selenium import webdriver
from selenium.webdriver.chrome.options import Options

from PyPDF2 import PdfFileMerger

# Update the following accordingly
list_topics = ['Interaction Design',
               'Machine Learning and Real-world Data'] # list of topics you want to download, topic names are specified by csv below
locked_solutions = [2018, 2019] # solutions only viewable by supervisors
flag_reverse_merge_pdf = True # merge papers in descending order of year
flag_merge_pdfs = True # merge questions into a single pdf
flag_excludeNoSols = True # exclude papers with no solutions accessible
download_wait_time = 15 # to adjust, or even reduce to 0, depending on your Internet speed

crsid = '' # raven username
raven_password = '' # raven password
download_folder = r"C:\Users\Admin\Downloads"
chromedriver_filepath = r'C:\Users\Admin\Downloads\chromedriver_win32\chromedriver.exe' # you will need to download chromedriver first: https://chromedriver.chromium.org/downloads
working_dir = r'C:\Users\Admin\Desktop\PYP' # where you want to download and merge the PDFs

csv = 'https://www.cl.cam.ac.uk/teaching/exams/pastpapers/index.csv'; # specifies questions and solutions available online

def site_login(list_pdf_url, download_folder):
   options = Options()
   options.add_experimental_option("prefs", {
     "download.default_directory": download_folder,
     "download.prompt_for_download": False,
     "download.directory_upgrade": True,
     "safebrowsing.enabled": True,
     "plugins.always_open_pdf_externally": True
   })
   # if chromedriver is not in your path, you’ll need to add it here
   driver = webdriver.Chrome(chromedriver_filepath, options=options)
   driver.get ('https://raven.cam.ac.uk/auth/login.html')
   driver.find_element_by_id('userid').send_keys(crsid)
   driver.find_element_by_id ('pwd').send_keys(raven_password)
   driver.find_element_by_name('submit').click()

   for url in list_pdf_url:
      driver.get(url)

def merge_pdf_each_folder (folder, output_dir=None, rename_pdfs=False, reverse=True):
   merger = PdfFileMerger()
   if rename_pdfs:
      for pdf in os.listdir(folder):
         year = pdf.split('p')[0]
         paper = 'p' + pdf.split('p')[1].split('q')[0].zfill(2)
         question = 'q' + pdf.split('q')[1].split('.')[0].zfill(2)
         os.rename(os.path.join(folder, pdf), os.path.join(folder, year+paper+question+'.pdf'))

   for pdf in sorted(os.listdir(folder), reverse=reverse):
      merger.append(os.path.join(folder,pdf))

   if output_dir==None:
      output_dir = os.path.dirname(os.path.dirname(folder))
      print("COMPLETED")
      print("merged from: " + folder)
      print("merged to: " + output_dir)
   if reverse:
      merger.write(os.path.join(output_dir, "_".join([os.path.basename(os.path.dirname(folder)),
                                                      os.path.basename(folder),
                                                      "rev_merged.pdf"])))
   else:
      merger.write(os.path.join(output_dir, "_".join([os.path.basename(os.path.dirname(folder)),
                                                  os.path.basename(folder),
                                                  "merged.pdf"])))

df = pd.read_csv(csv)
has_sol = pd.notnull(df['solutions'])
df = df[has_sol]
# assumes that number of years of locked_solutions is always 1 or 2
if (flag_excludeNoSols):
   df_sol = df[df['year']!=locked_solutions[0]]
   if len(locked_solutions) == 2:
      df_sol = df_sol[df_sol['year']!=locked_solutions[1]]
else:
   df_sol = df

folders = os.listdir(working_dir)

for topic in list_topics:
   df2 = df[df['topic']==topic]
   df3 = df_sol[df_sol['topic']==topic]
   site_login(['https://www.cl.cam.ac.uk/teaching/exams/pastpapers/' + x for x in df2['pdf'].tolist()], os.path.join(working_dir, topic, 'Questions'))
   site_login(df3['solutions'].tolist(), os.path.join(working_dir, topic, 'Solutions'))
   merge_pdf_each_folder(os.path.join(working_dir, topic, 'Questions'), rename_pdfs=True, reverse=flag_reverse_merge_pdf)
   try:
      merge_pdf_each_folder(os.path.join(working_dir, topic, 'Solutions'), reverse=flag_reverse_merge_pdf)
   except FileNotFoundError:
      print(topic, "was only started recently in %s, no solutions available", str(min(locked_solutions)))
   # to wait for the downloads to complete
   time.sleep(download_wait_time)
