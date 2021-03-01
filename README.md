# IsraeliLawParser
Parser &amp; Tagger for Israeli laws documents

Goal: Convert Israeli law documents from DOC format (random strings) into tagged XML files with proper heirarchy 

Input: PDF_DOC (input folder) (You can download the input folder from here https://drive.google.com/drive/folders/1bvjtlnGQZxYbSWumEKtDD4lKeqKImeBf)
Output: Israeli laws tagged as XML files (output files also are uploaded to our git)

Requirments: 
  1. Install LibreOffice in 'C:\Program Files\' OR change the path in project2.py to wherever you install it
  2. Install Python3 libraries (installion with PIP would be much easier):<br/>
    1. docx2python FROM https://pypi.org/project/docx2python/ <br/>
    2. fuzzywuzzy FROM https://pypi.org/project/fuzzywuzzy/ <br/>
    3. tqdm FROM https://pypi.org/project/tqdm/ <br/>
  3. Put all files in one folder that should contain: <br/>
    1. PDF_DOC (input folder) (You can download the input folder from here https://drive.google.com/drive/folders/1bvjtlnGQZxYbSWumEKtDD4lKeqKImeBf) <br/>
    2. output (output folder) <br/>
    3. project2.py (Python3 script) <br/>

Running the script: <br/>
  1. MAKE SURE YOUR PYTHON INTERPETER IS: Python 3.6 (DH3) <br/>
  2. Run project2.py and wait to finish <br/>
