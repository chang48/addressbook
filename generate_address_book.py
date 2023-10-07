#!/usr/bin/env python
# coding: utf-8

# Implement Google Sheet API for retrieving Address Book responses
# 
# Ref: 
# - Sheet API and sample: https://developers.google.com/sheets/api/quickstart/python
# - Drive API and sample: https://developers.google.com/drive/api/quickstart/python
# - Drive API manage downloads: https://developers.google.com/drive/api/guides/manage-downloads
# - Save downloaded image in bytes format to a file: https://stackoverflow.com/questions/18491416/pil-convert-bytearray-to-image

# In[19]:


from __future__ import print_function
import os.path
import numpy as np
import pandas as pd
import io
import re
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload
from math import ceil
from tqdm import notebook
import PIL.Image as Image


# In[26]:


# Need to pass a list for each service rather than just the string
# See: https://stackoverflow.com/questions/16633297/google-drive-oauth-2-flow-giving-invalid-scope-error
SCOPES = {
    'Sheet': ['https://www.googleapis.com/auth/spreadsheets.readonly'],
    'Drive': ['https://www.googleapis.com/auth/drive.readonly']
}

# -----------------------------------------------------------------------------------
# Service:
#   Drive - Google Drive
#   Sheet - Google Sheet
def retrieve_credential(service):
    token = 'token_{}.json'.format(service)
    
    # Initialize credential
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists(token):
        creds = Credentials.from_authorized_user_file(token, SCOPES[service])
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES[service])
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(token, 'w') as token:
            token.write(creds.to_json())
    
    return creds

# -----------------------------------------------------------------------------------
def download_file(creds, real_file_id):
    """Downloads a file
    Args:
        creds: credential object
        real_file_id: ID of the file to download
    Returns : IO object with location.
    """
    try:
        # create drive api client
        service = build('drive', 'v3', credentials=creds)

        file_id = real_file_id

        # pylint: disable=maybe-no-member
        request = service.files().get_media(fileId=file_id)
        file = io.BytesIO()
        downloader = MediaIoBaseDownload(file, request)
        done = False
        while done is False:
            status, done = downloader.next_chunk()
            # print(F'Download {int(status.progress() * 100)}%.')

    except HttpError as error:
        print(F'An error occurred: {error}')
        file = None

    return file.getvalue()


# -----------------------------------------------------------------------------------
# See: https://stackoverflow.com/questions/16259923/how-can-i-escape-latex-special-characters-inside-django-templates
def tex_escape(text):
    """
        :param text: a plain text message
        :return: the message escaped to appear correctly in LaTeX
    """
    conv = {
        '&': r'\&',
        '%': r'\%',
        '$': r'\$',
        '#': r'\#',
        '_': r'\_',
        '{': r'\{',
        '}': r'\}',
        '~': r'\textasciitilde{}',
        '^': r'\^{}',
        '\\': r'\textbackslash{}',
        '<': r'\textless{}',
        '>': r'\textgreater{}',
    }
    regex = re.compile('|'.join(re.escape(str(key)) for key in sorted(conv.keys(), key = lambda item: - len(item))))
    return regex.sub(lambda match: conv[match.group()], text)


# In[3]:


creds = retrieve_credential('Sheet')


# In[4]:


# The ID and range of the address book response spreadsheet.
SPREADSHEET_ID = '1hF_NiS2wqdRsRw5e8Md-kHjU0MpR6iIDu0eJLXnrPco'
RANGE_NAME = 'Splitted Responses'


# In[5]:


# Using Google Sheet API to fetch the splitted response sheet
service = build('sheets', 'v4', credentials=creds)
sheet = service.spreadsheets()
result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                            range=RANGE_NAME).execute()


# In[34]:


# Turn the results into a dataframe
columns = ['date', 'street number', 'city', 'state', 'zip', 'photo', 'name_en', 'name_zh', 'phone', 'email']
data = pd.DataFrame(result['values'][1:], columns=columns)

# Fill missing data (null) with empty string and drop the date
data = data.fillna('')
data['address'] = data['street number'] + " " + data['city'] + ", " + data['state'] + " " + data['zip']

# Retain a copy of the dataframe with only necessary columns
df = data[['address', 'street number', 'city', 'state', 'zip', 'name_en', 'name_zh', 'phone', 'email', 'photo']].copy()


# In[35]:


# Process escaping LaTeX special characters
df['phone'] = df['phone'].apply(tex_escape)
df['email'] = df['email'].apply(tex_escape)


# In[36]:


df.head(5)


# In[37]:


# Group the contact info dataframe by household address then fetch data for the first household member
df_tmp = df.groupby('address').first()

# Grab the last name
df_tmp['last_name'] = df_tmp['name_en'].str.split(',').str[0]

# Sort household by the last name in alphabetical order. The first reset_index() turns the groupby index
# into a column. The second reset_index() adds the index in the desired order
order = df_tmp.sort_values(by='last_name').reset_index().reset_index()[['address', 'last_name', 'index']]

# Join the sorted index back to the original dataframe
df = df.join(order.set_index('address'), on='address')


# In[38]:


# Now we group the dataframe by the index column. By defauly, groupby will sort the groups
# according to the order determined by the index column. This is exactly what we want
groups = df.groupby('index')
keys = groups.groups.keys()


# In[39]:


photo_header = """
    \\begin{{figure}}[H]
       \\includegraphics[height=0.13\\textheight, width=\\textwidth, keepaspectratio]{{{0}}}
    \\end{{figure}}"""


# In[40]:


creds = retrieve_credential('Drive')


# In[41]:


file_list = []
for key in notebook.tqdm(keys, position=0, leave=True):
    group = groups.get_group(key)
    fname = 'entries/{0:04d}-{1}.tex'.format(key, group.iloc[0]['last_name'])
    photoname = 'photos/{0:04d}-{1}.jpg'.format(key, group.iloc[0]['last_name'])

    url = group['photo'].drop_duplicates().values[0]
    if len(url) > 1:
        file_id = url.split('=')[1]
        photo = download_file(creds, file_id)
        photo = Image.open(io.BytesIO(photo))
        photo.save(photoname)
    
    file_list.append(fname)
    f = open(fname, 'w+')
    f.write('\\begin{minipage}[t][0.32\\textheight][t]{0.5\\textwidth}\n')
    if len(url) > 1:
        f.write(photo_header.format(photoname))
    else:
        f.write(photo_header.format('photos/DCCC_logo.jpg'))
    f.write('    \\vspace{-0.25cm}\n\n')
    f.write('    \\begin{tabular}{@{}l @{\hspace{3pt}} l @{}}\n')
    for _, row in group.iterrows():
        f.write('        {0:20s} & {1:5s} \\\\\n'.format(row['name_en'], row['name_zh']))
    f.write('    \\end{tabular}\n\n')
    
    f.write('    \\smallskip\n')
    tmp = group[['street number', 'city', 'state', 'zip']].drop_duplicates()
    f.write('    {0}\\\\\n'.format(tmp['street number'].values[0]))
    f.write('    {0}, {1} {2}\n\n'.format(tmp['city'].values[0],  tmp['state'].values[0], tmp['zip'].values[0]))
    f.write('')

    f.write('    \\smallskip\n')
    f.write('    \\begin{tabular}{@{}l @{\hspace{3pt}} l @{}}\n')
    for _, row in group.iterrows():
        if len(row['phone']) >= 1:
            f.write('        \\faPhone    & {{\\tt {0:12s}}} \\\\\n'.format(row['phone']))    
    for _, row in group.iterrows():
        if len(row['email']) >= 1:
            email = row['email'].replace('_', '\_')
            f.write('        \\faEnvelopeO & {{\\tt {0:12s}}} \\\\\n'.format(email))
    f.write('    \\end{tabular}\n')    
    f.write('\\end{minipage}\n')
    f.close()


# In[13]:


# For testing purpose
# file_list = file_list*3


# In[42]:


num_entry = len(file_list)
num_entry_per_page = 6
num_page = ceil(num_entry/num_entry_per_page)

if num_entry%num_entry_per_page != 0:
    diff = num_page*num_entry_per_page - num_entry


# In[43]:


addressbook = np.array(file_list + ['']*diff)
addressbook = addressbook.reshape(num_page, num_entry_per_page)


# In[44]:


print('total number of contact entries:', num_entry)
print('number of contacts per page:', num_entry_per_page)
print('total number of pages:', num_page)


# In[45]:


header = """
\\documentclass[12pt, letterpaper]{report}
% \\usepackage{xeCJK}
\\usepackage{geometry}
\\usepackage{graphicx}
\\usepackage{float}
\\usepackage{anyfontsize}
\\usepackage{everypage}
\\usepackage{fontawesome}
\\usepackage{ctex}  % ctex can process both simplified and traditional Chinese fonts
\\geometry{left=1.5cm, top=0.2cm, right=0.5cm, bottom=0.9cm, footskip=.3cm}

% Set up the Chinese font, see: https://jdhao.github.io/2018/03/29/latex-chinese/
% \\setCJKmainfont{Apple LiSung}

\\begin{document}
\\fontsize{12pt}{12pt}\\selectfont

"""


# In[46]:


f = open('addressbook.tex', 'w+')
f.write(header)

for idx, page in enumerate(addressbook):
    f.write('\\begin{minipage}{\\textwidth}\n')
    f.write('\\begin{tabular}{cc}\n')

    template = '    \\input{{{0:25s}}}  &  \\input{{{1:25s}}}\\\\\n'
    for row in range(3):
        f.write(template.format(page[row*2][:-4], page[row*2+1][:-4]))
    f.write('\\end{tabular}\n')
    f.write('\\end{minipage}\n')
    if (idx+1) != num_page:
        f.write('\n\\newpage\n')

f.write('\n\\end{document}')
f.close()


# In[ ]:




