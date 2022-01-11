import pandas as pd
import os
import shutil
from tkinter.filedialog import askopenfilename

file_path = askopenfilename()
new_file_path = './' + file_path.split(sep='/')[-1]
file_ext = file_path.split(sep='/')[-1].split(sep='.')[-1]
data_type = file_path.split(sep='/')[-1][:2]

if os.path.exists(new_file_path):
    os.remove(new_file_path)
if os.path.exists('tmp.txt'):
    os.remove('tmp.txt')

# copy file and change extension
shutil.copy2(file_path, new_file_path)

base = os.path.splitext(new_file_path)[0]
if not (os.path.exists(base + '.txt')) and file_ext != 'txt':
    os.rename(new_file_path, base + '.txt')

# delete headers
tmp = open('./tmp.txt', 'w+')
with open(file_path, 'r+') as f:
    d = f.readlines()
    for i in d:
        if i[0] != 'H':
            tmp.write(i)
tmp.close()


# In[56]:


def del_space_and_lcol(d):
    for i in range(len(d)):
        d[i] = ' '.join(d[i].split())
        last_dot_index = d[i].rfind('.')
        last_space_index = d[i].rfind(' ')
        if d[i][last_dot_index + 2] != ' ':
            # ---check for strange mistake
            if last_dot_index < last_space_index:
                d[i] = d[i][:last_space_index] + '0' + d[i][last_space_index + 1:]

            d[i] = d[i][:last_dot_index + 2] + ' ' + d[i][last_dot_index + 3:]

    return d


def sps_mistake(d):
    for i in range(len(d)):
        s = d[i][35:40]
        for j in range(len(s)):
            if s[j] == ' ':
                d[i] = d[i][:35] + s[:j] + '0' + s[j + 1:] + d[i][41:]
    return d


def old_type_mistake(d):
    for i in range(len(d)):
        if d[i][24] != ' ':
            d[i] = d[i][:24] + ' ' + d[i][24:]

        if d[i][1] != ' ':
            d[i] = d[i][0] + ' ' + d[i][1:]

    return d


def kt_mistake(data_list):
    for i in range(len(data_list)):
        last_space_index = data_list[i].rfind(' ')
        last_dot_index = data_list[i].rfind('.')
        if last_dot_index < last_space_index:
            data_list[i] = data_list[i][:last_space_index] + '0' + data_list[i][last_space_index + 1:]

    return data_list


# do list from .txt
with open('./tmp.txt', 'r+') as tmp:
    d = tmp.readlines()

# find specific file-types mistakes
if data_type == 'R ' or data_type == 'S ':
    d = old_type_mistake(d)
elif file_ext == 'sps':
    d = sps_mistake(d)
# if data_type == 'KT':
# d = kt_mistake(d)


# deleting ' ' and add last col
d = del_space_and_lcol(d)

# In[59]:


# create txt
with open('./tmp.txt', 'r+') as tmp:
    tmp.seek(0)
    for i in range(len(d)):
        # change . to ,
        # d[i] = d[i].replace('.', ',')

        tmp.write(d[i])
        tmp.write('\n')

    tmp.truncate()

# put spreadsheet to dataFrame
data = pd.read_csv('./tmp.txt', sep=' ', header=None, dtype=str)
cols_num = data.shape[1]

if cols_num == 9:
    data.columns = 'Type', 'Profile', 'Point', 'F1', 'F2', 'X', 'Y', 'H', 'F3'
elif cols_num == 10:
    data.columns = 'Type', 'Profile', 'Point', 'F1', 'F2', 'F3', 'X', 'Y', 'H', 'F4'
elif cols_num == 11:
    data.columns = 'Type', 'Profile', 'Point', 'F1', 'F2', 'F3', 'F4', 'X', 'Y', 'H', 'F5'
else:
    data.rename(columns={0: 'Type', 1: 'Profile'}, inplace=True)

# save to excel
data.to_excel(os.path.splitext(file_path)[0] + '.xlsx', index=None)

# delete tmps
os.remove('tmp.txt')
os.remove(base + '.txt')
