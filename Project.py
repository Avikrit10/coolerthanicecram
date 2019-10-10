# -*- coding: utf-8 -*-
"""
Created on Thu Oct 10 13:29:01 2019

@author: garv2
"""
import shutil
import os
import zipfile as zf
import random
files = zf.ZipFile("dataset-resized.zip",'r')
files.extractall()
files.close()

directory=r"C:\Users\garv2\Desktop\Internship\python\dataset-resized"
#print((os.listdir(directory)))
def split_indices(folder,seed1,seed2):    
    n = len(os.listdir(folder))
    full_set = list(range(1,n+1))
    print(n)
    random.seed(seed1)
    train = random.sample(list(range(1,n+1)),int(.5*n))
    remain = list(set(full_set)-set(train))
    random.seed(seed2)
    valid = random.sample(remain,int(.5*len(remain)))
    test = list(set(remain)-set(valid))
    
    return(train,valid,test)


def get_names(waste_type,indices):
    file_names = [waste_type+str(i)+".jpg" for i in indices]
    return(file_names)    

def move_files(source_files,destination_folder):
    for file in source_files:
        shutil.move(file,destination_folder)
subsets = ['train','valid']
waste_types = ['cardboard','glass','metal','paper','plastic','trash']


for subset in subsets:
    for waste_type in waste_types:
        folder = os.path.join('data',subset,waste_type)
        if not os.path.exists(folder):
            os.makedirs(folder)
            
if not os.path.exists(os.path.join('data','test')):
    os.makedirs(os.path.join('data','test'))
            
for waste_type in waste_types:
    source_folder = os.path.join('dataset-resized',waste_type)
    train_ind, valid_ind, test_ind = split_indices(source_folder,1,1)
    train_names = get_names(waste_type,train_ind)
    train_source_files = [os.path.join(source_folder,name) for name in train_names]
    train_dest = "data/train/"+waste_type
    move_files(train_source_files,train_dest)
    valid_names = get_names(waste_type,valid_ind)
    valid_source_files = [os.path.join(source_folder,name) for name in valid_names]
    valid_dest = "data/valid/"+waste_type
    move_files(valid_source_files,valid_dest)
    test_names = get_names(waste_type,test_ind)
    test_source_files = [os.path.join(source_folder,name) for name in test_names]
    move_files(test_source_files,"data/test")
path =r"C:\Users\garv2\Desktop\Internship\python\data"
print(path)