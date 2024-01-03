# -*- coding: utf-8 -*-
"""
Created on Wed Jan  3 10:31:34 2024

@author: hwxlenovo
"""

import sys
import os
import mhrc.automation
import pandas as pd
# Import other utilities to perform cool stuff
from win32com.client.gencache import EnsureDispatch as Dispatch
from mhrc.automation.utilities.word import Word
from mhrc.automation.utilities.file import File
import win32com.client
import shutil
import random
import time
import csv
import pickle
import numpy as np
import matplotlib.pyplot as plt
from scipy import stats
import math
from matplotlib import cm


# PSCAS automation library
pscad_version = 'PSCAD 4.6.2 (x64)'
fortran_version = 'GFortran 4.6.2'
fortran_ext = '.gf46'
project_name = '' 
outputfile_name = '' 
working_dir = os.getcwd() + "\\"
require_dir = os.getcwd() + "\\"



def save_variable(v,filename):
    f=open(filename,'wb')
    pickle.dump(v,f)
    f.close()
    return filename

def load_variable(filename):
    f=open('F:/' +filename,'r')
    r=pickle.load(f)
    f.close()
    return r

def specific_file(dir, str1, str2):
    import os, os.path
    filenames_selected = []
    for root, dirname, filenames in os.walk(dir):
        for filename in filenames:
            if str1 in os.path.splitext(filename)[0]:
                if str2 in os.path.splitext(filename)[1]:
                    filenames_selected.append(filename)
    print(filenames_selected)
    return filenames_selected

# PSO optimizing
class PSO():
    def __init__(self, pN, dim, max_iter):  
        self.w = 0.8
        self.c1 = 1.5
        self.c2 = 1.5
        self.pN = pN  
        self.dim = dim  
        self.max_iter = max_iter  
        self.X = np.zeros((self.pN, self.dim))  
        self.Xmax = 100
        self.Xmin = 1
        self.V =  np.zeros((self.pN, self.dim))  
        self.Vmax = 100
        self.Vmin = -100
        # self.Vmax = random.uniform(0, 1)
        # self.Vmin = random.uniform(-1, 0)
        self.pbest = np.zeros((self.pN, self.dim))  
        self.gbest = np.zeros((1, self.dim))  
        self.p_fit = np.zeros(self.pN) 
        self.fit = float(1e10)  

   
    def function(self, XX):#Custom loss function
       
       corr=corr1*((c1+c2)/2)    
        loss=100-corr   
        para3_corr.append(corr)         
    
    #===================================================       
      
        return loss
    
      
   
    def init_Population(self, region):              
        print("initialization")
        for i in range(self.pN):        
            for j in range(self.dim):  
                self.X[i][j] = random.uniform(region[j][0], region[j][1]) 
                self.V[i][j] = random.uniform(-50, 50)
                
            # for j in range(self.dim-1,self.dim): 
            #     self.V[i][j] = random.uniform(-1000, 1000) 
            self.pbest[i] = self.X[i]  
            
            tmp = self.function(self.X[i])    
            self.p_fit[i] = tmp  
            if (tmp < self.fit):  
                self.fit = tmp
                self.gbest = self.X[i]
                               
     
    def iterator(self, region):#iterated function
        fitness = []
        # print("")
        # for i in range(self.pN):         
        #     for j in range(self.dim):  
        #         self.X[i][j] = random.uniform(region[j][0], region[j][1]) 
        #         self.V[i][j] = random.uniform(-1, 1)  
        #     self.pbest[i] = self.X[i]  
            
        #     tmp = self.function(self.X[i])  
            
        #     self.p_fit[i] = tmp  
        #     if (tmp < self.fit):  
        #         self.fit = tmp
        #         self.gbest = self.X[i]
        for t in range(self.max_iter):
            print("Generations:",t)
            # end = time.time()
            # print('t:',end - start, 's')
            # start = time.time()

            for i in range(self.pN):
              
                self.V[i] = self.w * self.V[i] + self.c1 * random.uniform(0,1) * (self.pbest[i] - self.X[i]) + (self.c2 * random.uniform(0,1) * (self.gbest - self.X[i]))
                # print("V[i]:",self.V[i])

                for j in range(self.dim): #random.uniform(region[j][0], region[j][1])
                    if self.V[i][j] > self.Vmax:
                        self.V[i][j] = self.Vmax
                    elif self.V[i][j] < self.Vmin:
                        self.V[i][j] = self.Vmin
               
                self.X[i] = self.X[i] + self.V[i]
                
                for j in range(self.dim):
                    if self.X[i][j] > self.Xmax:
                        self.X[i][j] = random.uniform(region[j][0], region[j][1])
                    elif self.X[i][j]  < self.Xmin:
                        self.X[i][j]  = self.Xmin
                    # print("X[i][j]:",self.X[i][j])  
                    # if self.X[i][j] > self.Xmax:
                    #     self.X[i][j] = self.Xmax
                    # elif self.X[i][j]  < self.Xmin:
                    #     self.X[i][j]  = self.Xmin
                                           


            for i in range(self.pN):  
                print("Updating particle number:",i+1)
                temp = self.function(self.X[i])
               

                if (temp < self.p_fit[i]):  
                    self.pbest[i] = self.X[i]
                    self.p_fit[i] = temp

                
                if (temp < self.fit):  
                    self.gbest = self.X[i]
                    self.fit = temp
            fitness.append(self.fit)
            
            
            # print("pbest[i]:",self.pbest[i],"gbest[i]:",self.gbest[i])  
            # print("Best:", t, self.fit)  
        return fitness


def run():#Define optimization scope
    my_pso = PSO(pN=20, dim=3, max_iter=5)
    region = {0:(1, 120),1:(1, 8),2:(1, 200)} #L0 O B
    my_pso.init_Population(region)
    my_pso.iterator(region)
    # fitness = my_pso.iterator()
    
if __name__ == "__main__":
    print("Automation Library:", mhrc.automation.VERSION)
    pscad_version = 'PSCAD 4.6.2 (x64)'
    fortran_version = 'GFortran 4.6.2'
    fortran_ext = '.gf46'
    project_name = ''
    outputfile_name = ''

    working_dir = os.getcwd() + "\\"
    require_dir = os.getcwd() + "\\"

    src_folder = working_dir + project_name + fortran_ext
    csv_folder = working_dir + str(time.strftime('%Y-%m-%d %H-%M-%S', time.localtime(time.time())))
    

    try:
        shutil.rmtree(src_folder)
    except Exception as ignored:
        pass
    print("work directory：", working_dir)
    print("Enter Folder：", src_folder)
    print("Output Folder：", csv_folder)
    pscad = mhrc.automation.launch_pscad(pscad_version=pscad_version, fortran_version=fortran_version, certificate=False)
    if pscad:
        try:
    # Load Engineering========================================================
            pscad.load([working_dir + project_name + ".pscx"])
            project = pscad.project(project_name)
            project.focus()
            main = project.user_canvas('Main')  
            para3_value=[]
            para3_corr=[]
            ZZ=[]
            tmp=0            
            Icp0 = np.loadtxt('Input target waveform') #Input target waveform
            Icp0 = np.array(Icp0)
            Icp0 =Icp0.reshape(-1,1) 
            Icp0 = Icp0.flatten()       
            Icp0_gy=np.zeros((80,))
            for n in range(len(Icp0_gy)):
                Icp0_gy[n,]=(Icp0[n] - Icp0.min())/(Icp0.max() - Icp0.min())   
#----------------------------------------------------------------------
            L0=main.user_cmp("Module address1")           
            O=main.user_cmp("Module address2")
            B=main.user_cmp("Module address3")
            run()
#-----------------------Normalize the reference waveform---------------

            
        finally:
                pscad.quit()
        pass
    else:
        print("Failed to launch PSCAD")