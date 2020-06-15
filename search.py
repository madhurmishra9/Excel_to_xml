import win32api
import os, fnmatch

def get_drive():
    drives = win32api.GetLogicalDriveStrings()
    #print(drives)
    drives = drives.split('\0')[:-1]
    #print(drives)
    return drives

def file_search(filename):
    drives = get_drive()
    search_result = []
    for drive in drives:
        #print(drive)
        for root, dir, files in os.walk(drive):
            if filename in files:
                search_result.append(os.path.join(root,filename))
    #print (search_result)
    for i in range(0,len(search_result)):
        print('%d.%s'%(i+1,search_result[i]))
    if search_result == []:
        print('No file found with name %s'%filename)


