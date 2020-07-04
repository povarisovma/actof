import os
import datetime


def getlistfiles():
    path = 'local_acts'
    generallist = []
    filelist = os.listdir(path)
    if filelist:
        pdflist = []
        for i in filelist:
            if '.pdf' in i:
                pdflist.append(i)

        # temppdflist = [os.path.join(path, file) for file in pdflist]

        # for file in temppdflist:
        #     print(datetime.datetime.fromtimestamp(os.path.getmtime(file)).strftime("%d-%m-%Y %H.%M.%S"))


        for file in pdflist:
            generallist.append((file, datetime.datetime.fromtimestamp(os.path.getctime(os.path.join(path, file))).strftime("%d-%m-%Y %H.%M.%S"),
                               datetime.datetime.fromtimestamp(os.path.getmtime(os.path.join(path, file))).strftime("%d-%m-%Y %H.%M.%S")))
    return generallist


def getDictfiles():
    path = 'local_acts'
    generallist = {}
    filelist = os.listdir(path)
    if filelist:
        pdflist = []
        for i in filelist:
            if '.pdf' in i:
                pdflist.append(i)

        # temppdflist = [os.path.join(path, file) for file in pdflist]

        # for file in temppdflist:
        #     print(datetime.datetime.fromtimestamp(os.path.getmtime(file)).strftime("%d-%m-%Y %H.%M.%S"))

        for i in range(len(pdflist)):
            generallist[i] = ((pdflist[i], datetime.datetime.fromtimestamp(os.path.getctime(os.path.join(path, pdflist[i]))).strftime("%d-%m-%Y %H.%M.%S"),
                               datetime.datetime.fromtimestamp(os.path.getmtime(os.path.join(path, pdflist[i]))).strftime("%d-%m-%Y %H.%M.%S")))
    return generallist


def getDictFilesParam():
    path = 'local_acts'
    generallist = []
    filelist = os.listdir(path)
    if filelist:
        pdflist = []
        for i in filelist:
            if '.pdf' in i:
                pdflist.append(i)

        # temppdflist = [os.path.join(path, file) for file in pdflist]

        # for file in temppdflist:
        #     print(datetime.datetime.fromtimestamp(os.path.getmtime(file)).strftime("%d-%m-%Y %H.%M.%S"))


        for file in pdflist:
            generallist.append({"title": file, "creating": datetime.datetime.fromtimestamp(os.path.getctime(os.path.join(path, file))).strftime("%d-%m-%Y %H.%M.%S"),
                               "modifine": datetime.datetime.fromtimestamp(os.path.getmtime(os.path.join(path, file))).strftime("%d-%m-%Y %H.%M.%S")})
    return generallist

