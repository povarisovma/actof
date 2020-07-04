import os


def get_path():
    return r'E:\tested\Акты'


def get_number_act():
    filelist = os.listdir(get_path())
    numacts = []
    nextnumact = ''
    for i in range(len(filelist)):
        if '.docx' in filelist[i]:
            numacts.append(int(filelist[i].split('_')[0][3:]))
    numacts.sort(reverse=True)
    for i in range(len(numacts)):
        # print(i, numacts[i])
        if 0 <= (numacts[i] - numacts[i + 1]) <= 2:
            nextnumact = str(numacts[i] + 1)
            break
    return nextnumact
