import os

directory = os.getcwd()
files = os.listdir(directory)
countOfFiles = len(files)
files = filter(lambda x: x.endswith('.xls'), files)
os.chdir(directory)
print directory
print files

for i in range(0, countOfFiles):
    os.rename(directory + "\\" + files[i], str(i + 1) + '.html')
