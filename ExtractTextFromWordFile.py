import docx2txt
from os import listdir, makedirs
from os.path import join, exists, splitext, basename

def extractText(bookPath):
    outputPath = createOutputFolder(bookPath)
    dirs = listdir(bookPath)
    dirs = [item for item in dirs if item.endswith(".docx")]
    print("Found " + str(len(dirs)) + " books.")
    print("Starting to convert...")
    i = 0
    for file in listdir(bookPath):
        if file.endswith(".docx"):
            createTextFile(join(bookPath, file), outputPath)
            i = i+1
            print(str(i) + " of " + str(len(dirs)) + " done.")


def createOutputFolder(basePath):
    newpath = join(basePath, "output")
    if not exists(newpath):
        makedirs(newpath)
    return newpath

def createTextFile(file, outputPath):
    text = docx2txt.process(file)
    rawFileName = getRawFileName(file)
    txtFileName = rawFileName+".txt"
    text_file = open(join(outputPath, txtFileName), "w")
    n = text_file.write(text)
    text_file.close()


def getRawFileName(file):
    base=basename(file)
    return splitext(base)[0]


if __name__ == "__main__":
    print('Enter location of book folder:')
    path = input()
    extractText(path)
    print('Process finished.')