#read file
myFile = open("sample.txt", "w")

print("Name " + myFile.name)
print("Mode " + myFile.mode)
#write to file
myFile.write("GBJ : 100\nKHD : 99\nBBB : 89")
myFile.close()
#read the file
myFile = open("sample.txt", "r")
print("Reading...\n" + myFile.read())