#try to get the modules to create files
from msilib.schema import Error


try:
	from docx import Document #import the library to manipulate the docx
	from docx.shared import Cm #import the lenght unit
	from docx2pdf import convert #import the converter
	from os import chdir,listdir,getcwd #import OS directory manipulations
except ImportError as error:
	print(str(error) + " please install.")
	input("Type Enter to exit...")
	exit(None)


#creating document docx
document = Document()

#changing the layout of the document
sections = document.sections
for section in sections:
	section.top_margin = Cm(1.27)
	section.bottom_margin = Cm(1.27)
	section.left_margin = Cm(1.27)
	section.right_margin = Cm(1.27)


#name of the documents
name = input("Type the name of the document: ")
#puting the extension ".docx" in the file name
name_docx = name + ".docx"

#path to save the documents
path_document = input("Type the absolute path of the file: ")

#input of the directory of the imagens to seve in the file
path_images = input("Type the absolute path to the images: ")
print() #just to have another break row

#try to go to the directory of the pictures and save the documents to the specified path
try:
	#changing directory to the directory of the images
	chdir(path_images)

	#get the number of images
	images = listdir()
	num = len(images)

	#adding the images to the document in order, the name of the images shoud be "1.jpeg", "2.jpeg", "3.jpeg",...
	for i in range(1,num+1):
		pic_name = str(i) + ".jpeg" #build the picture name
		document.add_picture(pic_name, width=Cm(21-1.27), height=Cm(26.7-1.27)) #put the picture in the document

	#changing directory to the directory to save the files
	chdir(path_document)
	#save docx
	document.save(name_docx)
	print("Docx created!")

	#alterating the name with path to save the PDF
	name_pdf = getcwd() + "\\" + name + ".pdf"

	#convert docx to pdf
	convert(name_docx, name_pdf)
	print("PDF created!")

#if something goes wrong
except:
	print("Impossible to create document!")
	input("Type Enter to exit...")
	exit()