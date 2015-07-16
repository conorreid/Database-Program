#first, let's import our package of word document to read it, and re to use regular expressions
import re
import docx
#now, let's open and save the document itself
doc = docx.Document('Test.docx')
#let's read how many lines (paragraphs) will be in this document
n = len(doc.paragraphs)
#the first line of every document is the region
#because this will stay constant for the entire document, we can define it here
region = doc.paragraphs[0].text
#now, let's define our various functions for finding different variables
#first, we'll do hoard name
def hoard_name_match(strg, search=re.compile('^[\d]+[.]\s+[A-Z]+').search):
	return bool(search(strg))
	
#now, let's try for type
def type_match(strg, search=re.compile('([A-C])[.]\s[A-Z\s]+\s[(\d)]+').search):
	return bool(search(strg))
	
#and for dynasty
def dynasty_match(strg, search=re.compile('([IVXCL]+)[.]\s[A-Za-z\s]+\s[(\d)]+').search):
	return bool(search(strg))
	
#and now for the mint date and weight with comments line
def mint_line_match1(strg, search=re.compile('^[\d]+\s[A-Za][a-z]+').search):
	return bool(search(strg))
	
def mint_line_match2(strg, search=re.compile('^[\d]+\s[A-Za][a-z]+[-]+').search):
	return bool(search(strg))
	
#now let's make our for loop, looping n times where n is the number of lines in the document
for i in range(n):
	#this will create an easier variable to work with
	test_text = doc.paragraphs[i].text
	#this splits up the string pre-emptively by space
	test_text_split = test_text.split()
	#this tests for the length of the line
	y = len(test_text_split)
	#this checks if the line matches the hoard name case
	if type_match(test_text):
		type_name = test_text_split[1:(y-1)]
		print(type_name)
	elif dynasty_match(test_text):
		dynasty = test_text_split[1:(y-1)]
		print(dynasty)
	elif hoard_name_match(test_text):
		hoard_name = test_text_split[1:y]
		print(hoard_name)
	elif mint_line_match1(test_text) or mint_line_match2(test_text):
		mint_line = test_text
		print(test_text)