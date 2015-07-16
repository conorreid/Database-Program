#first, let's import our package of word document to read it, and re to use regular expressions
import re
import docx
#now let's get a real test mint line code
test_text = "	18 Su ra man ra 810/11 2$& 1.70g 1.90g 811/12 2.76g 812/13 4& 2.32g # 2.58g 2.93g 2.85g 813/14 2.70g 814/15 3& 2.45g # 2.70g 3.12g 815/16 2.53g # 816/17 5& 1.56g + 2.63g$ 2.71g 2.75g 2.61g 832 2.79g 815-830 2.94g"
#and let's split it up
test_text_split = test_text.split()
#and get a value
y = len(test_text_split)
#now let's define our detectors
#this is for the mint
mint = ""
def mint_match1(strg, search=re.compile('^[A-Za-z]+[-][A-Za-z]+').search):
	return bool(search(strg))
	
def mint_match2(strg, search=re.compile('^[A-Za-z]').search):
	return bool(search(strg))
	
def date_match1(strg, search=re.compile('^[\d]+/[\d]+').search):
	return bool(search(strg))
	
def date_match2(strg, search=re.compile('^[\d/]+[-][\d/]+').search):
	return bool(search(strg))
	
def date_match3(strg, search=re.compile('^[\d]+\Z').search):
	return bool(search(strg))
	
def weight_match(strg, search=re.compile('^[\d][.][\dg]+').search):
	return bool(search(strg))
	
def just_multiple(strg, search=re.compile('[\d]+[&]').search):
	return bool(search(strg))
	
def frag_multiple(strg, search=re.compile('[\d]+[$][&]').search):
	return bool(search(strg))
	
def just_frag(strg, search=re.compile('^[#]').search):
	return bool(search(strg))
	
	
for w in range(y):
	if mint_match1(test_text_split[w]) or mint_match2(test_text_split[w]):
		if mint == "":
			mint = test_text_split[w]
			print(mint)
		else:
			mint = mint + " " + test_text_split[w]
			print(mint)
	if date_match1(test_text_split[w]) or date_match2(test_text_split[w]) or date_match3(test_text_split[w]):
		date = test_text_split[w]
		print(date)
		if weight_match(test_text_split[w+1]) and just_frag(test_text_split[w+2]):
			weight = test_text_split[w+1]
			comment = "fragment"
			print(mint, date, weight, comment)
		elif weight_match(test_text_split[w+1]):
			weight = test_text_split[w+1]
			print(mint, date, weight)
		if just_multiple(test_text_split[w+1]):
			j = int(filter(str.isdigit, test_text_split[w+1]))
			print(j)
			z = 0
			h = 0
			while z < j:
				if weight_match(test_text_split[w+2+h]) and just_frag(test_text_split[w+3+h]):
					weight = test_text_split[w+2+h]
					comment = "fragment"
					print(mint, date, weight, comment)
					z = z + 1
					h = h + 1
				elif weight_match(test_text_split[w+2+h]):
					weight = test_text_split[w+2+h]
					print(mint, date, weight)
					z = z + 1
					h = h + 1
				else:
					h = h + 1
		if frag_multiple(test_text_split[w+1]):
			j = int(filter(str.isdigit, test_text_split[w+1]))
			print(j)
			z = 0
			h = 0
			while z < j:
				if weight_match(test_text_split[w+2+h]):
					weight = test_text_split[w+2+h]
					comment = "fragment"
					print(mint, date, weight, comment)
					z = z + 1
					h = h + 1
				else:
					h = h + 1