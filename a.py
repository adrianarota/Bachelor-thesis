# -*- coding: UTF-8 -*-
from bs4 import BeautifulSoup
import xlsxwriter 
import sys
import xml.etree.ElementTree as ET
import cssutils
import logging
import unicodedata
import re
from difflib import SequenceMatcher 
import math
from time import ctime


colorStrings=[["AliceBlue" , "#F0F8FF"],
["AntiqueWhite" , "#FAEBD7"],
["Aqua" , "#00FFFF"],
["Aquamarine" , "#7FFFD4"],
["Azure" , "#F0FFFF"],
["Beige" , "#F5F5DC"],
["Bisque" , "#FFE4C4"],
["Black" , "#000000"],
["BlanchedAlmond" , "#FFEBCD"],
["Blue" , "#0000FF"],
["BlueViolet" , "#8A2BE2"],
["Brown" , "#A52A2A"],
["BurlyWood" , "#DEB887"],
["CadetBlue" , "#5F9EA0"],
["Chartreuse" , "#7FFF00"],
["Chocolate" , "#D2691E"],
["Coral" , "#FF7F50"],
["CornflowerBlue" , "#6495ED"],
["Cornsilk" , "#FFF8DC"],
["Crimson" , "#DC143C"],
["Cyan" , "#00FFFF"],
["DarkBlue" , "#00008B"],
["DarkCyan" , "#008B8B"],
["DarkGoldenRod" , "#B8860B"],
["DarkGray" , "#A9A9A9"],
["DarkGrey" , "#A9A9A9"],
["DarkGreen" , "#006400"],
["DarkKhaki" , "#BDB76B"],
["DarkMagenta" , "#8B008B"],
["DarkOliveGreen" , "#556B2F"],
["DarkOrange" , "#FF8C00"],
["DarkOrchid" , "#9932CC"],
["DarkRed" , "#8B0000"],
["DarkSalmon" , "#E9967A"],
["DarkSeaGreen" , "#8FBC8F"],
["DarkSlateBlue" , "#483D8B"],
["DarkSlateGray" , "#2F4F4F"],
["DarkSlateGrey" , "#2F4F4F"],
["DarkTurquoise" , "#00CED1"],
["DarkViolet" , "#9400D3"],
["DeepPink" , "#FF1493"],
["DeepSkyBlue" , "#00BFFF"],
["DimGray" , "#696969"],
["DimGrey" , "#696969"],
["DodgerBlue" , "#1E90FF"],
["FireBrick" , "#B22222"],
["FloralWhite" , "#FFFAF0"],
["ForestGreen" , "#228B22"],
["Fuchsia" , "#FF00FF"],
["Gainsboro" , "#DCDCDC"],
["GhostWhite" , "#F8F8FF"],
["Gold" , "#FFD700"],
["GoldenRod" , "#DAA520"],
["Gray" , "#808080"],
["Grey" , "#808080"],
["Green" , "#008000"],
["GreenYellow" , "#ADFF2F"],
["HoneyDew" , "#F0FFF0"],
["HotPink" , "#FF69B4"],
["IndianRed" , "#CD5C5C"],
["Indigo" , "#4B0082"],
["Ivory" , "#FFFFF0"],
["Khaki" , "#F0E68C"],
["Lavender" , "#E6E6FA"],
["LavenderBlush" , "#FFF0F5"],
["LawnGreen" , "#7CFC00"],
["LemonChiffon" , "#FFFACD"],
["LightBlue" , "#ADD8E6"],
["LightCoral" , "#F08080"],
["LightCyan" , "#E0FFFF"],
["LightGoldenRodYellow" , "#FAFAD2"],
["LightGray" , "#D3D3D3"],
["LightGrey" , "#D3D3D3"],
["LightGreen" , "#90EE90"],
["LightPink" , "#FFB6C1"],
["LightSalmon" , "#FFA07A"],
["LightSeaGreen" , "#20B2AA"],
["LightSkyBlue" , "#87CEFA"],
["LightSlateGray" , "#778899"],
["LightSlateGrey" , "#778899"],
["LightSteelBlue" , "#B0C4DE"],
["LightYellow" , "#FFFFE0"],
["Lime" , "#00FF00"],
["LimeGreen" , "#32CD32"],
["Linen" , "#FAF0E6"],
["Magenta" , "#FF00FF"],
["Maroon" , "#800000"],
["MediumAquaMarine" , "#66CDAA"],
["MediumBlue" , "#0000CD"],
["MediumOrchid" , "#BA55D3"],
["MediumPurple" , "#9370DB"],
["MediumSeaGreen" , "#3CB371"],
["MediumSlateBlue" , "#7B68EE"],
["MediumSpringGreen" , "#00FA9A"],
["MediumTurquoise" , "#48D1CC"],
["MediumVioletRed" , "#C71585"],
["MidnightBlue" , "#191970"],
["MintCream" , "#F5FFFA"],
["MistyRose" , "#FFE4E1"],
["Moccasin" , "#FFE4B5"],
["NavajoWhite" , "#FFDEAD"],
["Navy" , "#000080"],
["OldLace" , "#FDF5E6"],
["Olive" , "#808000"],
["OliveDrab" , "#6B8E23"],
["Orange" , "#FFA500"],
["OrangeRed" , "#FF4500"],
["Orchid" , "#DA70D6"],
["PaleGoldenRod" , "#EEE8AA"],
["PaleGreen" , "#98FB98"],
["PaleTurquoise" , "#AFEEEE"],
["PaleVioletRed" , "#DB7093"],
["PapayaWhip" , "#FFEFD5"],
["PeachPuff" , "#FFDAB9"],
["Peru" , "#CD853F"],
["Pink" , "#FFC0CB"],
["Plum" , "#DDA0DD"],
["PowderBlue" , "#B0E0E6"],
["Purple" , "#800080"],
["RebeccaPurple" , "#663399"],
["Red" , "#FF0000"],
["RosyBrown" , "#BC8F8F"],
["RoyalBlue" , "#4169E1"],
["SaddleBrown" , "#8B4513"],
["Salmon" , "#FA8072"],
["SandyBrown" , "#F4A460"],
["SeaGreen" , "#2E8B57"],
["SeaShell" , "#FFF5EE"],
["Sienna" , "#A0522D"],
["Silver" , "#C0C0C0"],
["SkyBlue" , "#87CEEB"],
["SlateBlue" , "#6A5ACD"],
["SlateGray" , "#708090"],
["SlateGrey" , "#708090"],
["Snow" , "#FFFAFA"],
["SpringGreen" , "#00FF7F"],
["SteelBlue" , "#4682B4"],
["Tan" , "#D2B48C"],
["Teal" , "#008080"],
["Thistle" , "#D8BFD8"],
["Tomato" , "#FF6347"],
["Turquoise" , "#40E0D0"],
["Violet" , "#EE82EE"],
["Wheat" , "#F5DEB3"],
["White" , "#FFFFFF"],
["WhiteSmoke" , "#F5F5F5"],
["Yellow" , "#FFFF00"],
["YellowGreen" , "#9ACD32"]]

under12=0
under16=0
sizecnt=0
mlink=[]
fl= []
colorlist=[]
colorpalette=[]
fs=1
lineh=11.5
heights=[lineh]
fontsizes=[1]
ratioviolation=0

if len(sys.argv) < 2:
	sys.exit()
else:
	file = 	sys.argv[1]
	
try:
	soup = BeautifulSoup(open(file), 'html.parser')	
except:
	print "No file found. Exiting script"
	sys.exit()

cssutils.log.setLevel(logging.CRITICAL)	
tree = ET.parse('fonts.xml')
root = tree.getroot()


#Finds the longest substring between 2 strings and returns it if it's longer than 2/3 of the first one
#in: str1, str2 -strings
#out: string or None
def longestSubstring(str1,str2): 
	seqMatch = SequenceMatcher(None,str1.lower(),str2.lower()) 
	match = seqMatch.find_longest_match(0, len(str1), 0, len(str2))
	if (match.size>=len(str1)*2/3):
		return (str1[match.a: match.a + match.size])
	else:
		return (None )
		

#Find the desirability value for a font
# in: font - font name
#out: int value that represents the font's desirability, or 0 if not found
def findInXML(font):
	nodes = tree.find('.//Font[@name="'+font+'"]')
	if nodes!= None:
		x= nodes.attrib['desirability']
		return(x)
	else:
		return(0)
		

#Find the code for a specific standard color 
# in: color - color name
#out: string value that represents the color's code
def returncode(color):
	
	for l in colorStrings:
		if l[0].lower()== color.lower():
			return l[1]
	return None
				

#Fixes color format mistakes
# in: str - color string format
#out: string value with 2 bytes for each component	
def normalizecolor(str):
		
		if len(str)==4:
			
			str=str[1:]
			out = [(str[i:i+1]) for i in range(0, len(str))]	
			for i in range(0,3):
				out[i]=out[i]+out[i]
				
			return "#"+ out[0]+out[1]+out[2]
		else:
			return str
			

def parseSheet(sheet):
	global sizecnt
	global under12
	global under16
	
	print "Analyzing fonts ..."
	for rule in sheet:
		f=0
		h=0
		if rule.type == rule.STYLE_RULE:
        
				for property in rule.style:
					if property.name == 'font-family':
						
						x = property.value.split(",")
					
						for i in x:
							i= i.lstrip()
							i= i.replace('"', '')
							i= i.replace('\'', '')
							i=i.capitalize()
							r= findInXML(i)
							if [i,r] not in fl:
								fl.append([i,r])
					if property.name == 'font-size':
						x = property.value
						num = re.findall(r'\d+',x)
						sizecnt+=1
						
						fs=int(num[0]) 
						f=fs
						#fontsizes.append(fs)
						if fs < 16 and fs>=13:
							under16+=1
							
						if fs < 13:
							under12+=1
					if property.name == 'color' or property.name == 'background-color':

						if "#" in property.value:
							#r = property.value
							r=normalizecolor(property.value)
							
							if r not in colorlist:
								
								colorlist.append(r.strip())
						else:
							z=returncode(property.value)
							if z!=None:
								if z not in colorlist:
									colorlist.append(z)

					if property.name== "line-height":
						x = property.value
						
						l = re.findall(r'\d+',x)
						
						if(len(l)==1):
							lineh= int (l[0])
							h=lineh
							#heights.append(lineh)
		if h!=0 and f!=0:
			heights.append(h)
			fontsizes.append(f)

#Splits a RGB code into its component colors
#in: str- color in string format
#out: r, g, b- int color values
def parsecolor(str):

	#print str
	str=str[1:]
	out1 = [(str[i:i+2]) for i in range(0, len(str), 2)]		
	try:
		r=int(out1[0],16)
	except:
		r=0
	try:
		g=int(out1[1],16)
	except:
		g=0
	try:
		b=int(out1[2],16)
	except:
		b=0	
	return r , g, b


#Difference between 2 colors
# in: str1, str2 - 2 colors in string format
#out: Euclidian difference	
def findDifference(str1, str2):
	a1, b1, c1 = parsecolor(str1)
	a2, b2, c2 = parsecolor(str2)
	return math.sqrt(2*(a2-a1)**2+4*(b2-b1)**2+3*(c2-c1)**2)


#Transforms RGB to HST format
# in: r, g, b - values to transform
#out: h, s, l - transformed values		
def rgb_to_hsl(r, g, b):
	r=float(r)/255
	g=float(g)/255
	b=float(b)/255
	high = max(r, g, b)
	low = min(r, g, b)
   # h, s, v = ((high + low) / 2,)*3
	l= (high + low) /2.0
	d= high- low
	if d==0:
		h = 0.0
		s = 0.0
		return round(h), round(s*100), round(l*100)
	else:
	
		if l<0.5:
			s= d / (high + low)
		else:
			s = d / (2 - high - low) 	
	if r==high:

		h= (g - b) / d 
	elif g ==high:
		h= (b - r) / d + 2
	else: 
		h=(r - g) / d + 4
    
	h=h *60;
	if (h<0):
		h+=360

	return round(h), round(s*100), round(l*100)

#Helper function for transformation between HSL and RGB format
# in: p, q, hue - values needed 
#out: hue for the given color			
def hue_to_rgb(p, q, hue):
		
        if hue < 0:
			hue+=1
        if hue > 1:
			hue -= 1
        if hue*6 < 1: return p + (q - p) * 6 * hue
        if hue*2 < 1: return q
        if hue*3 < 2: return p + (q - p) * (2.0/3 - hue) * 6
        return p


#Transforms HSL to RGB format
# in: h, s, l - values to transform
#out: r, g, b - transformed values		
def hsl_to_rgb(h, s, l):
	l=l/100
	s=s/100
	if s == 0:
		r = l*255
		r=hex(int(round(r,0)))

		if int(r,16)<16:
			
			r= "0"+str(r)[2:]
		else:
			r= str(r)[2:]

		return "#"+r+r+r
	else:
		if ( l< 0.5 ): 
			var_2 = l * ( 1 + s )
		else:
			var_2 = ( l + s ) - ( s * l )

		var_1 = 2 * l - var_2
		
		h=h/360
	
		r =255* hue_to_rgb(var_1,var_2, h + 1.0/3)
		#print r
		
		g =255* hue_to_rgb(var_1,var_2, h)
	#	print g
		b =255* hue_to_rgb(var_1,var_2,h - 1.0/3)
	#	print b
	r=hex(int(round(r,0)))
	
	if int(r,16)<16:
		r= "0"+str(r)[2:]
	else:
		r= str(r)[2:]
		
	g=hex(int(round(g,0)))
	if int(g,16)<16:
		g= "0"+str(g)[2:]
	else:
		g= str(g)[2:]
			
	b=hex(int(round(b,0)))
	if int(b,16)<16:
		b= "0"+str(b)[2:]
	else:
		b= str(b)[2:]

	return "#"+r+g+b

	
#Generates complementary color palette
# in: h- hue
# out: generated hue from h	
def complementary(h):
	return abs(h +180)

	
#Generates splitcomplementary color palette
# in: h- hue
# out: h1, h2 -generated hues from h
def splitcomplementary(h):
	h1 = abs(h +150-360)
	h2 =abs(h +210-360)
	return h1, h2
	
	
#Generates triadic color palette
# in: h- hue
# out: h1, h2 -generated hues from h	
def triadic(h):
	h1 = abs(h +120-360)
	h2 =abs(h +240-360)
	return h1, h2

	
#Generates analogous color palette
# in: h- hue
# out: h1, h2, h3 -generated hues from h
def analogous(h):
	h1 = abs(h +30-360)
	h2 =abs(h +60-360)
	h3 =abs(h +90-360)	
	return h1, h2,  h3


#Analyzes the colors used and generates suggested color patterns
# in: none
# out: none 
def colorProcess():
	print "Analyzing the color palette..."
	for c1 in range(len(colorlist)):
		a, b, c = parsecolor(colorlist[c1])
		
		h, s, l = rgb_to_hsl(a,b,c)
		
	#	print h , s , l 
		hc = complementary(h)
		cc = hsl_to_rgb(hc, s, l)
	#	print cc
		
			
		colorpalette.append(cc)
		sc, sc2 = splitcomplementary(h)
		
		cc = hsl_to_rgb(sc, s, l)
	#	print cc
		
		cc2 = hsl_to_rgb(sc2, s, l)
	#	print cc2
		
		colorpalette.append(cc)
		colorpalette.append(cc2)
		
		sc, sc2 = triadic(h)
		
		cc = hsl_to_rgb(sc, s, l)
	#	print cc
		
		cc2 = hsl_to_rgb(sc2, s, l)
	#	print cc2
		colorpalette.append(cc)
		colorpalette.append(cc2)
		
		sc, sc2, sc3 = analogous(h)
		
		cc = hsl_to_rgb(sc, s, l)
	#	print cc
		
		cc2 = hsl_to_rgb(sc2, s, l)
	#	print cc2
		cc3 = hsl_to_rgb(sc3, s, l)
		
		colorpalette.append(cc)
		colorpalette.append(cc2)
		colorpalette.append(cc3)



#Parses CSS from a html file and then the rest of the file
#in: none
#out:none		
def parseCSS():
	print "Getting data..."
	x= soup.find('style',{"type" : "text/css"})

	if x!=None:
		x=str(x)
		x=x.replace('</style>','')
		x= x.replace('<style type="text/css">','')
		sheet = cssutils.parseString(x)
		parseSheet(sheet)
	parseHTML()
	findMysteryLink()
	colorProcess()				

		
#Parses CSS file
#in: none
#out:none
def parseCSSfile():
	print "Getting data..."

	parser = cssutils.CSSParser()
	sheet = parser.parseFile(file, 'utf-8')

	parseSheet(sheet)
	colorProcess()

	
								
# Parses attributes for information
# in: list of style attributes obtained with beautifulsoup
# out: none; generates lists of information about fonts, colors etc
def parseStyleAtrib(list):
	for a in list: 
		if a.has_attr('style'):
			f=0
			h=0		
			st=a["style"].split(';')
			for s in st:
				if "font-family" in s:
				
					s= s.split(":")
					x = s[1].split(",")
					for i in x:
						i= i.lstrip()
						i= i.replace('"', '')
						i= i.replace('\'', '')
						i=i.capitalize()
						r= findInXML(i)
						if [i,r] not in fl:
							fl.append([i,r])
				if "color" in s:

					r= s.split(":")[1]
						
					r = normalizecolor(r)
					if r not in colorlist:
							
						if "#" in r:
								
							colorlist.append(r)
						else:
							z=returncode(r)
							if z!=None:
								if r not in colorlist:
									colorlist.append(z)
				if 'font-size' in s:
					x = s.split(":")[1]
					num = re.findall(r'\d+',x)
					sizecnt+=1
					fs=int(num[0]) 
					f=fs
						#fontsizes.append(fs)
					if fs < 16 and fs>=13:
						under16+=1
					if fs < 13:
						under12+=1
				if "line-height" in s:
					x = s.split(":")[1]
					l = re.findall(r'\d+',x)
						
					if(len(l)==1):
						lineh= int (l[0])
						h=lineh
			if h!=0 and f!=0:
				heights.append(h)
				fontsizes.append(f)			
				
				
				
problempairs=[]
def processheights():
	global ratioviolation
	for i in range(0, len(heights)):
		if heights[i]/fontsizes[i]<1.4 or heights[i]/fontsizes[i]>1.6:
			ratioviolation+=1
			problempairs.append([heights[i], fontsizes[i]])
						
							
classes=[]						
#Parses html file, looks for info in different type of tags
# in: none
# out:none
def parseHTML():
	for element in soup.find_all(class_=True): #find classes in html file
		if element["class"] not in classes:
			classes.append(element["class"])

	l = soup.findAll('a')
	parseStyleAtrib(l)
	l = soup.findAll('div')
	parseStyleAtrib(l)
	l = soup.findAll('h1')
	parseStyleAtrib(l)
	l = soup.findAll('h2')
	parseStyleAtrib(l)
	l = soup.findAll('body')
	parseStyleAtrib(l)

#Looks for "mistery meat" links in html files using keywords
# in: none
# out: none; generates a list of suspicious links
def findMysteryLink():
	print "Analyzing links..."
	count =0
	links = soup.findAll('a')
	for a in links: 
		if a.has_attr('href'):
			if a.text.strip() != '':
				count =0
				subs = a.text.split(" ")
				for s in subs:
					if s.strip() != '':
						normal = unicodedata.normalize('NFKD', s).encode('ASCII', 'ignore')

						if  normal.strip() != '':
							ret= longestSubstring(normal,a["href"])
							if ret != None:
								count=count+1							
							
				if count==0:
					mlink.append([a["href"],a.text.strip()])



# Generates the report on the given file in an Excel document
# in: none
# out: none
def printReport():
	cpcouner=0
	
	workbook = xlsxwriter.Workbook('Report.xlsx') 
	
	worksheet = workbook.add_worksheet('Report') 
	cell_format = workbook.add_format()
	cell_format.set_bold()
	#cell_format.set_font_color('red')
	cell_format.set_font_size(15)
	worksheet.write(0, 0, "Report on file: " + sys.argv[1] + ", date: " + ctime(),cell_format) 
	row = 2
	column = 0
	cell_format = workbook.add_format()
	cell_format.set_bold(False)
	#cell_format.set_font_color('red')
	cell_format.set_font_size(15)
	worksheet.write(row , 0, "Rating of the fonts used: ", cell_format)
	row+=1
	
	cell_format_info = workbook.add_format()
	cell_format_info.set_bold(True)
	#cell_format.set_font_color('red')
	cell_format_info.set_font_size(10)
	worksheet.write(row , 0, "5")
	worksheet.write(row , 1, "Most desirable(eg.serif)", cell_format_info)
	row+=1
	worksheet.write(row , 0, "4")
	worksheet.write(row , 1, "Desirable(eg. Times New Roman)", cell_format_info)
	row+=1
	worksheet.write(row , 0, "3")
	worksheet.write(row , 1, "Neutral", cell_format_info)
	row+=1
	worksheet.write(row , 0, "2")
	worksheet.write(row , 1, "Generally illegible(eg. script fonts)", cell_format_info)
	row+=1
	worksheet.write(row , 0, "1")
	worksheet.write(row , 1, "Not a font(eg. wingbats)", cell_format_info)
	row+=1
	worksheet.write(row , 0, "0")
	worksheet.write(row , 1, "Not found in database", cell_format_info)
	row+=2
	
	

	for [item,n] in fl :
		worksheet.write(row, 0, item) 
	
		worksheet.write_number(row, 1, int(n))

		row += 1
    
	worksheet.write_formula(row,1, '=IFERROR(AVERAGEIF(B3:B%d,"<>0"),0)' % (row)) 
	
	row += 3

	worksheet.write(row , 0, "Potential mistery meat links: ", cell_format)
	row +=1
	worksheet.write(row , 0, "Link")	
	worksheet.write(row , 1, "Link name")
	
	row +=1
	for [i,k] in mlink:
		worksheet.write(row, 0, i) 
		
		worksheet.write(row, 1, k)
		row +=1
	
	row += 3
	worksheet.write(row , 0, "Report on font size: ", cell_format)
	row +=1
	cell_format = workbook.add_format()
	cell_format.set_bold(False)
	#cell_format.set_font_color('red')
	cell_format.set_font_size(12)
	worksheet.write(row , 0, "The best size for body text is at least 16px and for secondary text is at least 13px", cell_format_info)	
	row +=1
	worksheet.write(row , 1, "Classes with font-size under 16px" , cell_format)
	worksheet.write(row, 0, str(under16)+"/"+ str(sizecnt))
	row +=1
	worksheet.write(row , 1, "Classes with font-size under 12px", cell_format)
	worksheet.write(row, 0, str(under12)+"/"+ str(sizecnt))
	row +=2
	
	worksheet.write(row , 0, "The ratio between font size and line height should be around 1.5", cell_format_info)
	row +=1
	worksheet.write(row , 0, ratioviolation , cell_format)
	worksheet.write(row , 1, "Number of rule violation (ratio not between 1.4 and 1.6)", cell_format)
	
	row+=1
	
	row += 3
	cell_format = workbook.add_format()

	cell_format.set_font_size(15)
	worksheet.write(row , 0, "Report on colors: ", cell_format)
	row +=1
	worksheet.write(row , 0, "Colors used:")
	worksheet.write(row , 2, "Suggested color palettes: complementary, split-complementary, triadic, analogous")
	row +=1
	
	for k in range(0,len(colorlist)):
	#for c in colorlist:
	
		cell_format = workbook.add_format()
		cell_format.set_bg_color(colorlist[k])
		
		if int(colorlist[k][1:],16) >=0x7FFFFF:
			cell_format.set_font_color('black')
		if int(colorlist[k][1:],16) <0x7FFFFF:
			cell_format.set_font_color('white')
		
		worksheet.write(row, 0, colorlist[k], cell_format)
		cell_format = workbook.add_format()
		
		for j in range (0,8):
			cell_format_cp = workbook.add_format()
			
			cell_format_cp.set_bg_color(colorpalette[cpcouner+j])
			if int(colorpalette[cpcouner+j][1:],16) >=0x7FFFFF:
				cell_format_cp.set_font_color('black')
			if int(colorpalette[cpcouner+j][1:],16) <0x7FFFFF:
				cell_format_cp.set_font_color('white')
			worksheet.write(row, 2+j, colorpalette[cpcouner+j], cell_format_cp)
		
		cpcouner+=8
		
		row+=1
	if len(classes)!=0:
		worksheet.write(row, 0, "There are " + str(len(classes))+ " css classes in this html file")
		row+=1
		worksheet.write(row, 0, "If you don't find the information you are looking for, please run the script on the coresponding css file(s)",cell_format_info)
		
	try:
		workbook.close()
		print "Report completed"
	except:
		print "Close the opened report and run the script again"


		
if file.endswith(".css"):
	parseCSSfile()
	processheights()
	printReport()

elif file.endswith(".html"):
	parseCSS()
	processheights()
	printReport()

else: 
	print "Please enter the path to a css/html file"