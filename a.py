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


cssutils.log.setLevel(logging.CRITICAL)
	
tree = ET.parse('fonts.xml')
root = tree.getroot()
under12=0
under16=0
sizecnt=0
mlink=[]
mlinkname=[]	
fl= []
sl= []
colorlist=[]
colorpalette=[]

if len(sys.argv) < 2:
    print "Please give the path to the css/html file"
else:
	file = 	sys.argv[1]
	

def findInXML(font):
	nodes = tree.find('.//Font[@name="'+font+'"]')
	if nodes!= None:
		x= nodes.attrib['desirability']
		return(x)
	else:
		return(0)


def returncode(color):
	
	for l in colorStrings:
		if l[0].lower()== color.lower():
			return l[1]
	return None
				
soup = BeautifulSoup(open(file), 'html.parser')
def normalizecolor(str):
		
		if len(str)==4:
			print str
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
		
		if rule.type == rule.STYLE_RULE:
        
				for property in rule.style:
					if property.name == 'font-family':
						
						x = property.value.split(",")
						print x
						for i in x:
							i= i.lstrip()
							i= i.replace('"', '')
							r= findInXML(i)
							if [i,r] not in fl:
								fl.append([i,r])
					if property.name == 'font-size':
						x = property.value
						num = re.findall(r'\d+',x)
						sizecnt+=1
						if int(num[0]) < 16:
							under16+=1
							
						if int(num[0]) < 13:
							under12+=1
					if property.name == 'color':
						print property
						if "#" in property.value:
							#r = property.value
							r=normalizecolor(property.value)
							if r not in colorlist:
								colorlist.append(r.strip())
						else:
							z=returncode(property.value)
							if z!=None:
								if property.value not in colorlist:
									colorlist.append(z)
								


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

	
def findDifference(str1, str2):
	a1, b1, c1 = parsecolor(str1)
	a2, b2, c2 = parsecolor(str2)
	return math.sqrt((a2-a1)**2+(b2-b1)**2+(c2-c1)**2)




def rgb_to_hsl(r, g, b):
	r=float(r)/255
	g=float(g)/255
	b=float(b)/255
	high = max(r, g, b)
	low = min(r, g, b)
   # h, s, v = ((high + low) / 2,)*3
	l= (high + low) /2
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
	if r==max:
		h= (g - b) / d 
	elif g ==max:
		h= (b - r) / d + 2
	else: 
		h=(r - g) / d + 4
    
	h=h *60;
	if (h<0):
		h+=360

	return round(h), round(s*100), round(l*100)

def hue_to_rgb(p, q, hue):
		
        if hue < 0:
			hue+=1
        if hue > 1:
			hue -= 1
        if hue*6 < 1: return p + (q - p) * 6 * hue
        if hue*2 < 1: return q
        if hue*3 < 2: return p + (q - p) * (2.0/3 - hue) * 6
        return p
		
def hsl_to_rgb(h, s, l):
	l=l/100
	s=s/100
	if s == 0:
		r, g, b = l*255, l*255, l*255
		r=hex(int(r))
		if int(r,16)<16:
			
			r= "0"+str(r)[2:]
		else:
			r= str(r)[2:]
		
		g=hex(int(g))
		if int(g,16)<16:
			
			g= "0"+str(g)[2:]
		else:
			g= str(g)[2:]
				
		b=hex(int(b))
		if int(b,16)<16:
			
			b= "0"+str(b)[2:]
		else:
			b= str(b)[2:]
		
		return "#"+r+g+b
	else:
		if ( l< 0.5 ): 
			var_2 = l * ( 1 + s )
		else:
			var_2 = ( l + s ) - ( s * l )

		var_1 = 2 * l - var_2
		
		h=h/360
	
		
		r =255* hue_to_rgb(var_1,var_2, h + 1.0/3)
		g =255* hue_to_rgb(var_1,var_2, h)
		b =255* hue_to_rgb(var_1,var_2,h - 1.0/3)
		
	r=hex(int(r))
	if int(r,16)<16:
		r= "0"+str(r)[2:]
	else:
		r= str(r)[2:]
		
	g=hex(int(g))
	if int(g,16)<16:
		g= "0"+str(g)[2:]
	else:
		g= str(g)[2:]
			
	b=hex(int(b))
	if int(b,16)<16:
		b= "0"+str(b)[2:]
	else:
		b= str(b)[2:]

	return "#"+r+g+b

	
	
def complementary(h):
	return abs(h +180)

def splitcomplementary(h):
	h1 = abs(h +150-360)
	h2 =abs(h +210-360)
	return h1, h2

def colorProcess():
	print "Analyzing the color palette..."
	for c1 in range(len(colorlist)):

		a, b, c = parsecolor(colorlist[c1])
		
		h, s, l = rgb_to_hsl(a,b,c)
		print h , s , l 
		hc = complementary(h)
		cc = hsl_to_rgb(hc, s, l)
		print cc
		
			
		colorpalette.append(cc)
		sc, sc2 = splitcomplementary(h)
		
		cc = hsl_to_rgb(sc, s, l)
		print cc
		
		cc2 = hsl_to_rgb(sc2, s, l)
		print cc2
		
		colorpalette.append(cc)
		colorpalette.append(cc2)
		print "............."
	
def parseCSS():
	x= soup.find('style',{"type" : "text/css"})

	if x!=None:
		x=str(x)
		x=x.replace('</style>','')
		x= x.replace('<style type="text/css">','')
		sheet = cssutils.parseString(x)
		print x
		parseSheet(sheet)
								

parseCSS()

def parseCSSfile():
	print "Getting data..."
	#global sizecnt
#	global under12
#	global under16
	parser = cssutils.CSSParser()
	sheet = parser.parseFile(file, 'utf-8')

	parseSheet(sheet)
					
							
								
parseCSSfile()			

def parseStyleAtrib(list):
	for a in list: 
		if a.has_attr('style'):
				st=a["style"].split(';')
				for s in st:
					if "font-family" in s:
					
						s= s.split(":")
						x = s[1].split(",")
						for i in x:
							i= i.lstrip()
							i= i.replace('"', '')
							r= findInXML(i)
							if [i,r] not in fl:
								fl.append([i,r])
					if "color" in s:
					
						r= s.split(":")[1]
						print r
						print len(r)
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
						if int(num[0]) < 16:
							under16+=1
							
						if int(num[0]) < 13:
							under12+=1

def parseHTML():
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
			
parseHTML()
colorProcess()
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


findMysteryLink()


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
	
	for [item,n] in fl :
		worksheet.write(row, 0, item) 
	
		worksheet.write_number(row, 1, int(n))

		row += 1
    
	worksheet.write_formula(row,1, '=IFERROR(AVERAGEIF(B3:B%d,"<>0"),0)' % (row)) 
	
	row += 3

	worksheet.write(row , 0, "Potential mistery meat links: ", cell_format)
	row +=1
	worksheet.write(row , 0, "Link", cell_format)	
	worksheet.write(row , 1, "Link name", cell_format)
	
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
	worksheet.write(row , 0, "The best size for body text is at least 16px and for secondary text is at least 13px", cell_format)	
	row +=1
	worksheet.write(row , 0, "Classes with font-size under 16px:" , cell_format)
	worksheet.write(row, 1, str(under16)+"/"+ str(sizecnt))
	row +=1
	worksheet.write(row , 0, "Classes with font-size under 12px:", cell_format)
	worksheet.write(row, 1, str(under12)+"/"+ str(sizecnt))
	
	row += 3
	worksheet.write(row , 0, "Colors used:", cell_format)
	worksheet.write(row , 1, "Complementaries:", cell_format)
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
		
		for j in range (0,3):
			cell_format_cp = workbook.add_format()
			
			cell_format_cp.set_bg_color(colorpalette[cpcouner+j])
			if int(colorpalette[cpcouner+j][1:],16) >=0x7FFFFF:
				cell_format_cp.set_font_color('black')
			if int(colorpalette[cpcouner+j][1:],16) <0x7FFFFF:
				cell_format_cp.set_font_color('white')
			worksheet.write(row, 2+j, colorpalette[cpcouner+j], cell_format_cp)
		
		cpcouner+=3
		
		row+=1
	
	try:
		workbook.close()
		print "Report completed"
	except:
		print "Close the opened report and run the script again"

	
printReport()