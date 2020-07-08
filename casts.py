# casts.py
#	version 0.1.2
#		now handles the minority of cast lists organized as html tables rather than html lists.
#		minor additions to genderwords.json
#		excludes a few more "nonstandard" lines that are not actor/ character pairs.
#			problem: rising % of "unknown" gender over time (from 2.5% -> 5.5% 1930-2018, US)
#			problem: rising % of mismatched actor and character genders over time (1.5% -> 2.5% US)
#			problem: no case distinctions e.g., chick-> female; Chick-> male.
#			problem: multiple characters for same actor; divided sometimes by /, sometimes indented lists
#	1. reads a series of text files of cast lists from wikipedia movie .html files and parses each line into 3 segments
#		- actor
#		- character
#		- other description
#	2. assigns genders to actor and character names 
#		(female, mostly_female, andy, mostly_male, male, unknown)
#		counts how often actor and character genders match, or partly match
#	eventually will:
#		3. count how often the character's name (first, last, both) appear in plot summary
#			eventually even attribute "he","she","her", etc to that character
#		4. associate each character with one or more "occupations"
#			codes some ~30,000 "jobtitles" into 2010 3-digit U.S. Census codes 
#
#	arg= prefix is a filename prefix to output (& input) files
#		e.g. python3 casts.py NYT
#
#	input:
#		prefix+files.txt (e.g., USfiles.txt)= a file of filenames of text files (cast lists) to be read and coded for gender.
#		gendernames.json	revisions of gender_guesser, "name": "gender",
#
#	to compile casts.py:
#		uses python standard packages: sys json
#		uses python packages that must be downloaded and installed: gender_guesser BeautifulSoup
#			e.g. pip install BeautifulSoup4
#			e.g. pip install gender-guesser  (see:  https://pypi.org/project/gender-guesser/)
#
#	

import json
import sys
import gender_guesser.detector as gender
g = gender.Detector(case_sensitive=False)

# should be called with one argument, a prefix (e.g., US, Fr, SK, etc.)
# 	there should be one and only one argument in calling Python
Nargs= len(sys.argv)
if Nargs<2:
	print("You need to call casts.py with an argument giving the filename prefix")
	print("	e.g., python3 casts.py movies")
	print("	exiting...")
	sys.exit()
else:
	prefix= sys.argv[1]
	print("prefix: ", prefix)

# to strip html code from text:
from bs4 import BeautifulSoup

#
# Korean surnames that would be coded as English first names:
notgivenname= ['Gil', 'Kim', 'Lee', 'Park', 'Choi', 'Jung', 'Kang', 'Cho', 'Jo', 'Yoon', 
               'Jang', 'Iim', 'Han', 'Oh', 'Seo', 'Shin', 'Kwon', 'Hwang', 'Ahn', 'Song', 
               'Jeon', 'Hong', 'Yu', 'Ko', 'Moon', 'Yang', 'Bae', 'Baek', 'Heo', 'Yoo',
               'Nam', 'Sim', 'Rho', 'Ha', 'Kwak', 'Sung', 'Cha', 'Choo', 'Joo', 'Wu','Koo',
                'Min', 'Ku', 'Na', 'Keum', 'Chae', 'Chun', 'Bang', 'Kong', 'Yeon', 'Yeo',
               'On', 'Lm', 'Gi','Chu','Do','Seong','Seon','Shim','Ji','Bok','Bong','Goo',
              'Kil','Noh','Pyo','Ra','Ri','Ro','Uhm']
#
# first words in some nonstandard lines that will be dropped:
nonstandard= ['cast', '(cast', 'actor', 'additional', 'additionally', 'cameo', 'cameos', 'casting', 'categories', 'character', 'featuring', 'list', 'main', 'minor', 'notes', 'opening', 'other', 'others', 'retrieved', 'role', 'starring', 'supporting', 'unbilled', 'uncredited', 'voice', 'voices', '@media', '[note']
nonstandard2= ['The cast', 'The film', 'As appearing', 'See also']
uncredited= ['unbilled', 'uncredited']
cameos= ['cameo', 'cameos']
minor= ['additional', 'additionally', 'minor', 'other', 'others', 'supporting'] 
#
# most cast lists separate actor and character with " as "
#	a few cast lists separate actor and character with 
#		" ... " or more periods, 
# 		" - " dash with spaces around (not a hyphen)
#		"â" non-ascii
#		"–" endash
# 		" ... " or more periods
#	actorend is a dict with separation marker and # spaces to skip to beginning of character name
actorend =	{
	" as ": 4,
	" - ": 3, 
	"---": 3, 
	"...": 4, 
	":": 1, 
	"â": 6, 
	"—": 1, 
	"–": 1 
	}
# charend is a list of separation markers between the character name and additional material on that line
charend= [",", ":", ";", "–", "â", " - ", "(", " who " ]	
xtab= [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
####################################
# 3 output files:
#
# name-level summary: after proccessing all casts, list each name found & its counts
namesfile= prefix + "names.xls"	
namesf= open(namesfile, "w")
namesf.write ("iname	name	gender	Nfound	Nactor	Nchar	Nfemale	Nmfemale	Nandy	Nmmale	Nmale	Nunknown	PairScore	PctFemale	PctMale\n")
#
# detailed output: XXcharacters.xls= after processing each text file, writes one record for each character
#	writes as tab delimited file for input into spreadsheet/ stats program
charfile= prefix + "characters.xls"	
charf= open(charfile, "w")
charf.write ("filename" + "	status" + "	AName" + "	Actor" + "	AGender" + "	CName" + "	Character" + "	ChGender" + "	Final" + "	additional" + "\n" )
#
# nonstandard lines have no actor as character... OR actor - character...
#	print to separate file to figure out what is wrong
nonstandardfile= prefix + "nonstandard.xls"	
nonstandardf= open(nonstandardfile, "w")
nonstandardf.write ("filename" + "	" + "nonstandard line:\n")
#
# text-level summary: after processing each text, writes a count of #words, #sentences, etc., one line per text
#	need to check for Nchars=0 or =1;  that movie not processed correctly
textsfile= prefix + "casts.xls"	
textsf= open(textsfile, "w")
textsf.write ("Nlines	Nblanks	Nread	Nwords	Nactors	Nchars	Nmiss	Nextra	Nnonstandard	movie\n")

####################################
# 2 input files:
#
# genderwords.json a file of supplemental (corrected) names,titles,words each with assigned gender:
# 	Queen, Princess, King, Prince, -- already in gender_guesser
genderwords= open("genderwords.json").read()
# genderwords is a long string, make genderwords into a dict, namegender:
namegender= json.loads(genderwords)
print  ('\nnamegender (names with fixed gender)= ' + str(type(namegender)) + ' Nlines=' + str(len(namegender)) )
#
#
# XXfiles.txt: list of filenames from Wikipedia cast lists that casts.py will read
#	read list of filenames of input texts: all movies with cast lists
infilelist= prefix + "files.txt"
files=open(infilelist).read()
#	files is a long string with all file names:
#	number of files:
textfiles=files.split('\n')
Ntextfiles=len(textfiles)
Sumtextlines=0
#	strip off end of file line:
textfiles=textfiles[0:Ntextfiles-1]
print ('\nList of cast files to be analyzed: ' + infilelist + " " + str(type(textfiles)) + ", #files= " + str(len(textfiles)) )
#

foundnames= [""]
gendernames= ["unknown"]
countnames= [0]
countanames= [0]
countcnames= [0]
countmales= [0]
countmmales= [0]
countmfemales= [0]
countfemales= [0]
countandy= [0]
countunknown= [0]
Sumcharacters=0

for file in textfiles:
	Nactors=0
	Ncharacters=0
	Nnochar=0
	Nadditional=0
	Nnonstandard=0
	Nblanklines=0
	# status records the prominence of the actor / character
	#	eventually will be revised used "starring" from the infobox on each page
	#		starring vs. others may be the most important distinction
	#	another prominence indicator might be the order in the list (iactor)
	#	in the cast lists, wikipages are inconsistent with indicating prominence of the role; 
	#		some casts ignore seconday or uncredited roles; 
	#		others label them as uncredited etc.; 
	#		others include them with main actors undifferentiated.
	#	non-main characters are indicated in two ways:
	#		a separate line (e.g., Uncredited:) below which all the roles are uncredited (statustype=list)
	#		or on the line of the actor / character only (statustype=line)
	status="main"
	statustype="statuslist"
	# write blank line to character.xls file to separate each movie:
	charf.write ("\n")
	# todo: work with different codings
	#text0=open(file, encoding = "ISO-8859-1").read()
	text0=open(file).read()
	# drop html:
	#	first, to accodomdate the few movies that use tables instead of lists:
	text1=text0.replace('\n</td>\n','</td> : ')
	text1=text1.replace('</td>\n','</td> : ')
	text1= BeautifulSoup(text1, features="html.parser")
	textstring= text1.get_text()
	Nbytes=len(textstring)
	Nwords=textstring.count(" ")
	Nlines=textstring.count("\n")
	textstring= textstring.replace('/',' / ')	# often a marker between segments or multiple roles
	textstring= textstring.replace('[',' [')	# wikipedia footnotes: make it a separate word
	textstring= textstring.replace(':',' : ')	# wikipedia footnotes: make it a separate word
	textstring= textstring.replace('	', ' ') # tabs create problems for casts.py
	textstring= textstring.replace('  ',' ')
	#print ("\nafter BeautifulSoup: " + file + ":\n" + str(textstring) + "\nNlines= " + str(Nlines) + "\nNwords= " + str(Nwords))
	#
	# make each sentence a new string
	textlines= textstring.splitlines()
	Ntextlines= len(textlines)
	Sumtextlines= Sumtextlines + Ntextlines
	# process each line in the file:
	#
	iline=0
	Nblanklines=0
	for line in textlines:
		line= line.strip()
		line= line.replace('	', ' ')
		line= line.replace('  ',' ')
		words= line.split(" ")
		Nwordline= line.count(" ") + 1
		if statustype=="statusline":	# status changed just for previous line using "additional"
			status="main"
			statustype="statuslist"		# if statustype="statuslist" will keep whatever previous status
		if Nwordline>=1:
			wordtest= words[0].lower()
		if Nwordline>=2:
			words2= words[0] + " " + words[1]
		else:
			words2= ""
		lenline= len(line)
		actor=""
		char=""
		aname=""
		additional=""
		if lenline<=1 :
			Nblanklines= Nblanklines+1
		elif words2 in nonstandard2:
			Nnonstandard= Nnonstandard+1
			nonstandardf.write (file + "	common nonstandard2 line:	" + str(line) + "\n" )
		elif wordtest in nonstandard:
			Nnonstandard= Nnonstandard+1
			nonstandardf.write (file + "	common nonstandard line:	" + str(line) + "\n" )
			if wordtest in uncredited:
				status="uncredited"
				statustype="statuslist"
				#print (file, words[0],line)
			elif wordtest in minor:
				status="supporting"
				statustype="statuslist"
			elif wordtest in cameos:
				status="cameo"
				statustype="statuslist"
			elif wordtest == "main":
				status="main"
				statustype="statuslist"
		#elif words[0[0:-1]] in nonstandard:	# sometimes there is a colon, comma, or other punctuation after the nonstandard first word.
			#Nnonstandard= Nnonstandard+1
			#nonstandardf.write (file + "	common nonstandard line:	" + str(line) + "\n" )
		else:	# not a blank line and not a common nonstandard cast line
			iline= iline+1
			#
			# find the separation marker between actor name and character name segments
			# 	endactor= the position in the line where it is found
			#		if more than one separation marker, use the first one
			# 	startchar= value in dict to skip spaces to start of character name
			#
			endactor=lenline	# initialize endactor as the end of the line
			startchar=0
			for testmark in actorend:		# actorend is a dict with the possible separation markers
				endtest= line.find(testmark) # endtest is the position of testmark if testmark is found
				if endtest != -1:			# if testmark was not found in the line, find() returns -1
					if endtest < endactor:	# only if this marker occurred before any other marker or lenline
						endactor=endtest	# the position where the marker was found
						startchar= actorend[testmark]	# how many bytes to skip to start the next segment, character name
			# test if no marker found 
			#	this does not work if the cast is stored as a table; then there are no segment markers (maybe <td> but html is gone?)
			if endactor==lenline: 	
				#no actor/character segment marker found
				if Nwordline>=6:
					# too many words in line without actor/character segment makres
					# 	can't figure out what this line is; save in file of nonstandard lines to examine later
					#	keep actor and char as empty strings ""	
					#	go to next line
					nonstandardf.write (file + "	longline; endline=" + str(lenline) + "	" + line + "\n")
					Nnonstandard= Nnonstandard+1
				else: 
					#	probably only an actor listed with no character
					actor= line.strip()
					char="missing"
					cname="missing"
					cgender="missing"
					Nnochar=Nnochar+1
			else: 
				# found an actor/character segment marker
				if endactor>80:
					# segment marker is too deep into line; skip and write to nonstandard line file
					nonstandardf.write (file + "	longact; endactor=" + str(endactor) + "	" + line + "\n")
				else:
					actor=line[:endactor].strip()
					Nactors= Nactors+1
					#
					# next, find end of character name marker that separates character name segment from additional material.:
					#	forward slash usually means multiple characters for the same actor; deal with that later.
					#	character name segment ends with a , : , end of line, or some distinctive character
					#		begin find() only after actor's name since a few actor name segments have a comma ( , Jr. )
					startchar=endactor+startchar
					Nwordchar= line[startchar:].count(" ")
					endchar=lenline		# initialize end of character segment to end of line
					for testmark in charend:
						endtest= line[startchar:].find(testmark)
						if endtest != -1:
							if endtest < endchar:
								endchar=endtest
					#
					# sometimes a period separates the character from additional text.  This is more of a problem.
					#	it could be an abbreviation for a title (e.g., Mrs. Jr. Gen. etc.) and not a separation marker.
					#	or it could be a character's initials (e.g., J.R.  W.C.) and not a separation
					#endperiod= line[startchar:].find(".")
					#if endperiod!= -1:	# found a period because -1 == not found 
						#if endperiod < endchar	# if period is before another separation character, check whether the period is not a separation
					if endchar==lenline:	# no comma, colon, etc., probably no additional material
						if Nwordchar>=8 :
							# more than 8 words describing character (even if no , or :)
							nonstandardf.write (file + "	char 8words+:	" + str(line) + "\n" )
							Nnonstandard= Nnonstandard+1
						else:
							char= line[startchar:].strip()	# store character segment in char even if 8+ words.
							Ncharacters= Ncharacters+1
					else:
						# line has a comma, colon, semicolon, etc. segment marker between character and additional material
						char= line[startchar:endchar+startchar].strip()
						Ncharacters= Ncharacters+1
						if line[endchar+startchar]=="(":	# unlike other segment markers, keep the left paren in the additional string
							startchar= startchar-1
						additional= line[1+endchar+startchar:].strip()	# store additional material segment in additional
						chkadd= additional.lower()
						Nadditional= Nadditional+1
						if chkadd.find("uncredited")!=-1:
							status="uncredited"
							statustype="statusline"
						if chkadd.find("cameo")!=-1:
							status="cameo"
							statustype="statusline"
			#
			# line is now separated into 3 segments: full actor name, full character name (if any), additional words (if any)
			# 	fix some problems with markers that extend into char or additional
			lenchar= len(char)
			if lenchar>0:
				if char[-1]==".":
					char= char[0:-1]
			lenchar= len(char)
			if lenchar>0:
				for ichar in range(lenchar):
					if char[0]=="-" or char[0]==".":
						char= char[1:]
					else:
						break
			lenadditional= len(additional)
			if lenadditional>0:
				if additional[-1]==".":
					additional= additional[0:-1]
			lenadditional= len(additional)
			if lenadditional>0:
				for iadditional in range(lenadditional):
					if additional[0]=="-" or additional[0]==".":
						additional= additional[1:]
					else:
						break
			#
			# loop through the words in anames[] to find first gendered name (e.g., Mary), title (e.g. Mrs.), or word (e.g., mother)
			if actor!="":
				actor= actor.replace('	',' ')	# rarely a tab separates words in names, but tabs cause problems so replace with a space
				actor= actor.replace('"',' ')
				actor= actor.strip()
				anames= actor.split(" ")	# this will unfortunately separate some compound names
				if prefix=="SK":
					anames.reverse()
				agender= "unknown"			# initialize as unknown; loop will continue through anames[] until agender is no longer unknown
				for aname in anames:
					# some common surnames in Korean are given names in English;
					if prefix=="SK" and aname in notgivenname:
						agender= "unknown"
					elif len(aname)>0 and aname[0]=="[":	# wiki footnote usually at end of char section
						agender= "unknown"
					elif len(aname)>=2 and aname[-2]=="'" and aname[-1]=="s":	# possessive: skip 
						agender= "unknown"
					elif len(aname)>=2 and aname[-2]=="s" and aname[-1]=="'":	# plural possessive
						agender= "unknown"
					else:
						aname= aname.replace("'"," ")
							# delete quotes around nicknames, e.g., 'Doc', 
							# but would also replace possessives, so possessives skipped above
						# then check if aname is in revision dict, namegender{}, above:
						agender= namegender.get(aname.lower())
						if agender==None:	# aname not found in updated namegender json file
							# last place to look is in gender_guesser
							agender= g.get_gender(aname)
						if agender!= "unknown": 	#if gender is unknown, then continue for loop & go to next word; ?"andy"?
								break
					# end of going through words in anames
				if agender=="unknown":	# if at end of for loop, all anames are unknown gender, then use the first word
						# ? but maybe the second if the first is just a title or a stopword like "the"?
						aname=anames[0]
				# check whether this aname has already been found: 
				if aname in foundnames:
					# aname has been found before, so just increment count
					iname= foundnames.index(aname)
					countnames[iname]=  countnames[iname]+1
					countanames[iname]= countanames[iname]+1
				else:
					# aname has not been found before, so add to list of names found
					foundnames.append(aname)
					iname= len(foundnames)-1
					countnames.append(1)
					countanames.append(1)
					countcnames.append(0)
					gendernames.append(agender)
					countmales.append(0)
					countmmales.append(0)
					countandy.append(0)
					countmfemales.append(0)
					countfemales.append(0)
					countunknown.append(0)
				ianame= iname
				# 
				# loop through the words in cnames[] to find first gendered name | title
				if char =="":
					cgender="unknown"
					cname="missing"
					icname=0
				else:
					char= char.replace('	',' ')	# rarely a tab separates words in names
					char= char.replace('"',' ')	
					char= char.strip()
					cnames= char.split(" ")		# this will separate some compound names
					cgender= "nochar"
					if prefix=="SK":
						cnames.reverse()
					for cname in cnames:
						if prefix=="SK" and cname in notgivenname:
							cgender= "unknown"
						elif len(cname)>0 and cname[0]=="[":	# wiki footnote usually at end of char section
							cgender= "unknown"
						elif len(cname)>=2 and cname[-2]=="'" and cname[-1]=="s":	# possessive: skip 
							cgender= "unknown"
						elif len(cname)>=2 and cname[-2]=="s" and cname[-1]=="'":	# plural possessive
							cgender= "unknown"
						else:
							cname= cname.replace("'"," ")	
								# delete quotes around nicknames, e.g., 'Doc', 
								# but would also replace possessives, so possessives skipped above
							# then check if cname is in revision dict, namegender{}, above:
							cgender= namegender.get(cname.lower())
							if cgender==None:	# cname not found in updated namegender json file
								# last place to look is in gender_guesser
								cgender= g.get_gender(cname)
							if cgender!= "unknown": 	#if gender is unknown, then continue for loop & go to next word; ?"andy"?
								break
						# end of loop for finding gender of words in cnames
					if cgender=="unknown":	# if all cnames are unknown gender, then use the first word
						# ? but maybe the second if the first is just a title or stopword?
						cname=cnames[0]
					# check whether this cname has already been found: 
					if cname in foundnames:
						# cname has been found before, so just increment count
						iname= foundnames.index(cname)
						countnames[iname]=  countnames[iname]+1
						countcnames[iname]= countcnames[iname]+1
					else:
						# cname has not been found before, so add to list of names found
						foundnames.append(cname)
						iname= len(foundnames)-1
						countnames.append(1)
						countcnames.append(1)
						countanames.append(0)
						gendernames.append(cgender)
						countmales.append(0)
						countmmales.append(0)
						countandy.append(0)
						countmfemales.append(0)
						countfemales.append(0)
						countunknown.append(0)
					icname= iname
				#
				# increment xtab counts of actor genders for this character name
				if agender=="female":
					countfemales[icname]= countfemales[icname]+1
					xtabcell=0
				elif agender=="mostly_female":
					countmfemales[icname]= countmfemales[icname]+1
					xtabcell=7
				elif agender=="andy":
					countandy[icname]= countandy[icname]+1
					xtabcell=14
				elif agender=="mostly_male":
					countmmales[icname]= countmmales[icname]+1
					xtabcell=21
				elif agender=="male":
					countmales[icname]= countmales[icname]+1
					xtabcell=28
				elif agender=="unknown":
					countunknown[icname]= countunknown[icname]+1
					xtabcell=35
				else:
					xtabcell=42	# no actor
				# increment counts of character genders for this actor name
				if cgender=="female":
					countfemales[ianame]= countfemales[ianame]+1
					#xtabcell= xtabcell+0
				elif cgender=="mostly_female":
					countmfemales[ianame]= countmfemales[ianame]+1
					xtabcell= xtabcell+1
				elif cgender=="andy":
					countandy[ianame]= countandy[ianame]+1
					xtabcell= xtabcell+2
				elif cgender=="mostly_male":
					countmmales[ianame]= countmmales[ianame]+1
					xtabcell= xtabcell+3
				elif cgender=="male":
					countmales[ianame]= countmales[ianame]+1
					xtabcell= xtabcell+4
				elif cgender=="unknown":
					xtabcell= xtabcell+5
					countunknown[ianame]= countunknown[ianame]+1
				else:
					xtabcell= xtabcell+6
				xtab[xtabcell]= xtab[xtabcell]+1
				#
				# assign this actor/character line a best guess gender
				finalgender=""
				if agender=="female" or agender=="mostly_female":
					if cgender=="female" or cgender=="mostly_female" or cgender=="andy":
						finalgender="female"
					elif cgender=="unknown" or cgender=="missing":
						finalgender="likely_female"
					elif cgender=="male" or cgender=="mostly_male":
						finalgender="mismatch"
				elif agender=="male" or agender=="mostly_male":
					if cgender=="male" or cgender=="mostly_male" or cgender=="andy":
						finalgender="male"
					elif cgender=="unknown" or cgender=="missing":
						finalgender="likely_male"
					elif cgender=="female" or cgender=="mostly_female":
						finalgender="mismatch"
				elif agender=="andy":
					if cgender=="female" or cgender=="mostly_female":
						finalgender="female"
					elif cgender=="unknown" or cgender=="missing" or cgender=="andy":
						finalgender="unknown"
					elif cgender=="male" or cgender=="mostly_male":
						finalgender="male"
				elif agender=="unknown" or agender=="missing":
					if cgender=="female" or cgender=="mostly_female":
						finalgender="likely_female"
					elif cgender=="unknown" or cgender=="missing" or cgender=="andy":
						finalgender="unknown"
					elif cgender=="male" or cgender=="mostly_male":
						finalgender="likely_male"
				else:
					finalgender="unknown"
				charf.write (file + "	" + status + "	" + aname + "	" + actor + "	"  + agender + "	" + cname + "	" + char + "	" +cgender + "	" + finalgender + "	" + additional + "\n")
	# end of file loop through all lines in a file
	#Nnochar= Nactors-Ncharacters
	textsf.write (str(Nlines) + "	" + str(Nblanklines) + "	" + str(iline) + "	" + str(Nwords) + "	" + str(Nactors) + "	" + str(Ncharacters) + "	" + str(Nnochar) + "	" + str(Nadditional)+ "	" + str(Nnonstandard) + "	" + file + "\n")
	print ("Nlines=" + str(Nlines) + "	Nblanks=" + str(Nblanklines) + "	iline=" + str(iline) + " 	#words=" + str(Nwords) + "	#actors=" + str(Nactors) + "	#characters=" + str(Ncharacters) + "	#missing="  + str(Nnochar) + "	#additnl=" + str(Nadditional)+ "	#nonstandard lines=" + str(Nnonstandard) + "	"+ file )
	Sumcharacters= Sumcharacters+Ncharacters
print ("# movies(files)=" + str(len(textfiles)))
Sumnames= len(foundnames)
for iname in range(Sumnames):
	# pairscore is summary measure of whether the paired name (e.g. the character's when found name is the actor's) is female (negative scores=male)
	pairscore=2*countfemales[iname]+countmfemales[iname]-countmmales[iname]-2*countmales[iname]
	# pctfemale is % of foundnames that are "female" or "mostly_female";
	#	can be used to identify errors in male names
	if countnames[iname]>0:
		pctfemale=100*(countfemales[iname]+countmfemales[iname])/countnames[iname]
	else:
		pctfemale=0
	# pctmale is % of foundnames that are "male" or "mostly_male";
	#	can be used to identify errors in female names
	#	pctmale + pctfemale <=100 because andy, unknown
	if countnames[iname]>0:
		pctmale=100*(countmales[iname]+countmmales[iname])/countnames[iname]
	else:
		pctfemale=0
	lineout= str(iname) +"	"+ foundnames[iname] +"	"+ gendernames[iname] +"	"+ str(countnames[iname]) +"	"+ str(countanames[iname]) +"	"+ str(countcnames[iname]) +"	"+ str(countfemales[iname]) +"	"+ str(countmfemales[iname]) +"	"+ str(countandy[iname]) +"	"+ str(countmmales[iname]) +"	"+ str(countmales[iname]) +"	"+ str(countunknown[iname]) + "	" + str(pairscore) + "\n" 
	namesf.write (lineout)
print ("\n# unique names=", str(len(foundnames)))
xls= prefix+"names.xls;"
print ("See", xls, "sort on gender, pairscore to revise unknown and andy")
print ("See", xls, "sort on gender, pctmale or gender, pctfemale to check male&female errors\n")
#
print (str(Sumcharacters), "movie actor/character pairs out of", str(Sumtextlines), "lines in the", str(len(textfiles)), "cast files.")
#	crosstabs of actor gender X character gender
coltotal= [0,0,0,0,0,0,0]
print ("\n 	character gender:")
print ("actor:	female	mfemale	andy	male	mmale	unknown	miss	rowtotal")
rowlabel= ["female", "mfemale", "andy", "mmale", "male", "unknown", "miss", "rowtotal"]
for row in range(0,49,7):
	rowtotal= xtab[row+0] + xtab[row+1] + xtab[row+2] + xtab[row+3] + xtab[row+4] + xtab[row+5] + xtab[row+6]
	coltotal[0]= coltotal[0] + xtab[row+0]
	coltotal[1]= coltotal[1] + xtab[row+1]
	coltotal[2]= coltotal[2] + xtab[row+2]
	coltotal[3]= coltotal[3] + xtab[row+3]
	coltotal[4]= coltotal[4] + xtab[row+4]
	coltotal[5]= coltotal[5] + xtab[row+5]
	coltotal[6]= coltotal[6] + xtab[row+6]
	rowint=int(row/7)
	print (rowlabel[rowint], str(xtab[row+0]), str(xtab[row+1]), str(xtab[row+2]), str(xtab[row+3]), str(xtab[row+4]), str(xtab[row+5]), str(xtab[row+6]), str(rowtotal), sep="	")
total= coltotal[0] + coltotal[1] + coltotal[2] + coltotal[3] + coltotal[4] + coltotal[5] + coltotal[6]
print ("\ntotals", str(coltotal[0]), str(coltotal[1]), str(coltotal[2]), str(coltotal[3]), str(coltotal[4]), str(coltotal[5]), str(coltotal[6]), str(total), sep="	")
#
xls= prefix+"characters.xls"
print ("\nSee", xls, "for each movie character.")
print ("	Sort by Cname Chgender Agender Aname (or by Aname Agender Chgender Cname) to identify misclassified names.")
print ("\n	exiting...")
sys.exit()

