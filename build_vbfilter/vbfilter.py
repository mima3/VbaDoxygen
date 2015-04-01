#!/usr/bin/env python
# -*- coding: utf-8 -*-	
#
# This is a filter to convert Visual Basic v6.0 code
# into something doxygen can understand.
# Copyright (C) 2005  Basti Grembowietz
# 
# This program is free software; you can redistribute it and/or
# modify it under the terms of the GNU General Public License
# as published by the Free Software Foundation; either version 2
# of the License, or (at your option) any later version.
# 
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
# 
# You should have received a copy of the GNU General Public License
# along with this program; if not, write to the Free Software
# Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
# ------------------------------------------------------------------------- 
#
# This filter depends on following circumstances:
# in VB-code,
#  '! comments get converted to doxygen-class-comments (comments to a class)
#  '* comments get converted to doxygen-comments (function, sub etc)
#
#
# v0.1 - 2004-12-25
#  initial work
# v0.2 - 2004-12-30
#  added states
# v0.3 - 2004-12-31
#  removed states =)
# v0.4 - 2005-01-01
#  added class-comments
# v0.5 - 2005-01-03
#  changed default behaviour from "private" to "public"
#  + fixed re_doxy (whitespace now does not matter anymore)
#  + fixed re_sub and re_func (brackets inside brackets ok now)
# v0.6 - 2005-02-14
#  minor changes
# v0.7 - 2005-02-23
#  refactoring: removed double code.
#  + VB-Types are enabled now
#  + Doxygen-Comments can also start in the line of the feature which should be documented
# v0.8 - 2005-02-25
#  changed command line switches: now the usage is just "vbfilter.py filename".
# v0.9 - 2005-03-09
#  added handling of friends in vb.
# v0.10 - 2005-04-14
#  added handling of Property Let and Set
#  added recognition of default-values for parameters
# v0.11 - 2005-05-05
#  fixed handling of Property Get ( instead of Set ... )
# ========================================================================= 
# 2008/2/26 modified by Ryo Satsuki
#  modified handling of variable (Const, initial value, array)
#  modified handling of Function for Variant-return-function
#  added handling of End Function/Sub
#  added handling of Enum
#  added handling of blank line to keep comment block separation
# 2008/2/28 modified by Ryo Satsuki
#  modified handling of Function / Sub so as to format args
#  added handling of multiple divided lines
# 2008/4/9 modified by Ryo Satsuki
#  modified handling of comment for "'" in strings
#  added handling of a double quotation marks in a strings
#  modified handling of initial values so as to pass expressions
# 2008/8/27 modified by Ryo Satsuki
#  corrected handling of property procedure
#  modified handling of Sub so as to handle Property Set procedure
#  modified handling of Enum
import getopt          # get command-line options
import os.path         # getting extension from file
import string          # string manipulation
import sys             # output and stuff
import re              # for regular expressions

## stream to write output to
outfile = sys.stdout

# VB source encoding (added by R.S.)
src_encoding = "cp932"

# regular expression
## re to strip comments from file (modified by R.S.)
re_comments   = re.compile(r"((?:\"(?:[^\"]|\"\")*\"|[^\"'])*)'.*")
re_VB_Name    = re.compile(r"\s*Attribute\s+VB_Name\s+=\s+\"(\w+)\"", re.I)

## re to blank line (added by R.S.)
re_blank_line = re.compile(r"^\s*$")

## re to search doxygen-class-comments (modified by R.S.)
re_doxy_class = re.compile(r"(?:\"(?:[^\"]|\"\")*\"|[^\"'])*'!(.*)")
## re to search doxygen-comments (modified by R.S.)
re_doxy       = re.compile(r"(?:\"(?:[^\"]|\"\")*\"|[^\"'])*'\*(.*)")
## re to search for global variables members (used in bas-files)
re_globals    = re.compile(r"\s*Global\s+(Const\s+)?([^']+)", re.I)
## re to search for class-members (used in cls-files) (modified by R.S.)
re_members    = re.compile(r"\s*(Public\s+|Friend\s+|Private\s+|Static\s+){0,1}(?:(Const\s+)?(?:Dim\s+)?([\w]+(?:\([\w\s\(\)\+\-\*/\.]*\))?)\s+As\s+(\w+)\s*(?:=\s*(\"(?:[^\"]|\"\")*\"|[^']+))?|(?:Const\s+([\w\(\)]+)\s+=\s*(\"(?:[^\"]|\"\")*\"|[^']+)))", re.I)
re_array      = re.compile(r"([\w]+)\(([\w\s\(\)\+\-\*/\.]*)\)", re.I)
re_const_string	= re.compile(r"\"(?:[^\"]|\"\")*\"")
re_backslash	= re.compile(ur"\\")
re_doublequote	= re.compile(ur"(?=.)\"\"(?=.)")
## re to search Subs (modified by R.S.)
re_sub        = re.compile(r"\s*(Public\s+|Friend\s+|Private\s+|Static\s+){0,1}(Sub|Property\s+(?:Let|Set))\s+(\w+)\s*(\([\w\s=,\(\)\+\-\*/\.\"]*\))", re.I)
re_endSub  	  = re.compile(r"End\s+(?:Sub|Property)", re.I)
## re to search Functions (modified by R.S.)
re_function = re.compile(r"\s*(Public\s+|Friend\s+|Private\s+|Static\s+){0,1}(Function|Property\s+Get)\s+(\w+)\s*(\([\w\s=,\(\)\+\-\*/\.\"]*\))(?:\s+As\s+(\w+))?", re.I)
re_endFunction = re.compile(r"End\s+(?:Function|Property)", re.I)
## re to search args (added by R.S.)
re_arg      = re.compile(r"\s*(Optional\s+)?((?:ByVal\s+|ByRef\s+)?(?:ParamArray\s+)?)(\w+)(\(\s*\))?(?:\s+As\s+(\w+))?(?:\s*=\s*(\"(?:[^\"]|\"\")*\"|[^,\)]+))?", re.I)
## re to search for type-statements
re_type     = re.compile(r"\s*(Public\s+|Friend\s+|Private\s+|Static\s+){0,1}Type\s+(\w+)", re.I)
## re to search for type-statements
re_endType  = re.compile(r"End\s+Type", re.I)
## re to search for enum  (added by R.S.)
re_enum		= re.compile(r"\s*(Public\s+|Friend\s+|Private\s+|Static\s+){0,1}Enum\s+(\w+)", re.I)
re_endEnum  = re.compile(r"End\s+Enum", re.I)

# default "level" (private / public / protected) to take when not specified
def_level = "public:"

# strips vb-style comments from string
def strip_comments(str):
	global re_comments
	my_match = re_comments.match(str)
	if my_match is not None:
		return my_match.group(1)
	else:
		return str

# dumps the given file
def dump(filename):
	f = open(filename)
	r = f.readlines()
	f.close()
	for s in r:
		sys.stdout.write("."), 
		sys.stdout.write(s)

def processGlobalComments(r):
	global re_doxy_class
	# we have to look for global comments first!
	# they start with '!
	for s in r:
		gcom = re_doxy_class.match(s)
		if gcom is not None:
			# found global comment
			if (gcom.group(1) is not None):
				# write this comment to file
				outfile.write("/// " + gcom.group(1) + "\n")

def processClassName(r):
	global re_VB_Name
	sys.stderr.write("Searching for classname... ") 
	className = "dummy"
	for s in r:
		# now search for a class name
		cname = re_VB_Name.match(strip_comments(s))
		if cname is not None:
			# ok, className is found, so save it...
			sys.stderr.write("found! ") 
			className = cname.group(1)
			# ...and leave searching-loop
			break
	# ok, so let's start writing the pseudo-class
	sys.stderr.write(" using " + className + "\n") 
	outfile.write("\nclass " + className + "\n{\n") 

# pass blank lines to keep comment block separation
# added by R.S.
def checkBlankLine(s):
	global re_blank_line
	blank_line = re_blank_line.match(s)
	if (blank_line is not None):
		outfile.write("\n")

def checkDoxyComment(s):
	global re_doxy
	doxy = re_doxy.match(s)
	if (doxy is not None):
	# a comment was found... so write it
		outfile.write("/// " + doxy.group(1) + "\n")
		# 2005-01-03 : do not continue -> member-comments can now be in the same line as members
		#continue # and go to next line in source

# modified by R.S. for const, dim, array, initial value, and so on.
def foundMember(s):
	global re_members
	global re_array
	global re_const_string
	global re_backslash
	global re_doublequote
	global src_encoding
	member = re_members.match(strip_comments(s))
	if (member is not None):
		if (member.group(6) is not None):
			#	typeless const declaretion
			initval_str = ""
			if (member.group(7) is not None):
				if (re_const_string.match(member.group(7))):
					initval_str = u" = " + re_doublequote.sub(ur"\\\"", re_backslash.sub(ur"\\\\", unicode(member.group(7), src_encoding)))
					initval_str = initval_str.encode(src_encoding)
				else:
					initval_str = " = " + member.group(7)
			res_str = getAccessibility(member.group(1)) + " const " + member.group(6) + initval_str + ";"
		else:
			#	normal declaretion
			#	check const condition
			const_str = ""
			if (member.group(2) is not None):
				const_str = "const "
			#	check intial value
			initval_str = ""
			if (member.group(5) is not None):
				if (re_const_string.match(member.group(5))):
					initval_str = u" = " + re_doublequote.sub(ur"\\\"", re_backslash.sub(ur"\\\\", unicode(member.group(5), src_encoding)))
					initval_str = initval_str.encode(src_encoding)
				else:
					initval_str = " = " + member.group(5)
			#	check array
			valname_str = member.group(3)
			array_idfr = re_array.match(member.group(3))
			if (array_idfr is not None):
				valname_str = array_idfr.group(1) + "[" + array_idfr.group(2) + "]"
			#	produce resulting string
			res_str = getAccessibility(member.group(1)) + " " + const_str + " " + (member.group(4) or "") + " " + valname_str + initval_str + ";"
		# and deliver it
		outfile.write(res_str + "\n")
		return True
	else:
		return False

# added by R.S.
# modify arglist
def rearrangeArg(argstr):
	global re_const_string
	global re_backslash
	global re_doublequote
	global src_encoding
	# get type
	type_str = "Variant"
	if (argstr.group(5) is not None):
		type_str = argstr.group(5)
	# get arg name
	if (argstr.group(4) is not None):
		argname_str = argstr.group(3) + "[]"
	else:
		argname_str = argstr.group(3)
	# get default value
	dfltval_str = ""
	if ((argstr.group(1) is not None) and (argstr.group(6) is not None)):
		if (re_const_string.match(argstr.group(6))):
			dfltval_str = u" = " + re_doublequote.sub(ur"\\\"", re_backslash.sub(ur"\\\\", unicode(argstr.group(6), src_encoding)))
			dfltval_str = dfltval_str.encode(src_encoding)
		else:
			dfltval_str = " = " + argstr.group(6)
	return (argstr.group(1) or "") + " " +(argstr.group(2) or "") +" " + type_str + " " + argname_str + " " + dfltval_str

# modified by R.S. for variant type, and for scan inside function
def foundFunction(s):
	global re_function
	global re_arg
	s_func = re_function.match(strip_comments(s))	 # s_func == start_of_a_function
	if (s_func is not None):
		type_str = "Variant"
		if (s_func.group(5) is not None):
			type_str = s_func.group(5)
		# now make the resulting string
		# modified by R.S. to rearrange arglist
		res_str = getAccessibility(s_func.group(1)) + " " + type_str + " " + s_func.group(3) + re_arg.sub(rearrangeArg, s_func.group(4)) + "{"
		# and deliver this string
		outfile.write(res_str + "\n")
		return True
	else:
		return False

# added by R.S.	for scan inside function (now, only skip inside)
def processFunction(s):
	global re_endFunction
	vbEndFunction = re_endFunction.match(strip_comments(s))
	if (vbEndFunction is not None): # found End Function
		outfile.write("} \n") #write end of function
		return False
	else:
		# inside Sub
		return True

#  modified by R.S. for check inside sub
def foundSub(s):
	global def_level # for private/public/protected-issue
	global re_sub
	global re_arg
	s_sub = re_sub.match(strip_comments(s))
	if (s_sub is not None):
		#	produce resulting string
		# modified by R.S. to rearrange arglist
		res_str = getAccessibility(s_sub.group(1)) + " void " + s_sub.group(3) + re_arg.sub(rearrangeArg, s_sub.group(4))  + "{"
		# and deliver it
		outfile.write(res_str + "\n")
		return True
	else:
		return False

# added by R.S.	for scan inside sub (now, only skip inside)
def processSub(s):
	global re_endSub
	vbEndSub = re_endSub.match(strip_comments(s))
	if (vbEndSub is not None): # found End Sub
		outfile.write("} \n") #write end of function
		return False
	else:
		# inside Sub
		return True

def getAccessibility(s):
	accessibility = def_level
	if (s is not None):
		if (s.strip().lower() == "private"): accessibility = "private:"
		elif (s.strip().lower() == "public"): accessibility = "public:"
		elif (s.strip().lower() == "friend"): accessibility = "friend "
		elif (s.strip().lower() == "static"): accessibility = "static"
	return accessibility

# modified by R.S. for const, dim, array, initial value, and so on.
def foundMemberOfType(s):
	global re_members
	global re_array
	global re_const_string
	global re_backslash
	global re_doublequote
	global src_encoding
	member = re_members.match(strip_comments(s))
	if (member is not None):
		if (member.group(6) is not None):
			#	typeless const declaretion
			initval_str = ""
			if (member.group(7) is not None):
				if (re_const_string.match(member.group(7))):
					initval_str = u" = " + re_doublequote.sub(ur"\\\"", re_backslash.sub(ur"\\\\", unicode(member.group(7), src_encoding)))
					initval_str = initval_str.encode(src_encoding)
				else:
					initval_str = " = " + member.group(7)
			res_str = "const " + member.group(6) + initval_str + ";"
		else:
			#	normal declaretion
			#	check const condition
			const_str = ""
			if (member.group(2) is not None):
				const_str = "const "
			#	check intial value
			initval_str = ""
			if (member.group(5) is not None):
				if (re_const_string.match(member.group(5))):
					initval_str = u" = " + re_doublequote.sub(ur"\\\"", re_backslash.sub(ur"\\\\", unicode(member.group(5), src_encoding)))
					initval_str = initval_str.encode(src_encoding)
				else:
					initval_str = " = " + member.group(5)
			#	check array
			valname_str = member.group(3)
			array_idfr = re_array.match(member.group(3))
			if (array_idfr is not None):
				valname_str = array_idfr.group(1) + "[" + array_idfr.group(2) + "]"
			#	produce resulting string
			res_str = const_str + " " + (member.group(4) or "") + " " + valname_str + initval_str + ";"
		# and deliver it
		outfile.write(res_str + "\n")

def foundType(s):
	global re_type
	vbType = re_type.match(strip_comments(s))
	if (vbType is not None):
		#	produce resulting string
		res_str = getAccessibility(vbType.group(1)) + " struct " + vbType.group(2)  + " {"
		# and deliver it
		outfile.write(res_str + "\n")
		return True
	else:
		return False

def processType(s):
	global re_endType
	vbEndType = re_endType.match(strip_comments(s))
	if (vbEndType is not None): # found End Type
		outfile.write("}; \n") #write end of struct
		return False
	else:
		# match <var AS type>
		# write <type var;>
		foundMemberOfType(s)
		return True

# modified by R.S. for process enum
def foundEnum(s):
	global re_enum
	vbEnum = re_enum.match(strip_comments(s))
	if (vbEnum is not None):
		#	produce resulting string
		res_str = getAccessibility(vbEnum.group(1)) + " enum " + vbEnum.group(2)  + " {"
		# and deliver it
		outfile.write(res_str + "\n")
		return True
	else:
		return False

# modified by R.S. for process enum
def processEnum(s, notfirst = True):
	global re_endEnum
	vbEndEnum = re_endEnum.match(strip_comments(s))
	if (vbEndEnum is not None): # found End Enum
		outfile.write("}; \n") #write end of enum
		return False
	else:
		# inside enum
		#if notfirst:
		#	outfile.write(", " + strip_comments(s) + "\n")
		#else:
		#	outfile.write(strip_comments(s) + "\n")
		outfile.write(strip_comments(s) + ", " + "\n")
		return True

# filters .cls-files - VB-CLASS-FILES
def filterCLS(filename):
	global outfile ## get global variable
	global re_comments
	f = open(filename)
	r = f.readlines()
	f.close()
	outfile.write("\n// -- processed by [filterCLS] --\n") 

	processGlobalComments(r)

	processClassName(r)

	# now scan every line and look either for doxy-comments, members or functions/subs
	# searching for multiline-type-statements, we need a flag here:
	inTypeSearch = False

	# added by R.S. for scan inside function, sub, enum
	inSubSearch = False
	inFunctionSearch = False
	inEnumSearch = False
	notEnumFirst = False

	s = ""
	lineContinue = False
	for ln in r:
		# added by R.S. to connect multiple lines
		if lineContinue:
			s = s + ln
		else:
			s = ln
		if ((re_comments.match(s) is None) and (s[-3:] == " _\n")):
			s = s[:-2]
			lineContinue = True
			continue
		else:
			lineContinue = False
			
		# added by R.S. for pass blank lines to separate each comment block
		checkBlankLine(s)
		checkDoxyComment(s)

		if inTypeSearch:
			inTypeSearch = processType(s)
			continue

		# added by R.S. for scan inside sub
		if inSubSearch:
			inSubSearch = processSub(s)
			continue

		# added by R.S. for scan inside function
		if inFunctionSearch:
			inFunctionSearch = processFunction(s)
			continue

		# added by R.S. for scan inside enum
		if inEnumSearch:
			inEnumSearch = processEnum(s, notEnumFirst)
			notEnumFirst = True
			continue
		
		if foundType(s):
			inTypeSearch = True
			continue

		#	see if line contains a member
		if foundMember(s):
			continue # line could not contain anything more than a member

		# see if there is a function-statement
		if foundFunction(s):
			# added by R.S. for scan inside function
			inFunctionSearch = True
			continue

		# there was no match to a function - let's try a sub
		if foundSub(s):
			# added by R.S. for scan inside function
			inSubSearch = True
			continue

		# added by R.S.
		# see if there is an enum declaretion
		if foundEnum(s):
			inEnumSearch = True
			notEnumFirst = False
			continue

	outfile.write("}")  # for ending class
	outfile.write("\n// -- [/filterCLS] --\n") 


# filters .bas-files
def filterBAS(filename):
	global outfile ## get global variable
	global re_comments
	f = open(filename)
	r = f.readlines()
	outfile.write("\n// -- processed by [filterBAS] --\n") 

	processGlobalComments(r)

	processClassName(r)

	# now scan every line and look either for doxy-comments or functions/subs
	# or both
	
	# searching for multiline-type-statements, we need a flag here:
	inTypeSearch = False

	# added by R.S. for scan inside function, sub, enum
	inSubSearch = False
	inFunctionSearch = False
	inEnumSearch = False
	notEnumFirst = False

	s = ""
	lineContinue = False
	for ln in r:
		# added by R.S. to connect multiple lines
		if lineContinue:
			s = s + ln
		else:
			s = ln
		if ((re_comments.match(s) is None) and (s[-3:] == " _\n")):
			s = s[:-2]
			lineContinue = True
			continue
		else:
			lineContinue = False
			
		# added by R.S. for pass blank lines to separate each comment block
		checkBlankLine(s)
		checkDoxyComment(s)

		if inTypeSearch:
			inTypeSearch = processType(s)
			continue

		# added by R.S. for scan inside sub
		if inSubSearch:
			inSubSearch = processSub(s)
			continue

		# added by R.S. for scan inside function
		if inFunctionSearch:
			inFunctionSearch = processFunction(s)
			continue

		# added by R.S. for scan inside enum
		if inEnumSearch:
			inEnumSearch = processEnum(s, notEnumFirst)
			notEnumFirst = True
			continue
		
		if foundType(s):
			inTypeSearch = True
			continue

		#	see if line contains a member
		#	added by R.S. to proccess variables in BAS file
		if foundMember(s):
			continue # line could not contain anything more than a member

		# line is not a comment. 
		# see if there is a function-statement
		if foundFunction(s):
			# added by R.S. for scan inside function
			inFunctionSearch = True
			continue

		# there was no match to a function - let's try a sub
		if foundSub(s):
			# added by R.S. for scan inside function
			inSubSearch = True
			continue

		# added by R.S.
		# see if there is an enum declaretion
		if foundEnum(s):
			inEnumSearch = True
			notEnumFirst = False
			continue

	outfile.write("}")  # for ending class
	outfile.write("\n// -- [/filterBAS] --\n") 

## main filter-function ##
##
## this function decides whether the file is
## (*) a bas file  - module
## (*) a cls file  - class
## (*) a frm file  - frame
##
## and calls the appropriate function

def filter(filename, out=sys.stdout):
	global outfile
	outfile = out

	try:
		root, ext = os.path.splitext(filename)
		if (ext.lower() ==".bas"):
			## if it is a module call filterBAS
			filterBAS(filename)
		elif (ext.lower() ==".cls") or (ext.lower() == ".frm"):
			## if it is a class or frame call filterCLS
			filterCLS(filename)
		else:
			## if it is an unknown extension, just dump it
			dump(filename)

		sys.stderr.write("OK\n") 
	except IOError,e:
		sys.stderr.write(e[1]+"\n")

## main-entry ##
################

if len(sys.argv) != 2:
	print "usage: ", sys.argv[0], " filename"
	sys.exit(1)

# Filter the specified file and print the result to stdout
filename = sys.argv[1] 
filter(filename)
sys.exit(0)
