#!/usr/local/bin/python
# Required header that tells the browser how to render the text.
print "Content-Type: text/plain\n\n"
import string
import os
from os import environ, path
from types import IntType, ListType, StringType, TupleType
import cgi


WNHOME = environ.get('WNHOME', {'mac': ":", 'dos': "C:\wn16", 'nt': "C:\wn16"}.get(os.name, "/usr/home/yasu/wordnet-1.6"))
WNSEARCHDIR = environ.get('WNSEARCHDIR', path.join(WNHOME, {'mac': "Database"}.get(os.name, "dict")))
_FILE_OPEN_MODE = os.name in ('dos', 'nt') and 'rb' or 'r'	# work around a Windows Python bug


#Index File Format
#lemma pos poly_cnt p_cnt [ptr_symbol...] sense_cnt tagsense_cnt synset_offset [synset_offset...]

#Data File Format 
#synset_offset lex_filenum ss_type w_cnt word lex_id [word lex_id...] p_cnt [ptr...] [frames...] | gloss 


#IDX_FILE_ELEMENTS = ["lemma", "pos", "poly_cnt", "p_cnt", "ense_cnt","tagsense_cnt", "synset_offset" ]
		     

def main():

	form = cgi.FieldStorage()
	local = 0
	if form.has_key("word"):
		key = form["word"].value
	else:
		if os.name =="nt":
			key = "jack"
			local = 1
		
	sresult = printlines(key)

	if local:
		fout = open ("y:\\test\\test.xml", 'w') 
		fout.write(sresult)
		fout.close()
		print sresult
		
	else:
		print sresult
	
def printlines(key):
	speechparts = ["noun", "verb", "adj", "adv"]
	smallspeechparts = ["n", "v", "a", "r"]
	spcount = -1
	sresult = ""
	
	for part in speechparts:
		spcount = spcount + 1
		srsynset = ""
		idxfile = open(_indexFilePathname(part), _FILE_OPEN_MODE)
		line= binarySearchFile(idxfile, key)
		datfile = open(_dataFilePathname(part), _FILE_OPEN_MODE)
		if line != None:
			#print line
			tokens = string.split(line)
			howmany = int(tokens[2])
			fromwhere = int(tokens[3])  + 6
			#print howmany, fromwhere
			ints = map(int, tokens[fromwhere:fromwhere + howmany])
			#print ints

			sxmlword = ""

			for ofst in ints:
				
				datfile.seek(ofst)
				line = datfile.readline()
				line = xmlproperstring(line)
				
				
				sxmlword = sxmlword + xmlout(xmloutsynset(line,key), "s" ) + "\n"
				#sresult = sresult + datfile.readline()
				
			#srsynset = srsynset + xmlout( sxmlword, "p", 'type="' + smallspeechparts[spcount]+'"')+'\n'
			srsynset = srsynset + xmlout( xmlout(smallspeechparts[spcount],"t") + sxmlword, "p")
		sresult = sresult + srsynset
		datfile.close		
		idxfile.close
	if sresult !="":
		sresult = "\n" + xmlout(key,"k") + sresult
		sresult = xmlout(sresult, "idx")
		sresult = """
<?xml version='1.0'?>
<?xml-stylesheet type="text/xsl" href="idx.xsl" ?>
""" + sresult

		
	return sresult
def xmlproperstring(line):
	line = string.replace(line, '&', '&amp;')
	line = string.replace(line, '<', '&lt;')
	line = string.replace(line, '>', '&gt;')
	return line
def xmloutsynset(line,key):
	elms = string.split(line[:string.index(line, '|')])
	sresult = "\n"
	#sresult = sresult + xmlout(elms[0], "synsetoffset") + "\n"
	#sresult = sresult + xmlout(elms[1], "lexfilenum") + "\n"
	#sresult = sresult + xmlout(elms[2], "sstype") + "\n"
	#sresult = sresult + xmlout(elms[3], "wcnt") + "\n"
	wcnt = int(elms[3])
	cr = range(wcnt)
	for x in cr:
		wd = elms[4+x*2]
		if wd != key:
			sresult = sresult + xmlout(wd, "w") + "\n"
		#sresult = sresult + xmlout(elms[5+x*2], "lexid") + "\n"
	p_cnt = int(elms[4+2*wcnt])
	gloss = string.strip(line[string.index(line, '|') + 1:])
	sresult = sresult + xmlout(gloss, "g") + "\n"

	return sresult

def xmlout(line, tag, optarg=""):
	sresult = "<" + tag
	if optarg != "":
		sresult = sresult + " "
	sresult = sresult + optarg + ">"
	sresult = sresult + line
	sresult = sresult + "</" + tag + ">"
	return sresult
def _dataFilePathname(filenameroot):
	if os.name in ('dos', 'nt'):
		return path.join(WNSEARCHDIR, filenameroot + ".dat")
	else:
		return path.join(WNSEARCHDIR, "data." + filenameroot)
def _indexFilePathname(filenameroot):
	if os.name in ('dos', 'nt'):
		return path.join(WNSEARCHDIR, filenameroot + ".idx")
	else:
		return path.join(WNSEARCHDIR, "index." + filenameroot)
def binarySearchFile(file, key, cache={}, cacheDepth=-1):
	from stat import ST_SIZE
	key = key + ' '
	keylen = len(key)
	start, end = 0, os.stat(file.name)[ST_SIZE]
	currentDepth = 0
	prevmiddle = -1
	searchcount = 0
	
	while start < end:
		middle = (start + end) / 2
		searchcount = searchcount + 1

		
		############ infinite or heavy loop prevention
		if prevmiddle == middle:
			#break
			pass
		if searchcount > 100:
			break
		prevmiddle = middle
		###########


		if cache.get(middle):
			(offset, line) = cache[middle]
		else:
			file.seek(max(0, middle -1))
			if middle > 0:
				file.readline()
				offset, line = file.tell(), file.readline()

			if currentDepth < cacheDepth:
				cache[middle] = (offset, line)
		#print start, middle, end, offset, line, 
		if offset > end:
			assert end != middle - 1, "infinite loop"
			end = middle - 1
		elif line[:keylen] == key:# and line[keylen + 1] == ' ':
			return line
		elif line > key or line=="":
			assert end != middle - 1, "infinite loop"
			end = middle - 1
		elif line < key:
			start = offset + len(line) - 1
		currentDepth = currentDepth + 1
	return None
NOUN = 'noun'
VERB = 'verb'
ADJECTIVE = 'adjective'
ADVERB = 'adverb'
PartsOfSpeech = (NOUN, VERB, ADJECTIVE, ADVERB)

MORPHOLOGICAL_SUBSTITUTIONS = {
	NOUN: (('s', ''), ('ses', 's'), ('xes', 'x'), ('zes', 'z'), ('ches', 'ch'), ('shes', 'sh')),
	VERB: (('s', ''), ('ies', 'y'), ('es', 'e'), ('ed', 'e'), ('ed', ''), ('ing', 'e'), ('ing', 'e')),
	ADJECTIVE: (('er', ''), ('er', 'est'), ('er', 'e'), ('est', 'e')),
	ADVERB: None}

def morphy(form, pos='noun', collect=0):
	"""Recursively uninflect _form_, and return the first form found in the dictionary.
	If _collect_ is true, a sequence of all forms is returned, instead of just the first
	one.
	
	>>> morphy('dogs')
	'dog'
	>>> morphy('churches')
	'church'
	>>> morphy('aardwolves')
	'aardwolf'
	>>> morphy('abaci')
	'abacus'
	"""
	from wordnet import _normalizePOS, _dictionaryFor
	pos = _normalizePOS(pos)
	excfile = open(path.join(WNSEARCHDIR, {NOUN: 'noun', VERB: 'verb', ADJECTIVE: 'adj', ADVERB: 'adv'}[pos] + '.exc'))
	substitutions = MORPHOLOGICAL_SUBSTITUTIONS[pos]
	def trySubstitutions(trySubstitutions, form, substitutions, lookup=1, dictionary=_dictionaryFor(pos), excfile=excfile, collect=collect, collection=[]):
		import string
		exceptions = binarySearchFile(excfile, form)
		if exceptions:
			form = exceptions[string.find(exceptions, ' ')+1:-1]
		if lookup and dictionary.has_key(form):
			if collect:
				collection.append(form)
			else:
				return form
		elif substitutions:
			(old, new) = substitutions[0]
			substitutions = substitutions[1:]
			substitute = None
			if endsWith(form, old):
				substitute = form[:-len(old)] + new
				#if dictionary.has_key(substitute):
				#	return substitute
			form = 				trySubstitutions(trySubstitutions, form, substitutions) or \
				(substitute and trySubstitutions(trySubstitutions, substitute, substitutions))
			return (collect and collection) or form
		elif collect:
			return collection
	return trySubstitutions(trySubstitutions, form, substitutions)


main()
#if __name__ == '__main__':
	
#	main()
