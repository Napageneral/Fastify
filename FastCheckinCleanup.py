import re
import os
import sys
import shutil
import copy
import math

ScopeBeginEx = re.compile('\w+ Function |\w+ Sub |\w+ Class |\w+ Structure |\w+ Module|^Function |^Sub |^Class |^Structure |^Module | Function\(| Sub\(')
ScopeEndEx = re.compile('End Function|End Sub|End Class|End Structure|End Module')
class ContextTerm(object):
	def __init__(self, oldName, newName):
		self.oldName = oldName
		self.newName = newName

class ContextReplacer(object):
	def __init__(self):		
		self.scopeStack = None
		self.currentScopeTerms = []

	def ReadLine(self, line):
		if self.NewScopeCheck(line):			
			if self.scopeStack == None:
				self.scopeStack = []
			else:
				self.scopeStack.append(copy.deepcopy(self.currentScopeTerms))
				self.currentScopeTerms.clear()

		if self.EndScopeCheck(line):
			if not self.scopeStack:
				self.currentScopeTerms = []
			else:
				self.currentScopeTerms = self.scopeStack.pop()

		if self.scopeStack:
			for scope in self.scopeStack:
				for term in scope:					
					line = self._replaceTerm(term, line)
			for term in self.currentScopeTerms:
					line = self._replaceTerm(term, line)		
		return line

	def NewScopeCheck(self, line):
		if ScopeBeginEx.findall(line) == []:
			return False
		return True

	def EndScopeCheck(self, line):
		if ScopeEndEx.findall(line) == []:
			return False	
		return True

	def AddTerm(self, oldName, newName):		
		self.currentScopeTerms.append(ContextTerm(oldName, newName))

	def _replaceTerm(self, term, line):	
		global TermToReplace
		TermToReplace = term		
		newLineTermReplaceEx = re.compile('\A'+TermToReplace.oldName)
		termReplaceEx = re.compile('([\s,=\(])'+TermToReplace.oldName+'([,.\s\)])')
		if termReplaceEx.findall(line) != []:
			line = termReplaceEx.sub(self._termReplace, line)
		if newLineTermReplaceEx.findall(line) != []:
			line = newLineTermReplaceEx.sub(TermToReplace.newName, line)
		return line
	
	def _termReplace(self, match):
		termPrecede = match.group(1)
		termSuccede = match.group(2)
		return termPrecede+TermToReplace.newName+termSuccede



VariablePrefixCheckEx = re.compile('\w+ lng|\w+ lst|\w+ dct|\w+ str|\w+ byt|\w+ bln|\w+ dbl|\w+ cur|\w+ sng|\w+ fda|\w+ obj|\w+ sb', re.IGNORECASE)
ForEachPrefixCheckEx = re.compile('For each lng|For each lst|For each dct|For each str|For each byt|For each bln|For each dbl|For each cur|For each sng|For each fda|For each obj|For each sb', re.IGNORECASE)
MemberAndParameterPrefixCheckEx = re.compile('\w+ [mp]lng|\w+ [mp]lst|\w+ [mp]dct|\w+ [mp]str|\w+ [mp]byt|\w+ [mp]bln|\w+ [mp]dbl|\w+ [mp]cur|\w+ [mp]sng|\w+ [mp]fda|\w+ [mp]obj|\w+ [mp]sb', re.IGNORECASE)
#MemberVariablePrefixCheckEx = re.compile('\w+ mlst|\w+ mdct|\w+ mstr|\w+ mbyt|\w+ mbln|\w+ mdbl|\w+ mcur|\w+ msng|\w+ mfda|\w+ mobj|\w+ msb', re.IGNORECASE)
#ParameterPrefixCheckEx = re.compile('\w+ plst|\w+ pdct|\w+ pstr|\w+ pbyt|\w+ pbln|\w+ pdbl|\w+ pcur|\w+ psng|\w+ pfda|\w+ pobj|\w+ psb', re.IGNORECASE)

def PrefixCheck(line):
	if 'Property' in line:
		return False
	if VariablePrefixCheckEx.findall(line) != []:
		return False
	if MemberAndParameterPrefixCheckEx.findall(line) != []:
		return False	
	if ForEachPrefixCheckEx.findall(line) != []:
		return False
	return True


# def ReplaceNewKeyword(line, replacer):	
# 	varKeyword, varPrefix, varName, varType = MainReplace(NewKeywordReplaceEx.findall(line), replacer)
# 	replacement = varKeyword+varPrefix+varName[:1].capitalize()+varName[1:]+" as New "+varType
# 	return NewKeywordReplaceEx.sub(replacement, line)



NewKeywordReplaceEx = re.compile('(\w+) (\w+) as New (\S+)', re.IGNORECASE)
def NewKeywordReplace(match):
	varKeyword, varPrefix, varName, varType = MainReplace(match)	
	return varKeyword+varPrefix+varName[:1].capitalize()+varName[1:]+" as New "+varType

KeywordReplaceEx = re.compile('(\w+) (\w+) as (\S+)|(For each) (\w+) as (\S+)', re.IGNORECASE)
def KeywordReplace(match):
	varKeyword, varPrefix, varName, varType = MainReplace(match)	
	return varKeyword+varPrefix+varName[:1].capitalize()+varName[1:]+" as "+varType

VariablePrefixDict = {
					'List':'lst',
					'Dictionary':'dct',
					'String':'str',
					'Byte':'byt',
					'Boolean':'bln',
					'Double':'dbl',
					'Decimal':'cur',
					'Single':'sng',
					'ggFda.ggFdaCls':'fda',
					'Int32':'lng',
					'Integer':'lng',
					'fda':'fda'}

AlphaNumericEx = re.compile('\w+', re.IGNORECASE)
parameterKeywords = ['ByVal', 'ByRef']
ArraySet = ['String', 'Byte', 'Boolean', 'Double', 'Decimal', 'Single', 'Int32', 'Integer']
def MainReplace(match):
	if match.group(1) != None:		
		varKeyword = match.group(1)
		varName = match.group(2)
		varType = match.group(3)
	else:
		varKeyword = match.group(4)
		varName = match.group(5)
		varType = match.group(6)

	alphanumericVarType = AlphaNumericEx.search(varType).group()

	varKeyword = varKeyword+' '

	if varType == 'ggFda.ggFdaCls)':
		varType = 'fda'

	if alphanumericVarType in VariablePrefixDict:
		varPrefix = VariablePrefixDict.get(alphanumericVarType)
	else:
		varPrefix = 'obj'

	if abs(len(alphanumericVarType)-len(varType)) > 1:		
		if alphanumericVarType in ArraySet:			
			if varPrefix == 'byt':
				varPrefix = 'img'
			else:
				varPrefix = 'a'+varPrefix

	if varKeyword in parameterKeywords:
		varPrefix = 'p'+varPrefix

	if varType == 'fda':
		varType = 'ggfda.ggFdaCls)'

	if varType == 'Integer':
		varType = 'Int32'
	
	replacer.AddTerm(varName, varPrefix+varName[:1].capitalize()+varName[1:])
	return (varKeyword, varPrefix, varName, varType)




def main():
	print('At least we started')
	rootDir = sys.argv[1]
	newDir = sys.argv[2]

	for path, dirs, files in os.walk(rootDir):
		for file in files:				
			if file[-3:] == '.vb':			
				with open(path+'\\'+file, 'r') as iFile:
					print(path+'\\'+file)
					if not os.path.exists(newDir+path[len(rootDir):]):
						os.makedirs(newDir+path[len(rootDir):])
					with open(newDir+path[len(rootDir):]+'\\'+file, 'w') as oFile:					
						global replacer 
						replacer = ContextReplacer()
						for line in iFile:												
							if PrefixCheck(line):
								line = NewKeywordReplaceEx.sub(NewKeywordReplace, line)
							if PrefixCheck(line):
								line = KeywordReplaceEx.sub(KeywordReplace, line)
							line = replacer.ReadLine(line)							
							oFile.write(line)							
								
			else:
				if not os.path.exists(newDir+path[len(rootDir):]):
						os.makedirs(newDir+path[len(rootDir):])
				try:
					shutil.copy(path+'\\'+file, newDir+path[len(rootDir):]+'\\'+file)
				except Exception as e:
					print(e)


if __name__ == '__main__':	
	main()