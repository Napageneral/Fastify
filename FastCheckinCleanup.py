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
		self.scopeStack = []
		self.currentScopeTerms = []

	def ReadLine(self, line):		
		if self.scopeStack:			
			for scope in self.scopeStack:
				for term in scope:					
					line = self._replaceTerm(term, line)
			for term in self.currentScopeTerms:
					line = self._replaceTerm(term, line)		
		return line

	def _NewScopeCheck(self, line):
		if ScopeBeginEx.findall(line) == []:
			return False
		return True

	def NewScopeCheck(self, line):
		if self._NewScopeCheck(line):			
			if self.scopeStack == None:
				self.scopeStack = []
			else:
				self.scopeStack.append(copy.deepcopy(self.currentScopeTerms))
				self.currentScopeTerms.clear()

	def _EndScopeCheck(self, line):
		if ScopeEndEx.findall(line) == []:
			return False	
		return True

	def EndScopeCheck(self, line):
		if self._EndScopeCheck(line):
			if not self.scopeStack:
				self.currentScopeTerms = []
			else:
				self.currentScopeTerms = self.scopeStack.pop()

	def AddTerm(self, oldName, newName):		
		self.currentScopeTerms.append(ContextTerm(oldName, newName))

	def _replaceTerm(self, term, line):	
		global TermToReplace
		TermToReplace = term		
		newLineTermReplaceEx = re.compile('\A'+TermToReplace.oldName)		
		termReplaceEx = re.compile('([-\s,=\(\{])'+TermToReplace.oldName+'([,.\s\)\(\}])')
		if termReplaceEx.findall(line) != []:
			line = termReplaceEx.sub(self._termReplace, line)
		if newLineTermReplaceEx.findall(line) != []:
			line = newLineTermReplaceEx.sub(TermToReplace.newName, line)
		return line
	
	def _termReplace(self, match):
		termPrecede = match.group(1)
		termSuccede = match.group(2)
		return termPrecede+TermToReplace.newName+termSuccede



VariablePrefixCheckEx = re.compile('\w{3,} enm|\w{3,} lng|\w{3,} lst|\w{3,} dct|\w{3,} str|\w{3,} byt|\w{3,} bln|\w{3,} dbl|\w{3,} cur|\w{3,} sng|\w{3,} fda|\w{3,} obj|\w{3,} sb', re.IGNORECASE)
ForEachPrefixCheckEx = re.compile('For each lng|For each lst|For each dct|For each str|For each byt|For each bln|For each dbl|For each cur|For each sng|For each fda|For each obj|For each sb', re.IGNORECASE)
MemberAndParameterPrefixCheckEx = re.compile('\w{3,} [mp]enm| mdta |\w{3,} [mp]lng|\w{3,} [mp]lst|\w{3,} [mp]dct|\w{3,} [mp]str|\w{3,} [mp]byt|\w{3,} [mp]bln|\w{3,} [mp]dbl|\w{3,} [mp]cur|\w{3,} [mp]sng|\w{3,} [mp]fda|\w{3,} [mp]obj|\w{3,} [mp]sb|\w{3,} [mp]vnt', re.IGNORECASE)

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
					'fda':'fda',
					'Text':'sb',
					'ggSql':'con'}

AlphaNumericEx = re.compile('\w+', re.IGNORECASE)
parameterKeywords = ['ByVal ', 'ByRef ', 'ByVal', 'ByRef', 'Optional']
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

	if AlphaNumericEx.search(varKeyword).group() in parameterKeywords:		
		varPrefix = 'p'+varPrefix

	if varType == 'fda':
		varType = 'ggfda.ggFdaCls)'

	if varType == 'Integer':
		varType = 'Int32'

	
	if replacer.scopeStack:		
		if len(replacer.scopeStack) == 1:			
			varPrefix = 'm'+varPrefix
	
	replacer.AddTerm(varName, varPrefix+varName[:1].capitalize()+varName[1:])
	return (varKeyword, varPrefix, varName, varType)




def main():	
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
							replacer.NewScopeCheck(line)
							replacer.EndScopeCheck(line)												
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