import re
import os
import sys
import shutil
import copy

rootDir = sys.argv[1]
newDir = sys.argv[2]

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
					'fda':'fda'
}

ScopeExitDict = {
				'Sub':'End Sub',
				'Function':'End Function',
				'Class':'End Class',
				'Structure':'End Structure',
				'Module':'End Module'
}

ArrayPrefixDict = ['String', 'Byte', 'Boolean', 'Double', 'Decimal', 'Single', 'Int32', 'Integer']
ArraySubEx = re.compile('\w+')

FunctionScopeBegin1Ex = re.compile('(Private) Function |(Public) Function |(Protected) Function |(Friend) Function | Function\(\w+\)')
FunctionScopeBegin2Ex = re.compile('(Private) (\w+) Function |(Public) (\w+) Function |(Protected) (\w+) Function |(Friend) (\w+) Function ')
FunctionScopeEndEx = re.compile('End Function')

SubScopeBegin1Ex = re.compile('(Private) Sub |(Public) Sub |(Protected) Sub |(Friend) Sub | Sub(\w+)')
SubScopeBegin2Ex = re.compile('(Private) (\w+) Sub |(Public) (\w+) Sub |(Protected) (\w+) Sub |(Friend) (\w+) Sub ')
SubScopeEndEx = re.compile('End Sub')

ClassScopeBegin1Ex = re.compile('(Private) Class |(Public) Class |(Protected) Class |(Friend) Class |^Class')
ClassScopeBegin2Ex = re.compile('(Private) (\w+) Class |(Public) (\w+) Class |(Protected) (\w+) Class |(Friend) (\w+) Class ')
ClassScopeEndEx = re.compile('End Class')

StructureScopeBegin1Ex = re.compile('(Private) Structure |(Public) Structure |(Protected) Structure |(Friend) Structure ')
StructureScopeBegin2Ex = re.compile('(Private) (\w+) Structure |(Public) (\w+) Structure |(Protected) (\w+) Structure |(Friend) (\w+) Structure ')
StructureScopeEndEx = re.compile('End Structure')

ModuleScopeBegin1Ex = re.compile('(Private) Module |(Public) Module |(Protected) Module |(Friend) Module ')
ModuleScopeBegin2Ex = re.compile('(Private) (\w+) Module |(Public) (\w+) Module |(Protected) (\w+) Module |(Friend) (\w+) Module ')
ModuleScopeEndEx = re.compile('End Module')	

class ContextTerm(object):
	def __init__(self, oldName, newName):
		self.oldName = oldName
		self.newName = newName

	def __str__(self):
		return '('+self.oldName+'|'+self.newName+')'	

class ContextReplacer(object):

	def __init__(self):		
		self.scopeStack = None
		self.currentScopeTerms = []

	def ReadLine(self, line):
		#print(line)
		#print('#############')
		# wait = input("PRESS ENTER TO CONTINUE")
		if self.NewScopeCheck(line):
			#print('NEW SCOPE------------')
			if self.scopeStack == None:
				self.scopeStack = []
			else:
				self.scopeStack.append(copy.deepcopy(self.currentScopeTerms))
				# for scope in self.scopeStack:					
				# 	for term in scope:
				# 		print(term.oldName+'|'+term.newName)
				# 	print('SCOPE---------')

				self.currentScopeTerms.clear()

		if self.EndScopeCheck(line):
			#print('END SCOPE----------------')
			# for term in self.scopeStack[-1]:
			# 	print(term.oldName+'|'+term.newName)
			if not self.scopeStack:
				self.currentScopeTerms = []
			else:
				self.currentScopeTerms = self.scopeStack.pop()

		if self.scopeStack:
			for scope in self.scopeStack:
				for term in scope:
					term1 = ' '+term.oldName			
					term1Ex = re.compile(re.escape(term1))
					term2 = '('+term.oldName
					term2Ex = re.compile(re.escape(term2))
					term3 = ','+term.oldName
					term3Ex = re.compile(re.escape(term3))
					term4 = '='+term.oldName
					term4Ex = re.compile(re.escape(term4))
					if term1Ex.findall(line) != []:																		
						line = term1Ex.sub(' '+term.newName, line)
					if term2Ex.findall(line) != []:
						line = term2Ex.sub('('+term.newName, line)
					if term3Ex.findall(line) != []:
						line = term3Ex.sub(','+term.newName, line)
					if term4Ex.findall(line) != []:
						line = term4Ex.sub('='+term.newName, line)
			for term in self.currentScopeTerms:
				term1 = ' '+term.oldName			
				term1Ex = re.compile(term1)
				term2 = '('+term.oldName
				term2Ex = re.compile(re.escape(term2))
				term3 = ','+term.oldName
				term3Ex = re.compile(term3)
				term4 = '='+term.oldName
				term4Ex = re.compile(re.escape(term4))
				if term1Ex.findall(line) != []:																		
					line = term1Ex.sub(' '+term.newName, line)
				if term2Ex.findall(line) != []:
					line = term2Ex.sub('('+term.newName, line)
				if term3Ex.findall(line) != []:
					line = term3Ex.sub(','+term.newName, line)
				if term4Ex.findall(line) != []:
					line = term4Ex.sub('='+term.newName, line)
		
		return line

	def NewScopeCheck(self, line):
		if FunctionScopeBegin1Ex.findall(line) == [] and FunctionScopeBegin2Ex.findall(line) == [] \
	 	and SubScopeBegin1Ex.findall(line) == [] and SubScopeBegin2Ex.findall(line) == [] \
		and ClassScopeBegin1Ex.findall(line) == [] and ClassScopeBegin2Ex.findall(line) == [] \
		and StructureScopeBegin1Ex.findall(line) == [] and StructureScopeBegin2Ex.findall(line) == [] \
		and ModuleScopeBegin1Ex.findall(line) == [] and ModuleScopeBegin2Ex.findall(line) == []:
			return False
		return True

	def EndScopeCheck(self, line):
		if FunctionScopeEndEx.findall(line) == [] and SubScopeEndEx.findall(line) ==[] and ClassScopeEndEx.findall(line) == [] and StructureScopeEndEx.findall(line) == [] and ModuleScopeEndEx.findall(line) == []:
			return False	
		return True

	def AddTerm(self, oldName, newName):		
		self.currentScopeTerms.append(ContextTerm(oldName, newName))
		#print('Added Term - '+oldName + ' | '+ newName)
		#self.Print()
		
	def Print(self):
		for scope in self.scopeStack:
			for term in scope:				
				print(term.oldName + ' | ' + term.newName)				
		for term in self.currentScopeTerms:
				print(term.oldName + ' | ' + term.newName)	






##Parameters
ByValParametersCheck = re.compile('ByVal plst|ByVal pdct|ByVal pstr|ByVal pbyt|ByVal pbln|ByVal pdbl|ByVal pcur|ByVal psng|ByVal pfda|ByVal pobj|ByVal plng|ByVal pvnt', re.IGNORECASE)
def CheckByValParameters(match):
	if ByValParametersCheck.findall(match) != []:		
		return True	

ByRefParametersCheck = re.compile('ByRef plst|ByRef pdct|ByRef pstr|ByRef pbyt|ByRef pbln|ByRef pdbl|ByRef pcur|ByRef psng|ByRef pfda|ByRef pobj|ByRef plng|ByRef pvnt', re.IGNORECASE)
def CheckByRefParameters(match):
	if ByRefParametersCheck.findall(match) != []:
		return True

ByRefParametersReplaceEx = re.compile('ByRef (\w+) as (\S+)', re.IGNORECASE)
def replaceByRefParameters(match):	
	varName = match.group(1)
	varType = match.group(2)

	if varType == 'Integer':
		varType = 'Int32'

	if varType == 'ggFda.ggFdaCls)':
		varType = 'fda'

	if len(ArraySubEx.search(varType).group()) != len(varType):
		alphanumericVarType = ArraySubEx.search(varType).group()
		if alphanumericVarType in VariablePrefixDict:
			varPrefix = VariablePrefixDict.get(alphanumericVarType)
			if varPrefix == 'byt':
				replacer.AddTerm(varName, 'pimg'+varName[:1].capitalize()+varName[1:])
				return 'ByRef pimg'+varName[:1].capitalize()+varName[1:]+" as "+varType
			replacer.AddTerm(varName, 'pa'+varPrefix+varName[:1].capitalize()+varName[1:])
			return 'ByRef pa'+varPrefix+varName[:1].capitalize()+varName[1:]+" as "+varType

	if varType in VariablePrefixDict:
		varPrefix = VariablePrefixDict.get(varType)
	else:
		varPrefix = 'obj'

	if varType == 'fda':
		varType = 'ggfda.ggFdaCls)'

	replacer.AddTerm(varName, 'p'+varPrefix+varName[:1].capitalize()+varName[1:])

	return "ByRef p"+varPrefix+varName[:1].capitalize()+varName[1:]+" as "+varType

ByValParametersReplaceEx = re.compile('ByVal (\w+) as (\S+)', re.IGNORECASE)
def replaceByValParameters(match):		
	varName = match.group(1)
	varType = match.group(2)

	if varType == 'Integer':
		varType = 'Int32'

	if varType == 'ggFda.ggFdaCls)':
		varType = 'fda'

	if len(ArraySubEx.search(varType).group()) != len(varType):
		alphanumericVarType = ArraySubEx.search(varType).group()
		if alphanumericVarType in VariablePrefixDict:
			varPrefix = VariablePrefixDict.get(alphanumericVarType)
			if varPrefix == 'byt':
				replacer.AddTerm(varName, 'pimg'+varName[:1].capitalize()+varName[1:])
				return 'ByVal pimg'+varName[:1].capitalize()+varName[1:]+" as "+varType
			replacer.AddTerm(varName, 'pa'+varPrefix+varName[:1].capitalize()+varName[1:])
			return 'ByVal pa'+varPrefix+varName[:1].capitalize()+varName[1:]+" as "+varType

	if varType in VariablePrefixDict:
		varPrefix = VariablePrefixDict.get(varType)
	else:
		varPrefix = 'obj'

	if varType == 'fda':
		varType = 'ggfda.ggFdaCls)'

	replacer.AddTerm(varName, 'p'+varPrefix+varName[:1].capitalize()+varName[1:])

	return "ByVal p"+varPrefix+varName[:1].capitalize()+varName[1:]+" as "+varType

##Variables
#DIM
DimReplaceEx = re.compile('Dim (\w+) as (\S{4,})', re.IGNORECASE)
def dimReplace(match):
	varName = match.group(1)
	varType = match.group(2)

	if varType == 'Integer':
		varType = 'Int32'

	if varType == 'ggFda.ggFdaCls)':
		varType = 'fda'

	if len(ArraySubEx.search(varType).group()) != len(varType):
		alphanumericVarType = ArraySubEx.search(varType).group()
		if alphanumericVarType in VariablePrefixDict:
			varPrefix = VariablePrefixDict.get(alphanumericVarType)
			if varPrefix == 'byt':
				replacer.AddTerm(varName, 'img'+varName[:1].capitalize()+varName[1:])
				return 'Dim img'+varName[:1].capitalize()+varName[1:]+" as "+varType
			replacer.AddTerm(varName, 'a'+varPrefix+varName[:1].capitalize()+varName[1:])
			return 'Dim a'+varPrefix+varName[:1].capitalize()+varName[1:]+" as "+varType
	
	if varType in VariablePrefixDict:
		varPrefix = VariablePrefixDict.get(varType)
	else:
		varPrefix = 'obj'

	if varType == 'fda':
		varType = 'ggfda.ggFdaCls)'

	replacer.AddTerm(varName, varPrefix+varName[:1].capitalize()+varName[1:])	

	return 'Dim '+varPrefix+varName[:1].capitalize()+varName[1:]+" as "+varType

DimNewReplaceEx = re.compile('Dim (\w+) as New (\S+)', re.IGNORECASE)
def dimNewReplace(match):
	varName = match.group(1)
	varType = match.group(2)

	if varType == 'Integer':
		varType = 'Int32'

	if varType == 'ggFda.ggFdaCls)':
		varType = 'fda'

	if len(ArraySubEx.search(varType).group()) != len(varType):
		alphanumericVarType = ArraySubEx.search(varType).group()
		if alphanumericVarType in VariablePrefixDict:
			varPrefix = VariablePrefixDict.get(alphanumericVarType)
			if varPrefix == 'byt':
				replacer.AddTerm(varName, 'img'+varName[:1].capitalize()+varName[1:])
				return 'Dim img'+varName[:1].capitalize()+varName[1:]+" as New "+varType
			replacer.AddTerm(varName, 'a'+varPrefix+varName[:1].capitalize()+varName[1:])
			return 'Dim a'+varPrefix+varName[:1].capitalize()+varName[1:]+" as New "+varType	

	if varType in VariablePrefixDict:
		varPrefix = VariablePrefixDict.get(varType)
	else:
		varPrefix = 'obj'

	if varType == 'fda':
		varType = 'ggfda.ggFdaCls)'

	replacer.AddTerm(varName, varPrefix+varName[:1].capitalize()+varName[1:])	

	return 'Dim '+varPrefix+varName[:1].capitalize()+varName[1:]+" as New "+varType

DimPrefixCheck = re.compile('Dim lst|Dim dct|Dim str|Dim byt|Dim bln|Dim dbl|Dim cur|Dim sng|Dim fda|Dim obj', re.IGNORECASE)
def CheckDimPrefix(line):
	if DimPrefixCheck.findall(line) != []:
		return True

#FOR
ForReplaceEx = re.compile('For (\w+) as (\S+)', re.IGNORECASE)
def ForReplace(match):
	varName = match.group(1)
	varType = match.group(2)

	if varType == 'Integer':
		varType = 'Int32'

	if varType == 'ggFda.ggFdaCls)':
		varType = 'fda'

	if len(ArraySubEx.search(varType).group()) != len(varType):
		alphanumericVarType = ArraySubEx.search(varType).group()
		if alphanumericVarType in VariablePrefixDict:
			varPrefix = VariablePrefixDict.get(alphanumericVarType)
			if varPrefix == 'byt':
				replacer.AddTerm(varName, 'img'+varName[:1].capitalize()+varName[1:])
				return 'For img'+varName[:1].capitalize()+varName[1:]+" as "+varType
			replacer.AddTerm(varName, 'a'+varPrefix+varName[:1].capitalize()+varName[1:])
			return 'For a'+varPrefix+varName[:1].capitalize()+varName[1:]+" as "+varType

	if varType in VariablePrefixDict:
		varPrefix = VariablePrefixDict.get(varType)
	else:
		varPrefix = 'obj'

	if varType == 'fda':
		varType = 'ggfda.ggFdaCls)'

	replacer.AddTerm(varName, varPrefix+varName[:1].capitalize()+varName[1:])

	return 'For '+varPrefix+varName[:1].capitalize()+varName[1:]+" as "+varType

ForPrefixCheck = re.compile('For lst|For dct|For str|For byt|For bln|For dbl|For cur|For sng|For fda|For obj', re.IGNORECASE)
def CheckForPrefix(line):
	if ForPrefixCheck.findall(line) != []:
		return True

#FOR EACH
ForEachReplaceEx = re.compile('For Each (\w+) as (\S+)', re.IGNORECASE)
def ForEachReplace(match):
	varName = match.group(1)
	varType = match.group(2)

	if varType == 'Integer':
		varType = 'Int32'

	if varType == 'ggFda.ggFdaCls)':
		varType = 'fda'

	if len(ArraySubEx.search(varType).group()) != len(varType):
		alphanumericVarType = ArraySubEx.search(varType).group()
		if alphanumericVarType in VariablePrefixDict:
			varPrefix = VariablePrefixDict.get(alphanumericVarType)
			if varPrefix == 'byt':
				replacer.AddTerm(varName, 'img'+varName[:1].capitalize()+varName[1:])
				return 'For each img'+varName[:1].capitalize()+varName[1:]+" as "+varType
			replacer.AddTerm(varName, 'a'+varPrefix+varName[:1].capitalize()+varName[1:])
			return 'For each a'+varPrefix+varName[:1].capitalize()+varName[1:]+" as "+varType
	
	if varType in VariablePrefixDict:
		varPrefix = VariablePrefixDict.get(varType)
	else:
		varPrefix = 'obj'

	if varType == 'fda':
		varType = 'ggfda.ggFdaCls)'

	replacer.AddTerm(varName, varPrefix+varName[:1].capitalize()+varName[1:])

	return 'For Each '+varPrefix+varName[:1].capitalize()+varName[1:]+" as "+varType

ForEachPrefixCheck = re.compile('For Each lst|For Each dct|For Each str|For Each byt|For Each bln|For Each dbl|For Each cur|For Each sng|For Each fda|For Each obj', re.IGNORECASE)
def CheckForEachPrefix(line):
	if ForEachPrefixCheck.findall(line) != []:
		return True


#Private
PrivateReplaceEx = re.compile('Private (\w+) as (\S+)', re.IGNORECASE)
def PrivateReplace(match):
	varName = match.group(1)
	varType = match.group(2)

	if varType == 'Integer':
		varType = 'Int32'

	if varType == 'ggFda.ggFdaCls)':
		varType = 'fda'

	if len(ArraySubEx.search(varType).group()) != len(varType):
		alphanumericVarType = ArraySubEx.search(varType).group()
		if alphanumericVarType in VariablePrefixDict:
			varPrefix = VariablePrefixDict.get(alphanumericVarType)
			if varPrefix == 'byt':
				replacer.AddTerm(varName, 'img'+varName[:1].capitalize()+varName[1:])
				return 'Private img'+varName[:1].capitalize()+varName[1:]+" as "+varType
			replacer.AddTerm(varName, 'a'+varPrefix+varName[:1].capitalize()+varName[1:])
			return 'Private a'+varPrefix+varName[:1].capitalize()+varName[1:]+" as "+varType
	
	if varType in VariablePrefixDict:
		varPrefix = VariablePrefixDict.get(varType)
	else:
		varPrefix = 'obj'

	if varType == 'fda':
		varType = 'ggfda.ggFdaCls)'

	replacer.AddTerm(varName, varPrefix+varName[:1].capitalize()+varName[1:])

	return 'Private '+varPrefix+varName[:1].capitalize()+varName[1:]+" as "+varType

PrivatePrefixCheck = re.compile('Private lst|Private dct|Private str|Private byt|Private bln|Private dbl|Private cur|Private sng|Private fda|Private obj', re.IGNORECASE)
def CheckPrivatePrefix(line):
	if PrivatePrefixCheck.findall(line) != []:
		return True


PrivateArrayPrefixCheck = re.compile('Private alst|Private adct|Private astr|Private img|Private abln|Private adbl|Private acur|Private asng|Private afda|Private aobj', re.IGNORECASE)
def CheckPrivateArrayPrefix(line):
	if PrivateArrayPrefixCheck.findall(line) != []:
		return True

#Public
PublicReplaceEx = re.compile('Public (\w+) as (\S+)', re.IGNORECASE)
def PublicReplace(match):
	varName = match.group(1)
	varType = match.group(2)

	if varType == 'Integer':
		varType = 'Int32'

	if varType == 'ggFda.ggFdaCls)':
		varType = 'fda'

	if len(ArraySubEx.search(varType).group()) != len(varType):
		alphanumericVarType = ArraySubEx.search(varType).group()
		if alphanumericVarType in VariablePrefixDict:
			varPrefix = VariablePrefixDict.get(alphanumericVarType)
			if varPrefix == 'byt':
				replacer.AddTerm(varName, 'img'+varName[:1].capitalize()+varName[1:])
				return 'Public img'+varName[:1].capitalize()+varName[1:]+" as "+varType
			replacer.AddTerm(varName, 'a'+varPrefix+varName[:1].capitalize()+varName[1:])
			return 'Public a'+varPrefix+varName[:1].capitalize()+varName[1:]+" as "+varType

	
	if varType in VariablePrefixDict:
		varPrefix = VariablePrefixDict.get(varType)
	else:
		varPrefix = 'obj'

	if varType == 'fda':
		varType = 'ggfda.ggFdaCls)'

	replacer.AddTerm(varName, varPrefix+varName[:1].capitalize()+varName[1:])

	return 'Public '+varPrefix+varName[:1].capitalize()+varName[1:]+" as "+varType

PublicPrefixCheck = re.compile('Public lst|Public dct|Public str|Public byt|Public bln|Public dbl|Public cur|Public sng|Public fda|Public obj', re.IGNORECASE)
def CheckPublicPrefix(line):
	if PublicPrefixCheck.findall(line) != []:
		return True


for path, dirs, files in os.walk(rootDir):
	for file in files:				
		if file[-3:] == '.vb':			
			with open(path+'\\'+file, 'r') as iFile:
				print(path+'\\'+file)
				if not os.path.exists(newDir+path[len(rootDir):]):
					os.makedirs(newDir+path[len(rootDir):])
				with open(newDir+path[len(rootDir):]+'\\'+file, 'w') as oFile:
					lines = iFile.readlines()
					replacer = ContextReplacer()
					for i in range(len(lines)):
						line = replacer.ReadLine(lines[i])
						replacedFlag = False																		
						if not CheckDimPrefix(line):
							if DimNewReplaceEx.findall(line) != []:
								replacedFlag = True
								line = DimNewReplaceEx.sub(dimNewReplace, line)
							elif DimReplaceEx.findall(line) != []:
								replacedFlag = True
								line = DimReplaceEx.sub(dimReplace, line)
						if not CheckForPrefix(line):
							if ForReplaceEx.findall(line) != []:								
								replacedFlag = True
								line = ForReplaceEx.sub(ForReplace, line)
						if not CheckForEachPrefix(line):
							if ForEachReplaceEx.findall(line) != []:
								replacedFlag = True
								line = ForEachReplaceEx.sub(ForEachReplace, line)
						if not CheckPrivatePrefix(line):
							if PrivateReplaceEx.findall(line) != []:
								replacedFlag = True
								line = PrivateReplaceEx.sub(PrivateReplace, line)
						if not CheckPublicPrefix(line):
							if PublicReplaceEx.findall(line) != []:
								replacedFlag = True
								line = PublicReplaceEx.sub(PublicReplace, line)
						if not CheckByRefParameters(line):
							if ByRefParametersReplaceEx.findall(line) != []:
								replacedFlag = True							
								line = ByRefParametersReplaceEx.sub(replaceByRefParameters, line)
						if not CheckByValParameters(line):
							if ByValParametersReplaceEx.findall(line) != []:
								replacedFlag = True											
								line = ByValParametersReplaceEx.sub(replaceByValParameters, line)
						
						oFile.write(line)							
							
		else:
			if not os.path.exists(newDir+path[len(rootDir):]):
					os.makedirs(newDir+path[len(rootDir):])
			try:
				shutil.copy(path+'\\'+file, newDir+path[len(rootDir):]+'\\'+file)
			except Exception as e:
				print(e)


			
					
