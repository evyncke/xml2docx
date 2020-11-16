#!/usr/bin/env python
# -*- coding: utf-8 -*-

#   Copyright 2020, Eric Vyncke, evyncke@cisco.com
#
#   Licensed under the Apache License, Version 2.0 (the "License");
#   you may not use this file except in compliance with the License.
#   You may obtain a copy of the License at
#
#       http://www.apache.org/licenses/LICENSE-2.0
#
#   Unless required by applicable law or agreed to in writing, software
#   distributed under the License is distributed on an "AS IS" BASIS,
#   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
#   See the License for the specific language governing permissions and
#   limitations under the License.
   
   
from xml.dom import minidom, Node
import xml.dom
from pprint import pprint
import sys, getopt
import io, os
import zipfile
import tempfile, datetime

# Same states to be kept
rfcDate = None
rfcAuthors = []
rfcTitle = None
rfcKeywords = []

def printTree(front):
	print('All children:')
	for elem in front.childNodes:
		if elem.nodeType != Node.ELEMENT_NODE:
			continue
		print("\t", elem.nodeName)
		print("\tAttributes:")
		for i in range(elem.attributes.length):
			attrib = elem.attributes.item(i)
			print("\t\t", attrib.name, ' = ' , attrib.value)
		print("\tChildren:")
		for child in elem.childNodes:
			if child.nodeType == Node.ELEMENT_NODE:
				print("\t\tELEMENT: ",child.nodeName)
			elif child.nodeType == Node.TEXT_NODE:
				print("\t\tTEXT: ", child.nodeValue)
	print("\n----------\n")

def docxNewParagraph(textValue, style = 'Normal', justification = None, unnumbered = None, numberingID = None, indentationLevel = None):
	if textValue is None:
		return None
	docxP = docxRoot.createElement('w:p')
	
# First handle the style or justification
#	<w:pPr>
#			<w:pStyle w:val="Title"/>
#			<w:jc w:val="right"/>
#			<w:rPr>
#				<w:lang w:val="en-US"/>
#			</w:rPr>
#	</w:pPr>
	pPr = docxRoot.createElement('w:pPr')
	if style != None:
		pStyle =  docxRoot.createElement('w:pStyle')
		pStyle.setAttribute('w:val', style) 
		pPr.appendChild(pStyle)
	if justification != None:
		jc =  docxRoot.createElement('w:jc')
		jc.setAttribute('w:val', justification) 
		pPr.appendChild(jc)
	if unnumbered:  # Try to override the default numbering in the style
		numPr = docxRoot.createElement('w:numPr')
		ilvl = docxRoot.createElement('w:ilvl ')
		ilvl.setAttribute('w:val', 0)
		numPr.appendChild(ilvl)
		numId = docxRoot.createElement('w:numId')
		numId.setAttribute('w:val', 0)
		numPr.appendChild(numId)
		pPr.appendChild(numPr)
	elif numberingID != None and indentationLevel != None:
#				<w:numPr>
#					<w:ilvl w:val="0"/>
#					<w:numId w:val="2"/>
#				</w:numPr>
		numPr = docxRoot.createElement('w:numPr')
		ilvl = docxRoot.createElement('w:ilvl ')
		ilvl.setAttribute('w:val', indentationLevel)
		numPr.appendChild(ilvl)
		numId = docxRoot.createElement('w:numId')
		numId.setAttribute('w:val', numberingID)
		numPr.appendChild(numId)
		pPr.appendChild(numPr)
	docxP.appendChild(pPr)
	
# Then handle the actual text
#	<w:r w:rsidRPr="00C46909">
#		<w:rPr>
#			<w:lang w:val="en-US"/>
#		</w:rPr>
#		<w:t>Title</w:t>
#	</w:r>
	r = docxRoot.createElement('w:r')
	rPr = docxRoot.createElement('w:rPr')
	lang = 	docxRoot.createElement('w:lang')
	lang.setAttribute('w:val', 'en-US')
	rPr.appendChild(lang)
	r.appendChild(rPr)
	t = docxRoot.createElement('w:t')
	text = docxRoot.createTextNode(' '.join(textValue.split()))
	t.appendChild(text)
	r.appendChild(t) 
	docxP.appendChild(r)
	return docxP

def parseAbstract(elem):
	for child in elem.childNodes:
		if child.nodeType != Node.ELEMENT_NODE:
			continue
		elif child.nodeName == 't':
			parseText(child, style = 'IntenseQuote')
		else:
			print('Unexpected tagName in Abstract: ', child.nodeName)

def parseArea(elem):
	textValue = 'Area: '
	for text in elem.childNodes:
		if text.nodeType == Node.TEXT_NODE:
			textValue += text.nodeValue
		if elem.nodeType == Node.ELEMENT_NODE:
			if text.nodeName != '#text':
				print('!!!!! parseKeyword: Text is ELEMENT_NODE: ', text.nodeName)
	docxBody.appendChild(docxNewParagraph(textValue))

def parseAuthor(elem):	# Per https://tools.ietf.org/html/rfc7991#section-2.7
	global rfcAuthors

	# looking for the organization element as in https://tools.ietf.org/html/rfc7991#section-2.35 that can only contain text
	organization = ''
	for child in elem.childNodes:
		if child.nodeType != Node.ELEMENT_NODE:
			continue
		elif child.nodeName == 'organization':
			for grandchild in child.childNodes:
				if grandchild.nodeType == Node.TEXT_NODE:
					organization = ', ' + grandchild.nodeValue

	if elem.hasAttribute('asciiFullname'):
		docxBody.appendChild(docxNewParagraph(elem.getAttribute('asciiFullname') + organization, justification = 'right'))
		rfcAuthors.append(elem.getAttribute('asciiFullname') + organization)
	elif elem.hasAttribute('fullname'):
		docxBody.appendChild(docxNewParagraph(elem.getAttribute('fullname') + organization, justification = 'right'))
		rfcAuthors.append(elem.getAttribute('fullname') + organization)
	else:
		author = ''
		if elem.hasAttribute('initials'):
			author = author + elem.getAttribute('initials') + ' '
		if elem.hasAttribute('surname'):
			author = author + elem.getAttribute('surname')
		if author != '':
			docxBody.appendChild(docxNewParagraph(author + organization, justification = 'right'))
			rfcAuthors.append(author + organization)

def parseBcp14(elem):  # https://tools.ietf.org/html/rfc7991#section-2.9 only text
	if elem.nodeValue != None:
		print('Bcp14 nodeValue: ' , elem.nodeValue)
	if elem.nodeType == Node.TEXT_NODE:
		print('Bcp14 node is TEXT_NODE')
	for child in elem.childNodes:
		if child.nodeType == Node.TEXT_NODE:
			return child.nodeValue
		else:
			print('!!!! parseBcp14 unexpected nodeType: ' + child.nodeType)
	
def parseBoilerPlate(elem):
	for child in elem.childNodes:
		if child.nodeType != Node.ELEMENT_NODE:
			continue
		elif child.nodeName == 'section':
			print('parseBoilerPlate calling parseSection()')
			parseSection(child, 1)
		else:
			print('Unexpected tagName in BoilerPlate: ', child.nodeName)

def parseDate(elem):
	global rfcDate
	
	dateString = ''
	if elem.hasAttribute('day'):
		dateString = elem.getAttribute('day') + ' '
	if elem.hasAttribute('month'):
		dateString = dateString + elem.getAttribute('month') + ' '
	if elem.hasAttribute('year'):
		dateString = dateString + elem.getAttribute('year')
	if dateString != '':
		docxBody.appendChild(docxNewParagraph(dateString, justification = 'right'))
		rfcDate = dateString
	
def parseEref(elem):	# See also https://tools.ietf.org/html/rfc7991#section-2.24
	if elem.nodeValue != None:
		print('Eref nodeValue: ' , elem.nodeValue)
	if elem.hasAttribute('target'):	# one and only mandatory attribute
		return '[' + elem.getAttribute('target') + ']'
	# Only target attribute, so, quite useless to parse other attributes
	if elem.nodeType == Node.TEXT_NODE:
		print('Eref node is TEXT_NODE')
	for child in elem.childNodes:
		if child.nodeType == Node.TEXT_NODE:
			return child.nodeValue
		if child.nodeName == 't':
			print("parseEref recurse into t !!!")
			parseText(child)

def parseFigure(elem):
	print('Skipping a figure')
	
def parseKeyword(elem):
	global rfcKeywords
	
	textValue = 'Keyword: '
	for text in elem.childNodes:
		if text.nodeType == Node.TEXT_NODE:
			textValue += text.nodeValue
			rfcKeywords.append(text.nodeValue)
		if elem.nodeType == Node.ELEMENT_NODE:
			if text.nodeName != '#text':
				print('!!!!! parseKeyword: Text is ELEMENT_NODE: ', text.nodeName)
	docxBody.appendChild(docxNewParagraph(textValue))

def parseList(elem):  # See also https://tools.ietf.org/html/rfc7991#section-2.29
	for child in elem.childNodes:
		if child.nodeType != Node.ELEMENT_NODE:
			continue
		if child.nodeName == 't':
			parseText(child, style = 'ListParagraph', numberingID = '2', indentationLevel = '0')  # numID = 2 is defined in numbering.xml as bullet list
		else:
			print('!!!! parseList, unexpected child: ', child.nodeName)
		
def parseListItem(elem, style = 'ListParagraph', numberingID = None, indentationLevel = None):
	print("start LI ", elem)
	for i in range(elem.attributes.length):
		attrib = elem.attributes.item(i)
		if attrib.name == 'pn' or  attrib.name == 'anchor' or  attrib.name == 'derivedCounter': 	# Let's ignore this marking as no obvious requirement or support in Office OpenXML
			continue
		print("\tLI unexpected attribute: ", attrib.name, ' = ' , attrib.value)

	textValue = ''
	for text in elem.childNodes:
		if text.nodeType == Node.TEXT_NODE:
			textValue += text.nodeValue
		if elem.nodeType == Node.ELEMENT_NODE:
			if text.nodeName == 'bcp14':
				textValue = textValue + parseBcp14(text)
			elif text.nodeName == 'eref':
				textValue = textValue + parseXref(text)
			elif text.nodeName == 'xref':
				textValue = textValue + parseXref(text)
			elif text.nodeName != '#text':
				print('!!!!! parseListItem: Text is ELEMENT_NODE: ', text.nodeName)
#			else:
#				print('parseListItem ignoring Text is ELEMENT_NODE: ', text.nodeName)
	docxBody.appendChild(docxNewParagraph(textValue, style = style, numberingID = numberingID, indentationLevel = indentationLevel))
	print("end LI ", textValue)

def parseOList(elem):
	for child in elem.childNodes:
		if child.nodeType != Node.ELEMENT_NODE:
			continue
		if child.nodeName == 'li':
			parseListItem(child, numberingID = '1', indentationLevel = '0')  # numID = 1 is defined in numbering.xml as enumeration list
		else:
			print('!!!! Unexpected List child: ', child.nodeName)


def parseSection(elem, headingDepth):
	if elem.nodeType != Node.ELEMENT_NODE:
		return
	if elem.hasAttribute('numbered'):
		unnumbered = (elem.getAttribute('numbered') == 'false')
	else:
		unnumbered = None	
	sectionTitle = None
	if elem.hasAttribute('title'):
		sectionTitle = elem.getAttribute('title')
	elif elem.nodeName == 'section': # Can be the case for <front> <middle> .... that are also processed by this part
		# Look after a child node of tag "name"
		nameChild = elem.getElementsByTagName('name')
		if nameChild != None:
			if nameChild[0].nodeType == Node.ELEMENT_NODE:
				sectionTitle = nameChild[0].childNodes[0].nodeValue
		else:
			print('??? This section has not title...') 
	if sectionTitle != None:
		docxBody.appendChild(docxNewParagraph(sectionTitle, 'Heading' + str(headingDepth), unnumbered = unnumbered))
	sectionId = 0
	for child in elem.childNodes:
		if child.nodeType != Node.ELEMENT_NODE:
			continue
		if child.nodeName == 'section':
			sectionId = sectionId + 1 
			# Should create a docx Child ???
			parseSection(child, headingDepth + 1)
		elif child.nodeName == 'abstract':
			parseAbstract(child)
		elif child.nodeName == 'area':
			parseArea(child)
		elif child.nodeName == 'author':
			parseAuthor(child)
		elif child.nodeName == 'boilerplate':
			parseBoilerPlate(child)
		elif child.nodeName == 'date':
			parseDate(child)
		elif child.nodeName == 'figure':
			parseFigure(child)
		elif child.nodeName == 'keyword':
			parseKeyword(child)
		elif child.nodeName == 'name': # Already processed
			continue
		elif child.nodeName == 'ol':
				parseOList(child)
		elif child.nodeName == 't':
			parseText(child, style = None)
		elif child.nodeName == 'seriesInfo':
			parseSeriesInfo(child)
		elif child.nodeName == 'textable':
			parseTextTable(child)
		elif child.nodeName == 'title':
			parseTitle(child)
		elif child.nodeName == 'toc':
			print('Skipping the ToC')
		elif child.nodeName == 'ul':
				parseUList(child)
		elif child.nodeName == 'workgroup':
			parseWorkgroup(child)
		else:
			print('!!!!! Unexpected tag in parseSection: ' + child.tagName)
 
def parseSeriesInfo(elem):
	seriesInfoString = ''
	if elem.hasAttribute('name'):
		seriesInfoString = elem.getAttribute('name')
	if elem.hasAttribute('value'):
		seriesInfoString = seriesInfoString+ elem.getAttribute('value') + ' '
	else:
		seriesInfoString = seriesInfoString + ' '
	if elem.hasAttribute('stream'):
		seriesInfoString = seriesInfoString + ' (stream: ' + elem.getAttribute('stream') + ')'
	if seriesInfoString != '':
		docxBody.appendChild(docxNewParagraph(seriesInfoString, justification = 'right'))

		
def parseText(elem, style = None, numberingID = None, indentationLevel = None):
	# Mainly for debugging
	for i in range(elem.attributes.length):
		attrib = elem.attributes.item(i)
		if attrib.name == 'pn': 	# Let's ignore this marking as no obvious requirement or support in Office OpenXML
			continue
		if attrib.name == 'indent':	# TODO later if really required
			continue
		print("\tparseText unexpected attribute: ", attrib.name, ' = ' , attrib.value)

	textValue = ''
	for text in elem.childNodes:
		if text.nodeType == Node.TEXT_NODE:
			textValue += text.nodeValue
		if elem.nodeType == Node.ELEMENT_NODE:
			if text.nodeName == 'bcp14':
				textValue = textValue + parseBcp14(text)
			elif text.nodeName == 'eref':
				textValue = textValue + parseEref(text)
			elif text.nodeName == 'list':
				docxBody.appendChild(docxNewParagraph(textValue, style = style, numberingID = numberingID, indentationLevel = indentationLevel))  # Need to emit the first part of the text
				textValue = ''
				parseList(text)
			elif text.nodeName == 'xref':
				textValue = textValue + parseXref(text)
			elif text.nodeName != '#text':
				print('!!!!! parseText: Text is ELEMENT_NODE: ', text.nodeName)
	docxBody.appendChild(docxNewParagraph(textValue, style = style, numberingID = numberingID, indentationLevel = indentationLevel))

def parseTextTable(elem):
	print('!!!!! Cannot parse TextTable')
	
def parseTitle(elem):
	global rfcTitle
	
	textValue = ''
	for text in elem.childNodes:
		if text.nodeType == Node.TEXT_NODE:
			textValue += text.nodeValue
	docxBody.appendChild(docxNewParagraph(textValue, 'Title'))
	rfcTitle = textValue 

def parseUList(elem):
	for child in elem.childNodes:
		if child.nodeType != Node.ELEMENT_NODE:
			continue
		if child.nodeName == 'li':
			parseListItem(child, numberingID = '2', indentationLevel = '0')  # numID = 2 is defined in numbering.xml as bullet list
		else:
			print('!!!! Unexpected List child: ', child.nodeName)

def parseWorkgroup(elem):
	textValue = 'Workgroup: '
	for text in elem.childNodes:
		if text.nodeType == Node.TEXT_NODE:
			textValue += text.nodeValue
		if elem.nodeType == Node.ELEMENT_NODE:
			if text.nodeName != '#text':
				print('!!!!! parseKeyword: Text is ELEMENT_NODE: ', text.nodeName)
	docxBody.appendChild(docxNewParagraph(textValue))

def parseXref(elem):	# See also https://tools.ietf.org/html/rfc7991#section-2.66
	if elem.nodeValue != None:
		print('Xref nodeValue: ' , elem.nodeValue)
	if elem.hasAttribute('target'):	# One and only mandatory attribute
		return '[' + elem.getAttribute('target') + ']'
	if elem.nodeType == Node.TEXT_NODE:
		print('Xref node is TEXT_NODE')
	# Only target attribute, so, quite useless to parse further for more attributes
	for child in elem.childNodes:
		if child.nodeType == Node.TEXT_NODE:
			return child.nodeValue
		print('!!!! parseXref, unexpected child.nodeName: ' + child.nodeName)	# Only text is allowed
							

def processXML(inFilename, outFilename = 'xml2docx.xml'):
	global xmldoc
	global docxRoot, docxBody, docxDocument
		
	xmldoc = minidom.parse(inFilename)
	rfc = xmldoc.getElementsByTagName('rfc')

	front = xmldoc.getElementsByTagName('front')[0]
	middle = xmldoc.getElementsByTagName('middle')[0]
	back = xmldoc.getElementsByTagName('back')[0]

	domImplementation = xml.dom.getDOMImplementation()
	docxRoot = domImplementation.createDocument(None, None, None)

	docxDocument = docxRoot.createElement('w:document')
	docxDocument.setAttribute('xmlns:wpc', 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas') # To be repeated for all namespaces
	docxDocument.setAttribute('xmlns:cx', 'http://schemas.microsoft.com/office/drawing/2014/chartex') 
	docxDocument.setAttribute('xmlns:cx1', 'http://schemas.microsoft.com/office/drawing/2015/9/8/chartex') 
	docxDocument.setAttribute('xmlns:cx2', 'http://schemas.microsoft.com/office/drawing/2015/10/21/chartex') 
	docxDocument.setAttribute('xmlns:cx3', 'http://schemas.microsoft.com/office/drawing/2016/5/9/chartex') 
	docxDocument.setAttribute('xmlns:cx4', 'http://schemas.microsoft.com/office/drawing/2016/5/10/chartex') 
	docxDocument.setAttribute('xmlns:cx5', 'http://schemas.microsoft.com/office/drawing/2016/5/11/chartex') 
	docxDocument.setAttribute('xmlns:cx6', 'http://schemas.microsoft.com/office/drawing/2016/5/12/chartex') 
	docxDocument.setAttribute('xmlns:cx7', 'http://schemas.microsoft.com/office/drawing/2016/5/13/chartex') 
	docxDocument.setAttribute('xmlns:cx8', 'http://schemas.microsoft.com/office/drawing/2016/5/14/chartex') 
	docxDocument.setAttribute('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006') 
	docxDocument.setAttribute('xmlns:aink', 'http://schemas.microsoft.com/office/drawing/2016/ink') 
	docxDocument.setAttribute('xmlns:am3d', 'http://schemas.microsoft.com/office/drawing/2017/model3d') 
	docxDocument.setAttribute('xmlns:o', 'urn:schemas-microsoft-com:office:office') 
	docxDocument.setAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships') 
	docxDocument.setAttribute('xmlns:m', 'http://schemas.openxmlformats.org/officeDocument/2006/math') 
	docxDocument.setAttribute('xmlns:v', 'urn:schemas-microsoft-com:vml') 
	docxDocument.setAttribute('xmlns:wp14', 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing') 
	docxDocument.setAttribute('xmlns:wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing') 
	docxDocument.setAttribute('xmlns:w10', 'urn:schemas-microsoft-com:office:word') 
	docxDocument.setAttribute('xmlns:w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main') 
	docxDocument.setAttribute('xmlns:w14', 'http://schemas.microsoft.com/office/word/2010/wordml') 
	docxDocument.setAttribute('xmlns:w15', 'http://schemas.microsoft.com/office/word/2012/wordml') 
	docxDocument.setAttribute('xmlns:w16cex', 'http://schemas.microsoft.com/office/word/2018/wordml/cex') 
	docxDocument.setAttribute('xmlns:w16cid', 'http://schemas.microsoft.com/office/word/2016/wordml/cid') 
	docxDocument.setAttribute('xmlns:w16', 'http://schemas.microsoft.com/office/word/2018/wordml') 
	docxDocument.setAttribute('xmlns:w16se', 'http://schemas.microsoft.com/office/word/2015/wordml/symex') 
	docxDocument.setAttribute('xmlns:wpg', 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup') 
	docxDocument.setAttribute('xmlns:wpi', 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk') 
	docxDocument.setAttribute('xmlns:wne', 'http://schemas.microsoft.com/office/word/2006/wordml') 
	docxDocument.setAttribute('xmlns:wps', 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape') 
	docxDocument.setAttribute('mc:Ignorable', 'w14 w15 w16se w16cid w16 w16cex wp14')	
	docxDocument.setAttribute('xmlns:w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
	docxRoot.appendChild(docxDocument)
	
	docxBody = docxRoot.createElement('w:body')
	docxDocument.appendChild(docxBody)

		
	parseSection(front, 0)
	parseSection(middle, 0)
	
	sectPrElem = docxRoot.createElement('w:sectPr')
	
	pgSzElem = docxRoot.createElement('w:pgSz')
	pgSzElem.setAttribute('w:h', '15840')
	pgSzElem.setAttribute('w:w', '12240')
	sectPrElem.appendChild(pgSzElem)

	pgMarElem = docxRoot.createElement('w:pgMar')
	pgMarElem.setAttribute('w:gutter', '0')
	pgMarElem.setAttribute('w:footer', '708')
	pgMarElem.setAttribute('w:header', '708')
	pgMarElem.setAttribute('w:left', '1440')
	pgMarElem.setAttribute('w:bottom', '1400')
	pgMarElem.setAttribute('w:right', '1440')
	pgMarElem.setAttribute('w:top', '1440')
	sectPrElem.appendChild(pgMarElem)
	
	
	colsElem = docxRoot.createElement('w:cols')
	colsElem.setAttribute('w:space', '708')
	sectPrElem.appendChild(colsElem)

	docGrid = docxRoot.createElement('w:docGrid')
	docGrid.setAttribute('w:linePitch', '360')
	sectPrElem.appendChild(docGrid)
	
	docxBody.appendChild(sectPrElem)
	
	docxFile = io.open(outFilename, 'w', encoding="'utf8'")
	# Ugly but no other way to put attributes in the top XML 
	docxFile.write(docxRoot.toprettyxml().replace('<?xml version="1.0" ?>', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'))
	docxFile.close()
	print('OpenXML document.xml file is at', outFilename)

def myParseDate(s):
	try:
		# Let's first try with short month names
		date = datetime.datetime.strptime(s,'%d %b %Y')
	except ValueError:
		# Then try with full length month names
		try:
			date = datetime.datetime.strptime(s,'%d %B %Y')
		except ValueError:
			date = datetime.datetime.utcnow()  # Giving up...
	return date
	
def generateDocPropsCore():
	xmlcore = minidom.parse(templateDirectory + '/docProps/core.xml')

	if len(rfcAuthors) > 0:
		creatorElem = xmlcore.getElementsByTagName('dc:creator')[0]
		for child in creatorElem.childNodes:
			creatorElem.removeChild(child)
		text = xmlcore.createTextNode(', '.join(rfcAuthors))
		creatorElem.appendChild(text)
	if rfcDate != None:
		createdElem = xmlcore.getElementsByTagName('dcterms:created')[0]
		for child in createdElem.childNodes:
			createdElem.removeChild(child)
		createdDate = myParseDate(rfcDate)	
		text = xmlcore.createTextNode(createdDate.strftime('%Y-%m-%dT%H:%M:%SZ'))
		createdElem.appendChild(text)
	if len(rfcKeywords) > 0:
		keywordsElem = xmlcore.getElementsByTagName('cp:keywords')[0]
		for child in keywordsElem.childNodes:
			keywordsElem.removeChild(child)
		text = xmlcore.createTextNode(', '.join(rfcKeywords))
		keywordsElem.appendChild(text)
	if rfcTitle != None:
		titleElem = xmlcore.getElementsByTagName('dc:title')[0]
		for child in titleElem.childNodes:
			titleElem.removeChild(child)
		text = xmlcore.createTextNode(rfcTitle)
		titleElem.appendChild(text)
	# Now, let's say that this script did it ;-)
	modifiedByElem = xmlcore.getElementsByTagName('cp:lastModifiedBy')[0]
	for child in modifiedByElem.childNodes:
		modifiedByElem.removeChild(child)
	text = xmlcore.createTextNode('Xml2rfc')
	modifiedByElem.appendChild(text)
	modifiedElem = xmlcore.getElementsByTagName('dcterms:modified')[0]
	for child in modifiedElem.childNodes:
		modifiedElem.removeChild(child)
	now = datetime.datetime.utcnow()
	text = xmlcore.createTextNode(now.strftime('%Y-%m-%dT%H:%M:%SZ'))
	modifiedElem.appendChild(text)
	
	return xmlcore.toprettyxml().replace('<?xml version="1.0" ?>', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
	
def docxPackage(docxFilename, openXML, templateDirectory):
	print('Generating OpenXML packaging file', docxFilename)
	print("\tUsing template in" + templateDirectory)
	coreXML = generateDocPropsCore()
	with zipfile.ZipFile(docxFilename, 'w', compression=zipfile.ZIP_DEFLATED) as docx:
		files = [ '[Content_Types].xml', '_rels/.rels', 'docProps/app.xml', 
			# Should not move the output in template directory... 'word/document.xml', 	
			'word/fontTable.xml', 'word/settings.xml', 'word/numbering.xml', 'word/webSettings.xml',
			'word/styles.xml', 'word/theme/theme1.xml', 'word/_rels/document.xml.rels']
		for file in files:
			docx.write(templateDirectory + '/' + file, arcname = file)
		docx.write(openXML, arcname = 'word/document.xml')
		docx.writestr('docProps/core.xml', coreXML)

if __name__ == '__main__':
	inFilename = None 
	outFilename = None
	templateDirectory = None
	docxFilename = None
	try:
		opts, args = getopt.getopt(sys.argv[1:],"d:hi:o:t:",["ifile=","ofile=","template=", "docx="])
	except getopt.GetoptError:
		print('xml2docx.py -i <inputfile> -o <outputfile>')
		sys.exit(2)
	for opt, arg in opts:
		if opt == '-h':
			print('xml2docx.py -i <inputfile> [-o <outputfile>] [--docx <result.docx>]')
			sys.exit()
		elif opt in ("-i", "--ifile"):
			inFilename = arg
		elif opt in ("-o", "--ofile"):
			outFilename = arg
		elif opt in ("-t", "--template"):
			templateDirectory = arg
		elif opt in ("-d", "--docx"):
			docxFilename = arg
	if templateDirectory == None: 
		templateDirectory = os.path.dirname(os.path.abspath(sys.argv[0])) + '/template' # default template is in the executable directory
	if inFilename == None:
		print('Missing input filename')
		sys.exit(2)
	if outFilename == None:
		if docxFilename != None:
			outFilename = templateDirectory + '/word/document.xml'
		else:
			outFilename = 'xml2docx.xml'
	if docxFilename == None:
		docxFilename = inFilename.replace('.xml', '.docx')

	# Let's generate the openXML word processing 'document.xml' file
	processXML(inFilename, outFilename)

	# Now, let's generate the .DOCX file
	docxPackage(docxFilename, outFilename, templateDirectory)
