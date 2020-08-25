#!/opt/local/bin/python
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
import sys
import io

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


#print('All children in middle:')
#for elem in middle.childNodes:
#	if elem.nodeType != Node.ELEMENT_NODE:
#		continue
#	print(elem.nodeName)
#	print("Attributes:")
#	for i in range(elem.attributes.length):
#		attrib = elem.attributes.item(i)
#		print("\t", attrib.name, ' = ' , attrib.value)
#	print("Children:")
#	for child in elem.childNodes:
#		if child.nodeType != Node.ELEMENT_NODE:
#			continue
#		pprint(child.nodeName)
#	print("\n----------\n")


def docxNewParagraph(textValue, style = None, justification = None):
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

def parseAuthor(elem):
	if elem.hasAttribute('fullname'):
		docxBody.appendChild(docxNewParagraph(elem.getAttribute('fullname'), justification = 'right'))
	else:
		author = ''
		if elem.hasAttribute('initials'):
			author = author + elem.getAttribute('initials') + ' '
		if elem.hasAttribute('surname'):
			author = author + elem.getAttribute('surname')
		if author != '':
			docxBody.appendChild(docxNewParagraph(author, justification = 'right'))
	
def parseDate(elem):
	dateString = ''
	if elem.hasAttribute('day'):
		dateString = elem.getAttribute('day') + ' '
	if elem.hasAttribute('month'):
		dateString = dateString + elem.getAttribute('month') + ' '
	if elem.hasAttribute('year'):
		dateString = dateString + elem.getAttribute('year')
	if dateString != '':
		docxBody.appendChild(docxNewParagraph(dateString, justification = 'right'))
	
def parseFigure(elem):
	print('!!!!! Cannot parse figure')
	
def parseList(elem):
#	print('Begin parseList()')
	for child in elem.childNodes:
		if child.nodeType != Node.ELEMENT_NODE:
			continue
		if child.nodeName == 't':
			parseText(child, 'ListParagraph')
		else:
			print('!!!! Unexpected List child: ', child.nodeName)
#	print('End parseList()')
		
def parseSection(elem, headingDepth, headingPrefix):
	if elem.nodeType != Node.ELEMENT_NODE:
		return
	if elem.hasAttribute('title'):
#		print(headingPrefix, elem.getAttribute('title'))
		docxBody.appendChild(docxNewParagraph(elem.getAttribute('title'), 'Heading' + str(headingDepth)))
	else:
		print('??? This section has not title...')
	sectionId = 0
	for child in elem.childNodes:
		if child.nodeType != Node.ELEMENT_NODE:
			continue
		if child.nodeName == 'section':
			sectionId = sectionId + 1 
			if headingDepth == 1:
				headingSuffix = str(sectionId)
			else:
				headingSuffix = '.' + str(sectionId)
			# Should create a docx Child ???
			parseSection(child, headingDepth + 1, headingPrefix + headingSuffix)
		elif child.nodeName == 't':
			parseText(child, style = None)
		elif child.nodeName == 'figure':
			parseFigure(child)
		elif child.nodeName == 'textable':
			parseTextTable(child)
		elif child.nodeName == 'title':
			parseTitle(child)
		elif child.nodeName == 'author':
			parseAuthor(child)
		elif child.nodeName == 'date':
			parseDate(child)
		elif child.nodeName == 'abstract':
			parseAbstract(child)
		else:
			print('!!!!! Unexpected tag:' + child.tagName)
 
def parseText(elem, style = None):
	for i in range(elem.attributes.length):
		attrib = elem.attributes.item(i)
		print("\t", attrib.name, ' = ' , attrib.value)

	textValue = ''
	for text in elem.childNodes:
		if text.nodeType == Node.TEXT_NODE:
#			print('parseText: Text is TEXT_NODE')
#			print(text.nodeValue)
			textValue += text.nodeValue
		if elem.nodeType == Node.ELEMENT_NODE:
			if text.nodeName == 'list':
				docxBody.appendChild(docxNewParagraph(textValue, style))  # Need to emit the first part of the text
				textValue = ''
				parseList(text)
			elif text.nodeName == 'xref':
				textValue = textValue + parseXref(text)
			elif text.nodeName != '#text':
				print('!!!!! parseText: Text is ELEMENT_NODE: ', text.nodeName)
	docxBody.appendChild(docxNewParagraph(textValue, style))
#	print('End parseText')

def parseTextTable(elem):
	print('!!!!! Cannot parse TextTable')
	
def parseTitle(elem):
	textValue = ''
	for text in elem.childNodes:
		if text.nodeType == Node.TEXT_NODE:
			textValue += text.nodeValue
	docxBody.appendChild(docxNewParagraph(textValue, 'Title'))

def parseXref(elem):
#	print('Begin parseXref') 
	if elem.nodeValue != None:
		print('Xref nodeValue: ' , elem.nodeValue)
	if elem.hasAttribute('target'):
		return '[' + elem.getAttribute('target') + ']'
	if elem.nodeType == Node.TEXT_NODE:
		print('Xref node is TEXT_NODE')
	# Only target attribute, so, quite useless to parse attributes
	for child in elem.childNodes:
		if child.nodeType == Node.TEXT_NODE:
			return child.nodeValue
		if child.nodeName == 't':
			parseText(child)
#	print('End parseXref')
							
if __name__ == '__main__':
#	xmldoc = minidom.parse('draft-ietf-opsec-v6-22.xml')
	xmldoc = minidom.parse('draft-ietf-intarea-provisioning-domains-00.xml')
	
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

		
	parseSection(front, 0, '')
	parseSection(middle, 0, '')
	
# Still need to add page format
#	<w:sectPr w:rsidR="008B532F" w:rsidRPr="00C46909">
#		<w:pgSz w:w="12240" w:h="15840"/>
#		<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>
#		<w:cols w:space="708"/>
#		<w:docGrid w:linePitch="360"/>
#</w:sectPr>

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
	
	
#	print(docxRoot.toprettyxml())
	docxFile = io.open('xml2doc.xml', 'w', encoding='utf8')
	docxFile.write(docxRoot.toprettyxml())
#	docxFile.write(docxRoot.toprettyxml(encoding='UTF-8'))
	docxFile.close()