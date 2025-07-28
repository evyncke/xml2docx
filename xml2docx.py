#!/usr/bin/env python3
# -*- coding: utf-8 -*-

#   Copyright 2020-2025, Eric Vyncke, evyncke@cisco.com
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

# XML2RFC is in
# https://datatracker.ietf.org/doc/html/rfc7998 "xml2rfc" Version 3 Preparation Tool Description
# https://datatracker.ietf.org/doc/rfc7991/ The "xml2rfc" Version 3 Vocabulary
# TODO
# Handle external entities used notably for references...
# https://www.w3schools.com/xml/xml_dtd_entities.asp
#  example 
#    <?xml version='1.0'?>
#        <!DOCTYPE rfc SYSTEM 'rfc2629.dtd' [
#        <!ENTITY rfc2629 PUBLIC '' 'http://xml2rfc.ietf.org/public/rfc/bibxml/reference.RFC.2629.xml'>
#       ]>
   
from xml.dom import minidom, Node
import xml.dom
from pprint import pprint
import sys, getopt
import os

import datetime
import urllib.request
import docxWriter

class xmlWriter:
	# Same states to be kept
	metaData = {}  # A dict for slugs: authors, date, keywords, title
	abstract = []  # A list of paragraphs in the abstract
	filename = None  # The filename of the to-be-created file

	def __init__(self, filename = None):
		self.filename = filename

	def save(self):
		pass

	def getMetaData(self, slug):
		if slug in self.metaData:
			return self.metaData[slug]
		else:
			return None
	
	def setMetaData(self, slug, value):
		if slug in self.metaData:
			self.metaData[slug].append(value)
		else:
			self.metaData[slug] = [value]

	def newParagraph(self, textValue, style = 'Normal', justification = None, unnumbered = None, 
				  numberingID = None, indentationLevel = None, removeEmpty = True, 
				  language = 'en-US', cdataSection = None):
		# As parseText() is the same  for front and body elements 
		if style is not None and style == "Abstract":
			self.abstract.append(textValue)
		pass

# For debugging purpose
def printTree(front):
	print('All children:')
	for elem in front.childNodes:
		if elem.nodeType == Node.TEXT_NODE:
			print("\t TEXT: '", elem.nodeValue, "'")
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

def docxNewParagraph(textValue, style = 'Normal', justification = None, unnumbered = None, numberingID = None, indentationLevel = None, removeEmpty = True, language = 'en-US', cdataSection = None):
	return writer.newParagraph(textValue, style, justification, unnumbered, numberingID, indentationLevel, removeEmpty, language, cdataSection)

libsTable = { 
	'RFC': 'https://www.rfc-editor.org/refs/bibxml/',
	'I-D': 'http://xml2rfc.ietf.org/public/rfc/bibxml3/',
	'W3C': 'http://xml2rfc.ietf.org/public/rfc/bibxml4/',
	'SDO-3GPP': 'http://xml2rfc.ietf.org/public/rfc/bibxml5/',
	'IEEE': 'http://xml2rfc.ietf.org/public/rfc/bibxml6/',
	'DOI': 'http://xml2rfc.ietf.org/public/rfc/bibxml7/',
	'BCP': 'http://xml2rfc.ietf.org/public/rfc/bibxml9/',
	'FYI': 'http://xml2rfc.ietf.org/public/rfc/bibxml9/',
	'STD': 'http://xml2rfc.ietf.org/public/rfc/bibxml9/',
}

def includeExternal(referenceName):
	global libsTable
	
	referenceTokens = referenceName.split('.')
	if libsTable.get(referenceTokens[1]):
		libURL = libsTable.get(referenceTokens[1])
		print("Importing " + referenceName + " from " + libURL + referenceName + '.xml')
		try:
			response = urllib.request.urlopen(libURL + referenceName + '.xml')
			importedString = response.read()
			importedXML = minidom.parseString(importedString)
		except urllib.error.HTTPError as err:
			print("Cannot import XML from " +  libURL + referenceName + ".xml, error: ", err)
			return None
		except:
			print('Not found or invalid XML in ' + libURL)
			return None
		return importedXML.getElementsByTagName('reference')[0]
	print("Reference type " + referenceTokens[1] + " not supported...")
	return None
	
def parseAbstract(elem):
	for child in elem.childNodes:
		if child.nodeType != Node.ELEMENT_NODE:
			continue
		elif child.nodeName == 't':
			parseText(child, style = 'Abstract')
		else:
			print('Unexpected tagName in Abstract: ', child.nodeName)

def parseArea(elem):
	textValue = 'Area: '
	for text in elem.childNodes:
		if text.nodeType == Node.TEXT_NODE:
			textValue += text.nodeValue
		if elem.nodeType == Node.ELEMENT_NODE:
			if text.nodeName != '#text':
				print('!!!!! parseArea: Text is ELEMENT_NODE: ', text.nodeName)
	writer.newParagraph(textValue)

def parseArtWork(elem):	# See also https://tools.ietf.org/html/rfc7991#section-2.5
	# If there is no type attribute, let's process the element
	# If there is a type attribute, let's process the element only if type == ascii-art
	if (not elem.hasAttribute('type')) or (elem.hasAttribute('type') and (elem.getAttribute('type') == 'ascii-art' or elem.getAttribute('type') == '')):
		figureLines = ''
		for chunk in elem.childNodes:	
			text = chunk.nodeValue
			figureLines += text
		# Let's split this string into lines and print each line
		for line in figureLines.splitlines():
			writer.newParagraph(line.rstrip(" \t"), style = 'Code', removeEmpty = False, language = None, cdataSection = True)

def parseAuthor(elem):	# Per https://tools.ietf.org/html/rfc7991#section-2.7
	global writer

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
		writer.newParagraph(elem.getAttribute('asciiFullname') + organization, justification = 'right')
		writer.setMetaData('authors', elem.getAttribute('asciiFullname') + organization)
	elif elem.hasAttribute('fullname'):
		writer.newParagraph(elem.getAttribute('fullname') + organization, justification = 'right')
		writer.setMetaData('authors', elem.getAttribute('fullname') + organization)
	else:
		author = ''
		if elem.hasAttribute('initials'):
			author = author + elem.getAttribute('initials') + ' '
		if elem.hasAttribute('surname'):
			author = author + elem.getAttribute('surname')
		if author != '':
			writer.newParagraph(author + organization, justification = 'right')
			writer.setMetaData('authors', author + organization)

def parseBack(elem): # https://tools.ietf.org/html/rfc7991#section-2.8
	if elem.nodeType != Node.ELEMENT_NODE:
		return
	for child in elem.childNodes:
		if child.nodeType != Node.ELEMENT_NODE:
			continue
		if child.nodeName == 'displayreference':
			parseDisplayReference(child)
		elif child.nodeName == 'references':
			parseReferences(child, 1)
		elif child.nodeName == 'section':
			parseSection(child, 2)
		else:
			print('!!!! parseBack: unexpected nodeName: ' + child.nodeName)

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
	
def parseBlockQuote(elem): # See also https://tools.ietf.org/html/rfc7991#section-2.10 that is similar to old <list> items
	parseText(elem, style = 'Quote', numberingID = None, indentationLevel = None)

def parseBoilerPlate(elem):
	for child in elem.childNodes:
		if child.nodeType != Node.ELEMENT_NODE:
			continue
		elif child.nodeName == 'section':
			parseSection(child, 1)
		else:
			print('Unexpected tagName in BoilerPlate: ', child.nodeName)

def parseDate(elem):
	global writer
	
	dateString = ''
	if elem.hasAttribute('day'):
		dateString = elem.getAttribute('day') + ' '
	if elem.hasAttribute('month'):
		dateString = dateString + elem.getAttribute('month') + ' '
	if elem.hasAttribute('year'):
		dateString = dateString + elem.getAttribute('year')
	if dateString != '':
		writer.newParagraph(dateString, justification = 'right')
		writer.rfcDate = dateString
	
def parseDisplayReference(elem): # https://tools.ietf.org/html/rfc7991#section-2.19
	# Presentation only... skipping it for now
	return
#	print("parseDisplayReference not yet implemented")
	
def parseDList(elem):  # See also https://tools.ietf.org/html/rfc7991#section-2.20 
	for child in elem.childNodes:
	# If should be a serie of DT DD elements in the right order, the code is not resilient to out of order
		if child.nodeType != Node.ELEMENT_NODE:
#			print("parseDList unexpected node type...", child) # TODO sometimes it is CRLF + white spaces possibly for indentation ?
			continue
		if child.nodeName == 'dt':	# Definition Term https://tools.ietf.org/html/rfc7991#section-2.21
			# Can contain text + some other elements
			parseText(child)
		elif child.nodeName == 'dd': # Definition part https://tools.ietf.org/html/rfc7991#section-2.18
			# Can contain text + some other elements including complex ones
			parseText(child)
		else:
			print('!!!! parseDList, unexpected child: ', child.nodeName)

# TODO switch off language to avoid wrong typos ?
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

def parseFigure(elem): # See https://tools.ietf.org/html/rfc7991#section-2.25
	# Figure had preamble (deprecated but let's process it)
	preambleChildren = elem.getElementsByTagName('preamble')
	if preambleChildren.length > 0 and preambleChildren[0].childNodes.length > 0:
		if preambleChildren[0].nodeType == Node.ELEMENT_NODE:
			preamble = preambleChildren[0].childNodes[0].nodeValue
			writer.newParagraph(preamble)
	# Let's process a single artwork
	artworkChildren = elem.getElementsByTagName('artwork')
	for child in artworkChildren:
		parseArtWork(child)
	# Let's process the source code
	
	# Could have a title attribute rather than the name element (same as in section)
	if elem.nodeType != Node.ELEMENT_NODE:
		return
	figureTitle = None
	if elem.hasAttribute('title'):
		figureTitle = elem.getAttribute('title')
	else:
		nameChild = elem.getElementsByTagName('name')
		if nameChild.length > 0:
			if nameChild[0].nodeType == Node.ELEMENT_NODE:
				figureTitle = nameChild[0].childNodes[0].nodeValue
	if figureTitle != None:
		writer.newParagraph('Figure: ' + figureTitle, justification = 'center')
	# Figure had postamble (deprecated but let's process it)
	postambleChildren = elem.getElementsByTagName('postamble')
	if postambleChildren.length > 0  and postambleChildren[0].childNodes.length > 0:
		if postambleChildren[0].nodeType == Node.ELEMENT_NODE:
			postamble = postambleChildren[0].childNodes[0].nodeValue
			writer.newParagraph(postamble)
	
def parseKeyword(elem):
	
	for text in elem.childNodes:
		if text.nodeType == Node.TEXT_NODE:
			writer.setMetaData('keywords', text.nodeValue)
		if elem.nodeType == Node.ELEMENT_NODE:
			if text.nodeName != '#text':
				print('!!!!! parseKeyword: Text is ELEMENT_NODE: ', text.nodeName)

def parseList(elem):  # See also https://tools.ietf.org/html/rfc7991#section-2.29
	for child in elem.childNodes:
		if child.nodeType == Node.COMMENT_NODE:
			continue
		elif child.nodeType == Node.TEXT_NODE: # Unexpected, let's hope it is empty space
			if child.nodeValue.strip(" \t\r\n") == '':
				continue
			print("!!!! parseList non empty text = '" + child.nodeValue.strip(" \t\r\n") + "'")
			continue
		elif child.nodeType != Node.ELEMENT_NODE:
			print('!!!! parseList, unexpected child node type: ', child)
			continue
		if child.nodeName == 't':
			parseText(child, style = 'ListParagraph', numberingID = '2', indentationLevel = '0')  # numID = 2 is defined in numbering.xml as bullet list
		else:
			print('!!!! parseList, unexpected child: ', child.nodeName)
		
def parseListItem(elem, style = 'ListParagraph', numberingID = None, indentationLevel = None):
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
			elif text.nodeName == 'ol':
				writer.newParagraph(textValue, style = style, numberingID = numberingID, indentationLevel = indentationLevel)
				textValue = ''
				parseOList(text)
			elif text.nodeName == 't':
				writer.newParagraph(textValue, style = style, numberingID = numberingID, indentationLevel = indentationLevel)
				parseText(text)
			elif text.nodeName == 'ul':
				writer.newParagraph(textValue, style = style, numberingID = numberingID, indentationLevel = indentationLevel)
				parseUList(text)
			elif text.nodeName == 'xref':
				textValue = textValue + parseXref(text)
			elif text.nodeName != '#text':
				print('!!!!! parseListItem: Text is ELEMENT_NODE: ', text.nodeName)
#			else:
#				print('parseListItem ignoring Text is ELEMENT_NODE: ', text.nodeName)
	writer.newParagraph(textValue, style = style, numberingID = numberingID, indentationLevel = indentationLevel)

def parseNote(elem):  # See https://tools.ietf.org/html/rfc7991#section-2.33
	print("<note> is an unsupported tag")
	
# TODO should reset the numbering to 1... cfr draft-ietf-anima-autonomic-control-plane-29.xml
def parseOList(elem):
	for child in elem.childNodes:
		if child.nodeType != Node.ELEMENT_NODE:
			continue
		if child.nodeName == 'li':
			parseListItem(child, numberingID = '1', indentationLevel = '0')  # numID = 1 is defined in numbering.xml as enumeration list
		else:
			print('!!!! Unexpected List child: ', child.nodeName)

def parseReference(elem):  # See https://tools.ietf.org/html/rfc7991#section-2.40
	if elem.nodeType != Node.ELEMENT_NODE:
		return
	if elem.hasAttribute('anchor'):
		text = '[' + elem.getAttribute('anchor') + ']  '
	else:
		print('!!!! parseReference, missing anchor attribute')
		text = ''
	seriesInfoText = ''
	for serieInfo in elem.getElementsByTagName('seriesInfo'):
		if serieInfo.hasAttribute('name') and serieInfo.hasAttribute('value'):
			if serieInfo.getAttribute('value') == '': # Sometimes the value field is empty... no need to add a useless space
				seriesInfoText += serieInfo.getAttribute('name') + ' ' + serieInfo.getAttribute('value') + ', '
			else:
				seriesInfoText += serieInfo.getAttribute('name') + ', '
		else:
			print("!!!! parseReference, no name/value attribute in seriesInfo for " + text)
	frontElem = elem.getElementsByTagName('front')[0]
	if frontElem:
		for author in frontElem.getElementsByTagName('author'):
			authorName = '?' # Could also simply be in the child elemn <organization>
			if author.hasAttribute('surname'):
				if author.hasAttribute('initials'):
					authorName = author.getAttribute('surname') + ', ' + author.getAttribute('initials')
				else:
					authorName = author.getAttribute('surname')
			elif author.hasAttribute('fullname'):
				authorName = author.getAttribute('fullname')
			else:   # Let's find the <organization> element
				orgElems = frontElem.getElementsByTagName('organization')
				if orgElems:
					orgElem = orgElems[0]
					if orgElem:
						authorName = ''
						for child in orgElem.childNodes:
							if child.nodeType == Node.TEXT_NODE:
								authorName += child.nodeValue
			text += authorName + ', '
		if frontElem.getElementsByTagName('title'):
			titleElem = frontElem.getElementsByTagName('title')[0]
			for child in titleElem.childNodes:
				if child.nodeType == Node.TEXT_NODE:
					text += '"' + child.nodeValue + '", '
		# Insert seriesInfo if any
		text += seriesInfoText
		if frontElem.getElementsByTagName('date'):
			dateElem = frontElem.getElementsByTagName('date')[0]
			if dateElem.hasAttribute('year'):
				if dateElem.hasAttribute('month'):
					text += dateElem.getAttribute('month') + ' ' + dateElem.getAttribute('year') + ', '
				else:
					text += dateElem.getAttribute('year') + ', '
	else: # In the absence of <front> element
		text += seriesInfoText

	if elem.hasAttribute('target'):
		text += elem.getAttribute('target')
	# Let's remove any trailing comma
	if text[-2:] == ', ':
		text = text[:-2]
	text += '.'
	writer.newParagraph(text)

def parseReferences(elem, headingLevel = 1): # https://tools.ietf.org/html/rfc7991#section-2.42
	if elem.nodeType != Node.ELEMENT_NODE:
		return
	sectionTitle = None
	if elem.hasAttribute('title'):
		sectionTitle = elem.getAttribute('title')
	else:
		nameChild = elem.getElementsByTagName('name')
		if nameChild.length > 0:
			if nameChild[0].nodeType == Node.ELEMENT_NODE:
				sectionTitle = nameChild[0].childNodes[0].nodeValue
		else:
			print(elem)
			print('??? parseReferences: this references section has not title...') 
	if sectionTitle != None:
		writer.newParagraph(sectionTitle, 'Heading' + str(headingLevel), unnumbered = None)
	for child in elem.childNodes:
		if child.nodeType == Node.PROCESSING_INSTRUCTION_NODE: # in this location it is probably <?rfc include='reference.RFC.2119'?> or <?rfc include='reference.I-D.ietf-emu-eaptlscert'?> 
			if child.target == 'rfc' and (child.data[0:9] == "include='" or child.data[0:9] == 'include="'):
				includeName = child.data[9:-1]
				child = includeExternal(includeName)
				if child is None:
					continue
			else:
				print("parseReferences: skipping unknown processing instruction: target = " + child.target + ", data = " + child.data[0:9]) 
		if child.nodeType == Node.TEXT_NODE:  # Let's skip whitespace (assuming it is white space...)
			continue
		if child.nodeType != Node.ELEMENT_NODE:
			print('!!!! parseReferences: unexpected nodeType: ', child)
			continue
		if child.nodeName == 'reference':
			parseReference(child)
		elif child.nodeName == 'references':
			parseReferences(child, headingLevel + 1)
		elif child.nodeName != 'name': # <name> is already processed
			print('!!!! parseReferences: unexpected nodeName: ' + child.nodeName)

def parseName(elem): 
	pass # EVY ?

def parseRfc(elem):  # See also https://tools.ietf.org/html/rfc7991#section-2.45 
	if elem.nodeType != Node.ELEMENT_NODE:
		return
	rfcInfo = ''
	if elem.hasAttribute('category'):
		writer.setMetaData('category', elem.getAttribute('category'))
		# docxBody.appendChild(docxNewParagraph('Category: ' + elem.getAttribute('category')))
	if elem.hasAttribute('submissionType'):
		writer.setMetaData('submissiontype', elem.getAttribute('submissionType'))
		# docxBody.appendChild(docxNewParagraph('Submission type: ' + elem.getAttribute('submissionType')))
	if elem.hasAttribute('obsoletes'):
		writer.setMetaData('obsoletes', elem.getAttribute('obsoletes'))
		# docxBody.appendChild(docxNewParagraph('Obsoletes: ' + elem.getAttribute('obsoletes')))
	if elem.hasAttribute('updates'):
		writer.setMetaData('updates', elem.getAttribute('updates'))
		# docxBody.appendChild(docxNewParagraph('Updates: ' + elem.getAttribute('updates')))

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
		if nameChild.length > 0:
			if nameChild[0].nodeType == Node.ELEMENT_NODE:
				sectionTitle = nameChild[0].childNodes[0].nodeValue
		else:
			print('??? This section has not title...') 
	if sectionTitle != None:
		writer.newParagraph(sectionTitle, 'Heading' + str(headingDepth), unnumbered = unnumbered)
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
		elif child.nodeName == 'artwork':
			parseArtWork(child)
		elif child.nodeName == 'author':
			parseAuthor(child)
		elif child.nodeName == 'blockquote':
			parseBlockQuote(child)
		elif child.nodeName == 'boilerplate':
			parseBoilerPlate(child)
		elif child.nodeName == 'date':
			parseDate(child)
		elif child.nodeName == 'dl':
			parseDList(child)
		elif child.nodeName == 'figure':
			parseFigure(child)
		elif child.nodeName == 'keyword':
			parseKeyword(child)
		elif child.nodeName == 'name': # Already processed
			continue
		elif child.nodeName == 'note':
				parseNote(child)
		elif child.nodeName == 'ol':
				parseOList(child)
		elif child.nodeName == 't':
			parseText(child, style = None)
		elif child.nodeName == 'seriesInfo':
			parseSeriesInfo(child)
		elif child.nodeName == 'texttable':
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
 
# TODO handle wrongly formatted    <seriesInfo name="Internet-Draft" value="draft-ietf-anima-autonomic-control-plane-29"/>
def parseSeriesInfo(elem):
	seriesInfoString = ''
	if elem.hasAttribute('name'):
		seriesInfoString = elem.getAttribute('name') + ' '
	if elem.hasAttribute('value'):
		seriesInfoString = seriesInfoString + elem.getAttribute('value') + ' '
	else:
		seriesInfoString = seriesInfoString
	if elem.hasAttribute('stream'):
		seriesInfoString = seriesInfoString + ' (stream: ' + elem.getAttribute('stream') + ')'
	if seriesInfoString != '':
		writer.newParagraph(seriesInfoString, justification = 'right')

		
def parseText(elem, style = None, numberingID = None, indentationLevel = None, Verbose = None):  # See https://tools.ietf.org/html/rfc7991#section-2.53
	if Verbose:
		print("parseText start: ", elem)
	textValue = ''
	# Mainly for debugging
	for i in range(elem.attributes.length):
		attrib = elem.attributes.item(i)
		if attrib.name == 'hangText':
			textValue = attrib.value
			continue
		if attrib.name == 'pn': 	# Let's ignore this marking as no obvious requirement or support in Office OpenXML
			continue
		if attrib.name == 'indent':	# TODO later if really required
			continue
		if attrib.name == 'keepWithNext':	# TODO later if really required
			continue
		print("\tparseText unexpected attribute: ", attrib.name, '=' , attrib.value)

	for text in elem.childNodes:
		if text.nodeType == Node.TEXT_NODE:
			textValue += text.nodeValue
			if Verbose:
				print("parseText adding TEXT_NODE: '", text.nodeValue, "'")
		if elem.nodeType == Node.ELEMENT_NODE:
			if text.nodeName == 'bcp14':
				textValue = textValue + parseBcp14(text)
			elif text.nodeName == 'eref':
				textValue = textValue + parseEref(text)
			elif text.nodeName == 'figure':
				writer.newParagraph(textValue, style = style, numberingID = numberingID, indentationLevel = indentationLevel)
				textValue = ''
				parseFigure(text)
			elif text.nodeName == 'list':
				writer.newParagraph(textValue, style = style, numberingID = numberingID, indentationLevel = indentationLevel)
				textValue = ''
				parseList(text)
			elif text.nodeName == 'ol':
				writer.newParagraph(textValue, style = style, numberingID = numberingID, indentationLevel = indentationLevel)
				textValue = ''
				parseOList(text)
			elif text.nodeName == 't':
				writer.newParagraph(textValue, style = style, numberingID = numberingID, indentationLevel = indentationLevel)
				if Verbose:
					print("parseText found <t>: emitting '", textValue, "'")
				textValue = ''
				parseText(text, style = style, numberingID = numberingID, indentationLevel = indentationLevel, Verbose = Verbose)
			elif text.nodeName == 'vspace':
				writer.newParagraph(textValue, style = style, numberingID = numberingID, indentationLevel = indentationLevel)
				# Now force an empty paragraph
				writer.newParagraph('', style = style, removeEmpty = False)
				textValue = ''
			elif text.nodeName == 'ul':
				writer.newParagraph(textValue, style = style, numberingID = numberingID, indentationLevel = indentationLevel)
				textValue = ''
				parseUList(text)
			elif text.nodeName == 'xref':
				textValue = textValue + parseXref(text)
			elif text.nodeName == 'tt': # Fixed font
				# textValue = textValue + parseText(text, style = 'Code', numberingID = None, indentationLevel = None, Verbose = True)
				parseText(text, style = 'Code', numberingID = None, indentationLevel = None, Verbose = True)
#			elif text.nodeName == 'em': # italics font
#				textValue = textValue + parseText(text, style = 'Emphasis', numberingID = None, indentationLevel = None)
			elif text.nodeName != '#text' and text.nodeName != '#comment':
				print('!!!!! parseText: Text is ELEMENT_NODE: ', text.nodeName)
	writer.newParagraph(textValue, style = style, numberingID = numberingID, indentationLevel = indentationLevel)

def parseTextTable(elem):
	print('Skipping TextTable')
	writer.newParagraph('... a TextTable was not imported...', justification = 'center')
	
def parseTitle(elem):
	
	textValue = ''
	for text in elem.childNodes:
		if text.nodeType == Node.TEXT_NODE:
			textValue += text.nodeValue
	writer.newParagraph(textValue, style = 'Title')
	writer.setMetaData('title', textValue)

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
				print('!!!!! parseWorkgroup: Text is ELEMENT_NODE: ', text.nodeName)
	writer.newParagraph(textValue)

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
	
	if os.path.isfile(inFilename):
		xmldoc = minidom.parse(inFilename)
	else:
		try:
			response = urllib.request.urlopen('https://datatracker.ietf.org/doc/id/' + inFilename + '.xml')
		except:
			print("Cannot fetch the XML document from the IETF site...")
			sys.exit(1)
		draftString = response.read()
		xmldoc = minidom.parseString(draftString)
		print("Fetching the draft from the IETF site...")
		
	rfc = xmldoc.getElementsByTagName('rfc')[0]

	front = rfc.getElementsByTagName('front')[0]
	middle = rfc.getElementsByTagName('middle')[0]
	back = rfc.getElementsByTagName('back')[0]


	parseRfc(rfc)
	parseSection(front, 0)
	parseSection(middle, 0)
	parseBack(back)
	

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
			print('xml2docx.py -i <inputfile/draft-name> [-o <outputfile>] [--docx <result.docx>]')
			sys.exit()
		elif opt in ("-i", "--ifile"):
			inFilename = arg
		elif opt in ("-o", "--ofile"):
			outFilename = arg
		elif opt in ("-t", "--template"):
			templateDirectory = arg
		elif opt in ("-d", "--docx"):
			docxFilename = arg
	if inFilename == None:
		print('Missing input filename')
		sys.exit(2)

	if outFilename == None:
		if docxFilename != None:
			outFilename = templateDirectory + '/word/document.xml'
		else:
			outFilename = 'xml2docx.xml'
	if docxFilename == None:
		if inFilename[-4:] == '.xml':
			docxFilename = inFilename.replace('.xml', '.docx')
		else:
			docxFilename = inFilename + '.docx'

	# Assuming DOCX write from now		
	writer = docxWriter.docxWriter(docxFilename)
	writer.templateDirectory = templateDirectory

	# Let's generate the openXML word processing 'document.xml' file
	processXML(inFilename, outFilename)

	# Now, let's generate the .DOCX file
	writer.save()