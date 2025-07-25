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
   
# A lot of information in http://officeopenxml.com/anatomyofOOXML.php

import zipfile, os, sys, io
from xml2docx import xmlWriter, myParseDate
from pprint import pprint
#import xmlcore
from xml.dom import minidom
import xml.dom
import datetime

class docxWriter(xmlWriter):
  
    # This class is used to write the XML file in the docx format
    templateDirectory = None
    openXML = None  # file path for the core OpenXML document
    docxBody = None
    docxDocument = None

    def __init__(self, filename = None):
        super().__init__(filename)
        domImplementation = xml.dom.getDOMImplementation()
        self.docxRoot = domImplementation.createDocument(None, None, None)

        self.docxDocument = self.docxRoot.createElement('w:document')
        self.docxDocument.setAttribute('xmlns:wpc', 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas') # To be repeated for all namespaces
        self.docxDocument.setAttribute('xmlns:cx', 'http://schemas.microsoft.com/office/drawing/2014/chartex') 
        self.docxDocument.setAttribute('xmlns:cx1', 'http://schemas.microsoft.com/office/drawing/2015/9/8/chartex') 
        self.docxDocument.setAttribute('xmlns:cx2', 'http://schemas.microsoft.com/office/drawing/2015/10/21/chartex') 
        self.docxDocument.setAttribute('xmlns:cx3', 'http://schemas.microsoft.com/office/drawing/2016/5/9/chartex') 
        self.docxDocument.setAttribute('xmlns:cx4', 'http://schemas.microsoft.com/office/drawing/2016/5/10/chartex') 
        self.docxDocument.setAttribute('xmlns:cx5', 'http://schemas.microsoft.com/office/drawing/2016/5/11/chartex') 
        self.docxDocument.setAttribute('xmlns:cx6', 'http://schemas.microsoft.com/office/drawing/2016/5/12/chartex') 
        self.docxDocument.setAttribute('xmlns:cx7', 'http://schemas.microsoft.com/office/drawing/2016/5/13/chartex') 
        self.docxDocument.setAttribute('xmlns:cx8', 'http://schemas.microsoft.com/office/drawing/2016/5/14/chartex') 
        self.docxDocument.setAttribute('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006') 
        self.docxDocument.setAttribute('xmlns:aink', 'http://schemas.microsoft.com/office/drawing/2016/ink') 
        self.docxDocument.setAttribute('xmlns:am3d', 'http://schemas.microsoft.com/office/drawing/2017/model3d') 
        self.docxDocument.setAttribute('xmlns:o', 'urn:schemas-microsoft-com:office:office') 
        self.docxDocument.setAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships') 
        self.docxDocument.setAttribute('xmlns:m', 'http://schemas.openxmlformats.org/officeDocument/2006/math') 
        self.docxDocument.setAttribute('xmlns:v', 'urn:schemas-microsoft-com:vml') 
        self.docxDocument.setAttribute('xmlns:wp14', 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing') 
        self.docxDocument.setAttribute('xmlns:wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing') 
        self.docxDocument.setAttribute('xmlns:w10', 'urn:schemas-microsoft-com:office:word') 
        self.docxDocument.setAttribute('xmlns:w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main') 
        self.docxDocument.setAttribute('xmlns:w14', 'http://schemas.microsoft.com/office/word/2010/wordml') 
        self.docxDocument.setAttribute('xmlns:w15', 'http://schemas.microsoft.com/office/word/2012/wordml') 
        self.docxDocument.setAttribute('xmlns:w16cex', 'http://schemas.microsoft.com/office/word/2018/wordml/cex') 
        self.docxDocument.setAttribute('xmlns:w16cid', 'http://schemas.microsoft.com/office/word/2016/wordml/cid') 
        self.docxDocument.setAttribute('xmlns:w16', 'http://schemas.microsoft.com/office/word/2018/wordml') 
        self.docxDocument.setAttribute('xmlns:w16se', 'http://schemas.microsoft.com/office/word/2015/wordml/symex') 
        self.docxDocument.setAttribute('xmlns:wpg', 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup') 
        self.docxDocument.setAttribute('xmlns:wpi', 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk') 
        self.docxDocument.setAttribute('xmlns:wne', 'http://schemas.microsoft.com/office/word/2006/wordml') 
        self.docxDocument.setAttribute('xmlns:wps', 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape') 
        self.docxDocument.setAttribute('mc:Ignorable', 'w14 w15 w16se w16cid w16 w16cex wp14')	
        self.docxDocument.setAttribute('xmlns:w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
        self.docxRoot.appendChild(self.docxDocument)
        
        self.docxBody = self.docxRoot.createElement('w:body')
        self.docxDocument.appendChild(self.docxBody)
    
    def setMetaData(self, slug, value):
        super().setMetaData(slug, value)
        # Now it also needs to appear in the DOCX as capitalized slug
        self.newParagraph(slug.title() + ': ' + value)

    def _generateDocPropsCore(self):
        xmlcore = minidom.parse(self.templateDirectory + '/docProps/core.xml')

        if len(self.getMetaData('authors')) > 0:
            creatorElem = xmlcore.getElementsByTagName('dc:creator')[0]
            for child in creatorElem.childNodes:
                creatorElem.removeChild(child)
            text = xmlcore.createTextNode(', '.join(self.getMetaData('authors')))
            creatorElem.appendChild(text)
        if self.getMetaData('date') != None:
            createdElem = xmlcore.getElementsByTagName('dcterms:created')[0]
            for child in createdElem.childNodes:
                createdElem.removeChild(child)
            createdDate = myParseDate(self.getMetaData('date'))	
            text = xmlcore.createTextNode(createdDate.strftime('%Y-%m-%dT%H:%M:%SZ'))
            createdElem.appendChild(text)
        if len(self.getMetaData('keywords')) > 0:
            keywordsElem = xmlcore.getElementsByTagName('cp:keywords')[0]
            for child in keywordsElem.childNodes:
                keywordsElem.removeChild(child)
            text = xmlcore.createTextNode(', '.join(self.getMetaData('keywords')))
            keywordsElem.appendChild(text)
        if self.getMetaData('title') != None:
            titleElem = xmlcore.getElementsByTagName('dc:title')[0]
            for child in titleElem.childNodes:
                titleElem.removeChild(child)
            text = xmlcore.createTextNode(self.getMetaData('title'))
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

    def save(self): 
        super().save()
        sectPrElem = self.docxRoot.createElement('w:sectPr')
	
        pgSzElem = self.docxRoot.createElement('w:pgSz')
        pgSzElem.setAttribute('w:h', '15840')
        pgSzElem.setAttribute('w:w', '12240')
        sectPrElem.appendChild(pgSzElem)

        pgMarElem = self.docxRoot.createElement('w:pgMar')
        pgMarElem.setAttribute('w:gutter', '0')
        pgMarElem.setAttribute('w:footer', '708')
        pgMarElem.setAttribute('w:header', '708')
        pgMarElem.setAttribute('w:left', '1440')
        pgMarElem.setAttribute('w:bottom', '1400')
        pgMarElem.setAttribute('w:right', '1440')
        pgMarElem.setAttribute('w:top', '1440')
        sectPrElem.appendChild(pgMarElem)
        
        
        colsElem = self.docxRoot.createElement('w:cols')
        colsElem.setAttribute('w:space', '708')
        sectPrElem.appendChild(colsElem)

        docGrid = self.docxRoot.createElement('w:docGrid')
        docGrid.setAttribute('w:linePitch', '360')
        sectPrElem.appendChild(docGrid)
        
        self.docxBody.appendChild(sectPrElem)

        if self.templateDirectory == None: 
            self.templateDirectory = os.path.dirname(os.path.abspath(sys.argv[0])) + '/template' # default template is in the executable directory
        if self.openXML == None:
            self.openXML = self.templateDirectory + '/word/document.xml'

        docxFile = io.open(self.openXML, 'w', encoding="'utf8'")
        # Ugly but no other way to put attributes in the top XML 
        docxFile.write(self.docxRoot.toprettyxml().replace('<?xml version="1.0" ?>', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'))
        docxFile.close()
        print('OpenXML document.xml file is at', self.openXML)

        print('Generating OpenXML packaging file', self.openXML)
        print("\tUsing template in" + self.templateDirectory)
        coreXML = self._generateDocPropsCore()
        with zipfile.ZipFile(self.filename, 'w', compression=zipfile.ZIP_DEFLATED) as docx:
            files = [ '[Content_Types].xml', '_rels/.rels', 'docProps/app.xml', 
                # Should not move the output in template directory... 'word/document.xml', 	
                'word/fontTable.xml', 'word/settings.xml', 'word/numbering.xml', 'word/webSettings.xml',
                'word/styles.xml', 'word/theme/theme1.xml', 'word/_rels/document.xml.rels']
            for file in files:
                docx.write(self.templateDirectory + '/' + file, arcname = file)
            docx.write(self.openXML, arcname = 'word/document.xml')
            docx.writestr('docProps/core.xml', coreXML)

    def newParagraph(self, textValue, style = 'Normal', justification = None, unnumbered = None, 
				  numberingID = None, indentationLevel = None, removeEmpty = True, 
				  language = 'en-US', cdataSection = None):
        if textValue is None:
            return None
        if cdataSection == None:  # remove extra spaces only if CDATA is not requested
            textValue = ' '.join(textValue.split())
        if textValue == '' and removeEmpty:
            return None
        docxP = self.docxRoot.createElement('w:p')
        
    # First handle the style or justification
    #	<w:pPr>
    #			<w:pStyle w:val="Title"/>
    #			<w:jc w:val="right"/>
    #			<w:rPr>
    #				<w:lang w:val="en-US"/>
    #			</w:rPr>
    #	</w:pPr>
        pPr = self.docxRoot.createElement('w:pPr')
        if style != None:
            pStyle =  self.docxRoot.createElement('w:pStyle')
            pStyle.setAttribute('w:val', style) 
            pPr.appendChild(pStyle)
        if justification != None:
            jc =  self.docxRoot.createElement('w:jc')
            jc.setAttribute('w:val', justification) 
            pPr.appendChild(jc)
        if unnumbered:  # Try to override the default numbering in the style
            numPr = self.docxRoot.createElement('w:numPr')
            ilvl = self.docxRoot.createElement('w:ilvl ')
            ilvl.setAttribute('w:val', 0)
            numPr.appendChild(ilvl)
            numId = self.docxRoot.createElement('w:numId')
            numId.setAttribute('w:val', 0)
            numPr.appendChild(numId)
            pPr.appendChild(numPr)
        elif numberingID != None and indentationLevel != None:
    #				<w:numPr>
    #					<w:ilvl w:val="0"/>
    #					<w:numId w:val="2"/>
    #				</w:numPr>
            numPr = self.docxRoot.createElement('w:numPr')
            ilvl = self.docxRoot.createElement('w:ilvl ')
            ilvl.setAttribute('w:val', indentationLevel)
            numPr.appendChild(ilvl)
            numId = self.docxRoot.createElement('w:numId')
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
        r = self.docxRoot.createElement('w:r')
        rPr = self.docxRoot.createElement('w:rPr')
        if language != None:
            lang = 	self.docxRoot.createElement('w:lang')
            lang.setAttribute('w:val', language)
            rPr.appendChild(lang)
        elif style != None:  # Seems mandatory for figure ASCII art to repeat the style per run
            rStyle = self.docxRoot.createElement('w:rStyle')
            rStyle.setAttribute('w:val', style)
            rPr.appendChild(rStyle)
        r.appendChild(rPr)
        t = self.docxRoot.createElement('w:t')
        if cdataSection == None:
            text = self.docxRoot.createTextNode(textValue)
        else:
            t.setAttribute('xml:space', 'preserve')
            text = self.docxRoot.createTextNode(textValue)
    #		text = docxRoot.createCDATASection(textValue)   # xml:space is enough to keep leading spaces, CDATA adds 4 tabs after in the pretty printing :-(
        t.appendChild(text)
        r.appendChild(t) 
        docxP.appendChild(r)
        self.docxBody.appendChild(docxP)
