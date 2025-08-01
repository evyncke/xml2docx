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
   
# A lot of information in https://github.com/cabo/kramdown-rfc/wiki/Syntax2 

from xml2docx import xmlWriter
import textwrap

class mdWriter(xmlWriter):
  
    # This class is used to write the XML file in the Markdown format
    mdMiddleText = [] # A list of all the paragraphs in the middle part
    mdBackText = []  # A list of all the paragraphs in the back part

    def __init__(self, filename = None):
        super().__init__(filename)
    
    def setMetaData(self, slug, value):
        super().setMetaData(slug, value)

    def _saveFront(self, f):
        # Save the front matter in Markdown format
        f.write("---\ncoding: utf-8\nstand_alone: yes\npi: [toc, sortrefs, symrefs, comments]\n")
        for slug, value in self.metaData.items():
            if (slug == 'authors'):
                f.write('author:\n')
                for author in value:
                    f.write(f'- {author}\n')
                continue
            elif isinstance(value, list):
                textValue = ', '.join(value)
            elif isinstance(value, str):
                textValue = value
            else:
                textValue = str(value)
            f.write(f'{slug}: {textValue}\n')
        if len(self.normativeReferences) > 0:
            f.write('\nnormative:\n')
            for ref in self.normativeReferences:
                f.write(f"\t{ref}:\n")
        if len(self.informativeReferences) > 0:
            f.write('\ninformative:\n')
            for ref in self.informativeReferences:
                f.write(f"\t{ref}:\n")
        if len(self.abstract) > 0:
            abstractText = ' '.join(self.abstract)
            if len(abstractText) > 0:
                f.write(f'\n--- abstract\n\n')
                for paragraph in self.abstract:
                    f.write(textwrap.fill(paragraph, width=80))
                f.write('\n\n')
        f.write('\n--- middle\n\n')

    def save(self): 
        super().save()
        print('Generating kramdown file', self.filename)
        with open(self.filename, 'w', encoding='utf-8') as f:
            self._saveFront(f)
            for paragraph in self.mdMiddleText:
                f.write(textwrap.fill(paragraph, width=72))
                if paragraph.endswith('\n'):
                    f.write('\n')  # as textwrap.fill remove the trailing new line if any (e.g., tables/figures do not have a trailing new line)
                f.write('\n')  # Add one new line between paragraphs (not the newParagraph also adds a new line but not newTable)
            f.write('\n--- back\n\n')
            for paragraph in self.mdBackText:
                f.write(textwrap.fill(paragraph, width=72))
                f.write('\n')

    def newParagraph(self, textValue, style = 'Normal', justification = None, unnumbered = None, 
				  numberingID = None, indentationLevel = None, removeEmpty = True, 
				  language = 'en-US', cdataSection = None):
        if textValue is None:
            return
        if cdataSection is None:  # remove extra spaces only if CDATA is not requested
            textValue = ' '.join(textValue.split())
        if textValue == '' and removeEmpty:
            return
        if justification is not None:
            pass
        if unnumbered:  # Try to override the default numbering in the style
            pass
        elif numberingID is not None and indentationLevel is not None:
            pass
        if cdataSection is None:
            pass
        if style is not None and style == 'Abstract':
            self.abstract.append(textValue)
            return
        if style is not None and style == 'Title':
            self.title = textValue
            return
        if style is not None and style.startswith('Heading'):
            # Convert Heading styles to Markdown headings
            level = style.replace('Heading', '')
            if level.isdigit():
                level = int(level)
                if 1 <= level <= 6:
                    textValue = '#' * level + ' ' + textValue
        if self.inMiddle:
            self.mdMiddleText.append(textValue + '\n')
        else:
            self.mdBackText.append(textValue + '\n')

    def newTable(self, table):
        needHeaderSeparator = False
        for row in table.rows:
            if row.rowType == 'thead': 
                # Add a header row
                textValue = "| " + " | ".join([cell.text for cell in row.cells]) + " |"
                needHeaderSeparator = True
            elif row.rowType in ['tbody', 'tfoot']:  # markdown does not have a tfoot
                if needHeaderSeparator:
                    # Add a separator row
                    textValue = "| " + " | ".join(['---'] * len(row.cells)) + " |"
                    if self.inMiddle:
                        self.mdMiddleText.append(textValue)
                    else:
                        self.mdBackText.append(textValue)
                    needHeaderSeparator = False
                # Add a body row
                textValue = "| " + " | ".join([cell.text for cell in row.cells]) + " |"
            else: 
                print('newTable: rowType not handled:', row.rowType)
            if self.inMiddle:
                self.mdMiddleText.append(textValue)
            else:
                self.mdBackText.append(textValue)
        # Write the table caption if any
        if table.name:
            self.newParagraph(table.name, style = 'Caption', justification = 'center')

    def newFigure(self, figure):
        if self.inMiddle:
            self.mdMiddleText.append('{:fig: artwork-align="center"}')
            self.mdMiddleText.append('~~~~')
        else:
            self.mdBackText.append('{:fig: artwork-align="center"}')
            self.mdBackText.append('~~~~')
        for row in figure.rows:
            if self.inMiddle:
                self.mdMiddleText.append(row)
            else:
                self.mdBackText.append(row)
        # Write the table caption if any
        if self.inMiddle:
            self.mdMiddleText.append('~~~~')
        else:
            self.mdBackText.append('~~~~')
        if figure.name:
            if self.inMiddle:
                self.mdMiddleText.append('{:fig title="' + figure.name + '"}')
            else:
                self.mdBackText.append('{:fig title="' + figure.name + '"}')
