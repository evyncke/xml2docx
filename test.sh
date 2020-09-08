#!/bin/sh -e
./xml2docx.py -i $1
mv xml2docx.xml template/word/document.xml 
cd template/
zip -r ../test.docx *
cd -
zip --delete test.docx word/.DS_Store 
