#!/bin/sh -e
./xml2docx.py $1
mv xml2doc.xml template/word/document.xml 
cd template/
zip -r ../test.docx *
cd -
