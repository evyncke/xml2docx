<?php
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
#
#
if (!isset($_FILES['xmlfile']) or $_FILES['xmlfile']['error'] != '') {
	die("Cannot upload file: $_FILES[xmlfile][error]") ;
}
$remote_xmlfname = $_FILES['xmlfile']['name'] ;
$remote_docx = preg_replace('/.xml$/', '.docx', $remote_xmlfname) ;
$local_xmlfname = $_FILES['xmlfile']['tmp_name'] ;
$local_file_type = $_FILES['xmlfile']['type'] ;
$local_file_size = $_FILES['xmlfile']['size'] ;

$local_word_xml = tempnam(sys_get_temp_dir(), 'XML') . ".xml" ;
$local_docx = tempnam(sys_get_temp_dir(), 'DOC') . ".docx" ;

$shell_command = escapeshellcmd("/usr/bin/python3 ./xml2docx.py --docx $local_docx --ifile $local_xmlfname  --ofile $local_word_xml") ;
exec($shell_command, $output, $return_code) ;

# Send the right headers
header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
header("Content-Disposition: attachment; filename=\"$remote_docx\"");
readfile($local_docx) ;

exit ;
?>
<html>
<head>
<title>XML to Office OpenXML DOCX</title>
<?php
// Allow some local CSS, JS, ...
// if (is_readable('header.inc')) readfile('header.inc') ;
// ?>
</head>
<body language="en">
<h1>IETF XML2RFC file conversion into Office OpenXML .DOCX</h1>

<form enctype="multipart/form-data" action="process.php" method="post">
  <input type="hidden" name="MAX_FILE_SIZE" value="300000" />
  File to upload and convert to .DOCX : <input name="xmlfile" type="file" />
  <br/>
  <input type="submit" value="Convert the file" />
</form>

<hr>
<em>Copyright Eric Vyncke, 2020. Clone me at https://github.com/evyncke/xml2docx.git</em>
