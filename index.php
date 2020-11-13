<html>
<!--
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
#-->
<head>
<title>XML to Office OpenXML .DOCX</title>
<?php
// Allow some local CSS, JS, ...
if (is_readable('header.inc')) readfile('header.inc') ;
?>
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
