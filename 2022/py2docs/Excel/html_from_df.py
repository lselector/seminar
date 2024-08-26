"""
# test_df_to_html.py
# test script to convert a simple test Chinese/English pandas 
# DataFrame into an HTML page, while also embedding Excel 
# binary (base64-encoded), and providing a link to 
# download the binary Excel file
#
# Here are some links:
# - https://www.linkedin.com/pulse/how-transfer-binary-files-pdfimages-etc-json-rahul-budholiya/
# - http://rahulbudholiya.blogspot.com
# Note: We use utf-8 and base64 encoding (to avoid problems with browser)
# Here is how to convert base64 string to binary blob in Javascript:
#  - https://stackoverflow.com/questions/16245767/creating-a-blob-from-a-base64-string-in-javascript
# Here is how to save binary blob from javascript in browser to the file:
#  - https://stackoverflow.com/questions/25547475/save-to-local-file-from-blob
"""

import sys, os, base64, io
import pandas as pd
import numpy as np

# ------------------------------------------
df = pd.DataFrame(data={
    'English' : ['car','boat','book'], 
    'Chinese' : ['','',''], 
    'usd'     : ['10000','15000','19.95'], 
    }, 
    columns=['English','Chinese','usd'])

output = io.BytesIO()
writer = pd.ExcelWriter(output, engine='xlsxwriter')
df.to_excel(writer, sheet_name='Sheet1')
writer.save()
xlsx_data = output.getvalue()
my_b64_string = base64.b64encode(xlsx_data).decode("utf-8")

myhtml = """
<!DOCTYPE html>
<html lang="en"> 
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
<title>Example</title>
<script type="text/javascript">

    function b64toBlob(b64Data, contentType, sliceSize) {
      contentType = contentType || '';
      sliceSize = sliceSize || 512;
    
      var byteCharacters = atob(b64Data);
      var byteArrays = [];
    
      for (var offset = 0; offset < byteCharacters.length; offset += sliceSize) {
        var slice = byteCharacters.slice(offset, offset + sliceSize);
    
        var byteNumbers = new Array(slice.length);
        for (var i = 0; i < slice.length; i++) {
          byteNumbers[i] = slice.charCodeAt(i);
        }
    
        var byteArray = new Uint8Array(byteNumbers);
    
        byteArrays.push(byteArray);
      }
    
      var blob = new Blob(byteArrays, {type: contentType});
      return blob;
    }
</script>
</head>
<body>
<p>Hello</p>

<script type="text/javascript">

var contentType = 'image/png';
var b64Data = '__my_b64_string__';

var blob = b64toBlob(b64Data, contentType);
var blobUrl = URL.createObjectURL(blob);

var link = document.createElement("a"); // Or maybe get it from the current document
link.href = blobUrl;
link.download = "myreport.xlsx";
link.innerHTML = "Click here to download the file";
document.body.appendChild(link); // Or append it where-ever you want
</script>
<p>&nbsp;</p>
__MYDATA__


</body>
</html>
"""

# myhtml = myhtml.replace("__MYDATA__", my_html_data)
myhtml = myhtml.replace("__my_b64_string__", my_b64_string)

fh = open("files/test.html", "w")
fh.write(myhtml)
fh.close()


