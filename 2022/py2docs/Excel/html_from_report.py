"""
# test_df_to_html1.py
#  - read binary excel file, base64-encodes it
#  - read text html of a pandas DataFrame
#  - combines everything into one HTML page
#    which shows the table with data
#    and provides a link to download Excel binary.
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

import sys, os, base64

# ------------------------------------------
DATA_DIR=os.environ['DIR_JNDATA']
fname_in_xlsx = DATA_DIR+'/rep_partner.xlsx'
print("reading ", fname_in_xlsx)
fh = open(fname_in_xlsx, "rb")
my_b64_string = base64.b64encode(fh.read()).decode("utf-8")
fh.close()

fname_in_html = DATA_DIR+'/rep_partner.html'
print("reading ", fname_in_html)
fh = open(fname_in_html, "r")
my_html_data = fh.read()
fh.close()

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
<link href='style0.css' rel='stylesheet' type='text/css'>
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

myhtml = myhtml.replace("__MYDATA__", my_html_data)
myhtml = myhtml.replace("__my_b64_string__", my_b64_string)

fh = open("files/test.html", "w")
fh.write(myhtml)
fh.close()


