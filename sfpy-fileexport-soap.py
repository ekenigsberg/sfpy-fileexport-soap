__author__ = 'ezrakenigsberg.com'
__debug_mode__ = False

ATT = 'Attachment'
ATTS_HEADER = '"Id","Name","ContentType","ParentId","Parent.Name"\r\n'
ATTS_QUERY = "SELECT Id, Name, ContentType, ParentId, Parent.Name FROM Attachment"
DOC = 'Document'
DOCS_HEADER = '"Id","Name","Type","FolderId","Folder.Name"\r\n'
DOCS_QUERY = "SELECT Id, Name, Type, FolderId, Folder.Name FROM Document"

import base64
import beatbox
from time import gmtime, strftime
from string import find
import xmltramp

def main():
    copyright()
    sf = beatbox._tPartnerNS
    menu(sf, login(sf))

def login(sf):
    print 'SET USERNAME AND PASSWORD'
    # Get login credentials
    if __debug_mode__:
        st_username = 'jake@spy.com'
        st_password = 'WaKUdbjJf6Ut'
        st_token = '3jpjJo3quesH7nUOfLYY4DQR'
    else:
        st_username = str(raw_input("-                         Enter username: "))
        st_password = str(raw_input("- Enter password, without security token: "))
        st_token = str(raw_input("-        Enter security token, if needed: "))

    # Log in to Salesforce
    svc = beatbox.Client()
    st_password_and_token = st_password + st_token
    svc.login(st_username, st_password_and_token)
    return svc

def menu(sf, svc):
    print 'MAIN MENU'
    print '- What do you want to do:'
    print '- (a) Extract Attachments'
    print '- (d) Extract Documents'
    print '- (f) Extract Files'
    print '- (q) Quit'
    st_choice = str(raw_input("- Enter your menu option: "))
    st_choice = st_choice.lower()
    if st_choice == 'a':
        get_aords(sf, svc, ATT)
    elif st_choice == 'd':
        get_aords(sf, svc, DOC)
#   Haven't written this yet
#    elif st_choice == 'f':
#        get_files(sf, svc)
    elif st_choice == 'q':
        exit()
    print
    menu(sf, svc)

# Because Attachments and Documents are retrieved in almost the exact same way, I combined their functions.
def get_aords(sf, svc, st_aord):
    # Create the "Attachment Data"/"Document Data" file. The timestamp guarantees a unique filename
    f_data = open(st_aord + ' Data ' + strftime("%Y-%m-%d %H.%M.%S", gmtime()) + '.csv', 'wb')
    if st_aord == ATT:
        # Run the Attachment query and populate the file with the Attachment header
        res = svc.query(ATTS_QUERY)
        f_data.write(ATTS_HEADER)
    elif st_aord == DOC:
        # Run the Document query and populate the file with the Document header
        res = svc.query(DOCS_QUERY)
        f_data.write(DOCS_HEADER)
    else:
        print 'The function "get_aord" was handed a bad value (expected "' + ATT + '" or "' + DOC + '").'
        exit()

    # Iterate through every record retrieved
    for rec in res[sf.records:]:
        # Establish id and name variables
        st_id = str(rec[2])
        st_name = str(rec[3])
        # Establish variables 4 thru 6(they differ slightly for Attachments and Documents)
        st_four = str(rec[4])
        st_five = str(rec[5])
        # Define st_six - requires a little trimming
        if st_aord == ATT:
            st_six = str(rec[6])[4:] #the [4:] chops off the first four chars (the string "Name")
        if st_aord == DOC:
            st_six = str(rec[6])[6:] #the [6:] chops off the first six chars (the string "Folder")
        # Add extension to st_name if the extension's missing
        if st_aord == ATT and find(st_name[-6:], '.') == -1:
            # Use the dictionary at the bottom of this module
            st_name = st_name + DICT_CONTENT_TYPE[st_four]
        elif st_aord == DOC and find(st_name[-6:], '.') == -1 and st_four != '':
            # Use the "Type" field
            st_name = st_name + '.' + st_four

        # Create the file for st_id's record
        print 'Fetching ' + st_aord + ' "' + st_name + '" (' + st_id + ')...'
        get_aord(svc, st_id, st_name, st_aord)
        # Write the extracted record to the "Attachment Data"/"Document Data" file
        f_data.write('"' + st_id + '","' + st_name + '","' + st_four + '","' + st_five + '","' + \
                     st_six + '"\r\n')
    # Close the "Attachment Data"/"Document Data" file
    st_count = str(res[sf.size])
    st_s = ''
    if st_count != '1': st_s = 's'
    print
    print 'Done! Extracted ' + st_count + ' ' + st_aord + st_s + '.'
    f_data.close()

def get_aord(svc, st_id, st_name, st_aord):
    # Get the base64-encrypted Body from Salesforce
    res = svc.query("SELECT Body FROM " + st_aord + " WHERE Id='" + st_id + "'")

    # Decode the body & make the filename
    st_body = str(res.records[2]).decode('base64')
    st_path = st_id + '#' + st_name

    # Write & save the file
    f_aord = open(st_path, 'wb')
    f_aord.write(st_body)
    f_aord.close()

def copyright():
    print """ +---------------- SFPY-FILEEXPORT-SOAP: COPYRIGHT NOTICE ----------------+
 |                   Copyright (C) Ezra Kenigsberg 2015                   |
 | Permission is hereby granted, free of charge, to any person obtaining  |
 | a copy of this software and associated documentation files (the        |
 | "Software"), to deal in the Software without restriction, including    |
 | without limitation the rights to use, copy, modify, merge, publish,    |
 | distribute, sublicense, and/or sell copies of the Software, and to     |
 | permit persons to whom the Software is furnished to do so, subject to  |
 | the following conditions:                                              |
 |                                                                        |
 | The above copyright notice and this permission notice shall be         |
 | included in all copies or substantial portions of the Software.        |
 |                                                                        |
 | THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,        |
 | EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF     |
 | MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. |
 | IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY   |
 | CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,   |
 | TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE      |
 | SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.                 |
 +------------------------------------------------------------------------+"""
DICT_CONTENT_TYPE = {
    '"application/x-zip"':'.zip',
    '"application/zip"':'.zip',
    '.zip':'.zip',
    'application/applefile':'.png',
    'application/atom+xml':'.xml',
    'application/base64':'.jpg',
    'application/binary':'.TXT',
    'application/comma-separated-values':'.csv',
    'application/csv':'.csv',
    'application/docx':'.docx',
    'application/gif':'.gif',
    'application/gzip':'.gzip',
    'application/ics':'.ics',
    'application/ie-ignore-eps':'.eps',
    'application/java-archive':'.jar',
    'application/javascript':'.js',
    'application/jpg':'.jpg',
    'application/json':'.json',
    'application/msexcel':'.xls',
    'application/ms-excel':'.xls',
    'application/msonenote':'.one',
    'application/ms-tnef':'.dat',
    'application/msword':'.doc',
    'application/pdf':'.pdf',
    'application/pgp-signature':'.asc',
    'application/photoshop':'.psd',
    'application/pkcs7-mime':'.p7m',
    'application/pkcs7-signature':'.p7s',
    'application/pkix-cert':'.cer',
    'application/png':'.png',
    'application/postscript':'.EPS',
    'application/rar':'.rar',
    'application/rss+xml':'.xml',
    'application/rtf':'.rtf',
    'application/softgrid-jpg':'.jpg',
    'application/stuffit':'.zip',
    'application/vnd.ms-excel':'.xls',
    'application/vnd.ms-excel.12':'.xlsx',
    'application/vnd.ms-excel.sheet.binary.macroEnabled.12':'.xlsb',
    'application/vnd.ms-outlook':'.msg',
    'application/vnd.ms-powerpoint':'.ppt',
    'application/vnd.ms-powerpoint.presentation.12':'.pptx',
    'application/vnd.ms-powerpoint.presentation.macroenabled.12':'.pptm',
    'application/vnd.ms-publisher':'.pub',
    'application/vnd.ms-visio.drawing':'.vsd',
    'application/vnd.ms-word':'.doc',
    'application/vnd.ms-word.document.12':'.docx',
    'application/vnd.ms-word.document.macroEnabled.12':'.docm',
    'application/vnd.ms-xpsdocument':'.xps',
    'application/vnd.oasis.opendocument.spreadsheet':'.ods',
    'application/vnd.oasis.opendocument.text':'.odt',
    'application/vnd.openxmlformats-officedocument.presentationml.presentation':'.pptx',
    'application/vnd.openxmlformats-officedocument.presentationml.slideshow':'.ppsx',
    'application/vnd.openxmlformats-officedocument.presentationml.template':'.potx',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':'.xlsx',
    'application/vnd.openxmlformats-officedocument.word':'.docx',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document':'.docx',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.template':'.dotx',
    'application/x-7z-compressed':'.7z',
    'application/x-asp':'.asp',
    'application/x-chrome-extension':'.crx',
    'application/x-compressed':'.tgz',
    'application/x-compressed-tar':'.tgz',
    'application/x-crossover-pptx':'.pptx',
    'application/x-extension-eml':'.eml',
    'application/x-gzip':'.gz',
    'application/x-gzip-compressed':'.gz',
    'application/xhtml+xml':'.xhtml',
    'application/x-httpd-php':'.php',
    'application/x-java-archive':'.jar',
    'application/x-javascript':'.js',
    'application/xml':'.xml',
    'application/x-msexcel':'.xls',
    'application/x-mspublisher':'.pub',
    'application/x-ms-wmz':'.wmz',
    'application/x-msword':'.pdf',
    'application/x-msword-doc':'.doc',
    'application/x-pdf':'.pdf',
    'application/x-photoshop':'.psd',
    'application/x-pkcs7-signature':'.p7s',
    'application/x-rar':'.rar',
    'application/x-rar-compressed':'.rar',
    'application/x-secure-download':'.zip',
    'application/x-shockwave-flash':'.swf',
    'application/x-tar':'.tar',
    'application/x-webarchive':'.webarchive',
    'application/x-x509-ca-cert':'.cer',
    'application/x-xml':'.xml',
    'application/x-zip':'.zip',
    'application/x-zip-compressed':'.zip',
    'application/zip':'.zip',
    'audio/mp3':'.mp3',
    'audio/wav':'.wav',
    'audio/x-ms-wmv':'.wmv',
    'audio/x-wav':'.wav',
    'chemical/x-cerius':'.cer',
    'image/bmp':'.bmp',
    'image/gif':'.gif',
    'image/jpeg':'.jpg',
    'image/jpg':'.jpg',
    'image/pdf':'.pdf',
    'image/pjpeg':'.jpg',
    'image/png':'.png',
    'image/svg+xml':'.svg',
    'image/tiff':'.tif',
    'image/vnd.adobe.photoshop':'.psd',
    'image/x-bmp':'.bmp',
    'image/x-citrix-gif':'.gif',
    'image/x-citrix-jpeg':'.jpg',
    'image/x-icon':'.ico',
    'image/x-photoshop':'.psd',
    'image/x-png':'.png',
    'multipart/x-zip':'.zip',
    'pdf':'.pdf',
    'text/calendar':'.ics',
    'text/comma-separated-values':'.csv',
    'text/css':'.css',
    'text/csv':'.csv',
    'text/directory':'.vcf',
    'text/enriched':'.txt',
    'text/html':'.htm',
    'text/javascript':'.js',
    'text/php':'.php',
    'text/plain':'.txt',
    'text/rtf':'.rtf',
    'text/tab-separated-values':'.tsv',
    'text/text':'.txt',
    'text/x-log':'.log',
    'text/xml':'.xml',
    'text/x-php':'.php',
    'text/x-python-script':'.py',
    'text/xsd':'.xsd',
    'text/x-sql':'.sql',
    'text/x-vcard':'.vcf',
    'video/avi':'.avi',
    'video/mp4':'.mp4',
    'video/ogg':'.ogv',
    'video/quicktime':'.mov',
    'video/x-fl':'.png',
    'video/x-ms-wmv':'.wmv',
    'x-unknown/pdf':'.pdf'
}

main()