# rtftopdf-service
Python microservice to convert RTF (or any Word) documents to PDF using MS Word and OLE

Requirements:
MS Office 2010 or more recent.

Python libraries:
- flask
- optparse-pretty
- pywin32

Usage:
Start the service with:
python rtftopdf-service.py [-H hostname=127.0.0.1] [-P port=5000]

To convert a document, call the url:
http://hostname:port/rtftopdf?input=<absolute_path_of_the_document>&out=<absolute_path_of_the_output_file>

Note:
- The document must be accessible from the server where this script runs. No feature to upload a file. The ouput file is also stored on the server.
- Must run on Windows because of the OLE automation. The absolute_path might need to be escaped (C:/my/path/input_file.rtf).
