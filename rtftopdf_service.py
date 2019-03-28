import sys
import os,os.path
import optparse
from win32com.client.dynamic import Dispatch
from win32com.client import gencache
import pythoncom
from flask import Flask, request

gencache.EnsureDispatch('Word.Application')

#gencache.EnsureModule('{00020905-0000-0000-C000-000000000046}',0,8,1)
wdApplicationClass = gencache.GetClassForProgID('Word.Application')

app=Flask(__name__)

@app.route("/rtftopdf")
def convert_service():

	input=request.args.get('input')
	output=request.args.get('output')

	if input==None:
		return "missing Input filepath"

	print(input)
		
	return convert(input,output)


def convert(input,output):
		
	wdFormatPDF = 17

	current_dir = os.getcwd()
	
	in_file=input
	if output==None:
		out_file=in_file+".pdf"
	else:
		out_file = output
		
	if not os.path.exists(in_file):
		return("File "+in_file+" does not exist!")
		
	
	print("Converting file from "+in_file+" to "+out_file)
	
	ok=False
	
	pythoncom.CoInitializeEx(pythoncom.COINIT_MULTITHREADED)
	
	word = wdApplicationClass()
  #word = Dispatch('Word.Application')
	try:
		doc = word.Documents.Open(in_file, Visible=False, NoEncodingDialog=True)
		try:
			doc.SaveAs(out_file, FileFormat=wdFormatPDF)
			ok=True
		finally:
			doc.Close()
	finally:
		word.Quit()
		pythoncom.CoUninitialize()
	
	if ok:
		return "done"
	else:
		return "error"

def flaskrun(app, default_host="127.0.0.1", 
                  default_port="5000"):
    """
    Takes a flask.Flask instance and runs it. Parses 
    command-line flags to configure the app.
    """

    pythoncom.CoUninitialize()
    
    # Set up the command-line options
    parser = optparse.OptionParser()
    parser.add_option("-H", "--host",
                      help="Hostname of the Flask app " + \
                           "[default %s]" % default_host,
                      default=default_host)
    parser.add_option("-P", "--port",
                      help="Port for the Flask app " + \
                           "[default %s]" % default_port,
                      default=default_port)

    # Two options useful for debugging purposes, but 
    # a bit dangerous so not exposed in the help message.
    parser.add_option("-d", "--debug",
                      action="store_true", dest="debug",
                      help=optparse.SUPPRESS_HELP)
    parser.add_option("-p", "--profile",
                      action="store_true", dest="profile",
                      help=optparse.SUPPRESS_HELP)

    options, _ = parser.parse_args()

    # If the user selects the profiling option, then we need
    # to do a little extra setup
    if options.profile:
        from werkzeug.contrib.profiler import ProfilerMiddleware

        app.config['PROFILE'] = True
        app.wsgi_app = ProfilerMiddleware(app.wsgi_app,
                       restrictions=[30])
        options.debug = True

    app.run(
        debug=options.debug,
        host=options.host,
        port=int(options.port)
    )		
		
if __name__ =="__main__":
	flaskrun(app)
