"""
 Doc, Docx to PDF HTTP Server
 (c) 2011 Felipe Astroza Araya
 under GPLv3
"""
from BaseHTTPServer import BaseHTTPRequestHandler, HTTPServer
from SocketServer import ForkingMixIn
from convert import OOConverter
import shutil
import os
import cgi

listen_host = '0.0.0.0'
listen_port = 8080

class MultiThreadedHTTPServer(ForkingMixIn, HTTPServer):
	pass

class RequestHandler(BaseHTTPRequestHandler):

	def do_GET(self):
		response = "<h2>Doc to PDF Server</h2><p>Only POST requests</p>"
		self.send_response(200)
		self.send_header('Content-Length', len(response))
		self.send_header('Content-Type', 'text/html')
		self.end_headers()
		self.wfile.write(response)

	def do_POST(self):
		form = cgi.FieldStorage(
			fp = self.rfile,
			headers = self.headers,
			environ = {'REQUEST_METHOD': 'POST', 'CONTENT_TYPE':self.headers['Content-Type']}
		)

		filenameSplitted = form['document'].filename.split('.')
		slen = len(filenameSplitted)
		if slen == 1 or (filenameSplitted[slen-1] != 'doc' and filenameSplitted[slen-1] != 'docx'):
			self.send_response(404)
			self.end_headers()
		else:
			pdfFilename = filenameSplitted[slen-2] + '.pdf'
			extension = filenameSplitted[slen-1]

			# Rutas relativas al directorio actual (CWD)
			dir = os.getcwd() + '/' + 'work_' + str(self.wfile.fileno()) + '/'
			docPath = unicode(dir + 'document.' + extension)
			try:
				os.mkdir(dir)
			except:
				pass
			destFile = open(docPath, 'w')
			shutil.copyfileobj(form['document'].file, destFile)
			destFile.close()
			converter = OOConverter("localhost", 2002)
			pdfPath = unicode(dir + 'document.pdf')
			converter.doc2pdf(docPath, pdfPath)
			pdfFile = open(pdfPath, 'r')
			self.send_response(200)
			self.send_header('Content-type', 'application/pdf')
			self.send_header('Content-Disposition', 'attachment; filename=' + pdfFilename)
			self.end_headers()
			shutil.copyfileobj(pdfFile, self.wfile)
			pdfFile.close()
			shutil.rmtree(dir)
			# Los recursos utilizados por UNO son liberados al terminar este proceso forkeado
			# Al usar threading no logre que se desconectara del OpenOffice service, pero con forking
			# Todo los recursos son liberados al terminar de servir la solicitud (y se desconecta de OO)

print "Doc[x]2PDF Converter HTTP Server"
print "(c) 2011 Felipe Astroza Araya - email: felipe@astroza.cl"
httpd = MultiThreadedHTTPServer((listen_host, listen_port), RequestHandler)
if httpd:
	print "HTTP Server running on http://%s:%d" % (listen_host, listen_port)
	httpd.serve_forever()
