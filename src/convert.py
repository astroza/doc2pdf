"""
 Doc, Docx to PDF HTTP Server
 (c) 2011 Felipe Astroza Araya
 under GPLv3
"""

import uno
from com.sun.star.beans import PropertyValue

class OOConverter:
	def __init__(self, host, port):
		localContext = uno.getComponentContext()
		localServiceManager = localContext.ServiceManager
		resolver = localServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext)
		self.context = resolver.resolve("uno:socket,host=" + host + ",port=" + str(port) + ";urp;StarOffice.ComponentContext")
		self.svcManager = self.context.ServiceManager
		self.inputProps = (PropertyValue("Hidden", 0, True, 0),)

	def doc2pdf(self, inputFilename, outputFilename):
		ret = True
		outputProps = (
			PropertyValue("FilterName", 0, "writer_pdf_Export", 0),
			PropertyValue("Overwrite", 0, True, 0),
		)

		desktop = self.svcManager.createInstanceWithContext("com.sun.star.frame.Desktop", self.context)
		try: 
			doc = desktop.loadComponentFromURL("file://" + inputFilename, "_blank", 0, self.inputProps)
		except:
			print "Cannot open " + inputFilename
			ret = False
		if ret:
			try:
				doc.storeToURL("file:///" + outputFilename, outputProps)
			except:
				print "Cannot write " + outputFilename
				ret = False
			if doc:
				doc.dispose()
		return ret

if __name__ == "__main__":
	conv = OOConverter("localhost", 2002)
	conv.doc2pdf("/home/fastroza/doc2pdf_server/demo.doc", "/home/fastroza/doc2pdf_server/demo.pdf")
	conv = OOConverter("localhost", 2002)
	conv.doc2pdf("/home/fastroza/doc2pdf_server/demo.doc", "/home/fastroza/doc2pdf_server/demo2.pdf")
