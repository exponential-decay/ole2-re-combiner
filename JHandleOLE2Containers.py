# -*- coding: utf-8 -*-

import os
import sys
import uniqid

from jarray import zeros
from java.io import FileOutputStream, FileInputStream, ByteArrayOutputStream

from org.apache.poi.poifs.filesystem import NPOIFSFileSystem, DocumentInputStream
from org.apache.poi.hpsf import SummaryInformation, DocumentSummaryInformation, PropertySetFactory, PropertySet, UnexpectedPropertySetTypeException

class ReadWriteOLE2Containers:

   replacechar1 = '\x01'
   replacechar5 = '\x05'

   def __debugfos__(self, fos, bufsize):
      buf = zeros(bufsize, 'b')
      fin.read(buf)
      print buf

   def replaceDocumentSummary(self, ole2filename, blank=False):
      fin = FileInputStream(ole2filename)
      fs = NPOIFSFileSystem(fin)
      root = fs.getRoot()
      si = False
      siFound = False
      for obj in root:
         x = obj.getShortDescription()
         if x == (u"\u0005" + "DocumentSummaryInformation"):   
            siFound=True
            if blank == False:
               test = root.getEntry((u"\u0005" + "DocumentSummaryInformation")) 
               dis = DocumentInputStream(test);
               ps = PropertySet(dis);
               try:
                  si = DocumentSummaryInformation(ps)
               except UnexpectedPropertySetTypeException as e:
                  sys.stderr.write("Error writing old DocumentSymmaryInformation:" + str(e).replace('org.apache.poi.hpsf.UnexpectedPropertySetTypeException:',''))
                  sys.exit(1)
                  
      if blank == False and siFound == True:
         si.write(root, (u"\u0005" + "DocumentSummaryInformation"))
      else:
         ps = PropertySetFactory.newDocumentSummaryInformation()      
         ps.write(root, (u"\u0005" + "DocumentSummaryInformation"));
      
      out = FileOutputStream(ole2filename);
      fs.writeFilesystem(out);
      out.close();'''

   #https://poi.apache.org/hpsf/how-to.html#sec3
   def replaceSummaryInfo(self, ole2filename, blank=False):

      fin = FileInputStream(ole2filename)
      fs = NPOIFSFileSystem(fin)
      root = fs.getRoot()
      si = False
      siFound = False
      for obj in root:
         x = obj.getShortDescription()
         if x == (u"\u0005" + "SummaryInformation"):
            siFound = True
            if blank == False:
               test = root.getEntry((u"\u0005" + "SummaryInformation")) 
               dis = DocumentInputStream(test);
               ps = PropertySet(dis);
               #https://poi.apache.org/apidocs/org/apache/poi/hpsf/SummaryInformation.html
               si = SummaryInformation(ps);

      if blank == False and siFound == True:
         si.write(root, (u"\u0005" + "SummaryInformation"))
      else:
         ps = PropertySetFactory.newSummaryInformation()      
         ps.write(root, (u"\u0005" + "SummaryInformation"));
      
      out = FileOutputStream(ole2filename);
      fs.writeFilesystem(out);
      out.close();

   def __makeoutputdir__(self, ole2filename):
      dirname = ole2filename.split('.')[0]
      if not os.path.exists(dirname):
         os.makedirs(dirname)
      return dirname

   def extractContainer(self, ole2filename):
   
      fin = FileInputStream(ole2filename)
      fs = NPOIFSFileSystem(fin)
      root = fs.getRoot()

      outdir = self.__makeoutputdir__(ole2filename)

      for obj in root:         
         fname = obj.getShortDescription()
         
         #replace strange ole2 characters we can't save in filesystem, todo: check spec
         fname = fname.replace(self.replacechar1, '[1]').replace(self.replacechar5, '[5]')
         
         f = open(outdir + "/" + fname, "wb")
         size = obj.getSize()
         stream = DocumentInputStream(obj); 
         bytes = zeros(size, 'b')
         n_read = stream.read(bytes)
         data = bytes.tostring()         
         f.write(data)
         f.close()

   def writeContainer(self, containerfoldername, ext, outputfilename=False):
      written = False
      if outputfilename == False:
         outputfilename = containerfoldername.strip('/') + "-" + uniqid.uniqid() + "." + ext.strip('.')
      containerfoldername = containerfoldername
      #we have folder name, written earlier
      #foldername is filename!!   
      if os.path.isdir(containerfoldername):
         fname = outputfilename
         fs = NPOIFSFileSystem()
         root = fs.getRoot();
         #triplet ([Folder], [sub-dirs], [files])
         for folder, subs, files in os.walk(containerfoldername):
            if subs != []:
               #TODO: cant't yet write directories      
               break
            else:
               for f in files:
                  fin = FileInputStream(folder + '/' + f)
                  if fin.getChannel().size() == 0:
                     fin.close()
                     written = False
                     break
                  else:
                     root.createDocument(f, fin)
                     fin.close()
                     written = True
      else:
         sys.exit("Not a valid folder: " + containerfoldername)
            
      if written == True:
         fos = FileOutputStream(fname)
         fs.writeFilesystem(fos);
         fs.close()

      return written

