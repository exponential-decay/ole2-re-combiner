# -*- coding: utf-8 -*-

import os
import uniqid

from jarray import zeros
from java.io import FileOutputStream, FileInputStream, ByteArrayOutputStream

from org.apache.poi.poifs.filesystem import NPOIFSFileSystem, DocumentInputStream
from org.apache.poi.hpsf import SummaryInformation, PropertySetFactory

class ReadWriteOLE2Containers:

   def __debugfos__(self, fos, bufsize):
      buf = zeros(bufsize, 'b')
      fin.read(buf)
      print buf

   #https://poi.apache.org/hpsf/how-to.html#sec3
   def replaceSummaryInfo(self, ole2filename):

      fin = FileInputStream(ole2filename)
      fs = NPOIFSFileSystem(fin)
      root = fs.getRoot()
      for obj in root:
         x = obj.getShortDescription()
         if x == (u"\u0005" + "SummaryInformation"):
            test = root.getEntry((u"\u0005" + "SummaryInformation"))         
      ps = PropertySetFactory.newSummaryInformation()      
      ps.write(root, (u"\u0005" + "SummaryInformation"));
      
      out = FileOutputStream(ole2filename);
      fs.writeFilesystem(out);
      out.close();

   def extractContainer(self, ole2filename):
   
      fin = FileInputStream(ole2filename)
      fs = NPOIFSFileSystem(fin)
      root = fs.getRoot()

      for obj in root:
         fname = obj.getShortDescription()
         f = open(u"tmp/" + fname, "wb")
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

      containerfoldername = containerfoldername + "/"

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
                  fin = FileInputStream(folder + f)
                  if fin.getChannel().size() == 0:
                     fin.close()
                     written = False
                     break
                  else:
                     root.createDocument(f, fin)
                     fin.close()
                     written = True

         if written == True:
            fos = FileOutputStream(fname)
            fs.writeFilesystem(fos);
            fs.close()

      return written

