# -*- coding: utf-8 -*-

import os
import sys
import uniqid

from jarray import zeros
from java.io import FileOutputStream, FileInputStream, ByteArrayOutputStream

from org.apache.poi.poifs.filesystem import POIFSFileSystem, DocumentInputStream, DirectoryNode
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
      fs = POIFSFileSystem(fin)
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
      out.close();

   #https://poi.apache.org/hpsf/how-to.html#sec3
   def replaceSummaryInfo(self, ole2filename, blank=False):

      fin = FileInputStream(ole2filename)
      fs = POIFSFileSystem(fin)
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
      else:
         sys.exit("Directory to output OLE2 contents to already exists.")
      return dirname

   def recurse_dir(self, root, outdir):
      #Cache DirectoryNode and directory name
      dircache = {'object': False, 'directory': False}

      for obj in root:
         fname = obj.getShortDescription()

         if type(obj) is DirectoryNode:
            tmpoutdir = outdir + '/' + fname
            os.makedirs(tmpoutdir)
            if dircache['object'] is False:
               dircache['object'] = obj
               dircache['directory'] = tmpoutdir
            else:
               sys.stderr.write("Check container in 7-Zip, likely more dirs at a root DirectoryNode than expected.")
         else:
            #replace strange ole2 characters we can't save in filesystem, todo: check spec
            #this seems to be the convention in 7-Zip, and it seems to work...
            fname = fname.replace(self.replacechar1, '[1]').replace(self.replacechar5, '[5]')

            f = open(outdir + "/" + fname, "wb")
            size = obj.getSize()
            stream = DocumentInputStream(obj);
            bytes = zeros(size, 'b')
            n_read = stream.read(bytes)
            data = bytes.tostring()
            f.write(data)
            f.close()

      #only recurse if we have an object to recurse into after processing DocumentNodes
      if dircache['object'] != False:
         self.recurse_dir(dircache['object'], dircache['directory'])

   def extractContainer(self, ole2filename):

      fin = FileInputStream(ole2filename)
      fs = POIFSFileSystem(fin)
      root = fs.getRoot()
      outdir = self.__makeoutputdir__(ole2filename)
      self.recurse_dir(root, outdir)

   def writeContainer(self, containerfoldername, ext, outputfilename=False):
      written = False
      if outputfilename == False:
         outputfilename = containerfoldername.strip('/') + "-" + uniqid.uniqid() + "." + ext.strip('.')
      containerfoldername = containerfoldername
      #we have folder name, written earlier
      #foldername is filename!!
      if os.path.isdir(containerfoldername):
         fname = outputfilename
         fs = POIFSFileSystem()
         root = fs.getRoot();
         #triplet ([Folder], [sub-dirs], [files])
         for folder, subs, files in os.walk(containerfoldername):
            if subs != []:
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

