from JHandleOLE2Containers import ReadWriteOLE2Containers
import argparse
import sys

def main():

   #	Handle command line arguments for the script
   parser = argparse.ArgumentParser(description='Split or re-combine OLE2 files.')
   parser.add_argument('--ole2', help='Optional: OLE2 file to open.', default=False)
   parser.add_argument('--ext', help='Optional: Extension for OLE2.', default='')
   parser.add_argument('--dir', help='Optional: Directory to combine into OLE2.', default=False)

   #if len(sys.argv)==1:
      #parser.print_help()
      #sys.exit(1)

   test = ReadWriteOLE2Containers()
   #test.writeContainer("tmp", "doc")
   #test.extractContainer("rsdocfile.doc")
   test.replaceSummaryInfo("rsdocfile.doc")

if __name__ == "__main__":      
   main()
