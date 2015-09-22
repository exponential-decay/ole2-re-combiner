from JHandleOLE2Containers import ReadWriteOLE2Containers
import argparse
import sys

def arg_check(argcount, max):
   if argcount == max:
      return True
   else:
      sys.exit("Too many arguments for option.")

def main():

   #	Handle command line arguments for the script
   parser = argparse.ArgumentParser(description='Split or re-combine OLE2 files.')
   parser.add_argument('--extract', help='Optional: OLE2 file to open.', default=False)
   parser.add_argument('--combine', help='Optional: Directory to combine into OLE2.', default=False)
   parser.add_argument('--ext', help='Optional: Extension for OLE2.', default='')   
   parser.add_argument('--blanksummary', help='Optional: Replace SummaryInfo with a blank SummaryInfo.', default=False)
   parser.add_argument('--fixsummary', help='Optional: Try to retrieve old SummaryInfo and fix.', default=False)

   if len(sys.argv)==1:
      parser.print_help()
      sys.exit(1)

   ole2class = ReadWriteOLE2Containers()

   #	Parse arguments into namespace object to reference later in the script
   global args
   args = parser.parse_args()

   if args.fixsummary:
      if arg_check(len(sys.argv),3):
         ole2class.replaceSummaryInfo(args.fixsummary)
   elif args.blanksummary:
      if arg_check(len(sys.argv),3):
         ole2class.replaceSummaryInfo(args.blanksummary, True)      
   elif args.extract:
      if arg_check(len(sys.argv),3):
         ole2class.extractContainer(args.extract)
   elif args.combine:
      if not args.ext:
         ole2class.writeContainer(args.combine, '')
      else:
         ole2class.writeContainer(args.combine, args.ext)   

if __name__ == "__main__":      
   main()
