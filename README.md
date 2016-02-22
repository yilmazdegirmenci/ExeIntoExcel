# ExeIntoExcel
Python script to transport an Exe file into Hex and then embed it into a Excel vba macro code.

I tested the code with netcat application (nc.exe):

python ExeIntoExcel.py -f nc.exe

Running this code produces two text files: hex.txt and hexready.txt.
Just copy and paste the content of hexready.txt into an Excel file.

Then you can start your listener:
nc -lvp 443

Now you can run your Excel file (macro enabled) and you are in.
Congratulations, you converted an Excel file into a Trojan Dropper. :)
