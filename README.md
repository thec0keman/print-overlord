Print Overlord
========

Monitor's a folder for new files and calls the system print command on them
This is very useful when you have to print a lot of email attachments or documents: just click download and this script takes care of the rest
Currently only is setup to work on Windows

File types it monitors for:
- pdf
- docx
- doc
- html
- htm
- xlsx
- xls

##Command lines

- `-f` folder ( default is %user%/downloads )
- `-i` include additional file types (comma separated, no spaces)
- `-e` exclude default file types (comma separated, no spaces)
- `-o` override file types with a new list (comma separated, no spaces)
- `-r` enable folder recursion