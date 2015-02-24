

1. Create C:\ae\aeloader folder and add all old files
2. Create doc, src, src/xml, zip and pix folders
3. Create blank database from A2013
4. Rename to aeloader.mdb.accdb

Setup for Access Blank Database
> Enable Content (can also set trusted location)
> File => Option => Current Database
> Name AutoCorrect Option
> Uncheck Perform name Autocorrect => OK
> File => Compact and Repair Database
> File => Option => Current Database
> Name AutoCorrect Option
> Uncheck Track name Autocorrect info => OK
> File => Compact and Repair Database

IDE
> Tools => Properties
> Project Name: aeloader
> 

1. Import latest aegit modules
2. Create log module for aeloader
3. Export 
4. Add folder to git revision control
5. Add .gitignore and .gitattributes from aegit
6. Add adaept64.ico
7. Zip old versions and put in zip folder
8. Move old versions to 2do folder
9. Configure and run aeloader_EXPORT
10. Create aeloader project on github
11. Make first commit
12. 