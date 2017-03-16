# Raffle

**Preface**

There are some sneakers in this world that are more than sneakers. They are dreams, promises, art objects, reminders of a different time, and even status markers. Companies who sell these shoes understand their product deeply and purposefully keep supply low. So when a retailer sells these shoes, the anticipation can cause people to camp out on sidewalks for long hours. It can also incite physical violence.

In order to protect customers, stores, and employees, some of these shoe retailers are using a raffle process. Customers go to an online portal where they enter name, phone, email, shoe size, and desired pickup location. Once all the names are collected, the winners have to be selected.

This is where my scripts come in.

**To Run The Code**

You will need to download the .xlsx Macro sheet and the sample breakout. Unfortunately, I cannot provide a sample of the raw list of raffle entries without compromising the safety of our customers, so you will have to construct one. This particular raffle had about 600 entrants, which is considered a low number. Our most popular raffles will have upwards of 10,000 entries, not counting duplicates.

Once everything is downloaded and you have constructed a sample winner's list, It is time to fill in the information in the Macro sheet.

1) All relevant files should be in the same folder, which you will want to list in the Directory cell, C2. Make sure to end with a \ or the script will not be able to find your file. The Macro itself does not need to be in this folder, but my coworkers have found this helps them organize.
2) Put the name of the breakout file, including file extension, into any cell in the range C4 to C9. There are multiple spaces so that multiple size runs can be processed at the same time.
3) Press "Create Template". This will create a new workbook, with one sheet for each store that is receiving the shoes. Each store's sheet will have one line per pair of shoes it receives, with the size listed. Save the template.
4) Place the name of the template, including file extension, in the range G4 to G9. Place the name of your sample entrants into the range E4 to E9. Click the "Fill in Winners" button. This will take names from the entrants list and copy them into the template.
5) Please do not click the "Email Winners" button. It does not work yet.

**A Note on the "Fill in Winners" Algorithm**

I wondered, at first, whether to loop through each blank line on the template or the entrants list. Our biggest raffles so far have had up to 10,000 entrants and up to 4,000 pairs of shoes. If you loop through the shoes, your worst-case scenario is 4,000 x 10,000 = 40,000,000 steps. However, looping through the winner's list means you only have to search the one store that the entrant wanted to pickup shoes from. With roughly 30 stores, this comes to an average of 133 shoes per store. 10,000 x 133 = 1,330,000. This is why I chose to loop through the entrants list and not the stores, even though there are more entrants than shoes.

**Code Sources**

https://msdn.microsoft.com/en-us/library/office/aa221273(v=office.11).aspx
https://msdn.microsoft.com/en-us/library/office/ff839847.aspx
http://stackoverflow.com/questions/6716068/add-new-sheet-to-existing-excel-workbook-with-vb-code
http://stackoverflow.com/questions/6040164/excel-vba-if-worksheetwsname-exists
http://stackoverflow.com/questions/10232150/run-excel-macro-from-outside-excel-using-vbscript-from-command-line
http://stackoverflow.com/questions/31182054/vbscript-to-open-excel-then-add-a-vba-macro
http://stackoverflow.com/questions/7506270/how-to-remove-characters-from-a-string-in-vbscript
http://stackoverflow.com/questions/1271949/exit-a-while-loop-in-vbs-vba
http://www.rondebruin.nl/win/s1/outlook/mail.htm

