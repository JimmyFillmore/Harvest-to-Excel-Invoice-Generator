# Harvest-to-Excel-Invoice-Generator
This takes a Harvest Excel sheet and converts it into an invoice using VBA as a macro. This was one of my first projects and was hard coded for my previous company. At a request I can go back and modify this if someone wants.

Couple of notes:
1) The hourly rate is hardcoded to either 15 or 18. Can be changed
2) This code assumes that you will only have a maximum of 5 different project entries for a day. It will break if you have more.
3) The shortcut key is **ctrl+shift+y** when a cell is selected in the worksheet you wish to change

**Instructions**

1) Create a macro in excel and paste the code into a macro (to add a macro see step 4a)
2) Go to Harvest and see detailed report for either the desired month or semimonth
3) Click export and select Excel
4) On the Harvest Excel doc, click on to any cell and press "ctrl + shift + y" (ONLY PRESS ONCE)
    a) or you can run the macro through the developer tab
        1) Right click on your toolbar above (also called the Ribbon)
        2) Customize Ribbon							
        3) On the right side, check the Developer box							
        4) On the Harvest Excel, click on the Developer Tab							
        5) Click on Macros							
        6) Click on the macro that ends with Harvest_To_Invoice_Converter							
            a) Harvest_To_Invoice_Converter_Beta.xlsm!Harvest_To_Invoice_Converter				
        7) Click Run		
5) Enter your information when prompted									
6) The Harvest Excel file is automatically saved and named "Month, invoice #"						
7) You will be prompted to save the file as a PDF named the same									

