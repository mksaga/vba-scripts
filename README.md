# VBA Scripts

This repository contains a few VBA scripts I wrote to automate some pesky tasks in Excel.

### CopyAllCellsToNewSheet

This one was pretty straightforward. Excel has a feature that allows duplicating a sheet, but using it also makes a mess out of any named ranges/variables (Excel creates new variables with the *exact* same name as those on the duplicated sheet, but with worksheet scope—which cannot be updated from VBA).

So, this routine copies just the cells along with their formatting without creating any new named ranges.  

### UpdateDuplicatedNames

I was quite surprised to not find this online. Basically, I had a rather large Excel model with a sheet I wanted to duplicate (around 200 named ranges, 800+ rows). I wanted to create a duplicate of this sheet and tweak a few parameters to add a feature to the model, but duplicating the sheet directly created a mess, as described above.

This method will look through all variables on the old sheet, create a new variable for the new sheet that contains an updated substring (in my case, I needed to replace all instances of "gas" or "gaseous" in the variable name to "MH"), set the new variable name to point to the proper location, and update all the references contained in the formulas. Hope this will save some headache!