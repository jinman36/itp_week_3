# Project 1: Reading Data from a Spreadsheet
# Say you have a spreadsheet of data from the 2010 US Census and you have the boring task of going through its thousands of rows to count both the total population and the number of census tracts for each county. (A census tract is simply a geographic area defined for the purposes of the census.) Each row represents a single census tract. We’ll name the spreadsheet file censuspopdata.xlsx, and you can download it from https://nostarch.com/automatestuff2/. Its contents look like Figure 13-2.

# image
# Figure 13-2: The censuspopdata.xlsx spreadsheet

# Even though Excel can calculate the sum of multiple selected cells, you’d still have to select the cells for each of the 3,000-plus counties. Even if it takes just a few seconds to calculate a county’s population by hand, this would take hours to do for the whole spreadsheet.

# In this project, you’ll write a script that can read from the census spreadsheet file and calculate statistics for each county in a matter of seconds.

# This is what your program does:

# Reads the data from the Excel spreadsheet
# Counts the number of census tracts in each county
# Counts the total population of each county
# Prints the results
# This means your code will need to do the following:

# Open and read the cells of an Excel document with the openpyxl module.
# Calculate all the tract and population data and store it in a data structure.
# Write the data structure to a text file with the .py extension using the pprint module.






# Project2: Updating a Spreadsheet
# In this project, you’ll write a program to update cells in a spreadsheet of produce sales. Your program will look through the spreadsheet, find specific kinds of produce, and update their prices. Download this spreadsheet from https://nostarch.com/automatestuff2/. Figure 13-3 shows what the spreadsheet looks like.

# image
# Figure 13-3: A spreadsheet of produce sales

# Each row represents an individual sale. The columns are the type of produce sold (A), the cost per pound of that produce (B), the number of pounds sold (C), and the total revenue from the sale (D). The TOTAL column is set to the Excel formula =ROUND(B3*C3, 2), which multiplies the cost per pound by the number of pounds sold and rounds the result to the nearest cent. With this formula, the cells in the TOTAL column will automatically update themselves if there is a change in column B or C.

# Now imagine that the prices of garlic, celery, and lemons were entered incorrectly, leaving you with the boring task of going through thousands of rows in this spreadsheet to update the cost per pound for any garlic, celery, and lemon rows. You can’t do a simple find-and-replace for the price, because there might be other items with the same price that you don’t want to mistakenly “correct.” For thousands of rows, this would take hours to do by hand. But you can write a program that can accomplish this in seconds.

# Your program does the following:

# Loops over all the rows
# If the row is for garlic, celery, or lemons, changes the price
# This means your code will need to do the following:

# Open the spreadsheet file.
# For each row, check whether the value in column A is Celery, Garlic, or Lemon.
# If it is, update the price in column B.
# Save the spreadsheet to a new file (so that you don’t lose the old spreadsheet, just in case).