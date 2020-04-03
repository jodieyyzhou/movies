# movies
Movies from 1900s-2010s

#Purpose
1. Load the IMDB Movie Score Excel file.
2. Join the separate sheets into a single table .
3. Add an Net Earnings (Gross Earnings - Budget) column, if the row does not contain the required data the Net Earnings value should be null.
4. Create a bar plot of the top 10 movies by Net Earnings in descending order
5. Output data table and plot into a new timestamped excel file.
6. Output all rows that do not contain budget or earnings data into a new timestamped excel file (only include the title and year columns).

7. Make sure that the program can handle unexpected data types, for example, what happens if there's a string in "Budget" or "Gross Earnings" column.
8. Instead of presenting the plot to the user running the program include it inside of the outputted Excel file. Imagine the program running unsupervised in a server environment.
9. The output file should be a single sheet that combines all of the sheets from the source file and not split into the separate years.
10. By timestamped file I meant that the file name should include the ISO formatted timestamp when it was created as well as its regular title.
