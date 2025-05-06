# Introduction

## Purpose of this Course

Excel is a Swiss Army Knife, it can do almost everything but easily gets overpowered when working with large or complicated amounts of data. Its for this versatility that Excel is a key skill to always have handy as a Data Analyst. This course is designed to help prospective and current Data Analysts learn the level of Excel that they’ll need in order to be successful in their jobs. It would be impossible to design a short course that includes everything that every job in Analytics needs to know, the content here is based on my experience and that of some analysts I know. 

If you work in Finance then the level of Excel skills necessary will be beyond the scope of this course.

## What do I use Excel for

I generally use Excel for data cleaning and ad-hoc reporting. It’s also an easy tool to share analyses with stakeholders of all levels of technical competency. 

I have also used it as an initial project planning tool. 

Personal financial planning tool

## Excel for Mac vs PC

Excel is definitively a better product on PCs than on Macs. For the tasks that most data analysts will need to undergo, Excel for Mac will work just fine. This course can be completed in either system. The below article outlines some of the key differences between Excel for Mac and Excel for PC.

[Excel for Mac vs Excel for Windows - Pros & Cons](https://spreadsheeto.com/mac-vs-windows/)

**To summarize:**

Pivot Tables exist on the Mac version of Excel but not Pivot Charts.

PowerPivot doesn’t exist in Excel for Mac (I usually would just default to Python for most of PowerPivot’s functionality anyways)

VBA editing doesn’t work as well on Excel for Mac

## My Opinion on VBA

VBA is a great tool to automate various operations in Excel and help speed up analyses. If you’re in a job that explicitly requires it, then learning it might be useful but otherwise I argue that most data analysts won’t see a major benefit to their careers from learning VBA. You’re better off learning something like Python which will help you take on more advanced projects in your career. Check out my free Python course here: 

[Python for Data Analysts and Data Scientists](https://www.youtube.com/watch?v=sZDgJKI8DAM)

## What about Google Sheets?

Google Sheets offer a lot of the functionality of Microsoft Excel and most of what we go over in this course will apply to Google Sheets as well. Being an internet first tool, Google Sheets have a lot of functionality that allows it to interact with the internet natively and interact with Google Cloud services.

## Wait but what is Excel really?

Basically the only reason to buy Office 365 these days.

# Basics

## Hyperbasics

### Layout

If you’re using the standard Excel format then the file you open will be called a “Workbook” with individual “Worksheets” being what you typically interact with. In the below screenshot: 

- **Red -** The Ribbon, this is where all of your core Excel functionality lives
- **Brown -** This is an individual cell, it’s where your individual records of data exist
- **Orange -** The Excel formula bar, this where you input data into individual cells and edit any formulas that you may want to write
- **Blue -** The sheet, this is where you input and edit your data
- **Pink -** This is a tab that denotes which sheet you’re currently on



### File Formats

Excel has a number of file formats, the main ones are listed here. 

**.xls -** This is the old file format that Excel uses and a format you might encounter if you’re pulling data from old systems. Always convert these into a newer format when you can as their Legacy code has many security flaws that Microsoft has no obligation to fix.

**.xlsx -** This is the current version of the Excel file format and supports (almost) all of Excel’s current functions

**.xlsm -** This is the same as an .xlsx except it supports Macros. Macros are a way you can automate certain functions within Excel. 

**.csv -** This stands for “Comma Separated Values” and is a standard method for storing data. Excel does some funny things when opening .csv’s which we’ll go over in a bit so be careful when using Excel for this function. 

## Excel Data Types

### General

This tells Excel to take its best guess at what data type you’re trying to use. In many cases this can work but in many cases it will keep attempting to force a bad guess. To illustrate this, try to store a number with a leading 0 as text, such as the postal code for Bangor, Maine (04401). Excel will keep trying to convert this into a number and deleting the leading 0. You can force a text type by switching the type to Text or by adding an apostrophe prior to your input like so: `'04401`.

You can see that the General data type and Excel’s misinterpretation of certain inputs can lead to some very real consequences such as this case where Scientists had to rename genes to stop Excel from misreading them as dates.  

[Scientists rename human genes to stop Microsoft Excel from misreading them as dates](https://www.theverge.com/2020/8/6/21355674/human-genes-rename-microsoft-excel-misreading-dates)

### Number

This converts any data you have into a numerical representation. It will default to showing two decimal places although you can increase and decrease the number of decimal places as you please later using: 



There are other specified versions of the numerical data type including: 

- Percentage
- Fraction
- Scientific

### Currency

This will convert any number you might have into a representation with the specified currency. 

### Accounting

This is similar to the “Currency” format in Excel except that negative amounts are represented by parentheses instead of hyphens. It’s a formatting change that can be important depending on the stakeholder you’re working with. 

### Date

As in other systems, dates can be very difficult to worth with in Excel. I recommend storing your dates in a YYYY-MM-DD format so that they can easily be sorted even if they get converted to text. This is also the standard format that most databases will store dates in. 

One interesting property of Excel dates is that they can be converted to a 5-digit integer value which correlates to the number of days after ‘1899-12-31’. This is different from a lot of other applications which use 1970-01-01 as the day to start their date calculations from. 

### Time

This is similar to the Date datatype except for times. 

### Text

This is like a “string” data type in many 

### Special

These include data types that are usually special for a given geography. In the US we have types that allow us to format our postal codes and social security numbers correctly (although I can’t imagine storing social security numbers in an Excel workbook is at all secure. 



### Custom

This is where you can specify your own data types. 

## Functions

### Aggregation Functions

Aggregation functions are applied across a range of cells and summarize their values into a single cell. There are many different aggregation functions some of the most common ones being listed below.

**SUM**

An aggregation function that adds up all of the values of its inputs. It only really works with numbers and formats that can be readily converted into numbers (like dates).



**AVERAGE**

An aggregation function that outputs the average of all of its inputs. Like the SUM function, works well with numbers and formats that can be converted into numbers like dates.



**COUNT**

An aggregation function that counts the number of non-blank cells within its inputs. Works with any data format as it just counts for content in cells.



**MEDIAN**

An aggregation function that outputs the median of all of its inputs. Like the SUM function, works well with numbers and formats that can be converted into numbers like dates.



**MODE**

An aggregation function that outputs the most common value of its inputs.  Like the SUM function, works well with numbers and formats that can be converted into numbers like dates.



**MODE.SNGL**

This is the modern form of the MODE function and returns the only a single value like the MODE function. 



**MODE.MULT**

This outputs an array of values for every value that could be considered the mode of the dataset. 



**MAX/MIN**

MAX and MIN are aggregation functions that determine the largest or smallest values of their inputs. Like the SUM function, works well with numbers and formats that can be converted into numbers like dates.



### Lookups

Lookups allow you to match values from one range to references in another. 

**VLOOKUP**

VLOOKUPs are the most common type of lookup and can be used to being data from one range to match another. The parameters of a VLOOKUP are as follows. 



- lookup_value
    - The value that you want to look up
- table_array
    - The range that we want to pull values in from. Make sure that your lookup value is in the left-most column of the range
- col_index_num
    - The number of the column for the value you want to bring in. 1 refers to the left-most column in your range
- range_lookup
    - Should the VLOOKUP find an approximate match or an exact match. It defaults to TRUE but you'll usually want an exact match and therefore need to set it to FALSE

**HLOOKUP**

HLOOKUPs are basically the same thing except they work horizontally. 

![44403441-6A9E-44CD-90C0-76DCAF6F0A3F.jpeg](attachment:2c9ce3c4-bc5b-4eb1-a957-03617b7e40cc:44403441-6A9E-44CD-90C0-76DCAF6F0A3F.jpeg)

**XLOOKUP**

This is a newer version of the VLOOKUP/HLOOKUP formula. Depending on your version of Excel you might not have access to this function. 

XLOOKUP is a much more flexible version of VLOOKUP and HLOOKUP because it doesn’t require your lookup array (the list of values you’re matching) to be on the left-hand side of your return array (the list of values you’re inputting).



- `lookup_value`
    - The value you want to lookup
- `lookup_array`
    - The list of values that might contain the value you’re trying to match to
- `return_array`
    - The list of values (of the same length as the `lookup_array`) that you’re trying to return values from.
- `[if_not_found]`
    - A default value to return if you can’t find a match in between your `lookup_value` and your `lookup_array`.
- `[match_mode]`
    - Should Excel try and find an exact match or settle for an approximate one. Unlike the VLOOKUP and HLOOKUP, this defaults to an exact match which is what we usually want anyways.
    - There is also an option for a wildcard match which can be useful if you don’t have the full information in your lookup_value cell (like only having a last name instead of a full name)
- `[search_mode]`
    - What order should Excel search the `lookup_array` in

### INDEX MATCH... MATCH

The INDEX and MATCH function, when combined are very powerful and are used in many advanced Excel functions. It is considered by advanced Excel users to be better practice to use instead of VLOOKUPs if you can. Let’s first see what the individual functions do. 

INDEX will return any value in a 2D array based on a specified row_num and column_num. Remember in Excel that Row Numbers and Column Numbers start with 1 not 0. 

![9E63D5E8-649D-4BC0-AF11-D0A4FAFCDD06.jpeg](attachment:342caefa-1298-42bf-b8f3-0c19aa29f394:9E63D5E8-649D-4BC0-AF11-D0A4FAFCDD06.jpeg)

MATCH will return the index number of a specified value within an array. It can only search 1D arrays so you’ll need to use two MATCH functions in order to find an item in a 2D array. 

![8547A459-F8B2-4364-8D67-2DFA844E9F9D.jpeg](attachment:7fdd98d4-76bc-4d53-8a15-03401a4596a1:8547A459-F8B2-4364-8D67-2DFA844E9F9D.jpeg)

You can group together the two MATCH statements in the `row_num` and `column_num` parameters of the INDEX function like so:

- `=INDEX(B4:F7,MATCH(I7,A4:A7,0), MATCH(H7,B3:F3,0))`

### IF

One of the most powerful and used functions in Excel. IF allows you to set a condition and then output one value if that condition is met, and another if it isn’t.

There’s a common joke amongst developers that coding is basically a bunch of for loops and if statements. With this function you can replicate some of the functionality of a full-fledged programming language.

- `logical_test`
    - This is a conditional that you can assign. Usually you’ll be looking for a value to be equal to, greater than, or less than another value.
- `[value_if_true]`
    - This is the value that Excel should output if the condition you specified is met
- `[value_if_false]`
    - This is the value that Excel should output if the condition you specified is not met

Let’s say that you want to have multiple conditions, this is where IF statements in Excel can get a bit clunky. What you’ll want to do is use the AND operator inside the IF function:

- =IF(AND(A2>=3, A2<=5), ">=3 AND <= 5", "None")



**SUMIF(S)**





**COUNTIF**

![FA7DD0D7-0868-467F-B147-8BFC7127CCA3.jpeg](attachment:13529659-af22-4ae1-a0ad-69d60efe8bae:FA7DD0D7-0868-467F-B147-8BFC7127CCA3.jpeg)

![7653F832-485D-494A-8323-EA335F1B5028.jpeg](attachment:0a9860a5-b829-4bc5-9397-72a5de125f55:7653F832-485D-494A-8323-EA335F1B5028.jpeg)

### String Functions

**CONCAT**

CONCAT will take multiple TEXT formatted cells and combine them together. You can also input your own text strings inside the CONCAT function to dynamically create labels.

![1F3C1217-12DC-49E4-8579-B8080FC62B95.jpeg](attachment:25a401b4-5fb2-455d-be28-dba091c1ec8b:1F3C1217-12DC-49E4-8579-B8080FC62B95.jpeg)

**LEN**

Output’s the length of characters of its inputs.

![E0564C78-8622-41EE-8123-B5D41D5ED123.jpeg](attachment:7e94319b-4774-4b53-90f3-a3869dc98389:E0564C78-8622-41EE-8123-B5D41D5ED123.jpeg)

**TRIM**

Trims off any blanks on the ends of any strings you have. This can be particularly useful when you’re importing CSV’s into your workbook as some CSV’s are poorly formatted and have trailing or leading spaces between their headers.

A lot of CSV’s import with the extra spaces in the top headers. 



[Advanced Excel Formulas Must Know](https://corporatefinanceinstitute.com/resources/excel/study/advanced-excel-formulas-must-know/)

# Formatting

One of the biggest use cases I have for Excel is to present data from ad hoc analyses where building a dashboard might be too much of a lift. Usually, I’ll pull data using SQL queries and then copy it into an Excel to share. In these cases I’ve found that properly formatting your outputs can yield outsized results and make your stakeholders much happier with your work. 

Merge Cells

![Untitled](attachment:9f839016-9b7d-416e-9d44-872a55ba37c8:Untitled.png)

Center Across Selection

![Untitled](attachment:c606dd82-0410-4630-9b3c-4ae43f608323:Untitled.png)

Remove gridlines

![Untitled](attachment:e05f14a5-732d-4210-b8c4-02d920e3d853:Untitled.png)

Grouping / Ungrouping Cells

![Untitled](attachment:5cbb51a4-542c-4e05-b77f-fd0c92ea8093:Untitled.png)

Adding and Removing Decimals

![Untitled](attachment:408f9bce-5646-4aef-8e32-c18b220fd4fa:Untitled.png)

Data Validation

![Untitled](attachment:0c4b0250-7fe7-4223-b8b6-e961e7b31af6:Untitled.png)

### Conditional Formatting

![Untitled](attachment:56a02aa9-4ba5-4d28-91b5-fcaa9c83ece5:Untitled.png)

# Data Manipulation

Properly Importing CSVs

![Untitled](attachment:5cc2686b-4541-4d28-a1b9-bcc810b592f2:Untitled.png)

Remote Duplicates

![Untitled](attachment:f951f1fe-922a-42c8-8975-2435532bca8d:Untitled.png)

Text to Columns

![Untitled](attachment:b79d5160-5cf8-4aee-9aab-b5a59e20ef03:Untitled.png)

Filters

![Untitled](attachment:f0ef5022-d3ad-4105-80b0-9facecd311f2:Untitled.png)

# Tables

# Pivoting

Pivot Tables are one of Excel’s most useful features and allow you to pivot and transform data in different ways so that you can easily and quickly analyze different cuts of the data. 

![Untitled](attachment:fbc0f4c9-344f-4c79-aea1-d2781932c55f:Untitled.png)

## Excel Goal Seek

Excel Goal seek is a very cool feature of Excel that allows you to solve an equation by manipulating the value in a single cell. It’s great for when you’ve built a model in Excel and need to check what the solution to said model needs to be. I have an example on my Instagram of me using Excel Goal Seek:

[](https://www.instagram.com/reel/CaGbRxAAJlH/?utm_medium=copy_link)

I based my example off of the great work from Kevin Stratvert’s channel: 

[Excel Solver & Goal Seek Tutorial](https://www.youtube.com/watch?v=UD9e-gQCQsE)

![Untitled](attachment:892aef35-c888-400f-a78a-2c6ee1dd7b8f:Untitled.png)
## Excel Solver

The Excel solver is a very powerful tool but not one that I’ve used consistently. It is meant to solve complex models and equations that you’ve built out in Excel (oftentimes Financial or Supply Chain models) and give you a targeted or optimal solution. I’ve tagged a video here by Leila Gharani (or you can watch the video above by Kevin Stratvert) which cover usecases of the Excel solver. 

[Excel Solver - Example and Step-By-Step Explanation](https://www.youtube.com/watch?v=dRm5MEoA3OI)

## Macros

Macros allow you to record repetitive tasks that you might do yourself repeatedly and have Excel handle them for you. Kevin Stratvert has a great video explaining how Macros work, generally I prefer to offload programmatic tasks to Python instead of creating Macros, but depending on what your use case it, Macros might make more sense.

[How to Create Macros in Excel Tutorial](https://youtu.be/uyj_OljPlcU)

# Why do I believe Excel should only be learned/used in limited quantities

No history

Doesn’t work well with large data

A lot of data is in databases anyways

Handles data types VERY poorly

Collaboration tools are not as strong as those available for coding platforms

If you want to become a Data Analyst, the extra income you can earn by being very good at Excel is not as much as that which can be earned by just learning some coding

- If you disagree then please respectfully leave a comment in the comment section along with the job title that benefits monetarily by knowing Excel better
- [Data Analyst Roadmap](https://youtu.be/-AbAm4t4FYQ)

[How to Become a Data Analyst (Updated for 2022)](https://youtu.be/-AbAm4t4FYQ)
