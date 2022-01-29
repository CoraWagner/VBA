# VBA Macros
## Cora Wagner

[Access VBA Code](https://github.com/CoraWagner/VBA/blob/4c92d580e2708a17d2b36338adc8a5b30a6afc44/MacrosCode)

### How to use the macro
1. Copy the code from the GitHub link above.
2. Open Excel and create a new file, or open an existing file you would like to use the macro on.
3. Go to the *Developer* tab at the top of the Excel Workbook. 
![Developer Tab](Developer.png)
4. Select *Record Macro* under the *Code* subsection. Name your macro and create a shortcut key combination. Select *Okay*. ![Record Macro](RecordMacro.png)
5. Select *Stop Recording* where the *Record Macro* button was previously.
6. Select *Macros* under the *Code* subsecction. A VBA page will pop-up.
7. Paste the GitHub macro code in text box and save the macro. ![VBA Page](VBA.png)
8. You can now run the macro by using the shortcut key you created, or by selecting *Macro* and double clicking the name you created for the macro.

### How the macro works
#### Hide Columns
The first section of my Macro made in VBA is:

`Dim ran As range`

`Dim col As range`

`On Error Resume Next`

`Set ran = Application.Selection`

`Set ran = Application.InputBox("Select a range of Columns you would like to hide.", "Hide Columns", ran.Address, Type:=8)`

`For Each col In ran`

`col.EntireColumn.Hidden = True`

`Next col`

This block of code allows the user to selects the specific range of columns that they would like to hide form view. This is possible by setting the `ran` range to `Application.Selection` and then `Application.InputBox("Select a range of Columns you would like to hide.", "Hide Columns", ran.Address, Type:=8)`. 

The parameters in `Application.InputBox()` are:
1. The main message of the pop-up box
2. The main title of the pop-up box
3. The location of the cells being selected
4. The type that is being selected

In this case, the type being selected is a cell reference. There is then a for loop that iterates through all the selected cells and hides the entire column. 

**NOTE**: You only need to select one cell within each column. Selecting a whole column will cause Excel to crash.

#### Hide Rows
The second section of the macro is:

`Dim rnge As range`

`Dim row As range`

`On Error Resume Next`

`Set rnge = Application.Selection`

`Set rnge = Application.InputBox("Select a range of Rows you would like to hide.", "Hide Rows", rnge.Address, Type:=8)`

`For Each row In rnge`

`row.EntireRow.Hidden = True`

`Next row`

This block of code is similar to *Hide Columns* section as it allows the user to select which rows they would like to remove. The only difference, besides the variable names, is that instead of `EntireColumn.Hidden = True` it is `EntireRow.Hidden = True`.

**NOTE**: You only need to select one cell in the row you want to remove. Selecting a whole row will cause Excel to crash.

#### Hide Cell Content
The third section of the macro is:

`Dim rang As range`

`Dim cel As range`

`On Error Resume Next`

`Set rang = Application.Selection`

`Set rang = Application.InputBox("Select a range of Cells you would like to hide.", "Hide Cells", rang.Address, Type:=8)`

`For Each cel In rang`

`cel.Select`

`ActiveCell.NumberFormat = ";;;"`

`Next cel`

Like the previous sections, the user is able to select the cells that they want the contents to be hidden. The way that the content is hidden is by setting the number format to `";;;"` which is a custom formula that allows the cell to retain the information, but removes the text.

#### Remove Unwanted Hyperlinks
The last section of the macro is:

`Dim rng As range, cell As range`

`Dim xLink As Hyperlink`

`Set rng = Application.Selection`

`Set rng = Application.InputBox("Select a range of Hyperlinks you would like to remove.", "Remove Hyperlinks", rng.Address, Type:=8)`

`For Each cell In rng`

`cell.Select`

`ActiveCell.Font.ColorIndex = 1`

`ActiveCell.Font.Underline = False`

`ActiveCell.Interior.ColorIndex = -4142`

`ActiveCell.ClearHyperlinks`

`Next cell`

As in the preveious sections, the user can select which cells they would like the hyperlinks to be removed from. The text is then changed to black, the undeline is removed, and the background color is set back to default. Then the hyperlink is cleared from the cell.

### Example Of a Table Cleaned-up with the Macro
															
	Day	Topic	Due					In Data Science 2 						Homework	70
					https://classroom.google.com/u/0/c/NDQ0NzcyODkzNjk4										
18-Jan	1	What is Data Science 			https://arielcwebster.github.io/DataScience/			R	Apache Pig					Project	15
20-Jan	2	VBA	HW1 - Excel					Nueral Networks	Hadoop					Readings	10
25-Jan	3	Data Communication						SQL						Participation	5
27-Jan	4	Work Day	HW2 - VBA					D3 - Java Script - probably should have done first semester as part of a unit on HTML							100
1-Feb	5	Why are data visualizations important ?	Reading Due - Florence Nightengale					Tableau							
3-Feb	6	Tableau	COVID Risk Calculator					Julia							
8-Feb	7	How visualizations lie	Reading Due - Differnet Kinds of Data Visualization					Data cleaning part 2 - https://github.com/JohnDickerson/cmsc320-fall2018/tree/master/project1							
10-Feb	8	Work Day	HW 3 - Tableau												
15-Feb	9	Danielle	Reading Due - How Charts Lie					Sentiment Analysis							
17-Feb	10	R Intro						VADER Sentiment Analysis							
22-Feb	11	Doing Better Data Visualization (R and ggplots tutorisl)	Why Data is good for governments to provide					TextBlob Sentiment Analysis (In Book 12.2)							
24-Feb	12	Work Day	HW 4 - ggplots					In DS2 maybe make them do machine learning for sentiment analysis							
1-Mar	13	Sentiment Analysis - History and Types	Data Annonymity	https://www.science.org/doi/10.1126/science.1256297											
3-Mar	14	TextBlob	Reading Due - How to un annonymize data	Why Big Data Helps Science											
8-Mar	15	VADER	De-Annonymizing Data	Or Access and more Data base stuff				Coursera Data Science Ethics							
10-Mar	16	P-Hacking Reflection	HW 5 - Sentiment Analysis	Privacy Concerns with Big Data				Data Privacy			Statistics 				
15-Mar	Spring Break		More P-Hacking								Nueral Networks				
17-Mar								"Analyze data using tools like Spark, MongoDB and Cassandra."			Gradient Descent				
22-Mar	17	Random Forest													
24-Mar	18							Talk about the difference between supervised and unsupervised learning							
29-Mar	Advising Day														
31-Mar	19		HW 6 - Random Forest					Ethics							
5-Apr	20	Clustering - K Nearest Neighbors						"History, Concept of Informed Consent "							
7-Apr	21		Possible Reading - Proxy Discrimination - When AI find predictive proxies for race - because society is segregated in this way. 					Data Ownership 	Data Citation						
12-Apr	22							Privacy	Touched on in Semester one but no lectures						
14-Apr	23		HW 6 - Clustering					Anonymity							
19-Apr	24	Final Project						Data Validity							
21-Apr	25							Algorithmic Fairness 							
26-Apr	26							Societal Consequences 							
28-Apr	27							Code of ethics 	" - Write your own code of ethics for data science. Data science is still a young field and we are still trying to define the basic norms of socially acceptable behavior. Use what you have learned in this course to write your own norms around one of the following subfilds of data Science: Visualization, Data Aquisition, ...."						
									What are counter arguments for each ethical rule you propose						
															
		Additional Readings	Data Sets												
			Maryland Data												
			NYT COVID Data												


### Works Cited
[Hide Columns](https://www.educba.com/vba-hide-columns/)

[InputBox](https://www.wallstreetmojo.com/vba-inputbox/)

[Remove Underline](https://software-solutions-online.com/excel-vba-underline-font-style/)

[Change Font Color](https://www.educba.com/vba-font-color/)

[Clear Hyperlinks](https://www.extendoffice.com/documents/excel/2221-excel-remove-hyperlink-without-removing-formatting.html#:~:text=In%20Excel%2C%20there%20is%20no%20direct%20way%20to,open%20the%20Microsoft%20Visual%20Basic%20for%20Applications%20window.)

[Markdown Table Generator](https://www.tablesgenerator.com/markdown_tables)
