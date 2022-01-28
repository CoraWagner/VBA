# VBA Macros
## Cora Wagner

[Access VBA Code](https://github.com/CoraWagner/VBA/blob/4c92d580e2708a17d2b36338adc8a5b30a6afc44/MacrosCode)

### Hide Columns
The first section of my Macro made in VBA is:

`Dim ran As range
Dim col As range
On Error Resume Next
Set ran = Application.Selection
Set ran = Application.InputBox("Select a range of Columns you would like to hide.", "Hide Columns", ran.Address, Type:=8)
For Each col In ran
col.EntireColumn.Hidden = True
Next col`

This block of code allows the user to selects the specific range of columns that they would like to hide form view. This is possible by setting the `ran` range to `Application.Selection` and then `Application.InputBox("Select a range of Columns you would like to hide.", "Hide Columns", ran.Address, Type:=8)`. 

The parameters in `Application.InputBox()` are:
1. The main message of the pop-up box
2. The main title of the pop-up box
3. The location of the cells being selected
4. The type that is being selected

In this case, the type being selected is a cell reference. There is then a for loop that iterates through all the selected cells and hides the entire column. 

**NOTE**: You only need to select one cell within each column. Selecting a whole column will cause Excel to crash.

### Hide Rows
The second section of the macro is:
`Dim rnge As range
Dim row As range
On Error Resume Next
Set rnge = Application.Selection
Set rnge = Application.InputBox("Select a range of Rows you would like to hide.", "Hide Rows", rnge.Address, Type:=8)
For Each row In rnge
row.EntireRow.Hidden = True
Next row`

This block of code is similar to *Hide Columns* section as it allows the user to select which rows they would like to remove. The only difference, besides the variable names, is that instead of `EntireColumn.Hidden = True` it is `EntireRow.Hidden = True`.

**NOTE**: You only need to select one cell in the row you want to remove. Selecting a whole row will cause Excel to crash.

### Hide Cell Content
The third section of the macro is:
`Dim rang As range
Dim cel As range
On Error Resume Next
Set rang = Application.Selection
Set rang = Application.InputBox("Select a range of Cells you would like to hide.", "Hide Cells", rang.Address, Type:=8)
For Each cel In rang
cel.Select
ActiveCell.NumberFormat = ";;;"
Next cel`

Like the previous sections, the user is able to select the cells that they want the contents to be hidden. The way that the content is hidden is by setting the number format to `";;;"` which is a custom formula that allows the cell to retain the information, but removes the text.

### Remove Unwanted Hyperlinks
The last section of the macro is:
`Dim rng As range, cell As range
Dim xLink As Hyperlink
Set rng = Application.Selection
Set rng = Application.InputBox("Select a range of Hyperlinks you would like to remove.", "Remove Hyperlinks", rng.Address, Type:=8)
For Each cell In rng
cell.Select
ActiveCell.Font.ColorIndex = 1
ActiveCell.Font.Underline = False
ActiveCell.Interior.ColorIndex = -4142
ActiveCell.ClearHyperlinks
Next cell`

As in the preveious sections, the user can select which cells they would like the hyperlinks to be removed from. The text is then changed to black, the undeline is removed, and the background color is set back to default. Then the hyperlink is cleared from the cell.

### Example Of a Table Cleaned-up with the Macro
|  Date  |                    Topic                   |                    Due                    |
|:------:|:------------------------------------------:|:-----------------------------------------:|
| 18-Jan |  What is Data Science? [In Class Reading](http://jse.amstat.org/v23n2/witmer.pdf)   | Make sure you have Excel on your computer |
| 20-Jan |                 Excel & VBA                |              [Excel Homework](https://docs.google.com/document/d/1g8eOYNe9sDmrstRgvFRZBskxjaIaD7Za4lFXSgPPkVw/edit)             |
| 25-Jan | Excel Presentations and Writing about Data |                                           |
| 27-Jan |       Writing about Data and Work Day      |               [VBA Homework](https://docs.google.com/document/d/1bTkmUon_Kq6_DupNw2Szh-T4rFGqzeA2aIIBy7m1yhk/edit)              |
|  1-Feb |    Why is Data Visualization Important?    |           [Florence Nightengale](https://docs.google.com/forms/d/e/1FAIpQLSeL-qfdJBp5YGpPWiKXRBsypEZTh9TTMcv1g5TrqOWTx_NF7A/viewform?hr_submission=ChkIq_Gs4d8BEhAI0Zip0ZoNEgcIgoif9PgMEAE)          |
|  3-Feb |                   Tablau                   |                                           |
|  8-Feb |              Chart Readability             |               How Charts Lie              |
| 10-Feb |                   Tablau                   |            Tablau Homework Due            |
| 15-Feb |             Data Annonymization            |         How to UN-Annonymize Data         |
| 17-Feb |             Sentiment Analysis             |             First Presenation             |
| 22-Feb |                    VADER                   |                                           |
| 24-Feb |                      R                     |      Sentiment Analysis Homework Due      |
|  1-Mar |                   More R                   |                                           |
|  3-Mar |                   GCPLOTS                  |                                           |
|  8-Mar |                More GCPLOTS                |                                           |
| 10-Mar |               GCPLOTS FOREVER              |            GCPLOTS Homework Due           |
| 15-Mar |                SPRING BREAK                |                                           |
| 17-Mar |             SPRING BREAK - WOO             |                                           |
| 22-Mar |            Remember Statistics?            |                                           |
| 24-Mar |               Random Forests               |                                           |
| 29-Mar |                Advising Day                |                                           |
| 31-Mar |                Presentations               |        Random Forests Homework Due        |
|  5-Apr |      Clustering - K Nearest Neighbors      |                                           |
|  7-Apr |                                            |                                           |
| 12-Apr |                                            |          Clustering Homework Due          |
| 14-Apr |                                            |                                           |
| 19-Apr |                                            |                                           |
| 21-Apr |           Practice Presentations           |                 Milestone                 |
| 26-Apr |                                            |                                           |
| 28-Apr |            Project Presentations           |             Final Project Due             |
