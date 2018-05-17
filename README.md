# Excel-Workshop
The following is a short Excel Workshop intended for UCSD math students to be introduced to some of the many useful functions in Excel
[TOC]
## Files
The three main types you will generally encounter will be 
- csv: Comma-seperated values
- xls: Microsoft Excel format
- xlsm Excel Macro enabled format


As the names suggest, csv's are generally used to store abundances of data as these take the least amount of space and can be openend/imported by many applications. These files cannot save any thing extra other than the data values in the cells as suggested by the "Possible Data Loss" warning you will see when opening one of these files. If you don't believe me try even highlighting a column, save, and reload the file. 

xls files on the other hand can only be opened by certain applications and saves the many different abillities excel has like highlighting, tables, graphs, charts, etc. 

xlsm files take another step from xls files and give the user the abilitiy to run macros which we will go over later in this workshop. 
## Data
Although many times datasets come conveniently in csv files, there are ocassionaly times where one must copy paste a data set from a txt file where the values may be seperated by a combination of tabs,spaces,commas,etc. We follow the procedure:
1. copy paste the attached ExcelWorkshop.txt file contents(crtl/cmd a, ctrl/cmd c, ctrl/cmd v) into the first column of your excel sheet.  
2. Go to the data section on the excel ribbon on top 
3. Click next with Delimited selected
4. Select Comma for this example, however, as mentioned a combination of tabs,semicolon,etc may be used. 
5. Click finish and your data should be all nice and arranged. 

## Pivot tables
Now excel also has normal table functions, however, those are boring so we shall jump straight into pivot tables. 
1. Highlight all your data by either crl/cmd a or click the south east poiting error at the top left corner of the cell titles(if your data is the whole sheet). 
2. Go to the Insert section on the excel ribbon on top.
3. Click pivot table.


Since in this data set we are looking at trends for CPI (Customer Purchase Intensions), click the "CPI" checkbox under the field names of the pivot table and you should see it immediately show the Count. Now drag the "store" checkbox into the row section and you should be able to see the counts of CPI seperated by store numbers. You can change how you want CPI to be aggregated by clicking on the little "i" icon next to its name in the Values section of the pivot table. For this example we will change count to average. Now one might also want to know the standard deviations of the CPI, however, not want to lose the average column. All you need to do then is drag the CPI from the field's list again into the "Values" section. Click the "i" and change this one into "stdev". (If you are wondering the difference between "stedv" and "stedvp", "stedvp" assumes the data used is the total population and therefore does not use the bias correction). 
Next, we can look at these CPI values in specific years. Drag the "Date" variable into the column section to do this. 
Lastly, we can add filters. Drag the "IsHoliday" variable in to filters and notice the popup at the top left. You can use the drag down here to switch between looking at all the data or just the subset of the Isholiday categories. 
Double clicking on a specific cell will pull up the data table of the values used for that cell. 
Not a fan of pivot charts because R is way better, however, you can easily just make a bar chart under the insert ribbon. Notice that this chart is linked to the pivot table so if we remove a variable like "date", the bar graph will imideately change as well. 

## Vlookup
Another neat tool excel has is the ability to reference and match values based on a seperate table by linking a definition variable. 
We show case this example:
1. Adjust the pivot table created earlier to just the store numbers and average CPI
2. We shall now match the average CPI to each store in our original data
3. On column H cell 2, type into the cell =Vlookup()
4. Click on A2 since this is the definition variable we will be matching to
5. Press comma, then highlight the pivot table from the other sheet.
6. Press comma, then type 2 since we are looking to match values from column 2 in the pivot table
7. Close the parenthesis and your equation should look something like this: =VLOOKUP(A2,SheetX!A:B,2)
8. We use autofill to complete the formula for the rest of the row. (double click the little square at the bottom of the cell with your formula)

We use autofill again to get a column of differences from the mean for the next example.

## Conditional Formatting
Another neat tool in excel is to highlight specific rows or cells based on values that they might contain. For this example we shall highlight each row where the difference in CPI is negative. 
1. Make sure your selected cell is on the second row
2. Under the "Home" tab on the excel ribbon, click on conditional formatting, new rule, and change the style to Classic.
3. Select "Use a formula to determine which cells to format"
4. Type "= $I2 < 0" and click OK after choosing what format you would like
5. The cell that was selected should now be formatted in whatever option you chose
6. Go back to conditional formatting and click "Manage Rules"
7. Under Applies to and select all of your data.

$\underline{Note}$ that the $ sign infront of the I was to fix the equation so that whenever it is copied and adjusted for different cells it will not adjust the I and continue to look at the I column whereas it will adapt and change the 2 since it should be looking at I 3 when considering row 3 and etc. 
A neat thing about this function is that if we change one of the values, lets say I 95, to a value that secifies our rule, it will automaticaly change the format of that row. 

## Macros
Macro allow us to do everything we just did manually, with a click of a button. To showcase this, we shall create a macro for conditional formatting since we have that fresh in our minds. 
1. First go to manage rules and delete the rule we just made so that we can remake it. 
2. (skipable) If you do not have the developer ribbon, go to Excel ->  Preferences -> Ribbon & Toolbar -> Check the Developer bar under Main Tabs on the right.
3. Under the Developer section now, click record macro and create a conditional format for cells containing NA.(For some reason macros recording does not register customize formulas) 
4. We can now put this macro on a button!
5. Click the button from the develop tool then click anywhere on the excel sheet and select the macro you want it to be linked with. 
6. Go to manage rules and remove the rule. Click the button and you can see that it instantly recreates the rule! WAW
7. Instead of having an ugly button stuck on a specfic page we can also attach a macro directly to the ribbon. (to delete the button  right click and press cut cause idk how else)
8. Excel ->  Preferences -> Ribbon & Toolbar -> Choose commands from -> Macros - > Create a new group(on right) -> send macro to that group
9. (Note) on windows there is a modify button that allows you to put a picture on your macro buttons, however, I am currently unable to find this fucntion for mac. 

## VBA
Now that we have seen macros, VBA comes naturally since you can instantly view the code for your macro by clicking edit. The great thing about this is you can find the functions you want to use most of the time by just recording a macro and then looking at its code. Like every other language, VBA has the usual if-else, for and while loops, and integer,double,string variables. 
To initialize variables we use:
```
Dim i as Integer
Dim text as String 
Dim yes as Boolean
```
Note that Vba is not case sensitive so feel free to butcher your variables. JK don't but it wont penalize you if you forget a capital somewhere. 
InputBox allows to, yes you guessed it, make and input box. So we can easily customize the macro we just made so that it is adaptable to look for any strings and highlight them. 
```
    'Comments in VBA are done with this apostrophe.
    Dim i as String
    i =  InputBox("Enter String to find", "String")
    'Same code as before however we add a variable i
    Range("A1:I8191").Select
    Range("E5").Activate
    Selection.FormatConditions.Add Type:=xlTextString, String:= i, _ 'Change this to i
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub
```

But now we can also recreate the original conditional formatting example with the custom equation. Where xlExpression is the key distrinction between the codes. 
```
    Range("A1:I8191").Select
    Range("E5").Activate
    'Selection.FormatConditions.Add Type:=xlTextString, String:=i, _
        TextOperator:=xlContains
        
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$I2 < 0"
    
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub

```
We use & to concatenate strings so we can adjust the code section 
```
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$I2 <" & i
```
to be adjustable where we have 
```
Dim i as Integer
i = InputBox("Yes", "enter here")
```
initialized in the beginning. 
### Notations

Since the general logic of for/while loops and if/else statements are the same as other coding languages, we shall just include the VBA equivalent notations here.
#### if-else
```
        If selection = "A1" Then
            Console.WriteLine("How")
        ElseIf selection = "A2" Then
            Console.WriteLine("many")
        ElseIf selection = "A3" Then
            Console.WriteLine("shrimps")
        ElseIf selection = "A4" Then
            Console.WriteLine("do")
        Else
            Console.WriteLine("you have to eat " + selection)
        End If
        Console.ReadLine()
```
Where the notable difference will be "End If"
#### for loops
```
For i As Integer = 10 To 0 Step -1
    'Do something here
Next i  
```
#### While loops
```
While 'Condtion
  'Do stuff
Wend
```
#### With
The with statement we have already seen basically allows the user to omit rewriting the start of a variable.
i.e.
```
x.y.z
x.y.w
x.y.k
```
is equivalent to
```
With x.y
    .z
    .w
    .k
End With
```
