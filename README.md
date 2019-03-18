## Demystify-Formula

At work, I often find myself looking at huge excel fomrulae with lot of cell addresses being refered in them. Without any exageration, some of the formulae that I see have more than 50 other cell addresses being refered in them. It becomes extremely difficult to make sense of logic of these formula due to the information being lost in the cell addresses. That's why, I created this plug-in which helps me define name ranges for columns, and then convert all the formula to their variable representation to make them instantly readable with two clicks of mouse.

An example may be, when you look at something like "= $M9 * $N9 * $N9 * $O9 * 0.001" It becomes difficult to guess what is being done here. But using the plug-in, you can convert it to something like "= Capacitance_nF_Col_M * Voltage_V_Col_N * Voltage_V_Col_N * Frequency_MHz_Col_O * 0.001" . This instantly tells you, its a P=CV<sup>2</sup>f equation for power calculation. The macro also handles off-sheet-references which are not handled by "Apply Names" feature of Excel.

## Demo

![Demo DemystifyFormula](https://github.com/gandhidarshak/DemystifyFormula/blob/master/Demo/DemystifyFormulaDemo.gif)

## How to install?

### Download the plugin file
Download or git-clone the DemystifyFormula.xlam file. Keep it in a location which would not change that often.

### Add Plugin to Excel 
1. Make sure you have Developer's tab visible in Excel. If not visible, On the File tab, go to Options > Customize Ribbon. Under Customize the Ribbon and under Main Tabs, select the Developer check box.
2. Open an Excel workbook and Go to Developer –> Add-ins –> Excel Add-ins
3. In the Add-ins dialogue box, browse and locate the DemystifyFormula.xlam file that you saved, and click OK.
At this point CopyAsTable, CopyAsCSV, CopyAsYaml Macros are added to your excel.

### Add Macros to Ribbon
1. Right-click on any of the ribbon tabs and select Customize the Ribbon
2. Create a New Tab (MyMacros in the Demo above) or use some existing tab depending on where you want the buttons to show up.
3. Create a New Group (Demystify Formula in the Demo above). Make sure the tab and groups are visible by selecting their check boxes.
4. In Choose commands from dropdown list, pick Macros. 
5. Three macros should be visible there. Add them to the group created above. Rename and change icons if preferred. 
At this point, you should have the three ribbon buttons set-up and ready to use. Enjoy!



## One last thing ...
If you have to deal with Excel formulae as much as I do, do visit this awesome website which will significantly improve your ability to make sense of excel formula's by beautifying them: http://excelformulabeautifier.com


## License

DemystifyFormula uses the MIT license. See [LICENSE](https://github.com/gandhidarshak/DemystifyFormula/blob/master/LICENSE) for more details.

