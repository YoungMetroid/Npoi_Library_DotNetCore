# Npoi-Library-DotNetCore
This is a copy of the Npoi Library but in the DotNetCore FrameWork

## Table of Contents
* [Summary](#Summary)
* [Dependencies](#Dependencies)
* [NpoiExcelCreator](#NpoiExcelCreator)
  * [Create-Excel-Sheet](#Create-Excel-Sheet)
* [NpoiExcelReader](NpoiExcelReader)



## Summary 

This class allows you to create a new excel file and paste information into it. Your able to pass information thats in a `object[,]` or in a `List<List<object>>` just be sure to intialize the rows using the `createRowsInstance(int rowCount) function` since they are all null when creating a new Excel file. Your also able to configure what type of text font, font type or cell style you want with the following functions:

* setFontStyle(string font)
* setFontSize(int size)
* setFontColor(byte red, byte green, byte blue)
* setCellColor(byte red, byte green, byte blue)
* setCellStyle(int firstRow, int lastRow, int firstColumn, int lastColumn,byte red, byte green, byte blue, int fontSize, string fontType)

## Dependencies
To be able to use Npoi-Library you'll need to include the [Utility Library DotNetCore](https://github.azc.ext.hp.com/SLAC-Dev/UtilityLibrary_DotNetCore).
## NpoiExcelCreator

### Create-Excel-Sheet
The following is a code snippet on how to create a Excel Sheet and store info into. The columns for every list in the `List<List<object>>` all need to have the same count if not you'll get outof bounds exception. This function pastes the info in the excel sheet as if they were tables meaning that every column and row cannot be null.
```C#
      NpoiExcelCreator excel = new NpoiExcelCreator();
			excel.createSheet("Test Sheet");
			excel.setSheet(0);
			List<List<object>> names = new List<List<object>>();
			List<object> nameInfo = new List<object>();


			nameInfo.AddRange(new List<object> { "FirstName", "LastName" });
			names.AddRange(new List<List<object>> { nameInfo });

			nameInfo = new List<object>();
			nameInfo.AddRange(new List<object>{ "Bob","Jones1"});
			names.AddRange(new List<List<object>> { nameInfo });

			nameInfo = new List<object>();
			nameInfo.AddRange(new List<object>{ "Phillip","Jones2"});
			names.AddRange(new List<List<object>> { nameInfo });


			nameInfo = new List<object>();
			nameInfo.AddRange(new List<object>{ "Mine","Jones3"});
			names.AddRange(new List<List<object>> { nameInfo });

			nameInfo = new List<object>();
			nameInfo.AddRange(new List<object>{ "Craft","Jones4"});
			names.AddRange(new List<List<object>> { nameInfo });

			excel.createRowsInstance(names.Count);
			excel.WriteList_To_Excel(0, 0, 0, names.Count()-1, names, 0);
			excel.saveFile(@"C:\TestFolder\Testing.xlsx");

```
You should get the following result in the excel file:

![alt-text](https://github.azc.ext.hp.com/SLAC-Dev/Npoi-Library-DotNetCore/blob/master/ReadMe%20Resources/Example1.PNG)

## NpoiExcelReader

We will read the following excel file: [Animals.xlsx](https://github.azc.ext.hp.com/SLAC-Dev/Npoi-Library-DotNetCore/blob/master/ReadMe%20Resources/Animales.PNG)

![alt-text](https://github.azc.ext.hp.com/SLAC-Dev/Npoi-Library-DotNetCore/blob/master/ReadMe%20Resources/Example2.PNG)

In the following code snippet you'll see that we Instantiate the Logger and set the path and name for the log file. If any errors were to occur then you'll see the info in that file. After that you'll see that we Create the `NpoiExcelReader` object and instantiate it by passing a file by parameter. We then indicate what sheet its going to read. Then we'll read the file with the `readSheet_ReturnLists` function. This function has 4 parameters which indicate in the first 2 parameters if there is a offset in the Rows. The first parameter will indicate how many rows to skip and the second parameter will indicate how many rows you should not read from the last row. Meaning if there are 10 row and you pass 2 by parameter then you'll only read to the eighth row. The third parameter indicates the first row to read and the fourth row indicates the first column to read.
`In most cases you'll only need to use the third and fourth parameter the first and second is good when you know that the file has the info in a pivot table.`

```C#
static class Program
{
        private const string TestFolderPath = @"C:\TestFolder\";
        private static Logger logger = Logger.getInstance;
        static void Main(string[] args)
        {
            logger.setLogPathandFile(TestFolderPath, "Error.log");
            NpoiExcelReader npoiExcelReader = new NpoiExcelReader(TestFolderPath + "Animals.xlsx");
            npoiExcelReader.setSheet(0);
            List<List<object>> table = npoiExcelReader.readSheet_ReturnLists(0, 0, 0, 0);

            List<object> dogs = table.Select(x => x[0]).ToList();
            List<object> cats = table.Select(x => x[1]).ToList();
            List<object> birds = table.Select(x => x[2]).ToList();

            dogs.ForEach(x => Console.WriteLine(x));
            Console.WriteLine();
	    
            cats.ForEach(x => Console.WriteLine(x));
            Console.WriteLine();
	    
            birds.ForEach(x => Console.WriteLine(x));
            Console.WriteLine();
	    
            Console.ReadKey();
        }
}
```

You'll see in the code that we use linq to seperate the info by columns. We create a list of dogs, cats and birds. And then we print the info so you should be able to see the following result:

![alt-text](https://github.azc.ext.hp.com/SLAC-Dev/Npoi-Library-DotNetCore/blob/master/ReadMe%20Resources/Example3.PNG)
