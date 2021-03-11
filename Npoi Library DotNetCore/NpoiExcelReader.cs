using NPOI.HSSF.UserModel;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using UtilityLibrary.Loggers;

namespace Npoi_Library_DotNetCore
{
    public class NpoiExcelReader
    {
        protected IWorkbook workbook;
        protected ISheet sheet;
        const string alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        protected string numbers;
        protected string letters;
        protected int firstPivotRow = 0;
        protected int lastPivotRow = 0;
        protected int firstPivotColumn = 0;
        protected int lastPivotColumn = 0;
		private Logger logger;
        bool isXlsSheet;

		public NpoiExcelReader(string ExcelDocument)
		{   
            logger = Logger.getInstance;
			try
			{
				FileStream file = new FileStream(ExcelDocument, FileMode.Open, FileAccess.Read);
                if (ExcelDocument.Contains(".xlsx") || ExcelDocument.Contains(".xlsm"))
                {
                    workbook = new XSSFWorkbook(file);
                    isXlsSheet = false;
                }
                else if (ExcelDocument.Contains(".xls"))
                { 
                    workbook = new HSSFWorkbook(file);
                    isXlsSheet = true;
                }
            }
			catch(IOException ex)
			{
				Console.WriteLine(ex.Message);
				Console.WriteLine(ex.StackTrace);
				logger.logException(ex);
			}
			catch(Exception ex)
			{
				Console.WriteLine(ex.Message);
				logger.logException(ex);
			}
        }
		public bool isWorkBookOpen()
		{
			if (workbook == null)
				return false;
			return true;
		}
        public bool setSheet(string sheetName)
        {
			if(isXlsSheet)
			    sheet = workbook.GetSheet(sheetName) as HSSFSheet;
            else
			    sheet = workbook.GetSheet(sheetName) as XSSFSheet;
			resetVaules();
			if (sheet != null) return true;
			return false;
        }
        public bool setSheet(int sheetNumber)
        {
            if (isXlsSheet)
                sheet = workbook.GetSheetAt(sheetNumber) as HSSFSheet;
            else
                sheet = workbook.GetSheetAt(sheetNumber) as XSSFSheet;
            resetVaules();
            if (sheet != null) return true;
            return false;
        }
        public int getFirstPopulatedRow(string firstHeader)
        {
            for(int row = 0; row < 30; row++)
            {
                ICell cell = sheet.GetRow(row).GetCell(0);
               if(cell != null && cell.ToString().Trim() == firstHeader)
               {
                    return row;
               }
            }
            return -1;
        }
        public object[,] readSheet(int offSetRowPlus, int offSetRowMinus, int startingRow, int startingColumn)
        {
            Console.WriteLine(sheet.GetType());
            if (!isXlsSheet)
            {
                XSSFSheet sheetTemp = sheet as XSSFSheet;
                List<XSSFPivotTable> tables = sheetTemp.GetPivotTables();
                readPivotTable(sheetTemp);
            }
            
            if (lastPivotRow == 0)
            {
                firstPivotRow = startingRow;
                lastPivotColumn = sheet.GetRow(startingRow).LastCellNum;
				ICell currentLastCell=  sheet.GetRow(startingRow).GetCell(lastPivotColumn);
				while(currentLastCell  == null || string.IsNullOrWhiteSpace(currentLastCell.ToString()))
				{
					lastPivotColumn--;
					currentLastCell = sheet.GetRow(startingRow).GetCell(lastPivotColumn);
				}
				if(currentLastCell != null)
				{
					lastPivotColumn++;
				}
                //Console.WriteLine(sheet.LastRowNum);
                lastPivotRow = sheet.LastRowNum + 1;
                for (int counter = startingRow; counter < sheet.LastRowNum; counter++)
                {
                    IRow row = sheet.GetRow(counter);
                    if (row.GetCell(startingColumn) == null ||  string.IsNullOrEmpty(row.GetCell(startingColumn).ToString()))
                    {
                        lastPivotRow = counter;
                        break;
                    }
                }
            }
            object[,] tableInfo = new object[lastPivotRow-firstPivotRow - offSetRowMinus,lastPivotColumn];

            try
            {
                for (int rowCounter = firstPivotRow + offSetRowPlus, counter = 0; rowCounter < lastPivotRow - offSetRowMinus; rowCounter++, counter++)
                {
                    try
                    {
                        IRow row = sheet.GetRow(rowCounter);
                        for (int colCounter = 0; colCounter < lastPivotColumn; colCounter++)
                        {
                            try
                            {
                                if (row.GetCell(colCounter) != null)
                                    tableInfo[counter, colCounter] = row.GetCell(colCounter).ToString();
                            }
                            catch(Exception ex)
                            {
                                //The cell is null and there it no way to catch with the if statement
                                logger.logException(ex);
                                logger.addTextToLogFile("Row: " + rowCounter + " Column: " + colCounter + " Object not instantiated");
                            }
                        }
                    }
                    catch(Exception ex)
                    {
                        logger.logException(ex);
                        logger.addTextToLogFile("Row: " + rowCounter);
                    }
                }
            }
            catch(Exception ex)
            {
				logger.logException(ex);
            }
            #if Debug
                Console.WriteLine(tableInfo[tableInfo.Count - 1][0, 2]);
            #endif
            resetVaules();
            
            return tableInfo;
        }
        public List<List<object>> readSheet_ReturnLists(int offSetRowPlus, int offSetRowMinus, int startingRow, int startingColumn)
        {
            if (sheet.GetType() == typeof(XSSFWorkbook))
            {
                XSSFSheet sheetTemp = sheet as XSSFSheet;
                List<XSSFPivotTable> tables = sheetTemp.GetPivotTables();
                readPivotTable(sheetTemp);
            }

            if (lastPivotRow == 0)
            {
                firstPivotRow = startingRow;
                lastPivotColumn = sheet.GetRow(startingRow).LastCellNum;
                ICell currentLastCell = sheet.GetRow(startingRow).GetCell(lastPivotColumn);
                while (currentLastCell == null || string.IsNullOrWhiteSpace(currentLastCell.ToString()))
                {
                    lastPivotColumn--;
                    currentLastCell = sheet.GetRow(startingRow).GetCell(lastPivotColumn);
                }
                if (currentLastCell != null)
                {
                    lastPivotColumn++;
                }
                //Console.WriteLine(sheet.LastRowNum);
                lastPivotRow = sheet.LastRowNum + 1;
                for (int counter = startingRow; counter < sheet.LastRowNum; counter++)
                {

                    IRow row = sheet.GetRow(counter);
                    if (row == null || row.GetCell(startingColumn) == null || string.IsNullOrEmpty(row.GetCell(startingColumn).ToString()))
                    {
                        lastPivotRow = counter;
                        break;
                    }

                }
            }
            List<List<object>> tableInfo = new List<List<object>>();
            
            try
            {
                for (int rowCounter = firstPivotRow + offSetRowPlus, counter = 0; rowCounter < lastPivotRow - offSetRowMinus; rowCounter++, counter++)
                {
                    IRow row = sheet.GetRow(rowCounter);

                    List<object> tempRow = new List<object>();
                      DateTime date;
                    //DateTime.TryParse(row.GetCell(38).DateCellValue.ToString(), out date);

                    for (int colCounter = startingColumn; colCounter < lastPivotColumn; colCounter++)
                    {
                        if (row.GetCell(colCounter) != null)
                        {
                            tempRow.Add(row.GetCell(colCounter));
                        }
                        else
                            tempRow.Add(String.Empty);
                    }
                    tableInfo.Add(tempRow);
                }
            }
            catch (Exception ex)
            {
                logger.logException(ex);
            }
#if Debug
                Console.WriteLine(tableInfo[tableInfo.Count - 1][0, 2]);
#endif
            resetVaules();
            return tableInfo;
        }

        private void readPivotTable(XSSFSheet sheet)
        {
            List<XSSFPivotTable> tables = sheet.GetPivotTables();
            if (tables.Count > 0)
            {
                try
                {
                    CT_PivotTableDefinition cT_PivotTableDefinition = tables[0].GetCTPivotTableDefinition();
                    parseLocation(cT_PivotTableDefinition.location.@ref);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
					logger.logException(ex);
                }
            }
        }
        private void parseLocation(string location)
        {
            letters = Regex.Replace(location, @"[\d]", string.Empty);
            numbers = Regex.Replace(location, @"[^\d:]", string.Empty);
            firstPivotRow = Convert.ToInt32(Regex.Replace(numbers, @"[:][\d]+", string.Empty));
            lastPivotRow = Convert.ToInt32(Regex.Replace(numbers, @"[\d]+:", string.Empty));
            firstPivotColumn = convertColumn(Regex.Replace(letters, @"[:][\w]+", string.Empty));
            lastPivotColumn = convertColumn(Regex.Replace(letters, @"[\w]+[:]", string.Empty));
        }
        private void resetVaules()
        {
            firstPivotRow = 0;
            lastPivotRow = 0;
            firstPivotColumn = 0;
            lastPivotColumn = 0;
        }

        private int convertColumn(string letters)
        {
            List<int> letterPositionValue = new List<int>();
            for (int counter = 0; counter < letters.Length; counter++)
            {
                for (int alphabetCounter = 0; alphabetCounter < alphabet.Length; alphabetCounter++)
                {
                    if (letters[counter] == alphabet[alphabetCounter])
                    {
                        letterPositionValue.Add(alphabetCounter + 1);
                    }
                }
            }

            for (int letterCount = letters.Length; letterCount > 1; letterCount--)
            {
                letterPositionValue[letters.Length - letterCount] = letterPositionValue[letters.Length - letterCount] * alphabet.Length;
            }
            int totalValue = 0;
            foreach (int value in letterPositionValue)
            {
                totalValue += value;
            }
            return totalValue;
        }

        public void releaseMemory()
        {
            try
            {
                workbook.Close();
                workbook = null;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
				logger.logException(ex);
            }
        }
    }
}
