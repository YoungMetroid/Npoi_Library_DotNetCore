using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.XSSF.UserModel;
using UtilityLibrary.Loggers;

namespace Npoi_Library_DotNetCore
{
	public class NpoiExcelReadWrite
	{
		protected XSSFWorkbook workbook;
		protected XSSFSheet sheet;
		protected FileStream file;
		private Logger logger;
		private XSSFCellStyle defaultCellStyle;
		private string defaultDateFormat = "MM/dd/yyyy";
		public NpoiExcelReadWrite(string ExcelDocument)
		{
			try
			{
				file = new FileStream(ExcelDocument, FileMode.Open, FileAccess.ReadWrite);
				workbook = new XSSFWorkbook(file);
				file.Close();
				defaultCellStyle = (XSSFCellStyle)workbook.GetCellStyleAt(0);
			}
			catch(Exception ex)
			{
				
			}
		}
		public void setSheet(string sheetName)
		{
			sheet = workbook.GetSheet(sheetName) as XSSFSheet;
		}
		public void setSheet(int sheetNumber)
        {
			sheet = workbook.GetSheetAt(sheetNumber) as XSSFSheet;
        }
		public void saveFile(string name)
		{
			try
			{
				file = File.Create(name);
				workbook.Write(file);
				file.Close();
			}
			catch(Exception ex)
            {
				logger.logException(ex);
            }
		}
		public void saveFile2(string name)
        {
			try
			{
				using (FileStream file = new FileStream(name, FileMode.Open, FileAccess.Write))
				{
					workbook.Write(file);
					file.Close();
				}
			}
			catch(Exception ex)
            {
				logger.logException(ex);
            }
		}
		public void createRowsInstance(int rowCount)
		{
			for (int counter = 0; counter < rowCount + 100; counter++)
			{
				sheet.CreateRow(counter);
			}
		}
		public void createRowsRange(int startRow, int numberOfRows)
		{
			
			for (int counter = startRow; counter <= numberOfRows; counter++)
			{
				sheet.CreateRow(counter);
			}
		}
		public int getLastRow()
		{
			return sheet.LastRowNum;
		}
		public int WriteArray_To_Excel(int rowAvailableCell, int startingCol, string[,] infoArray)
		{
			if(rowAvailableCell+infoArray.GetUpperBound(0) > sheet.LastRowNum)
			{
				createRowsRange(rowAvailableCell, rowAvailableCell + infoArray.GetUpperBound(0));
			}
			for (int rowCounter = 0; rowCounter <= infoArray.GetUpperBound(0); rowCounter++)
			{
				for (int columnCounter = 0; columnCounter <= infoArray.GetUpperBound(1); columnCounter++)
				{
					try
					{
						sheet.GetRow(rowAvailableCell + rowCounter).CreateCell(columnCounter + startingCol);
						sheet.GetRow(rowAvailableCell + rowCounter).GetCell(columnCounter + startingCol).SetCellValue(infoArray[rowCounter, columnCounter]);
					}
					catch (Exception ex)
					{
						Console.WriteLine(ex.Message);
					}
				}
			}
			return startingCol + infoArray.GetUpperBound(1) + 1;
		}
		public void WriteList_To_Excel(int rowAvailableCell, int startingCol, int startingReadColumn, int lastReadColumn, List<List<object>> data, int type)
		{
			XSSFCellStyle cellStyle = (XSSFCellStyle)workbook.CreateCellStyle();
			cellStyle.CloneStyleFrom(defaultCellStyle);
			var newDataFormat = workbook.CreateDataFormat();

			for (int rowCounter = 0; rowCounter < data.Count; rowCounter++)
			{
				for (int columnCounter = 0, dataStartColumn = startingReadColumn; dataStartColumn <= lastReadColumn; columnCounter++, dataStartColumn++)
				{
					try
					{
						XSSFCell cell = (XSSFCell)sheet.GetRow(rowAvailableCell + rowCounter).CreateCell(columnCounter + startingCol);
						if (data[rowCounter][dataStartColumn] != null)
							switch (type)
							{
								case 0:
									double value;
									if (Double.TryParse(data[rowCounter][dataStartColumn].ToString(), out value))
										cell.SetCellValue(value);
									else cell.SetCellValue(data[rowCounter][dataStartColumn].ToString());
									cell.CellStyle = cellStyle;
									break;
								case 1:
									try
									{
										cell.SetCellValue(data[rowCounter][dataStartColumn].ToString());
										cell.CellStyle = cellStyle;
									}
									catch (Exception ex)
									{
										cell.SetCellValue("");
										cell.CellStyle = cellStyle;
									}
									break;
								case 2:
									int dateValue;
									DateTime dateTime;
									cell.CellStyle = cellStyle;
									cell.CellStyle.DataFormat = newDataFormat.GetFormat(defaultDateFormat);

									if (int.TryParse(data[rowCounter][dataStartColumn].ToString(), out dateValue))
									{
										cell.SetCellValue(dateValue);
									}
									else if (DateTime.TryParse(data[rowCounter][dataStartColumn].ToString(), /*CultureInfo.CreateSpecificCulture("en-US"), DateTimeStyles.None*/ out dateTime))
									{
										cell.SetCellValue(dateTime);
									}
									else
									{

										cell.SetCellValue(data[rowCounter][dataStartColumn].ToString());
									}
									break;
							}
					}
					catch (Exception ex)
					{
						logger.logException(ex);
					}
				}
			}
		}
	}
}
