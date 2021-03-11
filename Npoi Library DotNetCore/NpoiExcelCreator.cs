
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UtilityLibrary.Loggers;

namespace Npoi_Library_DotNetCore
{
    public class NpoiExcelCreator
    {
        protected XSSFWorkbook workbook;
		protected IWorkbook iworkbook;
        protected XSSFSheet sheet;
		private string fontStyle = "Arial";
		private byte[] fontRgb = new byte[3] { 0, 0, 0 };
		private byte[] cellRgb = new byte[3] { 255, 255, 255 };
		private int fontSize = 14;
		private string defaultDateFormat = "MM/dd/yyyy";
		private XSSFCellStyle defaultCellStyle;
		private Logger logger;
		private ConsoleLogger consoleLogger;

		public NpoiExcelCreator()
        {
			try
			{
				logger = Logger.getInstance;
				consoleLogger = ConsoleLogger.getInstance;
				//iworkbook = new HSSFWorkbook();
				workbook = new XSSFWorkbook();
				XSSFFont ffont = (XSSFFont)workbook.CreateFont();
				ffont.FontHeight = fontSize * fontSize;
				ffont.SetColor(new XSSFColor(fontRgb));
				ffont.FontName = fontStyle;

				defaultCellStyle = (XSSFCellStyle)workbook.GetCellStyleAt(0);
			}
			catch(Exception ex)
			{
				logger.logException(ex);
			}
			//defaultCellStyle.SetFillForegroundColor(new XSSFColor(cellRgb));
			//defaultCellStyle.FillPattern = FillPattern.SolidForeground;
			//defaultCellStyle.SetFont(ffont);
			//defaultCellStyle.BorderBottom = BorderStyle.Thin;
			//defaultCellStyle.BorderTop = BorderStyle.Thin;
			//defaultCellStyle.BorderLeft = BorderStyle.Thin;
			//defaultCellStyle.BorderRight= BorderStyle.Thin;
        }
        public void createSheet(string sheetName)
        {
			try
			{
				workbook.CreateSheet(sheetName);
			}
			catch(Exception ex)
			{
				logger.logException(ex);
			}
        }
        public void setSheet(string sheetName)
        {
			try
			{
				sheet = workbook.GetSheet(sheetName) as XSSFSheet;
			}
			catch(Exception ex)
			{
				logger.logException(ex);
			}
        }
		public void setSheet(int sheetNumber)
		{
			try
			{
				sheet = workbook.GetSheetAt(sheetNumber) as XSSFSheet;
			}
			catch(Exception ex)
			{
				logger.logException(ex);
			}
		}
        public void saveFile(string name)
        {
			try
			{
				FileStream sw = File.Create(name);
				workbook.Write(sw);
				sw.Close();
			}
			catch(FileLoadException ex)
			{
				logger.logException(ex);
				Console.ForegroundColor = ConsoleColor.Red;
				Console.WriteLine(ex.Message);
			}
			catch(Exception ex)
			{
				logger.logException(ex);
				Console.ForegroundColor = ConsoleColor.Red;
				Console.WriteLine(ex.Message);
			}
        }
        public void createRowsInstance(int rowCount)
        {
            for(int counter = 0; counter < rowCount; counter++)
            {
                sheet.CreateRow(counter);
            }
        }
		public void setFontStyle(string font)
		{
			fontStyle = font;

			XSSFFont ffont = (XSSFFont)workbook.CreateFont();
			ffont.FontHeight = fontSize * fontSize;
			ffont.SetColor(new XSSFColor(fontRgb));
			ffont.FontName = fontStyle;
			defaultCellStyle.SetFont(ffont);
		}
		public void setFontSize(int size)
		{
			fontSize = size;

			XSSFFont ffont = (XSSFFont)workbook.CreateFont();
			ffont.FontHeight = fontSize * fontSize;
			ffont.SetColor(new XSSFColor(fontRgb));
			ffont.FontName = fontStyle;
			defaultCellStyle.SetFont(ffont);
		}
		public void setFontColor(byte red, byte green, byte blue)
		{
			fontRgb[0] = red;
			fontRgb[1] = green;
			fontRgb[2] = blue;

			XSSFFont ffont = (XSSFFont)workbook.CreateFont();
			ffont.FontHeight = fontSize * fontSize;
			ffont.SetColor(new XSSFColor(fontRgb));
			ffont.FontName = fontStyle;
			defaultCellStyle.SetFont(ffont);
		}
		public void setCellColor(byte red, byte green, byte blue)
		{
			cellRgb[0] = red;
			cellRgb[1] = green;
			cellRgb[2] = blue;
			defaultCellStyle.SetFillForegroundColor(new XSSFColor(cellRgb));
		}
		public void setBorderStyle(short borderBottom, int borderTop, int borderLeft, int borderRight)
		{
			if (borderBottom >= 0 && borderBottom <= 13)
				defaultCellStyle.BorderBottom = (BorderStyle)borderBottom;
			else defaultCellStyle.BorderBottom = BorderStyle.Thin;

			if (borderTop >= 0 && borderTop <= 13)
				defaultCellStyle.BorderTop = (BorderStyle)borderTop;
			else defaultCellStyle.BorderTop = BorderStyle.Thin;

			if (borderLeft >= 0 && borderLeft <= 13)
				defaultCellStyle.BorderLeft = (BorderStyle)borderLeft;
			else defaultCellStyle.BorderLeft = BorderStyle.Thin;

			if (borderRight >= 0 && borderRight <= 13)
				defaultCellStyle.BorderRight = (BorderStyle)borderRight;
			else defaultCellStyle.BorderRight = BorderStyle.Thin;

		}
		public void setCellStyle(int firstRow, int lastRow, int firstColumn, int lastColumn)
		{
			XSSFCellStyle cellStyle = (XSSFCellStyle)workbook.CreateCellStyle();
			cellStyle.CloneStyleFrom(defaultCellStyle);
			for (int currentRow = firstRow; currentRow <= lastRow; currentRow++)
			{
				for (int currentColumn = firstColumn; currentColumn <= lastColumn; currentColumn++)
				{
					try
					{
						XSSFCell cell = (XSSFCell)sheet.GetRow(currentRow).GetCell(currentColumn);
						if (cell != null)
						{
							cell.CellStyle = cellStyle;
						}
						else
						{
							cell = (XSSFCell)sheet.GetRow(currentRow).CreateCell(currentColumn);
							cell.CellStyle = cellStyle;
						}
					}
					catch(Exception ex)
					{
						logger.logException(ex);
						consoleLogger.logError(ex.Message);
					}
				}
			}
			
		}
		public void setAutoFilters(int firstRow, int lastRow, int firstColumn, int lastColumn)
		{
			try
			{
				CellRangeAddress cellRange = new CellRangeAddress(firstRow, lastRow, firstColumn, lastColumn);
				sheet.SetAutoFilter(cellRange);
			}
			catch (Exception ex)
			{
				logger.logException(ex);
				consoleLogger.logError(ex.Message);
			}
		}
		public void setDateFormat(int dateFormat)
		{
			switch(dateFormat)
			{
				case 0:
					defaultDateFormat = "dd/MM/yyyy";
					break;
				case 1:
					defaultDateFormat = "MM/dd/yyyy";
					break;
				case 2:
					defaultDateFormat = "yyyy/dd/MM";
					break;
				case 3:
					defaultDateFormat = "yyyy/MM/dd";
					break;
				default:
					defaultDateFormat = "MM/dd/yyyy";
					break;
			}
		}
        public int WriteArray_To_Excel(int rowAvailableCell, int startingCol, DateTime[,] infoArray)
        {
			XSSFCellStyle cellStyle = (XSSFCellStyle)workbook.CreateCellStyle();
			cellStyle.CloneStyleFrom(defaultCellStyle);
			var newDataFormat = workbook.CreateDataFormat();
			for (int rowCounter = 0; rowCounter <= infoArray.GetUpperBound(0); rowCounter++)
            {
                for (int columnCounter = 0; columnCounter <= infoArray.GetUpperBound(1); columnCounter++)
                {
					if (infoArray[rowCounter, columnCounter] != null)
					{
						DateTime dateTime;
						if (DateTime.TryParse(infoArray[rowCounter, columnCounter].ToString(), out dateTime))
						{
							XSSFCell cell = (XSSFCell)sheet.GetRow(rowAvailableCell + rowCounter).CreateCell(columnCounter + startingCol);
							cell.SetCellValue(dateTime.Date);
							cell.CellStyle = cellStyle;
							cell.CellStyle.DataFormat = newDataFormat.GetFormat("MM/dd/yyyy");

						}
					}
                }
            }
            return startingCol + infoArray.GetUpperBound(1)+1;
        }
		public int WriteArray_To_Excel(int rowAvailableCell, int startingCol, DateTime?[,] infoArray)
		{
			XSSFCellStyle cellStyle = (XSSFCellStyle)workbook.CreateCellStyle();
			cellStyle.CloneStyleFrom(defaultCellStyle);
			var newDataFormat = workbook.CreateDataFormat();
			
			for (int rowCounter = 0; rowCounter <= infoArray.GetUpperBound(0); rowCounter++)
			{
				for (int columnCounter = 0; columnCounter <= infoArray.GetUpperBound(1); columnCounter++)
				{
					if (infoArray[rowCounter, columnCounter] != null)
					{
						DateTime dateTime;
						if (DateTime.TryParse(infoArray[rowCounter, columnCounter].ToString(), out dateTime))
						{
							XSSFCell cell = (XSSFCell)sheet.GetRow(rowAvailableCell + rowCounter).CreateCell(columnCounter + startingCol);
							cell.SetCellValue(dateTime.Date);
							cell.CellStyle = cellStyle;
							cell.CellStyle.DataFormat = newDataFormat.GetFormat("MM/dd/yyyy");
						}
					}
				}
			}
			return startingCol + infoArray.GetUpperBound(1)+1;
		}
		public int WriteArray_To_Excel(int rowAvailableCell, int startingCol,double[,] infoArray)
        {
			XSSFCellStyle cellStyle = (XSSFCellStyle)workbook.CreateCellStyle();
			cellStyle.CloneStyleFrom(defaultCellStyle);
			for (int rowCounter = 0; rowCounter <= infoArray.GetUpperBound(0); rowCounter++)
            {
                for (int columnCounter = 0; columnCounter <= infoArray.GetUpperBound(1); columnCounter++)
                {
                    try
                    {
						XSSFCell cell = (XSSFCell)sheet.GetRow(rowAvailableCell + rowCounter).CreateCell(columnCounter + startingCol);
						cell.SetCellValue(infoArray[rowCounter, columnCounter]);
						cell.CellStyle = cellStyle;
					}
                    catch(Exception ex)
                    {
						logger.logException(ex);
						consoleLogger.logError(ex.Message);
					}
                }
            }
            return startingCol + infoArray.GetUpperBound(1)+1;
        }
        public int WriteArray_To_Excel(int rowAvailableCell, int startingCol, string[,] infoArray)
        {
			XSSFCellStyle cellStyle = (XSSFCellStyle)workbook.CreateCellStyle();
			cellStyle.CloneStyleFrom(defaultCellStyle);
			for (int rowCounter = 0; rowCounter <= infoArray.GetUpperBound(0);rowCounter++)
            {
                for(int columnCounter = 0; columnCounter <= infoArray.GetUpperBound(1); columnCounter++)
                {
                    try
                    {
						XSSFCell cell = (XSSFCell)sheet.GetRow(rowAvailableCell + rowCounter).CreateCell(columnCounter + startingCol);
						cell.SetCellValue(infoArray[rowCounter, columnCounter]);
						cell.CellStyle = cellStyle;
                    }
                    catch(Exception ex)
                    {
						logger.logException(ex);
						consoleLogger.logError(ex.Message);
					}
                }
            }
            return startingCol + infoArray.GetUpperBound(1)+1;
        }
		public int WriteArray_To_ExcelFormulas(int rowAvailableCell, int startingCol, string[,] infoArray)
		{
			XSSFCellStyle cellStyle = (XSSFCellStyle)workbook.CreateCellStyle();
			cellStyle.CloneStyleFrom(defaultCellStyle);
			for (int rowCounter = 0; rowCounter <= infoArray.GetUpperBound(0); rowCounter++)
			{
				for (int columnCounter = 0; columnCounter <= infoArray.GetUpperBound(1); columnCounter++)
				{
					try
					{
						XSSFCell cell = (XSSFCell)sheet.GetRow(rowAvailableCell + rowCounter).CreateCell(columnCounter + startingCol);
						cell.SetCellType(CellType.Formula);
						cell.SetCellFormula(infoArray[rowCounter, columnCounter]);
						cell.CellStyle = cellStyle;
					}
					catch (Exception ex)
					{
						logger.logException(ex);
						consoleLogger.logError(ex.Message);
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
						if(data[rowCounter][dataStartColumn] != null)
						switch(type)
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
								catch(Exception ex)
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
								else if(DateTime.TryParse(data[rowCounter][dataStartColumn].ToString(), /*CultureInfo.CreateSpecificCulture("en-US"), DateTimeStyles.None*/ out dateTime))
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
						consoleLogger.logError(ex.Message);
					}
				}
			}
		}
		
		/*public void WriteHeader_List(int rowAvailableCell, int startingCol, List<string> header)
		{
			XSSFCellStyle cellStyle = (XSSFCellStyle)workbook.CreateCellStyle();
			cellStyle.CloneStyleFrom(defaultCellStyle);

			for (int columnCounter = 0, dataStartColumn = startingReadColumn; dataStartColumn < lastReadColumn; columnCounter++, dataStartColumn++)
			{
				try
				{
					XSSFCell cell = (XSSFCell)sheet.GetRow(rowAvailableCell).CreateCell(columnCounter + startingCol);
					cell.SetCellValue(header[columnCounter]);
					cell.CellStyle = cellStyle;
				}
				catch(Exception ex)
				{
					Console.WriteLine(ex.Message);
				}
			}
		}*/
	}
}
