package AESProgram;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.hssf.model.Workbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelOperations {

	// "Meditate" - 12/29/16
	public ArrayList<String> names = new ArrayList<String>();
	private ArrayList<String> namesForBlankFile;
	public Double[] metadata;
	Row rows;
	Cell cells;

	public void getNamesFromCCRFile(String importedCCRFile) {

		try 
		{
			FileInputStream file = new FileInputStream(new File(importedCCRFile));
			XSSFWorkbook workbook = new XSSFWorkbook(); // File
			XSSFSheet sheet = workbook.getSheetAt(workbook.getNumberOfSheets() - 1);

			Row getNames = sheet.getRow(0);
			for(int n = 2; n < getNames.getLastCellNum() - 1; n++)
			{
				Cell eachName = getNames.getCell(n);
				names.add(eachName + "");
			}

			file.close();
		} 
		catch (Exception e) 
		{
			System.out.println("CCR file was not detected!\nOr names were not found.");
			e.printStackTrace();
		}	
	}

	public boolean readAvayaFile(String importedAvayaFile) {

		metadata = new Double[names.size() * 6];

		try 
		{
			FileInputStream file = new FileInputStream(new File(importedAvayaFile));
			XSSFWorkbook workbook = new XSSFWorkbook(); // File
			XSSFSheet sheet = workbook.getSheetAt(0);

			int rowNum = sheet.getPhysicalNumberOfRows();
			int colNum = sheet.getRow(0).getLastCellNum();

			int d = 0, k = 0;
			for(int r = 1; r < rowNum; r++) 
			{
				Row row = sheet.getRow(r);
				Cell cell = row.getCell(0);
				if(cell == null)
				{
					break;
				}
				else if(cell.getStringCellValue().startsWith(names.get(k) + "") && cell.getStringCellValue().endsWith("_total"))
				{
					if(k == names.size())
					{
						break;
					}
					for(int c = 2; c < colNum; c++)
					{
						Cell cellValues = row.getCell(c);
						metadata[d] = cellValues.getNumericCellValue();
						d++;
					}
					k++;
				}
			}
			file.close();
			return true;
		} 
		catch (Exception e) 
		{
			System.out.println("Avaya file was not detected!");
			return false;
		}	
	}

	public void createExcelFile(String sheetName, String selectedRow, String weekInput, String exportFileName, String ccrFile) {

		Double[] data = metadata;
		ArrayList<String> theNames = names;

		FileInputStream file;
		XSSFWorkbook workbook;
		boolean sheetFound = false;

		try {
			file = new FileInputStream(ccrFile);
			workbook = new XSSFWorkbook(); // File
			XSSFSheet sheet;
			int s;
			for(s = 0; s < workbook.getNumberOfSheets(); s++)
			{
				if(workbook.getSheetAt(s).getSheetName().contentEquals(sheetName))
				{
					sheetFound = true;
					break;
				}	
			}
			if(sheetFound == true)
			{
				System.out.println("Sheet found");
				Font fixedTextFont = workbook.createFont();
				fixedTextFont.setFontHeightInPoints((short) 11);
				fixedTextFont.setFontName("TIMESNew Roman");
				((XSSFFont) fixedTextFont).setBold(true);
				CellStyle aquaBoldCenter = workbook.createCellStyle();
				aquaBoldCenter.setAlignment(CellStyle.ALIGN_CENTER);
				aquaBoldCenter.setFont(fixedTextFont);
				aquaBoldCenter.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
				aquaBoldCenter.setFillPattern(CellStyle.SOLID_FOREGROUND);
				aquaBoldCenter.setBorderTop(CellStyle.BORDER_MEDIUM);
				aquaBoldCenter.setBorderBottom(CellStyle.BORDER_THIN);
				aquaBoldCenter.setBorderLeft(CellStyle.BORDER_THIN);
				aquaBoldCenter.setBorderRight(CellStyle.BORDER_THIN);
				aquaBoldCenter.setDataFormat(workbook.createDataFormat().getFormat("0.00"));
				CellStyle centerColBorder = workbook.createCellStyle();
				centerColBorder.setFont(fixedTextFont);
				centerColBorder.setAlignment(CellStyle.ALIGN_CENTER);
				centerColBorder.setBorderTop(CellStyle.BORDER_MEDIUM);
				centerColBorder.setBorderBottom(CellStyle.BORDER_MEDIUM);
				centerColBorder.setBorderLeft(CellStyle.BORDER_MEDIUM);
				centerColBorder.setBorderRight(CellStyle.BORDER_MEDIUM);
				centerColBorder.setDataFormat(workbook.createDataFormat().getFormat("0.00"));
				
				sheet = workbook.getSheetAt(s);

				int row = Integer.parseInt(selectedRow) - 1, theCell = 2, d = 0;
				while(d < data.length)
				{
					if(row > 42)
					{
						row = Integer.parseInt(selectedRow) - 1;
						theCell++;
					}
					sheet.getRow(row).getCell(theCell).setCellValue(data[d]);
					row = row + 7;
					d++;
				}
				
				/* Set the date/week */
				int rowDate = Integer.parseInt(selectedRow) - 1;
				for(int wk = 7; wk < 43; wk = wk + 7)
				{
					sheet.getRow(rowDate).getCell(1).setCellValue(weekInput);
					rowDate = rowDate + 7;
				}

				/* Set the averages and sums */
				char letter = 'C';
				int firstC = 3, secC =7, roww = 7, col = 2;
				while(roww < 43)
				{
					Cell wkAVGs = sheet.getRow(roww).getCell(col);
					wkAVGs.setCellFormula("AVERAGE(" + letter + firstC + ":" + letter + secC + ")");
					wkAVGs.setCellStyle(aquaBoldCenter);
					letter++;
					col++;
					if(col > 18)
					{
						col = 2;
						letter = 'C';
						roww = roww + 7;
						firstC = firstC + 7;
						secC = secC + 7;
					}
					if(roww == 21)
					{
						roww = 28;
					}
					if(firstC == 17)
					{
						firstC = 24;
					}
					if(secC == 21)
					{
						secC = 28;
					}
				}

				char l = 'C';
				for(int x = 2; x < 19; x++)
				{
					Cell wkAVGs = sheet.getRow(21).getCell(x);
					wkAVGs.setCellFormula("SUM(" + l + "17:" + l + "21)");
					wkAVGs.setCellStyle(aquaBoldCenter);
					l++;
				}

				int eachRow = 2, k = 3;
				while(eachRow < 43)
				{
					if(eachRow != 8 || eachRow != 15 || eachRow != 29 || eachRow != 36 || k != 9 || k != 16 || k != 17 || k != 18 || k != 19 || k != 20 || k != 21 || k != 22 || k != 23 || k != 30 || k != 37)
					{
						Cell center = sheet.getRow(eachRow).getCell(19);
						center.setCellFormula("AVERAGE(C" + k + ":S" + k + ")");
						center.setCellStyle(centerColBorder);
					}
					if(eachRow == 14)
					{
						eachRow = 22;
						k = 23;
					}
					eachRow++;
					k++;
				}

				int sum = 17;
				for(int x = 16; x < 22; x++)
				{
					Cell center = sheet.getRow(x).getCell(19);
					center.setCellFormula("SUM(C" + sum + ":S" + sum + ")");
					center.setCellStyle(centerColBorder);
					sum++;
				}
			}
			else
			{
				String safeName = WorkbookUtil.createSafeSheetName(sheetName);
				sheet = workbook.createSheet(safeName);

				String[] fixedTotalOf = {"", "Total Available", "HOURS", "", "", "", "", "", "Total AUX", "HOURS", "", "", "", "", "", "Total Calls", "# OF CALLS", "", "", "", "", "", "Call Duration", "", "", "", "", "", "", "Total Hours Logged", "PHONES", "", "", "", "", "", "Total Hours Worked", "SET", "", "", "", "", ""}; //43
				CellStyle black = workbook.createCellStyle();
				black.setFillForegroundColor(HSSFColor.BLACK.index);
				black.setFillPattern(CellStyle.SOLID_FOREGROUND);
				Font fixedTextFont = workbook.createFont();
				fixedTextFont.setFontHeightInPoints((short) 11);
				fixedTextFont.setFontName("TIMESNew Roman");
				((XSSFFont) fixedTextFont).setBold(true);
				CellStyle fontStyleCenterBold = workbook.createCellStyle();
				fontStyleCenterBold.setFont(fixedTextFont);
				fontStyleCenterBold.setAlignment(CellStyle.ALIGN_CENTER);
				fontStyleCenterBold.setBorderTop(CellStyle.BORDER_THIN);
				fontStyleCenterBold.setBorderBottom(CellStyle.BORDER_THIN);
				fontStyleCenterBold.setBorderLeft(CellStyle.BORDER_THIN);
				fontStyleCenterBold.setBorderRight(CellStyle.BORDER_THIN);
				CellStyle borderCenter = workbook.createCellStyle();
				borderCenter.setBorderTop(CellStyle.BORDER_THIN);
				borderCenter.setBorderBottom(CellStyle.BORDER_THIN);
				borderCenter.setBorderLeft(CellStyle.BORDER_THIN);
				borderCenter.setBorderRight(CellStyle.BORDER_THIN);
				borderCenter.setAlignment(CellStyle.ALIGN_CENTER);
				borderCenter.setDataFormat(workbook.createDataFormat().getFormat("0.00"));
				CellStyle aquaBoldCenter = workbook.createCellStyle();
				aquaBoldCenter.setAlignment(CellStyle.ALIGN_CENTER);
				aquaBoldCenter.setFont(fixedTextFont);
				aquaBoldCenter.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
				aquaBoldCenter.setFillPattern(CellStyle.SOLID_FOREGROUND);
				aquaBoldCenter.setBorderTop(CellStyle.BORDER_MEDIUM);
				aquaBoldCenter.setBorderBottom(CellStyle.BORDER_THIN);
				aquaBoldCenter.setBorderLeft(CellStyle.BORDER_THIN);
				aquaBoldCenter.setBorderRight(CellStyle.BORDER_THIN);
				aquaBoldCenter.setDataFormat(workbook.createDataFormat().getFormat("0.00"));
				CellStyle centerColBorder = workbook.createCellStyle();
				centerColBorder.setFont(fixedTextFont);
				centerColBorder.setAlignment(CellStyle.ALIGN_CENTER);
				centerColBorder.setBorderTop(CellStyle.BORDER_MEDIUM);
				centerColBorder.setBorderBottom(CellStyle.BORDER_MEDIUM);
				centerColBorder.setBorderLeft(CellStyle.BORDER_MEDIUM);
				centerColBorder.setBorderRight(CellStyle.BORDER_MEDIUM);
				centerColBorder.setDataFormat(workbook.createDataFormat().getFormat("0.00"));

				for(int theRows = 0; theRows < 43; theRows++)
				{
					rows = sheet.createRow(theRows);
					for(int theCells = 0; theCells < 20; theCells++)
					{
						cells = sheet.getRow(theRows).createCell(theCells);
						cells.setCellStyle(borderCenter);
					}
				}
				for(int b = 0; b < 19; b++)
				{
					sheet.getRow(1).getCell(b + 1).setCellStyle(black);
					sheet.getRow(8).getCell(b + 1).setCellStyle(black);
					sheet.getRow(15).getCell(b + 1).setCellStyle(black);
					sheet.getRow(22).getCell(b + 1).setCellStyle(black);
					sheet.getRow(29).getCell(b + 1).setCellStyle(black);
					sheet.getRow(36).getCell(b + 1).setCellStyle(black);
					sheet.getRow(7).getCell(b + 1).setCellStyle(aquaBoldCenter);
					sheet.getRow(14).getCell(b + 1).setCellStyle(aquaBoldCenter);
					sheet.getRow(21).getCell(b + 1).setCellStyle(aquaBoldCenter);
					sheet.getRow(28).getCell(b + 1).setCellStyle(aquaBoldCenter);
					sheet.getRow(35).getCell(b + 1).setCellStyle(aquaBoldCenter);
					sheet.getRow(42).getCell(b + 1).setCellStyle(aquaBoldCenter);
				}
				Cell dateRange = sheet.getRow(0).getCell(1);
				dateRange.setCellValue("Date Range");
				dateRange.setCellStyle(fontStyleCenterBold);
				for(int i = 0; i < theNames.size(); i++)
				{
					Cell headerCells = sheet.getRow(0).getCell(i + 2);
					headerCells.setCellValue(theNames.get(i) + "");
					headerCells.setCellStyle(fontStyleCenterBold);
				}
				Cell lastCellCen = sheet.getRow(0).getCell(19);
				lastCellCen.setCellValue("CENTER");
				lastCellCen.setCellStyle(fontStyleCenterBold);
				for(int i = 0; i < 43; i++)
				{
					sheet.getRow(i).getCell(0).setCellValue(fixedTotalOf[i]);
					Cell cellsTotalOf = sheet.getRow(i).getCell(0);
					cellsTotalOf.setCellValue(fixedTotalOf[i]);
					cellsTotalOf.setCellStyle(fontStyleCenterBold);
				}
				
				sheet.autoSizeColumn(0);
				
				int rowDate = Integer.parseInt(selectedRow) - 1;
				for(int wk = 7; wk < 43; wk = wk + 7)
				{
					if(wk != 21)
					{
						Cell wkAVG = sheet.getRow(wk).getCell(1);
						wkAVG.setCellValue("WEEK AVG");
					}
					sheet.getRow(rowDate).getCell(1).setCellValue(weekInput);
					rowDate = rowDate + 7;
				}
				
				Cell wkSUM = sheet.getRow(21).getCell(1);
				wkSUM.setCellValue("WEEK SUM");
				
				int i = Integer.parseInt(selectedRow) - 1, c = 2, d = 0;
				while(d < data.length)
				{
					if(i > 42)
					{
						i = Integer.parseInt(selectedRow) - 1;
						c++;
					}
					sheet.getRow(i).getCell(c).setCellValue(data[d]);
					i = i + 7;
					d++;
				}

				char letter = 'C';
				int firstC = 3, secC = 7, row = 7, col = 2;
				while(row < 43)
				{
					Cell wkAVGs = sheet.getRow(row).getCell(col);
					wkAVGs.setCellFormula("AVERAGE(" + letter + firstC + ":" + letter + secC + ")");
					wkAVGs.setCellStyle(aquaBoldCenter);
					letter++;
					col++;
					if(col > 18)
					{
						col = 2;
						letter = 'C';
						row = row + 7;
						firstC = firstC + 7;
						secC = secC + 7;
					}
					if(row == 21)
					{
						row = 28;
					}
					if(firstC == 17)
					{
						firstC = 24;
					}
					if(secC == 21)
					{
						secC = 28;
					}
				}

				char l = 'C';
				for(int x = 2; x < 19; x++)
				{
					Cell wkAVGs = sheet.getRow(21).getCell(x);
					wkAVGs.setCellFormula("SUM(" + l + "17:" + l + "21)");
					wkAVGs.setCellStyle(aquaBoldCenter);
					l++;
				}

				int eachRow = 2, k = 3;
				while(eachRow < 43)
				{
					if(eachRow != 8 || eachRow != 15 || eachRow != 29 || eachRow != 36 || k != 9 || k != 16 || k != 17 || k != 18 || k != 19 || k != 20 || k != 21 || k != 22 || k != 23 || k != 30 || k != 37)
					{
						Cell cen = sheet.getRow(eachRow).getCell(19);
						cen.setCellFormula("AVERAGE(C" + k + ":S" + k + ")");
						cen.setCellStyle(centerColBorder);
					}
					if(eachRow == 14)
					{
						eachRow = 22;
						k = 23;
					}
					eachRow++;
					k++;
				}

				int sum = 17;
				for(int x = 16; x < 22; x++)
				{
					Cell cent = sheet.getRow(x).getCell(19);
					cent.setCellFormula("SUM(C" + sum + ":S" + sum + ")");
					cent.setCellStyle(centerColBorder);
					sum++;
				}

				autoSizeColumns(workbook);

				file.close();
			}
			FileOutputStream outFile;
			if(exportFileName.endsWith(".xlsx"))
			{
				outFile = new FileOutputStream(new File(exportFileName));
			}
			else 
			{
				outFile = new FileOutputStream(new File(exportFileName  + ".xlsx"));
			}
			workbook.write(outFile);
			outFile.close();
			System.out.println("Excel file was written successfully...");
		} 
		catch (FileNotFoundException e) 
		{
			System.out.println("File not found!");
			e.printStackTrace();
		} 
		catch (IOException e) 
		{
			System.out.println("IO Exception!");
			e.printStackTrace();
		}
	}

	public void autoSizeColumns(XSSFWorkbook workbook) {

		int numberOfSheets = workbook.getNumberOfSheets();
		for (int i = 0; i < numberOfSheets; i++) 
		{
			XSSFSheet sheet = workbook.getSheetAt(i);
			if (sheet.getPhysicalNumberOfRows() > 0) 
			{
				Row row = sheet.getRow(0);
				Iterator<Cell> cellIterator = row.cellIterator();
				while (cellIterator.hasNext()) 
				{
					Cell cell = cellIterator.next();
					int columnIndex = cell.getColumnIndex();
					sheet.autoSizeColumn(columnIndex);
				}
			}
		}
	}
	
	public void nameChanges(ArrayList<rowAndColObj> theNewNames) {
		
		if(theNewNames == null)
		{
			return;
		}
		
		for(int i = 0; i < theNewNames.size(); i++)
		{
			
			if(theNewNames.get(i).getName() == "" || theNewNames.get(i).getName() == " " || theNewNames.get(i).getName().isEmpty())
			{
				continue;
			}
			else 
			{
				names.set(Integer.parseInt(theNewNames.get(i).getCol()) - 1, theNewNames.get(i).getName());
			}
		}
		
	}

}