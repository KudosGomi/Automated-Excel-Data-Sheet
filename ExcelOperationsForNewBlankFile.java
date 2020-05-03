package AESProgram;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import javax.swing.JCheckBox;
import javax.swing.JOptionPane;
import javax.swing.JTextArea;
import javax.swing.JTextField;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelOperationsForNewBlankFile {

	private ArrayList<String> namesForBlankFile;
	private Row rows;
	private Cell cells;
	
	public void getNamesFromTextArea(JTextArea newExcelSheet) {

		String textArea = newExcelSheet.getText();
		namesForBlankFile = new ArrayList<String>();
		int index = 0;
		for(int c = 0; c < textArea.length(); c++)
		{
			char readChar = textArea.charAt(c);
			if(readChar == ',')
			{
				char nextChar = textArea.charAt(++c);
				if(nextChar == ' ')
				{
					namesForBlankFile.add(textArea.substring(index, --c));
					index = c + 2;
				}
				else
				{
					namesForBlankFile.add(textArea.substring(index, --c));
					index = c + 1;
				}
			}
			if(readChar == ' ' || readChar == '\n')
			{
				namesForBlankFile.add(textArea.substring(index, c));
				index = c + 1;
			}
		}
		
		namesForBlankFile.add(textArea.substring(index, textArea.length()));
	}
	
	public void createNewBlankExcelWithNames(JCheckBox sheet, JCheckBox file, String sheetName, String ccrFile, JTextField saveFileAs) {
		
		if(sheet.isEnabled())
		{
			FileInputStream fileInput;
			XSSFWorkbook workbook;
			boolean sheetFound = false;
			
			try {
				fileInput = new FileInputStream(ccrFile);
				workbook = new XSSFWorkbook(); //fileInput
				XSSFSheet sheetInput;
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
					JOptionPane.showMessageDialog(null, "Sheet already exists!");
				}
				else
				{
					String safeName = WorkbookUtil.createSafeSheetName(sheetName);
					sheetInput = workbook.createSheet(safeName);
					
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
						rows = sheetInput.createRow(theRows);
						for(int theCells = 0; theCells < 20; theCells++)
						{
							cells = sheetInput.getRow(theRows).createCell(theCells);
							cells.setCellStyle(borderCenter);
						}
					}
					for(int b = 0; b < 19; b++)
					{
						sheetInput.getRow(1).getCell(b + 1).setCellStyle(black);
						sheetInput.getRow(8).getCell(b + 1).setCellStyle(black);
						sheetInput.getRow(15).getCell(b + 1).setCellStyle(black);
						sheetInput.getRow(22).getCell(b + 1).setCellStyle(black);
						sheetInput.getRow(29).getCell(b + 1).setCellStyle(black);
						sheetInput.getRow(36).getCell(b + 1).setCellStyle(black);
						sheetInput.getRow(7).getCell(b + 1).setCellStyle(aquaBoldCenter);
						sheetInput.getRow(14).getCell(b + 1).setCellStyle(aquaBoldCenter);
						sheetInput.getRow(21).getCell(b + 1).setCellStyle(aquaBoldCenter);
						sheetInput.getRow(28).getCell(b + 1).setCellStyle(aquaBoldCenter);
						sheetInput.getRow(35).getCell(b + 1).setCellStyle(aquaBoldCenter);
						sheetInput.getRow(42).getCell(b + 1).setCellStyle(aquaBoldCenter);
					}
					Cell dateRange = sheetInput.getRow(0).getCell(1);
					dateRange.setCellValue("Date Range");
					dateRange.setCellStyle(fontStyleCenterBold);
					for(int i = 0; i < namesForBlankFile.size(); i++)
					{
						Cell headerCells = sheetInput.getRow(0).getCell(i + 2);
						headerCells.setCellValue(namesForBlankFile.get(i) + "");
						headerCells.setCellStyle(fontStyleCenterBold);
					}
					Cell lastCellCen = sheetInput.getRow(0).getCell(19);
					lastCellCen.setCellValue("CENTER");
					lastCellCen.setCellStyle(fontStyleCenterBold);
					for(int i = 0; i < 43; i++)
					{
						sheetInput.getRow(i).getCell(0).setCellValue(fixedTotalOf[i]);
						Cell cellsTotalOf = sheetInput.getRow(i).getCell(0);
						cellsTotalOf.setCellValue(fixedTotalOf[i]);
						cellsTotalOf.setCellStyle(fontStyleCenterBold);
					}
					
					sheetInput.autoSizeColumn(0);
					
					for(int wk = 7; wk < 43; wk = wk + 7)
					{
						if(wk != 21)
						{
							Cell wkAVG = sheetInput.getRow(wk).getCell(1);
							wkAVG.setCellValue("WEEK AVG");
						}
					}
					
					Cell wkSUM = sheetInput.getRow(21).getCell(1);
					wkSUM.setCellValue("WEEK SUM");

					char letter = 'C';
					int firstC = 3, secC = 7, row = 7, col = 2;
					while(row < 43)
					{
						Cell wkAVGs = sheetInput.getRow(row).getCell(col);
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
						Cell wkAVGs = sheetInput.getRow(21).getCell(x);
						wkAVGs.setCellFormula("SUM(" + l + "17:" + l + "21)");
						wkAVGs.setCellStyle(aquaBoldCenter);
						l++;
					}

					int eachRow = 2, k = 3;
					while(eachRow < 43)
					{
						if(eachRow != 8 || eachRow != 15 || eachRow != 29 || eachRow != 36 || k != 9 || k != 16 || k != 17 || k != 18 || k != 19 || k != 20 || k != 21 || k != 22 || k != 23 || k != 30 || k != 37)
						{
							Cell cen = sheetInput.getRow(eachRow).getCell(19);
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
						Cell cent = sheetInput.getRow(x).getCell(19);
						cent.setCellFormula("SUM(C" + sum + ":S" + sum + ")");
						cent.setCellStyle(centerColBorder);
						sum++;
					}

					autoSizeColumns(workbook);

					fileInput.close();
				}
				FileOutputStream outFile;
				if(saveFileAs.getText().endsWith(".xlsx"))
				{
					outFile = new FileOutputStream(new File(saveFileAs.getText()));
				}
				else 
				{
					outFile = new FileOutputStream(new File(saveFileAs.getText()  + ".xlsx"));
				}
				workbook.write(outFile);
				outFile.close();
				System.out.println("Excel file was written successfully...");
				
			}
			catch (FileNotFoundException e) {
				System.out.println("File not found!");
			}
			catch (IOException e) 
			{
				System.out.println("IO Exception!");
				e.printStackTrace();
			}
		}
		else if(file.isEnabled())
		{
			XSSFWorkbook workbook;
			boolean sheetFound = false;
			
			try {
				workbook = new XSSFWorkbook();
				XSSFSheet sheetInput;
			
				String safeName = WorkbookUtil.createSafeSheetName(sheetName);
				sheetInput = workbook.createSheet(safeName);

				String[] fixedTotalOf = {"", "Total Available", "HOURS", "", "", "", "", "", "Total AUX", "HOURS", "", "", "", "", "", "Total Calls", "# OF CALLS", "", "", "", "", "", "Call Duration", "", "", "", "", "", "", "Total Hours Logged", "PHONES", "", "", "", "", "", "Total Hours Worked", "SET", "", "", "", "", ""}; //43
				CellStyle black = workbook.createCellStyle();
				black.setFillForegroundColor(HSSFColor.BLACK.index);
				black.setFillPattern(CellStyle.SOLID_FOREGROUND);
				Font fixedTextFont = workbook.createFont();
				fixedTextFont.setFontHeightInPoints((short) 11);
				fixedTextFont.setFontName("TIMESNew Roman");
//				fixedTextFont.setBold(true);
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
					rows = sheetInput.createRow(theRows);
					for(int theCells = 0; theCells < 20; theCells++)
					{
						cells = sheetInput.getRow(theRows).createCell(theCells);
						cells.setCellStyle(borderCenter);
					}
				}
				for(int b = 0; b < 19; b++)
				{
					sheetInput.getRow(1).getCell(b + 1).setCellStyle(black);
					sheetInput.getRow(8).getCell(b + 1).setCellStyle(black);
					sheetInput.getRow(15).getCell(b + 1).setCellStyle(black);
					sheetInput.getRow(22).getCell(b + 1).setCellStyle(black);
					sheetInput.getRow(29).getCell(b + 1).setCellStyle(black);
					sheetInput.getRow(36).getCell(b + 1).setCellStyle(black);
					sheetInput.getRow(7).getCell(b + 1).setCellStyle(aquaBoldCenter);
					sheetInput.getRow(14).getCell(b + 1).setCellStyle(aquaBoldCenter);
					sheetInput.getRow(21).getCell(b + 1).setCellStyle(aquaBoldCenter);
					sheetInput.getRow(28).getCell(b + 1).setCellStyle(aquaBoldCenter);
					sheetInput.getRow(35).getCell(b + 1).setCellStyle(aquaBoldCenter);
					sheetInput.getRow(42).getCell(b + 1).setCellStyle(aquaBoldCenter);
				}
				Cell dateRange = sheetInput.getRow(0).getCell(1);
				dateRange.setCellValue("Date Range");
				dateRange.setCellStyle(fontStyleCenterBold);
				for(int i = 0; i < namesForBlankFile.size(); i++)
				{
					Cell headerCells = sheetInput.getRow(0).getCell(i + 2);
					headerCells.setCellValue(namesForBlankFile.get(i) + "");
					headerCells.setCellStyle(fontStyleCenterBold);
				}
				Cell lastCellCen = sheetInput.getRow(0).getCell(19);
				lastCellCen.setCellValue("CENTER");
				lastCellCen.setCellStyle(fontStyleCenterBold);
				for(int i = 0; i < 43; i++)
				{
					sheetInput.getRow(i).getCell(0).setCellValue(fixedTotalOf[i]);
					Cell cellsTotalOf = sheetInput.getRow(i).getCell(0);
					cellsTotalOf.setCellValue(fixedTotalOf[i]);
					cellsTotalOf.setCellStyle(fontStyleCenterBold);
				}

				sheetInput.autoSizeColumn(0);

				for(int wk = 7; wk < 43; wk = wk + 7)
				{
					if(wk != 21)
					{
						Cell wkAVG = sheetInput.getRow(wk).getCell(1);
						wkAVG.setCellValue("WEEK AVG");
					}
				}

				Cell wkSUM = sheetInput.getRow(21).getCell(1);
				wkSUM.setCellValue("WEEK SUM");

				char letter = 'C';
				int firstC = 3, secC = 7, row = 7, col = 2;
				while(row < 43)
				{
					Cell wkAVGs = sheetInput.getRow(row).getCell(col);
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
					Cell wkAVGs = sheetInput.getRow(21).getCell(x);
					wkAVGs.setCellFormula("SUM(" + l + "17:" + l + "21)");
					wkAVGs.setCellStyle(aquaBoldCenter);
					l++;
				}

				int eachRow = 2, k = 3;
				while(eachRow < 43)
				{
					if(eachRow != 8 || eachRow != 15 || eachRow != 29 || eachRow != 36 || k != 9 || k != 16 || k != 17 || k != 18 || k != 19 || k != 20 || k != 21 || k != 22 || k != 23 || k != 30 || k != 37)
					{
						Cell cen = sheetInput.getRow(eachRow).getCell(19);
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
					Cell cent = sheetInput.getRow(x).getCell(19);
					cent.setCellFormula("SUM(C" + sum + ":S" + sum + ")");
					cent.setCellStyle(centerColBorder);
					sum++;
				}

				autoSizeColumns(workbook);

				FileOutputStream outFile;
				if(saveFileAs.getText().endsWith(".xlsx"))
				{
					outFile = new FileOutputStream(new File(saveFileAs.getText()));
				}
				else 
				{
					outFile = new FileOutputStream(new File(saveFileAs.getText()  + ".xlsx"));
				}
				workbook.write(outFile);
				outFile.close();
				System.out.println("Excel file was written successfully...");
			}
			catch (FileNotFoundException e) {
				System.out.println("File not found!");
			}
			catch (IOException e) 
			{
				System.out.println("IO Exception!");
				e.printStackTrace();
			}
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
}
