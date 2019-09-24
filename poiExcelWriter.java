package com.clr.excel;

import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFCreationHelper;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ExcelWriter {
	private int curRow;
	private HSSFWorkbook curWB;
	private HSSFSheet curSheet;

	
	public boolean createExcelWorkbook() {
		try {
			this.curWB = new HSSFWorkbook();
			return true;
		} catch(Exception ex) {
			ex.printStackTrace();
			return(false);	
		}
	}
	
	public boolean writeExcelHeader(String[] rowValues, String sheetName) {
		try {
			this.curRow = 0;
			this.curSheet = this.curWB.createSheet(sheetName);
			
			HSSFCellStyle headerStyle = curWB.createCellStyle();
			HSSFFont headerFont = curWB.createFont();
			headerFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD); 
			headerStyle.setFont(headerFont);
			
			HSSFRow row = this.curSheet.createRow(this.curRow);

			for(int i = 0; i < rowValues.length; i++) {
				HSSFCell hCell = row.createCell(i);
				hCell.setCellValue(rowValues[i]);
				hCell.setCellStyle(headerStyle);
			}

			this.curRow++;
			
			//This will write the header row
			// return(this.writeToExcel(rowValues));
			return(true);
		} catch(Exception ex) {
			ex.printStackTrace();
			return(false);	
		}
		
	}

	public boolean writeToExcel(String[] rowValues) {

		try {
			HSSFRow row = this.curSheet.createRow(this.curRow);
			
			for(int i = 0; i < rowValues.length; i++) {
				if ( rowValues[i].matches("^2703+[0-9]{11,15}$") ) {
					row.createCell(i).setCellValue(new HSSFRichTextString(rowValues[i]));
				}
				else if ( rowValues[i].matches("^0$") ) {
					double amount = Double.parseDouble( rowValues[i] );
					row.createCell(i).setCellValue( amount );	
				}
				else if ( rowValues[i].matches("^0?\\.[0-9]+$" ) ) {
					double amount = Double.parseDouble( rowValues[i] );
					row.createCell(i).setCellValue( amount );
				}
				else if ( rowValues[i].matches("^-?[1-9]([0-9]+)?(\\.[0-9]+)?$") ) {
					double amount = Double.parseDouble( rowValues[i] );
					row.createCell(i).setCellValue( amount );
				} else {
					row.createCell(i).setCellValue(new HSSFRichTextString(rowValues[i]));
				}

			}

			this.curRow++;

		} catch(Exception ex) {
			ex.printStackTrace();
			return(false);
		}

		return(true);
	}
	
	public boolean writeToExcelV2(String[] rowValues) {

		try {
			HSSFCellStyle cellStyleDate = curWB.createCellStyle();
			HSSFCreationHelper createHelper = curWB.getCreationHelper();
			short dateFormat = createHelper.createDataFormat().getFormat("yyyy-MM-dd");
			cellStyleDate.setDataFormat(dateFormat);
				
			HSSFRow row = this.curSheet.createRow(this.curRow);
			HSSFCell cell;
			Date date1;
			
			SimpleDateFormat sdf1 = new SimpleDateFormat( "yyyy-MM-dd" );
			SimpleDateFormat sdf2 = new SimpleDateFormat( "dd-MMM-yyyy" );
			
			for(int i = 0; i < rowValues.length; i++) {
				date1 = null;
				if ( rowValues[i].matches("^0$") ) {
					double amount = Double.parseDouble( rowValues[i] );
					row.createCell(i).setCellValue( amount );	
				}
				else if ( rowValues[i].matches("^0?\\.[0-9]+$" ) ) {
					double amount = Double.parseDouble( rowValues[i] );
					row.createCell(i).setCellValue( amount );
				}
				else if ( rowValues[i].matches("^-?[1-9]([0-9]+)?(\\.[0-9]+)?$") ) {
					double amount = Double.parseDouble( rowValues[i] );
					row.createCell(i).setCellValue( amount );
				}  else if ( rowValues[i].matches("^\\d{4}\\-(0?[1-9]|1[012])\\-(0?[1-9]|[12][0-9]|3[01])$" ) ) {
					date1 = (Date) sdf1.parse( rowValues[i] );
					cell = row.createCell(i);
					cell.setCellValue( date1 );
					cell.setCellStyle( cellStyleDate );
				} else if ( rowValues[i].matches("^(([1-9])|([0][1-9])|([1-2][0-9])|([3][0-1]))\\-(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\\-\\d{4}$" ) ) {
					date1 = (Date) sdf2.parse( rowValues[i] );
					cell = row.createCell(i);
					cell.setCellValue( date1 );
					cell.setCellStyle( cellStyleDate );					
				} else {
					row.createCell(i).setCellValue(new HSSFRichTextString(rowValues[i]));
				}

			}

			this.curRow++;

		} catch(Exception ex) {
			ex.printStackTrace();
			return(false);
		}

		return(true);
	}


	public boolean writeExcelFile(String fileName) {
		try {
			//Auto fit content
			for(int i = 0; i < 20; i++) {
				this.curSheet.autoSizeColumn((short)i);
			}

			// Write the output to a file
			FileOutputStream fileOut = new FileOutputStream(fileName);
			this.curWB.write(fileOut);
			fileOut.close();
		} catch(Exception ex) {
			ex.printStackTrace();
			return(false);
		}
		
		return(true);

	}


}
