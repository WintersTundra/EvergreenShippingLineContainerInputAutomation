
import java.io.*;
import java.util.*;
import org.apache.pdfbox.text.*;
import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.pdfbox.pdmodel.*;

public class PDFToExcelPortTS {
	public static void main(String[] args) throws InterruptedException, IOException {
		PDFTextStripper pdfstrip = new PDFTextStripper();
		pdfstrip.setSortByPosition(true);
		pdfstrip.setStartPage(1);
		pdfstrip.setEndPage(1);
		PDDocument pdf = PDDocument.load(new File("input file name"));
		String[] pdfText = pdfstrip.getText(pdf).trim().split("\n");
		String[] shipmentInformation = new String[2];
		/*
		 * shipmentInformation[0] = Booking Number; shipmentInformation[1]=SOXXXX
		 */
		int containerCount = 0;
		for (int i = 1; i < 6; i++) {
			String temp = pdfText[i];
			if (temp.contains("Booking Number")) {
				int index = temp.indexOf("Number");
				shipmentInformation[0] = "BL RIDER " + temp.substring(7 + index);
			}
			if (temp.contains("Customer PO#")) {
				int index = temp.indexOf("PO#");
				shipmentInformation[1] = "SO" + temp.substring(index + 4, index + 9);
			}
			if (temp.contains("Containers")) {
				int index = temp.indexOf("Containers");
				containerCount = Integer.parseInt(temp.trim().substring(index + 11));
			}
		}

		String[][] containerValues = new String[containerCount + 2][4];
		containerValues[0][0] = "CONT";
		containerValues[0][1] = "SEAL";
		containerValues[0][2] = "ROLLS";
		containerValues[0][3] = "CARGO/KGS";
		containerValues[containerCount + 1][0] = "TOTAL";

		for (int i = 6; i < pdfText.length; i++) {
			if (pdfText[i].contains("Load Container")) {
				int index = 1;
				i++;
				while (!pdfText[i].contains("Totals")) {
					String[] temp = pdfText[i].split(" ");
					containerValues[index][0] = temp[1];
					containerValues[index][1] = temp[5];
					containerValues[index][2] = temp[10];
					containerValues[index][3] = temp[12];
					i++;
					index++;
				}
			}
			if (pdfText[i].contains("Totals")) {
				String[] temp = pdfText[i].split(" ");
				containerValues[containerCount + 1][1] = "";
				containerValues[containerCount + 1][2] = temp[8];
				containerValues[containerCount + 1][3] = temp[10];
				break;
			}
		}
		for (int i = 0; i < containerCount + 2; i++) {
			for (int j = 0; j < 4; j++) {
				System.out.print(containerValues[i][j] + "\t");
			}
			System.out.println();
		}
		pdf.close();

		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("BOOKING INFORMATION");
		DataFormatter dataFormatter = new DataFormatter();
		XSSFCellStyle style = (XSSFCellStyle) workbook.createCellStyle();
		style.setBorderTop(BorderStyle.MEDIUM);
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setBorderLeft(BorderStyle.MEDIUM);
		style.setBorderRight(BorderStyle.MEDIUM);

		Row shipmentInformationRow = sheet.createRow(3);
		for (int i = 1; i < 5; i++) {
			shipmentInformationRow.createCell(i).setCellStyle(style);
		}
		shipmentInformationRow.getCell(1).setCellValue(shipmentInformation[0]);
		shipmentInformationRow.getCell(4).setCellValue(shipmentInformation[1]);

		for (int i = 5; i < containerCount + 6; i++) {
			Row row = sheet.createRow(i);
			if (i != 5) {
				row.createCell(0).setCellValue(i - 5);
				row.getCell(0).setCellStyle(style);
			}
			for (int j = 0; j < containerValues[i - 5].length; j++) {
				row.createCell(j + 1).setCellValue(containerValues[i - 5][j]);
				row.getCell(j + 1).setCellStyle(style);
			}
		}

		int lastLine = containerCount + 6;
		Row totalRow = sheet.createRow(lastLine);
		for (int i = 0; i < containerValues[lastLine - 5].length; i++) {
			totalRow.createCell(i + 1).setCellValue(containerValues[lastLine - 5][i]);
			totalRow.getCell(i + 1).setCellStyle(style);
		}
		for (int i = 1; i < 13; i++) {
			sheet.autoSizeColumn(i);
		}
		//check author
		((POIXMLDocument) workbook).getProperties().getCoreProperties().setCreator("name"); 
		FileOutputStream out = new FileOutputStream(
				new File("input file name"));
		workbook.write(out);

		out.close();
	}
}
