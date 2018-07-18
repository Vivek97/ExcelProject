package com.action;

import java.io.BufferedInputStream;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelAction {

	public static void main(String[] args) {
		InputStream inputStream = null;
		String workbookPath= "Files//AMBBS.xlsx";

			try {
				inputStream = new BufferedInputStream(new FileInputStream(
						workbookPath));
				Workbook workbook = WorkbookFactory.create(inputStream);
				System.out.println("File found");
				inputStream.close();
				
				workbook.createSheet("Demo");
				
				
			} catch (FileNotFoundException e) {
				
				e.printStackTrace();
			} catch (EncryptedDocumentException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (InvalidFormatException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		

		
		
}

}
