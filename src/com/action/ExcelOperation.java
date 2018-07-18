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




public class ExcelOperation {

	
public static void main(String [] arg) throws EncryptedDocumentException, InvalidFormatException, IOException 
{
	InputStream inputStream = null;
	String workbookPath= "Files//AMBBS.xlsx";

		try {
			inputStream = new BufferedInputStream(new FileInputStream(workbookPath));
			Workbook workbook = WorkbookFactory.create(inputStream);
			System.out.println("File found");
		} catch (FileNotFoundException e) {
			
			e.printStackTrace();
		}


	System.out.println("File found");
	
}

}