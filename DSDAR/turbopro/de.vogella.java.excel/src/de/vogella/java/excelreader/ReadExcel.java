package de.vogella.java.excelreader;

import java.io.BufferedWriter;
import java.io.File;
import java.io.IOException;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import java.io.FileOutputStream;
import java.io.OutputStreamWriter;

// CLASS DEFINITION
public class ReadExcel {

	// MAIN
	public static void main(String[] args) throws IOException {

		try {

			ReadExcel turboPro = new ReadExcel();
			String fileName = "C:\\Users\\Dave\\Documents\\DSDAR\\GenerateSql\\MPNCompany2009.xls";
			System.out.println(fileName);

			turboPro.setInputFile(fileName);
			String scriptForFile = turboPro.readAndCreateScript();
			turboPro.writeScriptFile(scriptForFile);

		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	// VARIABLE DEFINITIONS FOR CLASS
	// private String masterScript = "";

	private String inputFile;

	// ===========
	// FUNCTIONS
	// ===========
	public void setInputFile(String inputFile) {
		this.inputFile = inputFile;

	}

	public String readAndCreateScript() throws IOException {
		String masterScript = "";
		File inputWorkbook = new File(inputFile);
		Workbook w;
		try {
			w = Workbook.getWorkbook(inputWorkbook);
			// Get the first sheet
			Sheet sheet = w.getSheet(0);
			// Loop over first 10 column and lines

			// put all your code to iterate through sheet and CONSTRUCT
			// masterScript.

			System.out.println("readAndCreateScript rows=" + sheet.getRows() + " columns=" + sheet.getColumns());

			// for (int j = 0; j < sheet.getColumns(); j++) {
			for (int i = 0; i < sheet.getRows(); i++) {
				Cell column0 = sheet.getCell(0, i);
				Cell column1 = sheet.getCell(1, i);
				Cell column2 = sheet.getCell(2, i);
				Cell column3 = sheet.getCell(3, i);
				Cell column4 = sheet.getCell(4, i);
				Cell column5 = sheet.getCell(5, i); // column f

				System.out.println((i + 1) + " " + column0.getContents() + "\t" + column1.getContents() + "\t"
						+ column2.getContents() + "\t" + column3.getContents() + "\t" + column4.getContents());

				masterScript = masterScript + this.buildScriptEntry(column0.getContents(), column1.getContents(),
						column2.getContents(), column3.getContents(), column4.getContents(), column5.getContents());

				if ((column0.getContents().length() < 1))
					break;
			} // for rows
			// } // for columns

		} catch (BiffException e) {
			e.printStackTrace();

		}

		return masterScript;
	}

	// THis function will simply write to your file.
	public String buildScriptEntry(String column0, String column1, String column2, String column3, String column4,
			String column5) throws IOException {

		String insertGl = "";
		String insertGlLinkage = "";
		String supportInsert = "";
		String sqlVariables = "";
		String aVar = "";
		String bVar = "";
		String eVar = "";
		String dVar = "";
		String fVar = "";
		// String filePath = "";
		String yesNoString = column4;
		String updateGl = "";
		

		System.out.println(
				"buildScriptEntry: " + column0 + " " + column1 + " " + column2 + " " + column3 + " " + column4);

		String scriptEntry = "";

		// copy all our your script building logic here

		System.out.println("[" + yesNoString + "]");

		if (yesNoString.equals("No")) {

			aVar = column0;
			bVar = column1;
			eVar = column2;
			dVar = column3;

			sqlVariables = "SET @ref='" + aVar + "'" + ";" + "\n";
			sqlVariables = sqlVariables + "SET @act='" + bVar + "'" + ";" + "\n";
			sqlVariables = sqlVariables + "Set @debit='" + eVar + "'" + ";" + "\n";
			sqlVariables = sqlVariables + "Set @credit='" + dVar + "'" + ";" + "\n";
			supportInsert = "-- select CONCAT('I ref=', @ref, 'act=', @act,'debit=', @debit, 'cred=', @credit) AS '';"
					+ "\n";
			supportInsert = supportInsert + "SET @coAccountid=(SELECT coAccountID FROM coAccount  WHERE Number=@act);"
					+ "\n";
			System.out.println(sqlVariables);

			insertGl = "INSERT INTO `glTransaction` (coFiscalPeriodId, period, pStartDate, pEndDate,  coFiscalYearId, fyear, yStartDate, yEndDate,"
					+ "\n";
			insertGl = insertGl
					+ "journalId, journalDesc,   entrydate, enteredBy, coAccountId, coAccountNumber, reference, transactionDesc,  transactionDate, debit, credit)"
					+ "\n";
			insertGl = insertGl
					+ "(SELECT coFiscalPeriodId,period,pStartDate,pEndDate,coFiscalYearId,fyear,yStartDate,yEndDate,journalId,journalDesc,entrydate,enteredBy,"
					+ "\n";
			insertGl = insertGl
					+ "@coAccountid, @act, @ref, transactionDesc,transactionDate, @debit, @credit FROM glTransaction WHERE reference=@ref limit 1);"
					+ "\n";
			System.out.println(insertGl);

			insertGlLinkage = "INSERT INTO glLinkage (entryDate,coLedgerSourceID,glTransactionId,veBillID,STATUS)"
					+ "\n";
			insertGlLinkage = insertGlLinkage
					+ "VALUES((select entrydate from glTransaction ORDER BY glTransactionId DESC LIMIT 1)," + "\n";
			insertGlLinkage = insertGlLinkage
					+ "(select s.coLedgerSourceID from glTransaction g,  coLedgerSource s  where g.JournalID = s.JournalID ORDER BY g.glTransactionId DESC LIMIT 1),"
					+ "\n";
			insertGlLinkage = insertGlLinkage
					+ "(select max(glTransactionId) from glTransaction),(select reference from glTransaction ORDER BY glTransactionId DESC LIMIT 1),0);"
					+ "\n";
			System.out.println(insertGlLinkage);
			try {

				String line = sqlVariables + "\n" + supportInsert + "\n" + insertGl + "\n" + insertGlLinkage + "\n";
				scriptEntry = line;

			} catch (Exception e) {
				e.printStackTrace();

				System.exit(0);
			}

		} else {
			System.err.println("Yes");

			// aVar = Worksheets("Sheet1").Range("A" & itterator)
			fVar = column5;
			// Worksheets("Sheet1").Range("F" & itterator)
			sqlVariables = "SET @ref='" + aVar + "'" + ";" + "\n";
			sqlVariables = sqlVariables + "Set @pId='" + fVar + "'" + ";" + "\n";
			updateGl = "Set @p = (select period from coFiscalPeriod where coFiscalPeriodID = @pId);" + "\n";
			updateGl = updateGl
					+ "Set @yId = (select coFiscalYearID from coFiscalPeriod where coFiscalPeriodID = @pId);" + "\n";
			updateGl = updateGl + "Set @y = (select fiscalYear from coFiscalYear where coFiscalYearID = @yId);" + "\n"
					+ "\n";
			System.out.println(sqlVariables);

			updateGl = updateGl
					+ "UPDATE `glTransaction`set coFiscalPeriodId=@pId,period=@p,pStartDate=CONCAT(@y,'-12-0100:00:00'),pEndDate=CONCAT(@y,'-12-31 00:00:00'),"
					+ "\n";
			updateGl = updateGl
					+ "coFiscalYearId=@yId,fyear=@y, yStartDate=CONCAT(@y,'-01-01 00:00:00'),yEndDate=CONCAT(@y,'-12-31 00:00:00')"
					+ "\n";
			updateGl = updateGl + "WHERE reference=@ref and journalId='JE';" + "\n";
			System.out.println(updateGl);
			try {

				String line = sqlVariables + "\n" + updateGl + "\n";

				scriptEntry = line;

			} catch (Exception e) {
				e.printStackTrace();

				System.exit(0);
			}

		}

		// end

		return scriptEntry;
	}

	// THis function will simply write to your file.
	public void writeScriptFile(String scriptForFile) throws IOException {

		

		System.out.println("writeScriptFile: " + scriptForFile);
		
		try {
			
			//open file
			File fout = new File("C:\\Users\\Dave\\Documents\\DSDAR\\GenerateSql\\Years\\MPN Company 2011 One Sided JE Update.txt"); 
			FileOutputStream fos = new FileOutputStream(fout);
			
			//write to file
			BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(fos));
			bw.write(scriptForFile);
			bw.newLine();
		
			
			//close file
			bw.close();
			
		} catch (IOException ioe) { // close try
			System.out.println("Exception caught " + ioe.getMessage());

		} // close catch

	} // close main
		

	
}
