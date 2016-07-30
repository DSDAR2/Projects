package insertGLTool;

import java.io.BufferedWriter;
import java.io.File;
import java.io.IOException;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import java.io.FileOutputStream;
import java.io.OutputStreamWriter;

public class insertGLTool {

/*
 * coFiscalPeriodId	journalId	entrydate	enteredBy	coAccountNumber	reference	transactionDesc	debit	credit	linkageID	N/A														
94	VB	2016-07-13 0:00:00	Admin	2010-00	72896	Multistack (Voided veBillID 70036 in DB)	3369.29	0	70036															
94	VB	2016-07-13 0:00:00	Admin	5500-00	72896	Multistack (Voided veBillID 70036 in DB)	0	169.29	70036															
 * 
 * 	
 */

	// CLASS DEFINITION

	// MAIN
	public static void main(String[] args) throws IOException {
		 for (String s: args) {
	            System.out.println(s);}

			try {

				insertGLTool turboPro = new insertGLTool();
				String fileName = "C:\\Users\\Dave\\Documents\\DSDAR\\GenerateSql\\MPNOneSidedJournalEntries.xls";
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
				
				if ((column0.getContents().length() < 1))
					break;
			
				Cell column1 = sheet.getCell(1, i);
				Cell column2 = sheet.getCell(2, i);
				Cell column3 = sheet.getCell(3, i);
				Cell column4 = sheet.getCell(4, i);
				Cell column5 = sheet.getCell(5, i); // column f
				Cell column6 = sheet.getCell(6, i);
				Cell column7 = sheet.getCell(7, i);
				Cell column8 = sheet.getCell(8, i);
				Cell column9 = sheet.getCell(9, i);
				//Cell column10 = sheet.getCell(10, i);
				System.out.println((i + 1) + " " + column0.getContents() + "\t" + column1.getContents() + "\t"
						+ column2.getContents() + "\t" + column3.getContents() + "\t" + column4.getContents() + "\t" + column5.getContents() + "\t" + column6.getContents() + "\t" + column7.getContents() + "\t" + column8.getContents() + "\t" + column9.getContents());

				masterScript = masterScript + this.buildScriptEntry(column0.getContents(), column1.getContents(),
						column2.getContents(), column3.getContents(), column4.getContents(), column5.getContents(), column6.getContents(), column7.getContents(), column8.getContents(), column9.getContents());

				
			} // for rows
				// } // for columns

		} catch (BiffException e) {
			e.printStackTrace();

		}

		return masterScript;
	}

	// THis function will simply write to your file.
	public String buildScriptEntry(String column0, String column1, String column2, String column3, String column4,
			String column5, String column6, String column7, String column8, String column9) throws IOException {
		
		String sqlVariables = "";

		sqlVariables = sqlVariables + "SET @coFPID='" + column0  + "';\n";
		sqlVariables = sqlVariables + "SET @JID='" + column1+ "';\n";
		sqlVariables = sqlVariables + "SET @entrydate='" + column2  + "';\n";
		sqlVariables = sqlVariables + "SET @enteredBy='" + column3  + "';\n";
		sqlVariables = sqlVariables + "SET @accountNumber='" + column4    + "';\n";
		sqlVariables = sqlVariables + "SET @reference='" + column5  + "';\n";
		sqlVariables = sqlVariables + "SET @tdesc='" + column6  + "';\n";
		sqlVariables = sqlVariables + "SET @debit='" + column7  + "';\n";
		sqlVariables = sqlVariables + "SET @credit='" + column8  + "';\n";
		sqlVariables = sqlVariables + "SET @vebill='" + column9  + "';\n";
		//sqlVariables = sqlVariables + "SET @cuInvoiceID='" + column10 + "';\n";
		
		String sqlMasterInsert= sqlVariables + "\n" +
		""
		+ "SET @coAccountid=(SELECT coAccountID FROM coAccount  WHERE Number=@accountNumber); \n"
		+ "SET @JIDNAME=(select Description from coLedgerSource where JournalID=@JID); \n"
		+ "\n"
		+ "INSERT INTO `glTransaction` (coFiscalPeriodId, period, pStartDate, pEndDate,  coFiscalYearId, fyear, yStartDate, yEndDate,\n"
		+ "journalId, journalDesc,   entrydate, enteredBy, coAccountId, coAccountNumber, reference, transactionDesc,  transactionDate, debit, credit)\n"
		+ "(select p.coFiscalPeriodId, p.period, p.StartDate, p.EndDate, p.coFiscalYearID, y.fiscalYear, y.StartDate, y.EndDate,\n"
		+ "@JID, @JIDNAME, @entrydate, @enteredBy, @coAccountid, @accountNumber, @reference, @tdesc, @entrydate, @debit, @credit\n"
		+ "from coFiscalPeriod p, coFiscalYear y  where p.coFiscalYearId = y.coFiscalYearId and p.coFiscalPeriodId=@coFPID);\n"
		+ "\n" 
		+ "INSERT INTO glLinkage (entryDate,coLedgerSourceID,glTransactionId,veBillID,STATUS)\n"
		+ "VALUES((select entrydate from glTransaction ORDER BY glTransactionId DESC LIMIT 1),\n"
		+ "(select s.coLedgerSourceID from glTransaction g,  coLedgerSource s  where g.JournalID = s.JournalID ORDER BY g.glTransactionId DESC LIMIT 1),\n"
		+ "(select max(glTransactionId) from glTransaction), @vebill,0);\n"
		+ "\n ";

		

		System.out.println(
				"buildScriptEntry: " + column0 + " " + column1 + " " + column2 + " " + column3 + " " + column4 + " " + column5 + " " + column6 + " " + column7 + " " + column8 + " " + column9 + " ");

		return sqlMasterInsert;
	}

	// THis function will simply write to your file.
	public void writeScriptFile(String scriptForFile) throws IOException {

		System.out.println("writeScriptFile: " + scriptForFile);

		try {

			// open file
			File fout = new File("C:\\Users\\Dave\\Documents\\DSDAR\\GenerateSql\\out.doc");
			FileOutputStream fos = new FileOutputStream(fout);

			// write to file
			BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(fos));
			bw.write(scriptForFile);
			bw.newLine();

			// close file
			bw.close();

		} catch (IOException ioe) { // close try
			System.out.println("Exception caught " + ioe.getMessage());

		} // close catch

	} // close main

}


