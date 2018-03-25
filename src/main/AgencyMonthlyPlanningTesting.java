package main;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import main.utils.CommonUtil;
import main.utils.XLSXReaderWriter;

public class AgencyMonthlyPlanningTesting {
	private final static String SHEET_ACTIVITY_PLAN = "Activity plan";
	private final static int SHEET_ACTIVITY_PLAN_SKIPPED_ROWS = 11;
	private final static int SHEET_ACTIVITY_PLAN_MAX_ROW = 3;
	
	private final static String SHEET_INDIVIDUALS_DETAIL_PLAN = "Individuals_Detail Plan";
	private final static int SHEET_INDIVIDUALS_DETAIL_PLAN_SKIPPED_ROWS = 10;
	
	private final static String SHEET_LEADERS_DETAIL_PLAN = "Leaders_Detail Plan";
	private final static int SHEET_LEADERS_DETAIL_PLAN_SKIPPED_ROWS = 13;
	
	private final static int TEAM_COLUMN_INDEX = 2;
	private final static int ACTIVITY_COLUMN_INDEX = 1;
	
	private static CommonUtil util = new CommonUtil();
	private static XLSXReaderWriter xlsxRW = new XLSXReaderWriter();
	private static String exportedFile = "d:\\workspace_excel\\PHUCANH\\Agency Monthly Sales Plan 201803\\output\\out.xlsx";
	private static String excelFilesFolder = "d:\\workspace_excel\\PHUCANH\\Agency Monthly Sales Plan 201803\\";
	private static String[] sheets = new String[] {"Indi", "Unit", "Sessions", "Attendees"};
	
	//test
	public static int write(XSSFWorkbook book, String sheetName, int rownum, ArrayList<Row> rows ) throws IOException  {
		XSSFSheet sheet = book.getSheet(sheetName);
		//1. Create the date cell style
		CellStyle cellStyle = book.createCellStyle();
		CreationHelper createHelper = book.getCreationHelper();
		cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd/mm/yyyy"));
		
		// writing data into XLSX file
		for (int i=0; i<rows.size(); i++) {
			Row row = rows.get(i);
			Row newRow = sheet.getRow(rownum+i);
			if (newRow == null) {
				newRow = sheet.createRow(rownum+i);
			}
			// Iterating over each column of Excel file
			Iterator<Cell> cellIterator = row.cellIterator();
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				Cell newCell = newRow.getCell(cell.getColumnIndex());
				if (newCell == null) {
					newCell = newRow.createCell(cell.getColumnIndex());
				}
				CellType cellTypeEnum = cell.getCellTypeEnum();
				switch (cellTypeEnum) {
				case STRING:
					newCell.setCellValue(cell.getStringCellValue());
					break;
				case NUMERIC:
					if (HSSFDateUtil.isCellDateFormatted(cell)) {
						newCell.setCellValue(cell.getDateCellValue());
						newCell.setCellStyle(cellStyle);
					} else {
						CellStyle cs = book.createCellStyle();
						cs.setDataFormat(createHelper.createDataFormat().getFormat( cell.getCellStyle().getDataFormatString()));
						newCell.setCellValue(cell.getNumericCellValue());
						newCell.setCellStyle(cs);
					}
					break;
				case BOOLEAN:
					newCell.setCellValue(cell.getBooleanCellValue());
					break;
				case FORMULA:
					 switch(cell.getCachedFormulaResultTypeEnum()) {
			            case NUMERIC:
			                newCell.setCellValue(cell.getNumericCellValue());
			                break;
			            case STRING:
			                newCell.setCellValue(cell.getRichStringCellValue());
			                break;
					default:
						break;
			        }
					break;
				default:
					break;
				}
			}
		}
		return rownum+rows.size();
	}
	//test
	private static ArrayList<Row> filter(ArrayList<Row> rows, String teamName, int filteredColIndex) throws IOException {
		ArrayList<Row> filteredRows = new ArrayList<Row>();
		// filter data by team name in the column 2 nd
		filteredRows.addAll(xlsxRW.filter(rows, teamName, filteredColIndex)); 
		return filteredRows;
	}
	
	/**
	 * 
	 * @param rowidx
	 * @param st: pathname string
	 * @param teamName: data from source will be filtered by team name
	 * @param exportedFile: destination file where the output data will be saved
	 * @return row index: the index of the last row in the excel output file
	 * @throws IOException
	 */
	private static int filterWriteData(ArrayList<Row> rows, String teamName, String exportedFile, int rowidx, String sheetname, int filteredColIndex) throws IOException {
		if (rowidx == 0) {
			Row headerRow = rows.get(0);
			ArrayList<Row> header = new ArrayList<Row>();
			header.add(headerRow);
			rowidx = xlsxRW.write(exportedFile, sheetname, rowidx, header);
		} 
		ArrayList<Row> newData = (xlsxRW.filter(rows, teamName, filteredColIndex)); // filter data by team name in the column 2 nd
		rowidx = xlsxRW.write(exportedFile, sheetname, rowidx, newData);
		return rowidx;
	}

	public static void main(String[] args) {
		long startTime = System.currentTimeMillis();
		try {
			xlsxRW.createXLSX(exportedFile, sheets);
			File excel = new File(exportedFile);
			FileInputStream fis = new FileInputStream(excel);
			XSSFWorkbook book = new XSSFWorkbook(fis);
			
			// get all files from the folder
			File[] files = util.getFiles(excelFilesFolder);
			int rowidxIndiSheet = 0;
			int rowidxUnitSheet = 0;
			int rowidxSessSheet = 0;
			int rowidxAtteSheet = 0;
			
			List<String> teamcolumn = new ArrayList<String>();
			teamcolumn.add("TEAM");
			for (int i=0; i<files.length; i++) {
				File f = files[i];
				if (!f.getName().startsWith("~")) {
					String st = f.getPath();
					String teamName = st.split(" _ ")[1].split(".xls")[0];
					System.out.println(String.format("%d. Reading for: %s", i, teamName));
					// Read data from the current excel file
					
					 Map<String, Object> dataDic = xlsxRW.read(f,
							new String[] { SHEET_INDIVIDUALS_DETAIL_PLAN, SHEET_LEADERS_DETAIL_PLAN, SHEET_ACTIVITY_PLAN },
							new int[] { SHEET_INDIVIDUALS_DETAIL_PLAN_SKIPPED_ROWS,	SHEET_LEADERS_DETAIL_PLAN_SKIPPED_ROWS, SHEET_ACTIVITY_PLAN_SKIPPED_ROWS },
							new int[] {-1, -1, SHEET_ACTIVITY_PLAN_MAX_ROW});

					 teamcolumn.add(teamName);
					
					Set<String> keySet = dataDic.keySet();
					for (String k: keySet) {
						@SuppressWarnings("unchecked")
						ArrayList<Row> rows = (ArrayList<Row>)dataDic.get(k);
						if (k.equals(SHEET_INDIVIDUALS_DETAIL_PLAN)) {
							ArrayList<Row> filteredRows = filter(rows, teamName, TEAM_COLUMN_INDEX);
							if (book.getSheet(sheets[0]).getLastRowNum()==0) {
								// add header
								filteredRows.add(0, rows.get(0));
							}
							rowidxIndiSheet = write(book, sheets[0], rowidxIndiSheet, filteredRows);
						} else if (k.equals(SHEET_LEADERS_DETAIL_PLAN)) {
							ArrayList<Row> filteredRows = filter(rows, teamName, TEAM_COLUMN_INDEX);
							if (book.getSheet(sheets[1]).getLastRowNum()==0) {
								// add header
								filteredRows.add(0, rows.get(0));
							}
							rowidxUnitSheet = write(book, sheets[1], rowidxUnitSheet, filteredRows);
						} else if (k.equals(SHEET_ACTIVITY_PLAN)) {
							ArrayList<Row> filteredRows = filter(rows, "Sessions", ACTIVITY_COLUMN_INDEX);
							if (book.getSheet(sheets[2]).getLastRowNum()==0) {
								// add header
								filteredRows.add(0, rows.get(0));
							}
							rowidxSessSheet = write(book, sheets[2], rowidxSessSheet, filteredRows);
							filteredRows = filter(rows, "Attendees", ACTIVITY_COLUMN_INDEX);
							if (book.getSheet(sheets[3]).getLastRowNum()==0) {
								// add header
								filteredRows.add(0, rows.get(0));
							}
							rowidxAtteSheet = write(book, sheets[3], rowidxAtteSheet, filteredRows);
						}
					}
				}
			}
			
			// open an OutputStream to save written data into Excel file
			FileOutputStream os = new FileOutputStream(excel);
			book.write(os);
			// Close workbook, OutputStream and Excel file to prevent leak
			os.close();
			book.close();
			fis.close();
			
			xlsxRW.writeColumn(exportedFile, sheets[2], 0, 0, teamcolumn.toArray());
			xlsxRW.writeColumn(exportedFile, sheets[3], 0, 0, teamcolumn.toArray());
			xlsxRW.resizeCol(exportedFile, sheets);
			long stopTime = System.currentTimeMillis();
			long elapsedTime = stopTime - startTime;
			// Converting Milliseconds to Minutes and Seconds
			long minutes = TimeUnit.MILLISECONDS.toMinutes(elapsedTime);
			long seconds = TimeUnit.MILLISECONDS.toSeconds(elapsedTime);
			System.out.println(String.format("Done in %d minute(s) %d seconds.", minutes, seconds));
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
