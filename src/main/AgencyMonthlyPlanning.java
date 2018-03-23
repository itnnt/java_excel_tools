package main;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Row;

import main.utils.CommonUtil;
import main.utils.XLSXReaderWriter;

public class AgencyMonthlyPlanning {
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
					ArrayList<Row> activities = xlsxRW.read(f, SHEET_ACTIVITY_PLAN, SHEET_ACTIVITY_PLAN_SKIPPED_ROWS, SHEET_ACTIVITY_PLAN_MAX_ROW);
					Map<String, Object> dataDic = xlsxRW.read(f,
							new String[] { SHEET_INDIVIDUALS_DETAIL_PLAN, SHEET_LEADERS_DETAIL_PLAN },
							new int[] { SHEET_INDIVIDUALS_DETAIL_PLAN_SKIPPED_ROWS,	SHEET_LEADERS_DETAIL_PLAN_SKIPPED_ROWS });
					// TODO: test
					/*
					 * Map<String, Object> dataDic = xlsxRW.read(f,
							new String[] { SHEET_INDIVIDUALS_DETAIL_PLAN, SHEET_LEADERS_DETAIL_PLAN, SHEET_ACTIVITY_PLAN },
							new int[] { SHEET_INDIVIDUALS_DETAIL_PLAN_SKIPPED_ROWS,	SHEET_LEADERS_DETAIL_PLAN_SKIPPED_ROWS, SHEET_ACTIVITY_PLAN_SKIPPED_ROWS },
							new int[] {-1, -1, SHEET_ACTIVITY_PLAN_MAX_ROW});
					 */
					rowidxSessSheet = filterWriteData(activities, "Sessions", exportedFile, rowidxSessSheet, sheets[2], ACTIVITY_COLUMN_INDEX);
					rowidxAtteSheet = filterWriteData(activities, "Attendees", exportedFile, rowidxAtteSheet, sheets[3], ACTIVITY_COLUMN_INDEX);
					teamcolumn.add(teamName);
					
					Set<String> keySet = dataDic.keySet();
					for (String k: keySet) {
						@SuppressWarnings("unchecked")
						ArrayList<Row> rows = (ArrayList<Row>)dataDic.get(k);
						if (k.equals(SHEET_INDIVIDUALS_DETAIL_PLAN)) {
							rowidxIndiSheet = filterWriteData(rows, teamName, exportedFile, rowidxIndiSheet, sheets[0], TEAM_COLUMN_INDEX);
						} else if (k.equals(SHEET_LEADERS_DETAIL_PLAN)) {
							rowidxUnitSheet = filterWriteData(rows, teamName, exportedFile, rowidxUnitSheet, sheets[1], TEAM_COLUMN_INDEX);
						} else if (k.equals(SHEET_ACTIVITY_PLAN)) {
							rowidxSessSheet = filterWriteData(activities, "Sessions", exportedFile, rowidxSessSheet, sheets[2], ACTIVITY_COLUMN_INDEX);
							rowidxAtteSheet = filterWriteData(activities, "Attendees", exportedFile, rowidxAtteSheet, sheets[3], ACTIVITY_COLUMN_INDEX);
						}
					}
				}
			}
			
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
