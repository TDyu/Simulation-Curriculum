import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
//import java.util.Iterator;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.impl.piccolo.io.FileFormatException;

public class SimulationCurriculum extends Course {

	private String department; // class diagram miss
	private String id; // class diagram miss
	private int semester; // class diagram miss
	private String[] compulsory; // Code(int)(4)+Credits(String)(1)+Day(String)(1)+Start(String)(2)+End(String)(2)+Day2(String)(2)+Start2(String)(2)+End2(String)(2)+Name(String)(length-16) 當學期可選 
	private String[] elective; // Code(int)(4)+Credits(String)(1)+Day(String)(1)+Start(String)(2)+End(String)(2)+Day2(String)(2)+Start2(String)(2)+End2(String)(2)+Name(String)(length-16) 當學期可選
	private String[] common; // Code(int)(4)+Credits(String)(1)+Day(String)(1)+Start(String)(2)+End(String)(2)+Day2(String)(2)+Start2(String)(2)+End2(String)(2)+Name(String)(length-16) 當學期可選
	private String[] nonmajor; // Code(int)(4)+Credits(String)(1)+Day(String)(1)+Start(String)(2)+End(String)(2)+Day2(String)(2)+Start2(String)(2)+End2(String)(2)+Name(String)(length-16) 當學期可選
	private String[] credits = {"0", "0", "0", "0"}; // 目前已選學分數  //class diagram miss 
	private ArrayList<String> course = new ArrayList<String>(); // 已選 code+name // class diagram miss 
	private int choice;
	private Quire quireRecognize; // class diagram miss
	// private Quire quireHistoricCurriculum; // sequence diagram modify : delete
	private ArrangeCredits arrange; // class diagram miss

	public SimulationCurriculum(String department, String id, int semester) throws FileNotFoundException, FileFormatException, IOException {
		super();
		this.department = department;
		this.id = id;
		this.semester = semester;
		this.quireRecognize = new Quire(department);
		this.arrange = new ArrangeCredits(this.id, this.credits);
	}

	@SuppressWarnings("deprecation")
	private void pushLastHistoricTable(boolean first) throws FileNotFoundException, FileFormatException, IOException { // diagram modify : add boolean first

		String path = "D:/[Programming]/JAVA/Simulation Curriculum/data/History/" + this.id + ".xlsx"; // @@@
		File file = new File(path);
		if (!file.exists() && first) { // check
			// create new
			this.createTable(this.id, path);
			this.pushCompulsory();
			show(path, this.semester, "=");
		} else if (first) {
			XSSFWorkbook wb = null;
			try {
				InputStream is = new FileInputStream(path);
				wb = new XSSFWorkbook(is);
				XSSFSheet sheet = wb.getSheetAt(this.semester - 1);
				XSSFCell cell;
				for (int i=1; i<15; i++) {
					XSSFRow row = sheet.getRow(i);
					for (int j=1; j<8; j++) {
						cell = row.getCell(j);
						//{{判斷取值
						String currentCell;
						if (true) {
							cell.setCellType(Cell.CELL_TYPE_STRING);
						}

						if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
							currentCell = String.valueOf(cell.getBooleanCellValue());
						} else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
							currentCell = String.valueOf(cell.getNumericCellValue());
						} else {
							currentCell = String.valueOf(cell.getStringCellValue());
						}
						//}}
						if (!currentCell.equals("x")) {
							boolean have = false;
							for (int k=0; k<this.course.size(); k++) { // 不重複加入
								if (currentCell.equals(this.course.get(k))) {
									have = true;
								}
							}
							if (!have)  this.course.add(currentCell); // 加入已選
						} 
					}
				}
				//{{放學分
				for (int i=0; i<this.course.size(); i++) {
					
					boolean find = false;
					for (int cre=0; cre<this.compulsory.length; cre++) {
						if (this.course.get(i).substring(0, 4).equals(this.compulsory[cre].substring(0, 4))) {
							this.credits[0] = Integer.toString(Integer.parseInt(this.credits[0]) + Integer.parseInt(this.compulsory[cre].substring(4, 5))); 
							find = true;
							break;
						}
					}
					if (!find) {
						for (int cre=0; cre<this.elective.length; cre++) {
							if (this.course.get(i).substring(0, 4).equals(this.elective[cre].substring(0, 4))) {
								this.credits[1] = Integer.toString(Integer.parseInt(this.credits[1]) + Integer.parseInt(this.elective[cre].substring(4, 5))); 
								find = true;
								break;
							}
						}
					}
					if (!find) {
						for (int cre=0; cre<this.common.length; cre++) {
							if (this.course.get(i).substring(0, 4).equals(this.common[cre].substring(0, 4))) {
								this.credits[2] = Integer.toString(Integer.parseInt(this.credits[2]) + Integer.parseInt(this.common[cre].substring(4, 5))); 
								find = true;
								break;
							}
						}
					}
					if (!find) {
						for (int cre=0; cre<this.nonmajor.length; cre++) {
							if (this.course.get(i).substring(0, 4).equals(this.nonmajor[cre].substring(0, 4))) {
								this.credits[3] = Integer.toString(Integer.parseInt(this.credits[3]) + Integer.parseInt(this.nonmajor[cre].substring(4, 5))); 
								find = true;
								break;
							}
						}
					}
				}
				//}}
			} catch (Exception e) {
				e.printStackTrace();
			} finally {
				if (wb != null) {
					try {
						wb.close();
					} catch (IOException e) {
						e.printStackTrace();
					}
				}
			}
			show(path, this.semester, "=");
		} else if (!first) {
			show(path, this.semester, "=");
		}
	}

	@SuppressWarnings("deprecation")
	private void pushCompulsory() throws IOException {
		// 預備檔案處理
		String pathCom = "D:/[Programming]/JAVA/Simulation Curriculum/data/Compulsory/" + this.department + ".xlsx"; // @@@
		String pathCurr = "D:/[Programming]/JAVA/Simulation Curriculum/data/History/" + this.id + ".xlsx"; // @@@
		XSSFWorkbook wbCom = null;
		XSSFWorkbook wbCurr = null;
		
		try {
			//{{開必修檔案
			InputStream com = new FileInputStream(pathCom);
			wbCom = new XSSFWorkbook(com);
			//}}
			//{{開模擬課表
			InputStream curr = new FileInputStream(pathCurr);
			wbCurr = new XSSFWorkbook(curr);
			//}}
			
			XSSFSheet sheetCom, sheetCurr;
			XSSFCell cellCom, cellCurr;
			
			for (int numSheet = 0; numSheet < 8; numSheet++) { // 掃描sheet
				sheetCom = wbCom.getSheetAt(numSheet);
				sheetCurr = wbCurr.getSheetAt(numSheet);
				
				if (sheetCurr == null && sheetCom == null) {
					continue;
				}
				
				for (int i = 1; i < 15; i++) { // 掃row 除了title
					XSSFRow rowCom = sheetCom.getRow(i); 
					XSSFRow rowCurr = sheetCurr.getRow(i);
					for (int j = 1; j < 8; j++) { // 掃column 除了title
						cellCom = rowCom.getCell(j);
						cellCurr = rowCurr.getCell(j);
						
						//{{判斷取值
						String currentCell;
						if (true) {
							cellCom.setCellType(Cell.CELL_TYPE_STRING);
						}

						if (cellCom.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
							currentCell = String.valueOf(cellCom.getBooleanCellValue());
						} else if (cellCom.getCellType() == Cell.CELL_TYPE_NUMERIC) {
							currentCell = String.valueOf(cellCom.getNumericCellValue());
						} else {
							currentCell = String.valueOf(cellCom.getStringCellValue());
						}
						
						String storeCell;
						if (true) {
							cellCurr.setCellType(Cell.CELL_TYPE_STRING);
						}

						if (cellCurr.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
							storeCell = String.valueOf(cellCurr.getBooleanCellValue());
						} else if (cellCurr.getCellType() == Cell.CELL_TYPE_NUMERIC) {
							storeCell = String.valueOf(cellCurr.getNumericCellValue());
						} else {
							storeCell = String.valueOf(cellCurr.getStringCellValue());
						}
						//}}
						
						if (!storeCell.equals(currentCell)) cellCurr.setCellValue(currentCell); // 複製必修到模擬課表
						if (!currentCell.equals("x") && numSheet == this.semester - 1) {
							boolean have = false;
							for (int k=0; k<this.course.size(); k++) { // 不重複加入
								if (currentCell.equals(this.course.get(k))) {
									have = true;
								}
							}
							if (!have)  this.course.add(currentCell); // 加入已選
						}  
					}
				}
			}
			//{{放學分
			for (int i=0; i<this.course.size(); i++) {
				boolean find = false;
				for (int cre=0; cre<this.compulsory.length; cre++) {
					if (this.course.get(i).substring(0, 4).equals(this.compulsory[cre].substring(0, 4))) {
						this.credits[0] = Integer.toString(Integer.parseInt(this.credits[0]) + Integer.parseInt(this.compulsory[cre].substring(4, 5))); 
						find = true;
						break;
					}
				}
				if (!find) {
					for (int cre=0; cre<this.elective.length; cre++) {
						if (this.course.get(i).substring(0, 4).equals(this.elective[cre].substring(0, 4))) {
							this.credits[1] = Integer.toString(Integer.parseInt(this.credits[1]) + Integer.parseInt(this.elective[cre].substring(4, 5))); 
							find = true;
							break;
						}
					}
				}
				if (!find) {
					for (int cre=0; cre<this.common.length; cre++) {
						if (this.course.get(i).substring(0, 4).equals(this.common[cre].substring(0, 4))) {
							this.credits[2] = Integer.toString(Integer.parseInt(this.credits[2]) + Integer.parseInt(this.common[cre].substring(4, 5))); 
							find = true;
							break;
						}
					}
				}
				if (!find) {
					for (int cre=0; cre<this.nonmajor.length; cre++) {
						if (this.course.get(i).substring(0, 4).equals(this.nonmajor[cre].substring(0, 4))) {
							this.credits[3] = Integer.toString(Integer.parseInt(this.credits[3]) + Integer.parseInt(this.nonmajor[cre].substring(4, 5))); 
							find = true;
							break;
						}
					}
				}
			}
			//}}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (wbCom != null) {
				try {
					wbCom.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
			if (wbCurr != null) {
				try {
					this.storeTable(pathCurr, wbCurr);
					wbCurr.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}

	@SuppressWarnings("deprecation")
	private void show(String path, int semester, String sign) throws FileNotFoundException, FileFormatException, IOException {

		XSSFWorkbook workbook = null;
		try {
			InputStream is = new FileInputStream(path);
			workbook = new XSSFWorkbook(is);

			XSSFSheet sheet = workbook.getSheetAt(semester - 1);
			
			XSSFCell cell;
			for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
				XSSFRow row = sheet.getRow(i); 
				if (row != null) {
					for (int j = 0; j < 8; j++) { 
						cell = row.getCell(j);
						
						//{{判斷取值
						String currentCell = null;
						if (true) {
							cell.setCellType(Cell.CELL_TYPE_STRING);
						}

						if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
							currentCell = String.valueOf(cell.getBooleanCellValue());
						} else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
							currentCell = String.valueOf(cell.getNumericCellValue());
						} else {
							currentCell = String.valueOf(cell.getStringCellValue());
						}
						//}}
						
						if (j == 0) { // row title
							System.out.print(sign + currentCell + '\t' + "|");
						} else if (i == 0 && j != 7) { // column title
							System.out.print(currentCell + "        " + '\t' + "|");
						} else if (i == 0 && j == 7) { // column title
							System.out.print(currentCell + "          " + sign);
						} else if (j == 7 && currentCell.equals("x")) {
							System.out.print("             " + sign);
						} else if (j != 7 && currentCell.equals("x")) {
							System.out.print("           " + '\t' + "|");
						} else {
							if (currentCell.length() < 12) {
								for (int k=0; k<(12-currentCell.length()); k++) {
									currentCell += " ";
								}
							}
							System.out.print(currentCell + '\t' + "|");// 取出j列j行的值
						}
					}
					System.out.println();
					if (i != 14)  System.out.println(sign + "---------------------------------------------------------------------------------------------------------------------" + sign);
				}

			}

		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (workbook != null) {
				try {
					workbook.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}

	private void createTable(String fileName, String path) throws IOException {

		XSSFWorkbook curriculum = new XSSFWorkbook();

		for (int i = 1; i <= 8; i++) {
			String sheetName = "" + i;
			XSSFSheet sheet = curriculum.createSheet(sheetName);

			for (int r = 0; r < 15; r++) { // 14節+1空白 row
				XSSFRow row = sheet.createRow(r);

				for (int c = 0; c < 8; c++) { // 7天+1空白 column
					XSSFCell cell = row.createCell(c);
					String[] week = { "MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN" };
					if (c != 0 && r == 0) { // write column name
						cell.setCellValue(week[c - 1]);
					} else if (c == 0 && r != 0 && r < 10) { // write row name
						cell.setCellValue("0" + r);
					} else if (c == 0 && r != 0 && r >= 10) { // write row name
						cell.setCellValue(r + " ");
					} else if (c == 0 && r == 0) {
						cell.setCellValue(" ");
					} else {
						cell.setCellValue("x");
					}
				}
			}
		}

		storeTable(path, curriculum);
	}

	private void storeTable(String path, XSSFWorkbook wb) throws IOException {

		try {
			FileOutputStream fileOut = new FileOutputStream(path);
			wb.write(fileOut);
			fileOut.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private boolean closeSet(boolean close, Scanner cin) throws IOException { // diagram modify : void to boolean & add boolean close & add Scanner cin

		return this.arrange.compareSet(close, cin);
	}

	@SuppressWarnings({ "deprecation", "static-access" })
	public String[] Switch() throws FileNotFoundException, FileFormatException, IOException { // diagram modify : void to String[]
		
		//{{載入選分狀況
		int[] need = new int[4];
		int[] get = new int[4];
		String pathCre = "D:/[Programming]/JAVA/Simulation Curriculum/data/Credits/" + this.id + ".xlsx"; // @@@
		XSSFWorkbook credit = null;
		try {
			InputStream is = new FileInputStream(pathCre);
			credit = new XSSFWorkbook(is);
			XSSFSheet sheet = credit.getSheetAt(0);
			
			XSSFRow row = sheet.getRow(1);
			XSSFCell cell;
			for (int i=1; i<5; i++) {
				cell = row.getCell(i);
				need[i-1] = (int)cell.getNumericCellValue();
			}
			row = sheet.getRow(2);
			for (int i=1; i<5; i++) {
				cell = row.getCell(i);
				get[i-1] = (int)cell.getNumericCellValue();
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (credit != null) {
				try {
					credit.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		//}}
		
		//{{載入可選
		ArrayList<String> com = new ArrayList<String>();
		ArrayList<String> ele = new ArrayList<String>();
		ArrayList<String> comm = new ArrayList<String>();
		ArrayList<String> non = new ArrayList<String>();
		String pathCourse = "D:/[Programming]/JAVA/Simulation Curriculum/data/Course/" + this.department + ".xlsx"; // @@@
		XSSFWorkbook course = null;
		try {
			InputStream is = new FileInputStream(pathCourse);
			course = new XSSFWorkbook(is);
			XSSFSheet sheet;
			if (this.semester % 2 == 0)  sheet = course.getSheetAt(1);
			else sheet = course.getSheetAt(0);
			XSSFRow row = sheet.getRow(0);
			XSSFCell cell;
			for (int i=1; i<sheet.getPhysicalNumberOfRows(); i++) {
				row = sheet.getRow(i);
				cell = row.getCell(1);
				XSSFCell code = row.getCell(0);
				//{{判斷取值
				String codeCell = null;
				if (true) {
					code.setCellType(Cell.CELL_TYPE_STRING);
				}

				if (code.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
					codeCell = String.valueOf(code.getBooleanCellValue());
				} else if (code.getCellType() == code.CELL_TYPE_NUMERIC) {
					codeCell = String.valueOf(code.getNumericCellValue());
				} else {
					codeCell = String.valueOf(code.getStringCellValue());
				}
				//}}
				XSSFCell name = row.getCell(2);
				//{{判斷取值
				String nameCell = null;
				if (true) {
					name.setCellType(name.CELL_TYPE_STRING);
				}

				if (name.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
					nameCell = String.valueOf(name.getBooleanCellValue());
				} else if (name.getCellType() == name.CELL_TYPE_NUMERIC) {
					nameCell = String.valueOf(name.getNumericCellValue());
				} else {
					nameCell = String.valueOf(name.getStringCellValue());
				}
				//}}
				XSSFCell cre = row.getCell(3);
				//{{判斷取值
				String creCell = null;
				if (true) {
					cre.setCellType(cre.CELL_TYPE_STRING);
				}

				if (cre.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
					creCell = String.valueOf(cre.getBooleanCellValue());
				} else if (cre.getCellType() == cre.CELL_TYPE_NUMERIC) {
					creCell = String.valueOf(cre.getNumericCellValue());
				} else {
					creCell = String.valueOf(cre.getStringCellValue());
				}
				//}}
				XSSFCell day = row.getCell(4);
				//{{判斷取值
				String dayCell = null;
				if (true) {
					day.setCellType(day.CELL_TYPE_STRING);
				}

				if (day.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
					dayCell = String.valueOf(day.getBooleanCellValue());
				} else if (day.getCellType() == day.CELL_TYPE_NUMERIC) {
					dayCell = String.valueOf(day.getNumericCellValue());
				} else {
					dayCell = String.valueOf(day.getStringCellValue());
				}
				//}}
				XSSFCell start = row.getCell(5);
				//{{判斷取值
				String startCell = null;
				if (true) {
					start.setCellType(start.CELL_TYPE_STRING);
				}

				if (start.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
					startCell = String.valueOf(start.getBooleanCellValue());
				} else if (start.getCellType() == start.CELL_TYPE_NUMERIC) {
					startCell = String.valueOf(start.getNumericCellValue());
				} else {
					startCell = String.valueOf(start.getStringCellValue());
				}
				//}}
				XSSFCell end = row.getCell(6);
				//{{判斷取值
				String endCell = null;
				if (true) {
					end.setCellType(end.CELL_TYPE_STRING);
				}

				if (end.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
					endCell = String.valueOf(end.getBooleanCellValue());
				} else if (end.getCellType() == end.CELL_TYPE_NUMERIC) {
					endCell = String.valueOf(end.getNumericCellValue());
				} else {
					endCell = String.valueOf(end.getStringCellValue());
				}
				//}}
				XSSFCell day2 = row.getCell(7);
				//{{判斷取值
				String day2Cell = null;
				if (true) {
					day2.setCellType(day2.CELL_TYPE_STRING);
				}

				if (day2.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
					day2Cell = String.valueOf(day2.getBooleanCellValue());
				} else if (day2.getCellType() == day2.CELL_TYPE_NUMERIC) {
					day2Cell = String.valueOf(day2.getNumericCellValue());
				} else {
					day2Cell = String.valueOf(day2.getStringCellValue());
				}
				//}}
				XSSFCell start2 = row.getCell(8);
				//{{判斷取值
				String start2Cell = null;
				if (true) {
					start2.setCellType(start2.CELL_TYPE_STRING);
				}

				if (start2.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
					start2Cell = String.valueOf(start2.getBooleanCellValue());
				} else if (start2.getCellType() == start2.CELL_TYPE_NUMERIC) {
					start2Cell = String.valueOf(start2.getNumericCellValue());
				} else {
					start2Cell = String.valueOf(start2.getStringCellValue());
				}
				//}}
				XSSFCell end2 = row.getCell(9);
				//{{判斷取值
				String end2Cell = null;
				if (true) {
					end2.setCellType(end2.CELL_TYPE_STRING);
				}

				if (end2.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
					end2Cell = String.valueOf(end2.getBooleanCellValue());
				} else if (end2.getCellType() == end2.CELL_TYPE_NUMERIC) {
					end2Cell = String.valueOf(end2.getNumericCellValue());
				} else {
					end2Cell = String.valueOf(end2.getStringCellValue());
				}
				//}}
				if (cell.getStringCellValue().equals("必修")) {
					com.add(codeCell + creCell + dayCell + startCell + endCell + day2Cell + start2Cell + end2Cell + nameCell);
				} else if (cell.getStringCellValue().equals("選修")) {
					ele.add(codeCell + creCell + dayCell + startCell + endCell + day2Cell + start2Cell + end2Cell + nameCell);
				} else if (cell.getStringCellValue().equals("通識")) {
					comm.add(codeCell + creCell + dayCell + startCell + endCell + day2Cell + start2Cell + end2Cell + nameCell);
				} else if (cell.getStringCellValue().equals("外系")) {
					non.add(codeCell + creCell + dayCell + startCell + endCell + day2Cell + start2Cell + end2Cell + nameCell);
				}
			}
			this.compulsory = new String[com.size()];
			for (int i=0; i<com.size(); i++)  this.compulsory[i] = com.get(i);
			this.elective = new String[ele.size()];
			for (int i=0; i<ele.size(); i++)  this.elective[i] = ele.get(i);
			this.common = new String[comm.size()];
			for (int i=0; i<comm.size(); i++)  this.common[i] = comm.get(i);
			this.nonmajor = new String[non.size()];
			for (int i=0; i<non.size(); i++)  this.nonmajor[i] = non.get(i);
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (course != null) {
				try {
					course.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		//}}
		
		int code, semesterQuire, setGoal, index;
		Scanner cin = new Scanner(System.in);
		boolean flag = true;
		boolean open = false;
		boolean first = true;
		while (flag) {
			// 模擬課表
			System.out.println("+++++++++++++++++++++++++++++++++++++++++++++++++此為模擬課表頁面上端+++++++++++++++++++++++++++++++++++++++++++++++++++++++++");
			System.out.println("=======================================================================================================================");
			System.out.println("=當前學期模擬課表                                                                                                                                                                                                                                                                                                                                                                                      =");
			this.pushLastHistoricTable(first);
			System.out.println("=======================================================================================================================");
			// 學分狀況
			System.out.println("########################");
			System.out.println("#學分狀況                                                          #");
			System.out.println("#      |必修 |選修 |通識 |外系  #");
			System.out.println("#----------------------#");
			System.out.print("#畢業所需 |");
			for (int i=0; i<4; i++) {
				if (i != 3)  {
					if (need[i] >= 10)  System.out.print(need[i] + " |");
					else  System.out.print(need[i] + "  |");
				} else {
					if (need[i] >= 10)  System.out.print(need[i] + " ");
					else System.out.print(need[i] + "  ");
				}  
			}
			System.out.println("#");
			System.out.print("#目前已獲 |");
			for (int i=0; i<4; i++) {
				if (i != 3)  {
					if (get[i] >= 10)  System.out.print(get[i] + " |");
					else  System.out.print(get[i] + "  |");
				} else {
					if (get[i] >= 10)  System.out.print(get[i] + " ");
					else System.out.print(get[i] + "  ");
				} 
			}
			System.out.println("#");
			System.out.print("#目前預獲 |");
			for (int i=0; i<4; i++) {
				if (i != 3) {
					if (this.credits == null)  System.out.print("0  |");
					else {
						if (Integer.parseInt(this.credits[i]) >= 10)  System.out.print(this.credits[i] + " |");
						else  System.out.print(this.credits[i] + "  |");
					}
				} else {
					if (this.credits == null)  System.out.print("0  ");
					else {
						if (Integer.parseInt(this.credits[i]) >= 10)  System.out.print(this.credits[i] + " ");
						else  System.out.print(this.credits[i] + "  ");
					}
				}  
			}
			System.out.println("#");
			System.out.println("########################");
			// 學分安排系統
			if (open)  this.openArrange();;
			
			// 選單
			System.out.println("1: 模擬選課");
			System.out.println("2: 歷史模擬課表查詢");
			System.out.println("3: 外系承認查詢");
			System.out.println("4: 學分安排系統開啟(預設關)");
			System.out.println("5: 學分安排系統設定(開啟狀態可用)");
			System.out.println("6: 學分安排系統關閉(開啟狀態可用)");
			System.out.println("7: 離開模擬");
			System.out.println("請輸入以上選項: ");
			choice = cin.nextInt();

			switch (choice) {
			case 1:
				System.out.print("請輸入選課代碼: ");
				code = cin.nextInt();
				this.chooseCouse(code);
				break;
			case 2:
				System.out.print("請選擇所要查詢之學期: ");
				semesterQuire = cin.nextInt();
				while (semesterQuire > this.semester) {
					System.out.print("請勿填未來學期，請重新選擇: ");
					semesterQuire = cin.nextInt();
				}
				this.openHistoricTable(semesterQuire);
				break;
			case 3:
				this.openLink();
				break;
			case 4:
				if (!open)  open = true;
				break;
			case 5:
				if (open) {
					System.out.println("請選擇設定哪一格(1:必修 2:選修 3:通識 4:外系選修): ");
					index = cin.nextInt() - 1;
					while (index < 0 || index > 3) {
						System.out.println("請輸入合理選項(1:必修 2:選修 3:通識 4:外系選修): ");
						index = cin.nextInt() - 1;
					}
					System.out.println("請輸入目標值(>0&<=30整數): ");
					setGoal = cin.nextInt();
					while (setGoal < 0 || setGoal > 30) {
						System.out.println("請輸入合理之整數: ");
						setGoal = cin.nextInt();
					}
					this.arrange.setGoalValue(setGoal, index);
				} else {
					System.out.println("未開啟學分安排系統");
				}
				break;
			case 6:
				if (open) {
					this.closeSet(false, cin);
					open = false;
					
					String[] empty = {"0", "0", "0", "0"};
					this.arrange.setNowCredits(empty);
				} else {
					System.out.println("未開啟學分安排系統");
				}
				break;
			case 7:
				if (open) {
					flag = !this.closeSet(true, cin);
					if (!flag) {
						open = false;
						String[] empty = {"0", "0", "0", "0"};
						this.arrange.setNowCredits(empty);
					} 
				} else {
					flag = false;
				}
				break;
			default:
				System.out.println("請填正確選項");
			}
			System.out.println("+++++++++++++++++++++++++++++++++++++++++++++++++此為模擬課表頁面下端+++++++++++++++++++++++++++++++++++++++++++++++++++++++++");
			System.out.println();
			first = false;
		}
		cin.close();
		String[] re = new String[this.course.size()];
		for (int i=0; i<this.course.size(); i++) {
			re[i] = this.course.get(i);
		}
		this.course.clear();
		return re;
	}

	public void openHistoricTable(int semester) throws FileNotFoundException, FileFormatException, IOException { // diagram modify : add int semester

		String path = "D:/[Programming]/JAVA/Simulation Curriculum/data/History/" + this.id + ".xlsx"; // @@@

		System.out.println();
		System.out.println("***************************************************歷史模擬課表視窗上端**********************************************************");
		System.out.println("*第" + semester + "學期歷史模擬課表                                                                                                                                                                                                                                                                                                                                                                             *");
		this.show(path, semester, "*");
		System.out.println("***************************************************歷史模擬課表視窗下端**********************************************************");
	}

	public void openLink() {
		this.quireRecognize.loadLink();
	}

	public void openArrange() throws FileNotFoundException, FileFormatException, IOException {
		// 以下會分散在各地 不再這裡
		this.arrange.setNowCredits(this.credits);
		this.arrange.show();
	}

	@SuppressWarnings({ "deprecation", "static-access" })
	public void chooseCouse(int code) throws IOException {
		
		//{{check重複選
		if (!this.course.isEmpty()) {
			for (int i=0; i<this.course.size(); i++) {
				int codeAl= Integer.parseInt(this.course.get(i).substring(0, 4));
				if (code == codeAl) {
					System.out.println("重複選課");
					return;
				}
			}
		}
		//}}
		//{{找屬性
		int kind = -1; //0:com 1:ele 2:comm 3:non 
		int dayInt = 0, day2Int = 0;
		String credit = null, name = null, day = null, start = null, end = null, day2 = null, start2 = null, end2 = null; 
		boolean find = false;
		for (int i=0; i<this.compulsory.length; i++) {
			int codeAl= Integer.parseInt(this.compulsory[i].substring(0, 4));
			if (code == codeAl) {
				kind = 0;
				credit = this.compulsory[i].substring(4, 5);
				day = this.compulsory[i].substring(5, 6);
				start = this.compulsory[i].substring(6, 8);
				end = this.compulsory[i].substring(8, 10);
				if (!this.compulsory[i].substring(10, 12).equals("xx")) {
					day2 = this.compulsory[i].substring(10, 12);
					start2 = this.compulsory[i].substring(12, 14);
					end2 = this.compulsory[i].substring(14, 16);
				} 
				name = this.compulsory[i].substring(16, this.compulsory[i].length());
				find = true;
				break;
			}
		}
		if (!find) {
			for (int i=0; i<this.elective.length; i++) {
				int codeAl= Integer.parseInt(this.elective[i].substring(0, 4));
				if (code == codeAl) {
					kind = 0;
					credit = this.elective[i].substring(4, 5);
					day = this.elective[i].substring(5, 6);
					start = this.elective[i].substring(6, 8);
					end = this.elective[i].substring(8, 10);
					if (!this.elective[i].substring(10, 12).equals("xx")) {
						day2 = this.elective[i].substring(10, 12);
						start2 = this.elective[i].substring(12, 14);
						end2 = this.elective[i].substring(14, 16);
					} 
					name = this.elective[i].substring(16, this.elective[i].length());
					find = true;
					break;
				}
			}
		}
		if (!find) {
			for (int i=0; i<this.common.length; i++) {
				int codeAl= Integer.parseInt(this.common[i].substring(0, 4));
				if (code == codeAl) {
					kind = 0;
					credit = this.common[i].substring(4, 5);
					day = this.common[i].substring(5, 6);
					start = this.common[i].substring(6, 8);
					end = this.common[i].substring(8, 10);
					if (!this.common[i].substring(10, 12).equals("xx")) {
						day2 = this.common[i].substring(10, 12);
						start2 = this.common[i].substring(12, 14);
						end2 = this.common[i].substring(14, 16);
					} 
					name = this.common[i].substring(16, this.compulsory[i].length());
					find = true;
					break;
				}
			}
		}
		if (!find) {
			for (int i=0; i<this.nonmajor.length; i++) {
				int codeAl= Integer.parseInt(this.nonmajor[i].substring(0, 4));
				if (code == codeAl) {
					kind = 0;
					credit = this.nonmajor[i].substring(4, 5);
					day = this.nonmajor[i].substring(5, 6);
					start = this.nonmajor[i].substring(6, 8);
					end = this.nonmajor[i].substring(8, 10);
					if (!this.nonmajor[i].substring(10, 12).equals("xx")) {
						day2 = this.nonmajor[i].substring(10, 12);
						start2 = this.nonmajor[i].substring(12, 14);
						end2 = this.nonmajor[i].substring(14, 16);
					} 
					name = this.nonmajor[i].substring(16, this.compulsory[i].length());
					find = true;
					break;
				}
			}
		}
		if (!find) {
			System.out.println("選課代碼錯誤");
			return;
		}
		
		dayInt = Integer.parseInt(day);
		switch (Integer.parseInt(day)) {
		case 1:
			day = "MON";
			break;
		case 2:
			day = "TUE";
			break;
		case 3:
			day = "WED";
			break;
		case 4:
			day = "THU";
			break;
		case 5:
			day = "FRI";
			break;
		case 6:
			day = "SAT";
			break;
		case 7:
			day = "SUN";
			break;
		}
		if (day2 != null) {
			day2Int = Integer.parseInt(day2);
			switch (Integer.parseInt(day2)) {
			case 1:
				day = "MON";
				break;
			case 2:
				day = "TUE";
				break;
			case 3:
				day = "WED";
				break;
			case 4:
				day = "THU";
				break;
			case 5:
				day = "FRI";
				break;
			case 6:
				day = "SAT";
				break;
			case 7:
				day = "SUN";
				break;
			}
		}
		//}}
		String path = "D:/[Programming]/JAVA/Simulation Curriculum/data/History/" + this.id + ".xlsx"; // @@@
		XSSFWorkbook workbook = null;
		
		try {
			//{{開課表檔案
			InputStream is = new FileInputStream(path);
			workbook = new XSSFWorkbook(is);
			XSSFSheet sheet = workbook.getSheetAt(this.semester - 1);
			//}}
			
			XSSFRow row;
			XSSFCell cell;
			
			//{{check衝堂
			for (int i=1; i<15; i++) {
				row = sheet.getRow(i);
				for (int j=1; j<8; j++) {
					cell = row.getCell(j);
					//{{判斷取值
					String currentCell = null;
					if (true) {
						cell.setCellType(cell.CELL_TYPE_STRING);
					}

					if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
						currentCell = String.valueOf(cell.getBooleanCellValue());
					} else if (cell.getCellType() == cell.CELL_TYPE_NUMERIC) {
						currentCell = String.valueOf(cell.getNumericCellValue());
					} else {
						currentCell = String.valueOf(cell.getStringCellValue());
					}
					//}}
					if (!currentCell.equals("x")) {
						XSSFRow temp = sheet.getRow(0);
						XSSFCell week = temp.getCell(j);
						//{{判斷取值
						String weekCell = null;
						if (true) {
							week.setCellType(week.CELL_TYPE_STRING);
						}

						if (week.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
							weekCell = String.valueOf(week.getBooleanCellValue());
						} else if (week.getCellType() == week.CELL_TYPE_NUMERIC) {
							weekCell = String.valueOf(week.getNumericCellValue());
						} else {
							weekCell = String.valueOf(week.getStringCellValue());
						}
						//}}
						XSSFCell time = row.getCell(0);
						//{{判斷取值
						String timeCell = null;
						if (true) {
							time.setCellType(time.CELL_TYPE_STRING);
						}

						if (time.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
							timeCell = String.valueOf(time.getBooleanCellValue());
						} else if (time.getCellType() == time.CELL_TYPE_NUMERIC) {
							timeCell = String.valueOf(time.getNumericCellValue());
						} else {
							timeCell = String.valueOf(time.getStringCellValue());
						}
						//}}
						
						if (weekCell.equals(day)) {
							
							if (Integer.parseInt(end) - Integer.parseInt(start) >= 2) {
								String mid = Integer.toString(Integer.parseInt(start) + 1);
								if (Integer.parseInt(mid) < 10)  mid = "0" + mid;
								if (timeCell.equals(start) || timeCell.equals(end) || timeCell.equals(mid)) {
									System.out.println("衝堂");
									return;
								}
							} else {
								if (timeCell.equals(start) || timeCell.equals(end)) {
									System.out.println("衝堂");
									return;
								}
							}
						} 
						if (day2 != null) {
							if (weekCell.equals(day2)) {
								if (Integer.parseInt(end2) - Integer.parseInt(start2) >= 2) {
									String mid = Integer.toString(Integer.parseInt(start2) + 1);
									if (Integer.parseInt(mid) < 10)  mid = "0" + mid;
									if (timeCell.equals(start2) || timeCell.equals(end2) || timeCell.equals(mid)) {
										System.out.println("衝堂");
										return;
									}
								} else {
									if (timeCell.equals(start2) || timeCell.equals(end2)) {
										System.out.println("衝堂");
										return;
									}
								}
							}	
						}
					}
				}
			}
			//}}
			
			//{{modify
			this.credits[kind] = Integer.toString(Integer.parseInt(this.credits[kind]) + Integer.parseInt(credit));
			String store = Integer.toString(code) + name;
			this.course.add(store);
			this.arrange.setNowCredits(this.credits);
			//}}
			//{{display
			if (Integer.parseInt(end) - Integer.parseInt(start) >= 2) {
				row = sheet.getRow(Integer.parseInt(start));
				cell = row.getCell(dayInt);
				cell.setCellValue(store);
				row = sheet.getRow(Integer.parseInt(start) + 1);
				cell = row.getCell(dayInt);
				cell.setCellValue(store);
				row = sheet.getRow(Integer.parseInt(end));
				cell = row.getCell(dayInt);
				cell.setCellValue(store);
			} else {
				row = sheet.getRow(Integer.parseInt(start));
				cell = row.getCell(dayInt);
				cell.setCellValue(store);
				row = sheet.getRow(Integer.parseInt(end));
				cell = row.getCell(dayInt);
				cell.setCellValue(store);
			}
			if (day2 != null) {
				if (Integer.parseInt(end2) - Integer.parseInt(start2) >= 2) {
					row = sheet.getRow(Integer.parseInt(start2));
					cell = row.getCell(day2Int);
					cell.setCellValue(store);
					row = sheet.getRow(Integer.parseInt(start2) + 1);
					cell = row.getCell(day2Int);
					cell.setCellValue(store);
					row = sheet.getRow(Integer.parseInt(end2));
					cell = row.getCell(day2Int);
					cell.setCellValue(store);
				} else {
					row = sheet.getRow(Integer.parseInt(start2));
					cell = row.getCell(day2Int);
					cell.setCellValue(store);
					row = sheet.getRow(Integer.parseInt(end2));
					cell = row.getCell(day2Int);
					cell.setCellValue(store);
				}
			}
			//}}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (workbook != null) {
				try {
					this.storeTable(path, workbook); // 存關檔案
					workbook.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}
}
