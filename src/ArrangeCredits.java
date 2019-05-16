import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.impl.piccolo.io.FileFormatException;

public class ArrangeCredits {
	
	private String[] goalCredits = {"--", "--", "--", "--"}; // class diagram modify : int to String
	private String[] nowCredits = {"0", "0", "0", "0"}; // class diagram modify : int to String
	private String id; // class diagram miss
	
	public ArrangeCredits(String id, String[] nowCredits) throws FileNotFoundException, FileFormatException, IOException{
		super();
		this.id = id;
		this.loadLastSetValue();
		this.nowCredits = nowCredits;
	}

	@SuppressWarnings("deprecation")
	private void loadLastSetValue() throws FileNotFoundException, FileFormatException, IOException {
		String path = "D:/[Programming]/JAVA/Simulation Curriculum/data/Arrange/" + this.id + ".xlsx"; // @@@

		File file = new File(path);
		if (!file.exists()) { // check
			//{{新建
			XSSFWorkbook arrange = new XSSFWorkbook();

			String sheetName = "arrange";
			XSSFSheet sheet = arrange.createSheet(sheetName);

			for (int r = 0; r < 2; r++) { // 1個設定+1空白 row
					XSSFRow row = sheet.createRow(r);

				for (int c = 0; c < 5; c++) { // 4種+1空白 column
					XSSFCell cell = row.createCell(c);
					String[] style = { "Com", "Sele", "Comm", "Un"};  // 必修  選修  通識  外系
					if (c != 0 && r == 0) { // write column name
						cell.setCellValue(style[c - 1]);
					} else if (c == 0 && r != 0) { // write row name
						cell.setCellValue("Goal");
					} else if (c == 0 && r == 0) {
						cell.setCellValue(" ");
					} else {
						cell.setCellValue("--");
					}
				}
			}
			//}}
			//{{設定值
			XSSFCell cell;
			XSSFRow row = sheet.getRow(1); 
			for (int i=1; i<5; i++) {
				cell = row.getCell(i);

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
					
				this.goalCredits[i-1] = currentCell;
			}
			//}}
			this.storeSetValue(path, arrange); // 存檔
		} else {
			//{{載入&設定值
			XSSFWorkbook wb = null;
			try {
				// 載入
				InputStream is = new FileInputStream(path);
				wb = new XSSFWorkbook(is);
				
				//{{設定值
				XSSFSheet sheet = wb.getSheetAt(0);
				XSSFCell cell;
				XSSFRow row = sheet.getRow(1); 
				for (int i=1; i<5; i++) {
					cell = row.getCell(i);
						
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
					
					this.goalCredits[i-1] = currentCell;
				}
				//}}
			}  catch (Exception e) {
				e.printStackTrace();
			} finally {
				if (wb != null) {
					try {
						this.storeSetValue(path, wb);
						wb.close();
					} catch (IOException e) {
						e.printStackTrace();
					}
				}
			}
			//}}
		}
	}
	
	public void setNowCredits(String[] nowCredits) { // diagram modify : private to public && add String[] nowCredits
		this.nowCredits = nowCredits;
	}
	
	private void storeSetValue(String path, XSSFWorkbook wb) throws IOException { // class diagram modify : add String path, XSSFWorkbook wb 
		
		try {
			//{{update
			XSSFSheet sheet = wb.getSheetAt(0);
			XSSFCell cell;
			XSSFRow row = sheet.getRow(1);
			for (int i=1; i<5; i++) {
				cell = row.getCell(i);
				cell.setCellValue(this.goalCredits[i-1]);
			}
			//}}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (wb != null) {
				try {
					//{{Store
					try {
						FileOutputStream fileOut = new FileOutputStream(path);
						wb.write(fileOut);
						fileOut.close();
					} catch (Exception e) {
						e.printStackTrace();
					}
					//}}
					wb.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}
	
	public boolean compareSet(boolean close, Scanner cin) throws IOException { // class diagram modify : void to boolean & private to public & add boolean close & add Scanner cin
		
		String path = "D:/[Programming]/JAVA/Simulation Curriculum/data/Arrange/" + this.id + ".xlsx"; // @@@
		InputStream is = new FileInputStream(path);
		XSSFWorkbook workbook = new XSSFWorkbook(is);
		boolean flag = false;
		boolean re = true;
		if (close) { // 關閉模擬
			// 確認足額與否
			int goal, now;
			for (int i=0; i<4; i++) {
				if (!this.goalCredits[i].equals("--") && !this.nowCredits[i].equals("--")) {
					goal = Integer.parseInt(this.goalCredits[i]);
					now = Integer.parseInt(this.nowCredits[i]);
					if (goal > now) {
						flag = true;
						break;
					} 
				} else if (!this.goalCredits[i].equals("--") && this.nowCredits[i].equals("--")) {
					flag = true;
					break;
				} else {
					flag = false;
				}
			}
			
			//不足
			if (flag) {
				int sure;
				boolean out = true;
				while (out) {
					System.out.println("///////彈跳視窗詢問上端///////");
					System.out.print("您目前預計獲得學分數不足目標值，是否確定退出模擬選課(0:No / 1:Yes): ");
					sure = cin.nextInt();
					switch(sure) {
						case 0:
							out = false;
							re = false;
							break;
						case 1:
							try {
								this.storeSetValue(path, workbook);
							} catch (Exception e) {
								e.printStackTrace();
							} finally {
								workbook.close();
							}
							out = false;
							re = true;
							break;
						default:
							System.out.println("請填正確選項");
					}
				}
				System.out.println("///////彈跳視窗詢問下端///////");
			}
		} else { // 僅關閉學分安排
			try {
				this.storeSetValue(path, workbook);
			} catch (Exception e) {
				e.printStackTrace();
			} finally {
				workbook.close();
			}
			re = false;
		}
		return re;
	}
	
	public void setGoalValue(int goal, int index) { // diagram modify : add int index
		String goalCre = Integer.toString(goal);
		this.goalCredits[index] = goalCre;
	}
	
	public void show() throws FileNotFoundException, FileFormatException, IOException { // diagram miss
		
		System.out.println("########################");
		System.out.println("#學分安排系統                                                 #");
		System.out.println("#      |必修 |選修 |通識 |外系  #");
		System.out.println("#----------------------#");
		System.out.print("#目前已選  |");
		for (int i=0; i<4; i++) {
			if (i != 3)  {
				if (!this.nowCredits[i].equals("--")) {
					if (Integer.parseInt(this.nowCredits[i]) < 10)  System.out.print(this.nowCredits[i] + "  |");
					else  System.out.print(this.nowCredits[i] + " |");
				} else {
					System.out.print(this.nowCredits[i] + " |");
				}
			} else {
				if (!this.nowCredits[i].equals("--")) {
					if (Integer.parseInt(this.nowCredits[i]) < 10)  System.out.print(this.nowCredits[i] + "  ");
					else  System.out.print(this.nowCredits[i] + " ");
				} else {
					System.out.print(this.nowCredits[i] + " ");
				}
			}  
		}
		System.out.println("#");
		System.out.print("#目標設定  |");
		for (int i=0; i<4; i++) {
			if (i != 3)  {
				if (!this.goalCredits[i].equals("--")) {
					if (Integer.parseInt(this.goalCredits[i]) < 10)  System.out.print(this.goalCredits[i] + "  |");
					else  System.out.print(this.goalCredits[i] + " |");
				} else {
					System.out.print(this.goalCredits[i] + " |");
				}
			} else {
				if (!this.goalCredits[i].equals("--")) {
					if (Integer.parseInt(this.goalCredits[i]) < 10)  System.out.print(this.goalCredits[i] + "  ");
					else  System.out.print(this.goalCredits[i] + " ");
				} else {
					System.out.print(this.goalCredits[i] + " ");
				}
			}  
		}
		System.out.println("#");
		System.out.println("########################");
	}


}
