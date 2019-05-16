
public class Quire {
	private String department;
	private int semester; // X
	
	public Quire(String department) {
		super();
		this.department = department;
	}

	public void loadLink() { // diagram modify : private to public
		//{{path
		String path = null;
		if (this.department.equals("IECS")) {
			path = "http://www.iecs.fcu.edu.tw/wSite/publicfile/Data/f1470638754357.pdf";
		} else if (this.department.equals("EE")) {
			path = "http://www.ee.fcu.edu.tw/wSite/publicfile/Attachment/f1443152104382.pdf";
		} else if (this.department.equals("AUTO")) {
			path = "http://www.auto.fcu.edu.tw/wSite/ct?xItem=142235&ctNode=4640&mp=390101&idPath=4601_4639";
		} else if (this.department.equals("CE")) {
			path = "http://www.ce.fcu.edu.tw/wSite/publicfile/Attachment/f1468998370632.pdf";
		}//}}
		
		//{{調用預設瀏覽器
		if (java.awt.Desktop.isDesktopSupported()) {
		    try {
		        java.net.URI uri = java.net.URI.create(path); // construct URI instance
		        java.awt.Desktop dp = java.awt.Desktop.getDesktop(); // 獲取當前桌面擴展
		        if (dp.isSupported(java.awt.Desktop.Action.BROWSE)) { // 檢查系統桌面是否支持
		            dp.browse(uri); // 以預設瀏覽器打開URI
		        }
		    } catch (java.lang.NullPointerException e) { // URI為空
		        System.out.println("連結為空"); // diagram modify
		    } catch (java.io.IOException e) { // 無法調用系統預設瀏覽器
		    	System.out.println("無法調用系統預設瀏覽器");
		    }
		}
	}
	
	public void setDepartment(String department) {
		this.department = department;
	}

	public void choose(int semester) { // X
		this.semester = semester;
		if (this.semester == 0) {}
	}
	
}
