import java.util.ArrayList;

//import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Course {
	private String[] courseName;
	private int[] courseCode;
	//private int[] favorite;

	public Course() {
		super();
		//favorite = new int[5];
	}

	public String[] getCourseName() {
		return courseName;
	}

	public void setCourseName(String[] chosenCourse) {
		//{{
		ArrayList<Integer> code = new ArrayList<Integer>();
		ArrayList<String> name = new ArrayList<String>();
		for (int i=0; i<chosenCourse.length; i++) {
			code.add(Integer.parseInt(chosenCourse[i].substring(0, 4)));
			name.add(chosenCourse[i].substring(4));
		}
		int[] codeInt = new int[code.size()];
		for (int i=0; i<code.size(); i++) {
			codeInt[i] = code.remove(i);
		}
		this.setCourseCode(codeInt);
		code.clear();
		String[] nameStr = new String[name.size()];
		for (int i=0; i<name.size(); i++) {
			nameStr[i] = name.remove(i);
		}
		this.courseName = nameStr;
		name.clear();
		//}}
	}

	public int[] getCourseCode() {
		return courseCode;
	}

	public void setCourseCode(int[] courseCode) {
		this.courseCode = courseCode;
	}

	/*public void setFavorite(int[] favorite) {
		this.favorite = favorite;
	}
	
	public void display_course_information() {}
	public void display_favorite_information() {}
	public void enter_rewrite_favorite_course() {}*/
}
