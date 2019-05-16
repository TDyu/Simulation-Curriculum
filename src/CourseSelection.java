import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.xmlbeans.impl.piccolo.io.FileFormatException;

public class CourseSelection {

	public static void main(String[] arg) throws FileNotFoundException, FileFormatException, IOException {
		Course course = new Course();
		SimulationCurriculum test = new SimulationCurriculum("IECS", "D0000000", 4); // ("科系縮寫名", "學號", "目前哪個學期(1 2 3 4 5 6 7 8)")
		
		String[] chosenCourse = test.Switch();
		course.setCourseName(chosenCourse); // 模擬所選放到Course 以couse.setCourseName進入但其實courseName和courseCode都設定了
		
		// 測試
		int[] code = course.getCourseCode();
		String[] name = course.getCourseName();
		System.out.println(code[0]+name[0]);
	}

}
