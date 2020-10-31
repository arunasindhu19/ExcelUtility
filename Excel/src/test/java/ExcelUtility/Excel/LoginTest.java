package ExcelUtility.Excel;

import java.io.IOException;
import java.util.ArrayList;

public class LoginTest {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		ExcelTest data = new ExcelTest();
		ArrayList<String> dat = data.getData("Aruna");
		System.out.println(dat);

	}

}
