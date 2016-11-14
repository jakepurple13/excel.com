package com.excel.excel.com;

import java.io.IOException;
import java.util.ArrayList;
import java.util.concurrent.atomic.AtomicInteger;

public class Test {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		
		ArrayList<String> als;
		try {
			GoogleSheets gs = new GoogleSheets();

			als = gs.getList();

			for (int i = 0; i < als.size(); i++) {

				if (i % 2 == 0) {
					System.err.println(i + ": " + als.get(i));
				} else {
					System.out.println(i + ": " + als.get(i));
				}

			}

			System.err.println("--------------");

			ArrayList<Student> alStudent = gs.getStudentList();
			
			for(int i=0;i<alStudent.size();i++) {
				System.out.println(alStudent.get(i));
			}
			
			System.err.println("--------------");

			for (Student s : alStudent) {

				System.out.println(s);

			}

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
