package com.excel.excel.com;

import java.io.IOException;
import java.util.ArrayList;

public class Test {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		ArrayList<String> als;
		try {
			GoogleSheets gs = new GoogleSheets();
			
			als = gs.getList();
			
			for(int i=0;i<als.size();i++) {
				
				if(i%2==0) {
					System.err.println(i + ": " + als.get(i));
				} else {
					System.out.println(i + ": " + als.get(i));
				}
				
			}
			
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}

}
