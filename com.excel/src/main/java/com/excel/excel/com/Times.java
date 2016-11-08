package com.excel.excel.com;
import java.time.LocalTime;

public class Times {

	String dayOfWeek;
	LocalTime timeIn;
	LocalTime timeOut;

	public Times(String day, LocalTime timeIn, LocalTime timeOut) {
		
		dayOfWeek = day;
		this.timeIn = timeIn;
		this.timeOut = timeOut;
		
	}
	
	public String getDay() {
		return dayOfWeek;
	}
	
	
	/*public boolean equals() {
		
	}*/
	
	@Override
	public String toString() {
		return dayOfWeek + " Start: " + timeIn + " End: " + timeOut; 
	}
	

}
