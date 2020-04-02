package com.yilnz.excelhandler;

public class Line {
	Integer number;
	String pin;
	String dir;

	@Override
	public String toString() {
		return dir + "," + pin;
	}
}
