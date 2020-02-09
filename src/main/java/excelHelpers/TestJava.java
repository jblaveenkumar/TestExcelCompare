package excelHelpers;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.TreeMap;

public class TestJava {
	static TreeMap tMap;
	public static void main(String[] args) {
		ArrayList list = new ArrayList();

		for (int i = 1; i <= 5; i++) {
			tMap = new TreeMap();
			tMap.put("firstName", "first" + i);
			list.add(tMap);
		}
		System.out.println(list);
	}
}
