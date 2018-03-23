package test;

public class testing {

	public static void main(String[] args) {
		String repeated = new String(new char[10]).replace("\0", "abcd"+"separator");
		String[] split = repeated.split("separator");
		
		for (String s: split) {
			System.out.println(s);
		}

	}

}
