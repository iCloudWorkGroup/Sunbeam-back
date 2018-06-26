package test;

public class PastTest {
	public static void main(String[] args) {
           String a = "你好\\t你好\t\n 是的";
           String[] b = a.split("\t");
           System.out.println(b.length);
	}
}
