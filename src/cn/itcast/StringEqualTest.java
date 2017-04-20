package cn.itcast;

public class StringEqualTest {
	public static void main(String[] args) {
        String s1 = "Programming";
        String s2 = new String("Programming");
        String s3 = "Program" + "ming";
        System.out.println(s1 == s2);
        System.out.println(s1 == s3);
        System.out.println(s1 == s1.intern());

        System.out.println(s1);
        System.out.println(s2);
        System.out.println(s3);
        System.out.println(s1.intern());
    }
}
