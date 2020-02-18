public class Test {
    ExcelReader a = new ExcelReader("d:\\temp\\Book.xlsx", 0);
    void foo () {

        //System.out.println(a.getStringValue(1, 1));
        //System.out.println(a.getStringValue(3, 3));
        //System.out.println(a.getStringValue(5, 6));

        System.out.println(a.getNumericValue(1, 2));
        System.out.println(a.getNumericValue(4, 3));
        System.out.println(a.getNumericValue(5, 6));
    }
}
