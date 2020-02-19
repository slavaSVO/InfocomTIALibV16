public class Test {
    ExcelReader a = new ExcelReader("d:\\temp\\Book.xlsx", 0);
    void foo () {

        System.out.println(a.getNumericValue(1, 2));
        System.out.println(a.getNumericValue(4, 3));
        System.out.println(a.getNumericValue(5, 6));
    }
}
