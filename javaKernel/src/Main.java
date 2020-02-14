public class Main {
    public static void main(String[] args) {
        //Read data from excel file.
        try {
            SheetObjList reader = new SheetObjList("d:\\project\\IA_TIALibV16\\proj\\javaKernel\\Kernel.xlsx", 1);
            reader.isTypeEquals("XV", 49);
        } catch (Exception e) {
            //Do something.
        }
    }
}
