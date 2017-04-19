package excleToHtml;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class TohtmlTest {
    public static void main(String[] args) throws IOException {  
    	//！！以GBK输出  
    	File f=new File("D:\\out1.html");
    	FileOutputStream out=new FileOutputStream(f);
    	ExcelToHtml toHtml = ExcelToHtml
    			.create("D:\\3232"//"D:\\test.xls"//xlsx
    					, out,"UTF-8");  
         toHtml.setCompleteHTML(true);
         toHtml.setTitel("3232.xls");
         toHtml.printPage();
         System.out.println("end");
  
    }
}
