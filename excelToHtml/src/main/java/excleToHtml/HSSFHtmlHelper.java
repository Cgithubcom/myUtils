package excleToHtml;

import java.util.Formatter;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
public class HSSFHtmlHelper implements HtmlHelper{
	private final HSSFWorkbook wb;  
	  
	   private static final Map<Integer, HSSFColor> colors = HSSFColor.getIndexHash();  
	   // private static final Map<Integer, HSSFColor> colors = new XSSFColor().;  
	    public HSSFHtmlHelper(HSSFWorkbook wb) {  
	        this.wb = wb;  
	    }  
	  
	    public void colorStyles(CellStyle style, Formatter out) { 
	        HSSFCellStyle cs = (HSSFCellStyle) style;  
	        styleColor(out, "background-color", cs.getFillForegroundColorColor());  
	        
			styleColor(out, "color", cs.getFont(wb).getHSSFColor(wb));//#00b050
			//[255, 204, 153]//FFFF:CCCC:9999FFFF:CCCC:9999
	    }  
	  
	    private void styleColor(Formatter out, String attr, HSSFColor color) {  
	        if (color == null||color.getIndex()==HSSFColor.AUTOMATIC.index)  
	            return;  
	        short[] rgb = color.getTriplet(); 
	        if (rgb == null) { 
	            return; 
	        }
	        out.format("  %s: #%02x%02x%02x;%n", attr, rgb[0], rgb[1], rgb[2]); 
	        //out.format(" %s:#%s;%n",attr, color.getHexString());  
	        // This is done twice -- rgba is new with CSS 3, and browser that don't  
	        // support it will ignore the rgba specification and stick with the  
	        // solid color, which is declared first  
	       // out.format("  %s: #%02x%02x%02x;%n", attr, rgb[0], rgb[1], rgb[2]);  
	/*        out.format("  %s: rgba(0x%02x, 0x%02x, 0x%02x, 0x%02x);%n", attr,color.g 
	                rgb[3], rgb[0], rgb[1], rgb[2] );*/  
	    }  
}
