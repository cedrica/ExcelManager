import org.apache.poi.hssf.util.HSSFColor;

import excelmanager.annotations.XLS;
import excelmanager.annotations.XlsAdditionalInformation;
import excelmanager.annotations.XlsColumn;
import excelmanager.annotations.XlsStyler;
import excelmanager.enums.Location;
import excelmanager.enums.Orientation;

@XLS(sheetsname="test", xlsAdditionalInformation = @XlsAdditionalInformation(text="je suis une info additionnelle", location=Location.TOP, colspan=4))
public class TestObject {

	@XlsColumn(customname="ID")
	private int id;
	@XlsColumn(styler = @XlsStyler(bgColor=HSSFColor.GREEN.index, isBold=true))
	private String	vorname;
	@XlsColumn(styler = @XlsStyler(bgColor=HSSFColor.BLUE.index, isBold=true, orientation=Orientation.OBLIC))
	private String	nachname;
	@XlsColumn(styler = @XlsStyler(bgColor=HSSFColor.RED.index, isBold=true))
	private int		alt;
	@XlsColumn(styler = @XlsStyler(bgColor=HSSFColor.YELLOW.index, isBold=true))
	private String	addresse;
	
	public TestObject(int id , String vorname, String nachname, int alt, String addresse){
		this.id = id;
		this.nachname = nachname;
		this.vorname = vorname;
		this.addresse = addresse;
		this.alt = alt;
	}
	
	public int getId() {
		return id;
	}

	
	public void setId(int id) {
		this.id = id;
	}

	public String getVorname() {
		return vorname;
	}
	
	public void setVorname(String vorname) {
		this.vorname = vorname;
	}
	
	public String getNachname() {
		return nachname;
	}
	
	public void setNachname(String nachname) {
		this.nachname = nachname;
	}
	
	public int getAlt() {
		return alt;
	}
	
	public void setAlt(int alt) {
		this.alt = alt;
	}
	
	public String getAddresse() {
		return addresse;
	}
	
	public void setAddresse(String addresse) {
		this.addresse = addresse;
	}
	
	
}
