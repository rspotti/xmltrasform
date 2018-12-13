package xmltrasform;

import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamConstants;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;

/**
 * 
 * @author roberto
 * Class use StaxXML processing.
 */
public class XLS2CSV2 {
    private static String fieldSep="|";
    private static String rowSep="\r\n"; 
    private static boolean useQuotes=false;
    private static XMLInputFactory xmlInputFactory=null;
    private static FileWriter out;
    private static XMLStreamReader xmlStreamReader=null;
    
	
    private static XMLStreamReader StaxXMLStreamOpen(String fileName, String OutFile) throws IOException, XMLStreamException {
    	if (xmlInputFactory==null )
    		xmlInputFactory= XMLInputFactory.newInstance();
    	out = new FileWriter(OutFile);
        return xmlInputFactory.createXMLStreamReader(new FileInputStream(fileName));
    }
    
    private static void StaxXMLStreamClose() throws XMLStreamException, IOException {
    	if ( null != out ) {
    		out.close();
    		out=null;
    	}
    	if ( null != xmlStreamReader ) {
    		xmlStreamReader.close();
    		xmlStreamReader=null;
    	}
    }
    
    private static void writeQuotes(String testo) throws IOException {
    	if (useQuotes || testo.contains("\"") ||  testo.contains(fieldSep)) {
			out.write("\""+testo.replaceAll("\"", "\"\"")+"\"");
		} else {
			out.write(testo);
		}
    }
    
	private static void processRow() throws XMLStreamException, IOException {
		boolean uscita = false;
		int event;
		boolean first = true;
		while (!uscita && xmlStreamReader.hasNext()) {
			event = xmlStreamReader.next();
			switch (event) {
			case XMLStreamConstants.START_ELEMENT:
				if (xmlStreamReader.getLocalName().equals("Cell")) {
					if (first) {
						first = false;
					} else {
						out.write(fieldSep);
					}
					processCell();
				} else {
					throw new XMLStreamException("ERRORE atteso Cell trovato" + xmlStreamReader.getLocalName());
				}
				break;

			case XMLStreamConstants.END_ELEMENT:
				if (xmlStreamReader.getLocalName().equals("Row")) {
					uscita = true;
				} else 	if (!xmlStreamReader.getLocalName().equals("Cell")) {
					throw new XMLStreamException("ERRORE atteso Cell trovato" + xmlStreamReader.getLocalName());
				}
				break;
			}
		}
	}

	private static void processCell() throws XMLStreamException, IOException {
		boolean uscita = false;
		boolean isdate = false;
		DateTimeFormatter aFormatter = DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm"); // 22/11/2018 10:40

		int event;
		String testo="";
		while (!uscita && xmlStreamReader.hasNext()) {
			event = xmlStreamReader.next();
			switch (event) {
			case XMLStreamConstants.START_ELEMENT:
				if (!xmlStreamReader.getLocalName().equals("Data")) {
					throw new XMLStreamException("ERRORE" + xmlStreamReader.getLocalName());
				}
				isdate = xmlStreamReader.getAttributeValue(0).equalsIgnoreCase("DateTime");
				break;
			case XMLStreamReader.CHARACTERS:
			case XMLStreamReader.CDATA:
				if (isdate) {
					LocalDateTime dt = LocalDateTime.parse(xmlStreamReader.getText());
					testo += dt.format(aFormatter);
				} else {
					testo += xmlStreamReader.getText(); 
				}
				break;
			case XMLStreamConstants.END_ELEMENT:
				uscita = true;
				if (!xmlStreamReader.getLocalName().equals("Data")) {
					throw new XMLStreamException("ERRORE" + xmlStreamReader.getLocalName());
				}
				break;
			}
		}
		writeQuotes(testo);
//		if (testo.contains("\"") ||  testo.contains(fieldSep)) {
//			out.write("\""+testo.replaceAll("\"", "\"\"")+"\"");
//		} else {
//			out.write(testo);
//		}
			
	}

    /**
	 * Set the default field separator as the supplied parameter field 
	 *  if not specified it uses separator is '|'.
	 * @param field
	 */
	public static void setFieldSeparator(char field) {
		fieldSep=""+field;
	}
	
	/**
	 * Set true to surround text field with quotes. 
	 * 
	 * @param useQuotes
	 */
	public static void quoteText(boolean useQuotes) {
		XLS2CSV2.useQuotes=useQuotes;
	}
	
	/**
	 * 
	 * @param xlsFile
	 * @param csvFile
	 * @param field
	 * 
	 * @return true if succeed, false otherwise.
	 */
	public static boolean transform(String xlsFile, String csvFile,char field) {
		setFieldSeparator(field);
		return transform(xlsFile, csvFile);
	}
	
	/**
	 * 
	 * @param xlsFile
	 * @param csvFile
	 * @return
	 */
    public static boolean transform(String xlsFile, String csvFile) { 	
    	try {
    		int event;
    		xmlStreamReader=StaxXMLStreamOpen(xlsFile, csvFile);
			while ( xmlStreamReader.hasNext() ) {
				event = xmlStreamReader.next();// xmlStreamReader.getEventType();				
				if (XMLStreamConstants.START_ELEMENT == event && 
                    xmlStreamReader.getLocalName().equals("Row")) { 
					processRow();
					out.write(rowSep);
				}
			}
			StaxXMLStreamClose();
		} catch (XMLStreamException | IOException e) {
			e.printStackTrace();
			return false;
		}
    	return true;
    }
}
