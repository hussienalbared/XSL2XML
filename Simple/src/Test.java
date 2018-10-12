
import java.awt.Choice;
import java.awt.Frame;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.math.BigInteger;
import java.net.URL;
import java.nio.file.FileSystems;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.security.CodeSource;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Optional;
import java.util.Scanner;

import javax.swing.JFileChooser;
import javax.swing.filechooser.FileSystemView;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.commons.codec.digest.DigestUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class Test {

	public static void main(String[] args) throws IOException {
		 	     

		       
		    	   
		File source = getSourceFolder();
		File dest = getDestFolder();
		ArrayList<ExcelProperties> excelProperties=new ArrayList<ExcelProperties>();
		if (source != null && dest != null) {

			File[] listOfFiles = source.listFiles();
			for (int ii = 0; ii < listOfFiles.length; ii++) {
				
				
			if(listOfFiles[ii].isFile()&&listOfFiles[ii].getName().equals("properties.txt"))
			{
				excelProperties=getProperties(listOfFiles[ii]);
				break;
			}
				
			}
			
			
			
			for (int i = 0; i < listOfFiles.length; i++) {
				if (listOfFiles[i].isFile() && isSupportedExtension(listOfFiles[i])) {
					System.out.println("File " + listOfFiles[i].getName());
					try {
						convertFile(listOfFiles[i], dest.getAbsolutePath(),excelProperties);
					} catch (InvalidFormatException e) {

						e.printStackTrace();
					} catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					} catch (ParserConfigurationException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					} catch (TransformerException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				}
			}
	}
	}

	private static boolean isSupportedExtension(File file) {
		String extension = FilenameUtils.getExtension(file.getName());
		
		return extension.equalsIgnoreCase("xls") || extension.equalsIgnoreCase("xlsx");
	}

	public static String convertFile(File xlsFileName, String destination,ArrayList<ExcelProperties> properties)
			throws IOException, InvalidFormatException, ParserConfigurationException, TransformerException {
		SimpleDateFormat format = new SimpleDateFormat("yyyyMM");
		Calendar cal = Calendar.getInstance();
		
		cal.add(Calendar.MONTH, -1);
		String dateString=format.format(cal.getTime());

		
		// Creating a Workbook from an Excel file (.xls or .xlsx)
		Workbook workbook = WorkbookFactory.create(xlsFileName);


		// 1. You can obtain a sheetIterator and iterate over it
		Iterator<Sheet> sheetIterator = workbook.sheetIterator();
		while (sheetIterator.hasNext()) {
			Sheet sheet = sheetIterator.next();
			System.out.println("=> " + sheet.getSheetName());
		}
		List<String> headers = new ArrayList<String>();
		// 2. Or you can use a for-each loop

		for (int ii = 0; ii < workbook.getNumberOfSheets(); ii++) {
			headers.clear();
			DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
			DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
			Document doc = dBuilder.newDocument();
			Element rootElement = doc.createElement("document");
			doc.appendChild(rootElement);
			Element headerElement = doc.createElement("header");
			Element bankCode=doc.createElement("bankCode");
			headerElement.appendChild(bankCode);
			bankCode.appendChild(doc.createTextNode("4200"));
			rootElement.appendChild(headerElement);
			
			Sheet sheet = workbook.getSheetAt(ii);
			Element bodyElement = doc.createElement(sheet.getSheetName());
			rootElement.appendChild(bodyElement);


			// Create a DataFormatter to format and get each cell's value as String
			DataFormatter dataFormatter = new DataFormatter();

			// 1. You can obtain a rowIterator and columnIterator and iterate over them
			Iterator<Row> rowIterator = sheet.rowIterator();
			int i = 0;
			while (rowIterator.hasNext()) {

				Row row = rowIterator.next();

				// Now let's iterate over the columns of the current row
				Iterator<Cell> cellIterator = row.cellIterator();
				ArrayList<String> rowData = new ArrayList<String>();
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					String cellValue = dataFormatter.formatCellValue(cell);
					rowData.add(cellValue);

				}
				if (i == 0) {
					headers.addAll(rowData);
				} else {
					String rowName="row";
					String sheetName=sheet.getSheetName();
					if(sheetName.charAt(sheetName.length()-1)=='s'||sheetName.charAt(sheetName.length()-1)=='S') {
						rowName=sheetName.substring(0, sheetName.length()-1);
					}
					
					
					Element rowelment = doc.createElement(rowName);
					bodyElement.appendChild(rowelment);
					for (int col = 0; col < headers.size(); col++) {

						String header = headers.get(col);
						Object value = null;

						if (col < rowData.size()) {

							value = rowData.get(col);

						} else {
							value = "";
						}
						
						ExcelProperties columnProperties=getColumnInfo(header, properties);
						if(columnProperties==null)
						{
							Element cellelment = doc.createElement(header);
							rowelment.appendChild(cellelment);
							cellelment.appendChild(doc.createTextNode(value.toString()));

						}
						else {
							Element cellelment = doc.createElement(columnProperties.getTagName());
							rowelment.appendChild(cellelment);
							if(columnProperties.isHashed()&&(!value.toString().isEmpty()||value.toString().trim().length()>0))
							{
								String hashed = getMd5(value.toString());

								cellelment.appendChild(doc.createTextNode(hashed));

							}
							else
							{
								cellelment.appendChild(doc.createTextNode(value.toString()));
							}
							

						}


					}
				}
				i++;
			}
		Element rowsNo = doc.createElement("noOf"+sheet.getSheetName());
			Element month = doc.createElement("month");
			headerElement.appendChild(month);
			headerElement.appendChild(rowsNo);

			rowsNo.appendChild(doc.createTextNode(String.valueOf(i - 1)));
			month.appendChild(doc.createTextNode(dateString));
			Transformer transformer = TransformerFactory.newInstance().newTransformer();
			transformer.setOutputProperty(OutputKeys.INDENT, "yes");
			transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "2");
			// initialize StreamResult with File object to save to file
			String fileNameWithoutExtension = "";
			String pathXmlFile = "";

			fileNameWithoutExtension = xlsFileName.getName().substring(0, xlsFileName.getName().lastIndexOf("."));
			pathXmlFile = fileNameWithoutExtension + "_" + sheet.getSheetName() + ".xml";
			Path fileStorageLocation = Paths.get(destination).toAbsolutePath().normalize();
			Path xmlPath = fileStorageLocation.resolve(pathXmlFile).normalize();

			StreamResult result = new StreamResult(new File(xmlPath.toString()));
			DOMSource source = new DOMSource(doc);
			transformer.transform(source, result);
			System.out
					.println("File " + xlsFileName.getName() + " is saved successfully to path " + xmlPath.toString());
		}
		// Closing the workbook
		workbook.close();
		String xmlString = "result.getWriter().toString()";

		return xmlString;
	}

	public static String getMd5(String input) {

		// Static getInstance method is called with hashing MD5
		MessageDigest md;
		try {
			md = MessageDigest.getInstance("MD5");

			// digest() method is called to calculate message digest
			// of an input digest() return array of byte
			byte[] messageDigest = md.digest(input.getBytes());

			// Convert byte array into signum representation
			BigInteger no = new BigInteger(1, messageDigest);

			// Convert message digest into hex value
			String hashtext = no.toString(16);
			while (hashtext.length() < 32) {
				hashtext = "0" + hashtext;
			}
			return hashtext;
		} catch (NoSuchAlgorithmException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return input;

	}

	private static File getSourceFolder() {
		JFileChooser chooser = new JFileChooser();
		// chooser.setCurrentDirectory(new java.io.File("."));
		chooser.setDialogTitle("Select source");
		chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

		// disable the "All files" option.
		//
		chooser.setAcceptAllFileFilterUsed(false);
		//
		if (chooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
			return chooser.getSelectedFile();

		}
		return null;

	}

	private static File getDestFolder() {
		JFileChooser chooser = new JFileChooser();
		// chooser.setCurrentDirectory(new java.io.File("."));
		chooser.setDialogTitle("Select destination");
		chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

		// disable the "All files" option.
		//
		chooser.setAcceptAllFileFilterUsed(false);
		//
		if (chooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {

			return chooser.getSelectedFile();

		}
		return null;

	}
	
//	private static File getPropertiesFolder() {
//		JFileChooser jfc = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
//
//		int returnValue = jfc.showOpenDialog(null);
//		// int returnValue = jfc.showSaveDialog(null);
//
//		if (returnValue == JFileChooser.APPROVE_OPTION) {
//			File selectedFile = jfc.getSelectedFile();
//			System.out.println(selectedFile.getAbsolutePath());
//		}
//	}
	public static ExcelProperties getColumnInfo(String columnName,ArrayList<ExcelProperties> excelProperties) {
		Optional<ExcelProperties> columnInfo=excelProperties.stream().filter(a ->a.getColumnName().equals(columnName)).findFirst();
		if(columnInfo.isPresent())
		{
			return columnInfo.get();
		}
		return null;
	}
	public static ArrayList<ExcelProperties> getProperties(File file) throws IOException {
	       
//	       
//	       URL jarLocationUrl = Test.class.getProtectionDomain().getCodeSource().getLocation();
//	       System.out.println((jarLocationUrl.getFile()));
	       
	     
	       
	       BufferedReader br = new BufferedReader(new FileReader(file)); 
	       ArrayList<ExcelProperties> excelProperties=new ArrayList<>();
	       String st; 
	       br.readLine();
	       while ((st = br.readLine()) != null) 
	        
	       {
	    	   String []row=st.trim().split(",");
	    	  if(row.length==3)
	    	  {
	    		  ExcelProperties p=new ExcelProperties();
	    		  p.setColumnName(row[0]);
	    		  p.setTagName(row[1]);
	    		  p.setHashed(Boolean.valueOf(row[2]));
	    		  excelProperties.add(p);
	    	  }
	       }
	    	  
	     return excelProperties;
	}
}
