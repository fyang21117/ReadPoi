package com.example.readpoi;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipException;
import java.util.zip.ZipFile;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Picture;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.hwpf.usermodel.TableCell;
import org.apache.poi.hwpf.usermodel.TableIterator;
import org.apache.poi.hwpf.usermodel.TableRow;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.util.CellRangeAddress;
import org.xmlpull.v1.XmlPullParser;
import org.xmlpull.v1.XmlPullParserException;
import android.graphics.Bitmap;
import android.graphics.BitmapFactory;
import android.util.Log;
import android.util.Xml;

public class FR {
	private String nameStr;
	public Range range = null;
	public HWPFDocument hwpf = null;
	public String htmlPath;
	public String picturePath;
	public String filename;
	public List pictures;
	public TableIterator tableIterator;
	public int presentPicture = 0;
	public int screenWidth;
	public FileOutputStream output;
	public File myFile;
	StringBuffer lsb = new StringBuffer();
	String returnPath = "";
	 static final int BUFFER = 2048;  
	public FR(String namepath) {
		// this.screenWidth =
		// this.getWindowManager().getDefaultDisplay().getWidth() -
		// 10;//���ÿ��Ϊ��Ļ���-10
		this.nameStr = namepath;
		this.filename=getFileName(namepath);
		read();
	}

	public void read() {
		File sdFile = android.os.Environment
				.getExternalStorageDirectory();// ��ȡ��չ�豸���ļ�Ŀ¼
		String path = sdFile.getAbsolutePath() + File.separator
				+ "library";// �õ�sd��(��չ�豸)�ľ���·��+"/"+xiao
		File myFile = new File(path + File.separator +filename+ ".html");// ��ȡmy.html�ĵ�ַ
		if(!myFile.exists()){
		    if (this.nameStr.endsWith(".doc")) {
		    	this.getRange();
		    	this.makeFile();
		    	this.readDOC();			
		     }
		    if (this.nameStr.endsWith(".docx")) {
		    	this.makeFile();
		    	this.readDOCX();
		    }
		    if (this.nameStr.endsWith(".xls")) {
		    	try {
		    		this.makeFile();
		    		this.readXLS();
			    } catch (Exception e) {
			    	// TODO Auto-generated catch block
			    	e.printStackTrace();
		    	}
		    }
		    if (this.nameStr.endsWith(".xlsx")) {
		    	try{
		    	this.makeFile();
		    	this.readXLSX();
		    	}catch (Exception e) {
			    	// TODO Auto-generated catch block
			    	e.printStackTrace();
			    }
		    }		
		}
		returnPath = "file:///" + myFile;
		// this.view.loadUrl("file:///" + this.htmlPath);
		System.out.println("htmlPath" + this.htmlPath);

	}

	/* ��ȡword�е�����д��sdcard�ϵ�.html�ļ��� */
	public void readDOC() {

		try {
			myFile = new File(htmlPath);
			output = new FileOutputStream(myFile);
			presentPicture=0;
			String head = "<html><meta charset=\"utf-8\"><body>";
			String tagBegin = "<p>";
			String tagEnd = "</p>";
			output.write(head.getBytes());
			int numParagraphs = range.numParagraphs();// �õ�ҳ�����еĶ�����
			for (int i = 0; i < numParagraphs; i++) { // ����������
				Paragraph p = range.getParagraph(i); // �õ��ĵ��е�ÿһ������
				if (p.isInTable()) {
					int temp = i;
					if (tableIterator.hasNext()) {
						String tableBegin = "<table style=\"border-collapse:collapse\" border=1 bordercolor=\"black\">";
						String tableEnd = "</table>";
						String rowBegin = "<tr>";
						String rowEnd = "</tr>";
						String colBegin = "<td>";
						String colEnd = "</td>";
						Table table = tableIterator.next();
						output.write(tableBegin.getBytes());
						int rows = table.numRows();
						for (int r = 0; r < rows; r++) {
							output.write(rowBegin.getBytes());
							TableRow row = table.getRow(r);
							int cols = row.numCells();
							int rowNumParagraphs = row.numParagraphs();
							int colsNumParagraphs = 0;
							for (int c = 0; c < cols; c++) {
								output.write(colBegin.getBytes());
								TableCell cell = row.getCell(c);
								int max = temp + cell.numParagraphs();
								colsNumParagraphs = colsNumParagraphs
										+ cell.numParagraphs();
								for (int cp = temp; cp < max; cp++) {
									Paragraph p1 = range.getParagraph(cp);
									output.write(tagBegin.getBytes());
									writeParagraphContent(p1);
									output.write(tagEnd.getBytes());
									temp++;
								}
								output.write(colEnd.getBytes());
							}
							int max1 = temp + rowNumParagraphs;
							for (int m = temp + colsNumParagraphs; m < max1; m++) {
								temp++;
							}
							output.write(rowEnd.getBytes());
						}
						output.write(tableEnd.getBytes());
					}
					i = temp;
				} else {
					output.write(tagBegin.getBytes());
					writeParagraphContent(p);
					output.write(tagEnd.getBytes());
				}
			}
			String end = "</body></html>";
			output.write(end.getBytes());
			output.close();
		} catch (Exception e) {
			
			System.out.println("readAndWrite Exception:"+e.getMessage());
			e.printStackTrace();
		}
	}

	public void readDOCX() {
		String river = "";
		try { 
			this.myFile = new File(this.htmlPath);
			this.output = new FileOutputStream(this.myFile);
			presentPicture=0;
			String head = "<!DOCTYPE><html><meta charset=\"utf-8\"><body>";
			String end = "</body></html>";
			String tagBegin = "<p>";
			String tagEnd = "</p>";
			String tableBegin = "<table style=\"border-collapse:collapse\" border=1 bordercolor=\"black\">";
			String tableEnd = "</table>";
			String rowBegin = "<tr>";
			String rowEnd = "</tr>";
			String colBegin = "<td>";
			String colEnd = "</td>";
			String style = "style=\"";
			this.output.write(head.getBytes());
			ZipFile xlsxFile = new ZipFile(new File(this.nameStr));
			ZipEntry sharedStringXML = xlsxFile.getEntry("word/document.xml");
			InputStream inputStream = xlsxFile.getInputStream(sharedStringXML);
			XmlPullParser xmlParser = Xml.newPullParser();
			xmlParser.setInput(inputStream, "utf-8");
			int evtType = xmlParser.getEventType();
			boolean isTable = false;
			boolean isSize = false;
			boolean isColor = false;
			boolean isCenter = false;
			boolean isRight = false;
			boolean isItalic = false;
			boolean isUnderline = false;
			boolean isBold = false;
			boolean isR = false;
			boolean isStyle = false;
			int pictureIndex = 1;
			while (evtType != XmlPullParser.END_DOCUMENT) {
				switch (evtType) {

				// ��ʼ��ǩ
				case XmlPullParser.START_TAG:
					String tag = xmlParser.getName();

					if (tag.equalsIgnoreCase("r")) {
						isR = true;
					}
					if (tag.equalsIgnoreCase("u")) {
						isUnderline = true;
					}
					if (tag.equalsIgnoreCase("jc")) {
						String align = xmlParser.getAttributeValue(0);
						if (align.equals("center")) {
							this.output.write("<center>".getBytes());
							isCenter = true;
						}
						if (align.equals("right")) {
							this.output.write("<div align=\"right\">"
									.getBytes());
							isRight = true;
						}
					}

					if (tag.equalsIgnoreCase("color")) {

						String color = xmlParser.getAttributeValue(0);

						this.output
								.write(("<span style=\"color:" + color + ";\">")
										.getBytes());
						isColor = true;
					}
					if (tag.equalsIgnoreCase("sz")) {
						if (isR == true) {
							int size = decideSize(Integer.valueOf(xmlParser
									.getAttributeValue(0)));
							this.output.write(("<font size=" + size + ">")
									.getBytes());
							isSize = true;
						}
					}
					if (tag.equalsIgnoreCase("tbl")) {
						this.output.write(tableBegin.getBytes());
						isTable = true;
					}
					if (tag.equalsIgnoreCase("tr")) {
						this.output.write(rowBegin.getBytes());
					}
					if (tag.equalsIgnoreCase("tc")) {
						this.output.write(colBegin.getBytes());
					}

					  if (tag.equalsIgnoreCase("pic")) {
					      String entryName_jpeg = "word/media/image"
					        + pictureIndex + ".jpeg";
					      String entryName_png = "word/media/image"
					        + pictureIndex + ".png";
					      String entryName_gif = "word/media/image"
					        + pictureIndex + ".gif";
					      String entryName_wmf = "word/media/image"
							        + pictureIndex + ".wmf";
					      ZipEntry sharePicture = null;
					      InputStream pictIS = null;
					      sharePicture = xlsxFile.getEntry(entryName_jpeg);

					      if (sharePicture == null) {
					       sharePicture = xlsxFile.getEntry(entryName_png);
					      }
					      if(sharePicture == null){
					       sharePicture = xlsxFile.getEntry(entryName_gif);
					      }
					      if(sharePicture == null){
						       sharePicture = xlsxFile.getEntry(entryName_wmf);
						      }
					      
					      if(sharePicture != null){
					    	  pictIS = xlsxFile.getInputStream(sharePicture);
						      ByteArrayOutputStream pOut = new ByteArrayOutputStream();
						      byte[] bt = null;
						      byte[] b = new byte[1000];
						      int len = 0;
						      while ((len = pictIS.read(b)) != -1) {
						       pOut.write(b, 0, len);
						      }
						      pictIS.close();
						      pOut.close();
						      bt = pOut.toByteArray();
						      Log.i("byteArray", "" + bt);
						      if (pictIS != null)
						       pictIS.close();
						      if (pOut != null)
						       pOut.close();
						      writeDOCXPicture(bt);
					      }
					      
					      pictureIndex++;
					     }

					if (tag.equalsIgnoreCase("b")) {
						isBold = true;
					}
					if (tag.equalsIgnoreCase("p")) {
						if (isTable == false) {
							this.output.write(tagBegin.getBytes());
						}
					}
					if (tag.equalsIgnoreCase("i")) {
						isItalic = true;
					}
					// ��⵽ֵ ��ǩ
					if (tag.equalsIgnoreCase("t")) {
						if (isBold == true) {
							this.output.write("<b>".getBytes());
						}
						if (isUnderline == true) {
							this.output.write("<u>".getBytes());
						}
						if (isItalic == true) {
							output.write("<i>".getBytes());
						}
						river = xmlParser.nextText();
						this.output.write(river.getBytes());
						if (isItalic == true) {
							this.output.write("</i>".getBytes());
							isItalic = false;
						}
						if (isUnderline == true) {
							this.output.write("</u>".getBytes());
							isUnderline = false;
						}
						if (isBold == true) {
							this.output.write("</b>".getBytes());
							isBold = false;
						}
						if (isSize == true) {
							this.output.write("</font>".getBytes());
							isSize = false;
						}
						if (isColor == true) {
							this.output.write("</span>".getBytes());
							isColor = false;
						}
						if (isCenter == true) {
							this.output.write("</center>".getBytes());
							isCenter = false;
						}
						if (isRight == true) {
							this.output.write("</div>".getBytes());
							isRight = false;
						}
					}
					break;
				// ������ǩ
				case XmlPullParser.END_TAG:
					String tag2 = xmlParser.getName();
					if (tag2.equalsIgnoreCase("tbl")) {
						this.output.write(tableEnd.getBytes());
						isTable = false;
					}
					if (tag2.equalsIgnoreCase("tr")) {
						this.output.write(rowEnd.getBytes());
					}
					if (tag2.equalsIgnoreCase("tc")) {
						this.output.write(colEnd.getBytes());
					}
					if (tag2.equalsIgnoreCase("p")) {
						if (isTable == false) {
							this.output.write(tagEnd.getBytes());
						}
					}
					if (tag2.equalsIgnoreCase("r")) {
						isR = false;
					}
					break;
				default:
					break;
				}
				evtType = xmlParser.next();
			}
			this.output.write(end.getBytes());
		} catch (ZipException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (XmlPullParserException e) {
			e.printStackTrace();
		}
		if (river == null) {
			river = "";
		}
	}

	public StringBuffer readXLS() throws Exception {

		myFile = new File(htmlPath);
		output = new FileOutputStream(myFile);
		lsb.append("<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:x='urn:schemas-microsoft-com:office:excel' xmlns='http://www.w3.org/TR/REC-html40'>");
		lsb.append("<head><meta http-equiv=Content-Type content='text/html; charset=utf-8'><meta name=ProgId content=Excel.Sheet>");
		HSSFSheet sheet = null;

		String excelFileName = nameStr;
		try {
			HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(
					excelFileName));

			for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
                sheet = workbook.getSheetAt(sheetIndex);
                String sheetName = workbook.getSheetName(sheetIndex);
                if (workbook.getSheetAt(sheetIndex) != null) {
                    sheet = workbook.getSheetAt(sheetIndex);
                    if (sheet != null) {
                        int firstRowNum = sheet.getFirstRowNum();
                        int lastRowNum = sheet.getLastRowNum();
                        lsb.append("<table width=\"100%\" style=\"border:1px solid #000;border-width:1px 0 0 1px;margin:2px 0 2px 0;border-collapse:collapse;\">");

                        for (int rowNum = firstRowNum; rowNum <= lastRowNum; rowNum++) {
                            if (sheet.getRow(rowNum) != null) {
                                HSSFRow row = sheet.getRow(rowNum);
                                short firstCellNum = row.getFirstCellNum();
                                short lastCellNum = row.getLastCellNum();
                                int height = (int) (row.getHeight() / 15.625);
                                lsb.append("<tr height=\""
                                        + height
                                        + "\" style=\"border:1px solid #000;border-width:0 1px 1px 0;margin:2px 0 2px 0;\">");
                                for (short cellNum = firstCellNum; cellNum <= lastCellNum; cellNum++) {
                                    HSSFCell cell = row.getCell(cellNum);
                                    if (cell != null) {
                                        if (cell.getCellType() == HSSFCell.CELL_TYPE_BLANK) {
                                            continue;
                                        } else {
                                            StringBuffer tdStyle = new StringBuffer(
                                                    "<td style=\"border:1px solid #000; border-width:0 1px 1px 0;margin:2px 0 2px 0; ");
                                            HSSFCellStyle cellStyle = cell
                                                    .getCellStyle();
                                            HSSFPalette palette = workbook
                                                    .getCustomPalette();
                                            HSSFColor hColor = palette
                                                    .getColor(cellStyle
                                                            .getFillForegroundColor());
                                            HSSFColor hColor2 = palette
                                                    .getColor(cellStyle
                                                            .getFont(workbook)
                                                            .getColor());

                                            String bgColor = convertToStardColor(hColor);
                                            short boldWeight = cellStyle
                                                    .getFont(workbook)
                                                    .getBoldweight();
                                            short fontHeight = (short) (cellStyle
                                                    .getFont(workbook)
                                                    .getFontHeight() / 2);
                                            String fontColor = convertToStardColor(hColor2);
                                            if (bgColor != null
                                                    && !"".equals(bgColor
                                                    .trim())) {
                                                tdStyle.append(" background-color:"
                                                        + bgColor + "; ");
                                            }
                                            if (fontColor != null
                                                    && !"".equals(fontColor
                                                    .trim())) {
                                                tdStyle.append(" color:"
                                                        + fontColor + "; ");
                                            }
                                            tdStyle.append(" font-weight:"
                                                    + boldWeight + "; ");
                                            tdStyle.append(" font-size: "
                                                    + fontHeight + "%;");
                                            lsb.append(tdStyle + "\"");

                                            int width = (int) (sheet.getColumnWidth(cellNum) / 35.7);

                                            int cellReginCol = getMergerCellRegionCol(
                                                    sheet, rowNum, cellNum);
                                            int cellReginRow = getMergerCellRegionRow(
                                                    sheet, rowNum, cellNum);
                                            String align = convertVerticalAlignToHtml(cellStyle
                                                    .getAlignment()); //
                                            String vAlign = convertVerticalAlignToHtml(cellStyle
                                                    .getVerticalAlignment());

                                            lsb.append(" align=\"" + align
                                                    + "\" valign=\"" + vAlign
                                                    + "\" width=\"" + width
                                                    + "\" ");

                                            lsb.append(" colspan=\""
                                                    + cellReginCol
                                                    + "\" rowspan=\""
                                                    + cellReginRow + "\"");
                                            lsb.append(">" + getCellValue(cell)
                                                    + "</td>");
                                        }
                                    }
                                }
                                lsb.append("</tr>");
                            }
                        }
                    }

                }

            }
            output.write(lsb.toString().getBytes());
        } catch (FileNotFoundException e) {
            throw new Exception(excelFileName );
        } catch (IOException e) {
            throw new Exception(excelFileName
                    + e.getMessage() + ")!");
        }
        return lsb;
	}

	public void readXLSX() {
		try {
			this.myFile = new File(this.htmlPath);
			this.output = new FileOutputStream(this.myFile);
			String head = "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.01 Transitional//EN\"\"http://www.w3.org/TR/html4/loose.dtd\"><html><meta charset=\"utf-8\"><head></head><body>";// ����ͷ�ļ�,�����������utf-8,��Ȼ���������
			String tableBegin = "<table style=\"border-collapse:collapse\" border=1 bordercolor=\"black\">";
			String tableEnd = "</table>";
			String rowBegin = "<tr>";
			String rowEnd = "</tr>";
			String colBegin = "<td>";
			String colEnd = "</td>";
			String end = "</body></html>";
			this.output.write(head.getBytes());
			this.output.write(tableBegin.getBytes());
			String str = "";
			String v = null;
			boolean flat = false;
			List<String> ls = new ArrayList<String>();
			try {
				ZipFile xlsxFile = new ZipFile(new File(this.nameStr));
				ZipEntry sharedStringXML = xlsxFile
						.getEntry("xl/sharedStrings.xml");
				InputStream inputStream = xlsxFile
						.getInputStream(sharedStringXML);
				XmlPullParser xmlParser = Xml.newPullParser();
				xmlParser.setInput(inputStream, "utf-8");
				int evtType = xmlParser.getEventType();
				while (evtType != XmlPullParser.END_DOCUMENT) {
					switch (evtType) {
					case XmlPullParser.START_TAG:
						String tag = xmlParser.getName();
						if (tag.equalsIgnoreCase("t")) {
							ls.add(xmlParser.nextText());
						}
						break;
					case XmlPullParser.END_TAG:
						break;
					default:
						break;
					}
					evtType = xmlParser.next();
				}
				ZipEntry sheetXML = xlsxFile
						.getEntry("xl/worksheets/sheet1.xml");
				InputStream inputStreamsheet = xlsxFile
						.getInputStream(sheetXML);
				XmlPullParser xmlParsersheet = Xml.newPullParser();
				xmlParsersheet.setInput(inputStreamsheet, "utf-8");
				int evtTypesheet = xmlParsersheet.getEventType();
				this.output.write(rowBegin.getBytes());
				int i = -1;
				while (evtTypesheet != XmlPullParser.END_DOCUMENT) {
					switch (evtTypesheet) {
					case XmlPullParser.START_TAG:
						String tag = xmlParsersheet.getName();
						if (tag.equalsIgnoreCase("row")) {
						} else {
							if (tag.equalsIgnoreCase("c")) {
								String t = xmlParsersheet.getAttributeValue(
										null, "t");
								if (t != null) {
									flat = true;
									System.out.println(flat);
								} else {
									this.output.write(colBegin.getBytes());
									this.output.write(colEnd.getBytes());
									System.out.println(flat );
									flat = false;
								}
							} else {
								if (tag.equalsIgnoreCase("v")) {
									v = xmlParsersheet.nextText();
									this.output.write(colBegin.getBytes());
									if (v != null) {
										if (flat) {
											str = ls.get(Integer.parseInt(v));
										} else {
											str = v;
										}
										this.output.write(str.getBytes());
										this.output.write(colEnd.getBytes());
									}
								}
							}
						}
						break;
					case XmlPullParser.END_TAG:
						if (xmlParsersheet.getName().equalsIgnoreCase("row")
								&& v != null) {
							if (i == 1) {
								this.output.write(rowEnd.getBytes());
								this.output.write(rowBegin.getBytes());
								i = 1;
							} else {
								this.output.write(rowBegin.getBytes());
							}
						}
						break;
					}
					evtTypesheet = xmlParsersheet.next();
				}
				System.out.println(str);
			} catch (ZipException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			} catch (XmlPullParserException e) {
				e.printStackTrace();
			}
			if (str == null) {
				str = "";
			}
			this.output.write(rowEnd.getBytes());
			this.output.write(tableEnd.getBytes());
			this.output.write(end.getBytes());
		} catch (Exception e) {
			System.out.println("readAndWrite Exception");
		}
	}

	/**
	 * ȡ�õ�Ԫ���ֵ
	 * 
	 * @param cell
	 * @return
	 * @throws IOException
	 */
	private static Object getCellValue(HSSFCell cell) throws IOException {
		Object value = "";
		if (cell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
			value = cell.getRichStringCellValue().toString();
		} else if (cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
			if (HSSFDateUtil.isCellDateFormatted(cell)) {
				Date date = (Date) cell.getDateCellValue();
				SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
				value = sdf.format(date);
			} else {
				double value_temp = (double) cell.getNumericCellValue();
				BigDecimal bd = new BigDecimal(value_temp);
				BigDecimal bd1 = bd.setScale(3, bd.ROUND_HALF_UP);
				value = bd1.doubleValue();

				DecimalFormat format = new DecimalFormat("#0.###");
				value = format.format(cell.getNumericCellValue());

			}
		}
		if (cell.getCellType() == HSSFCell.CELL_TYPE_BLANK) {
			value = "";
		}
		return value;
	}


	private static int getMergerCellRegionCol(HSSFSheet sheet, int cellRow,
			int cellCol) throws IOException {
		int retVal = 0;
		int sheetMergerCount = sheet.getNumMergedRegions();
		for (int i = 0; i < sheetMergerCount; i++) {
			CellRangeAddress cra = (CellRangeAddress) sheet.getMergedRegion(i);
			int firstRow = cra.getFirstRow();
			int firstCol = cra.getFirstColumn();
			int lastRow = cra.getLastRow();
			int lastCol = cra.getLastColumn();
			if (cellRow >= firstRow && cellRow <= lastRow) {
				if (cellCol >= firstCol && cellCol <= lastCol) {
					retVal = lastCol - firstCol + 1;
					break;
				}
			}
		}
		return retVal;
	}


	private static int getMergerCellRegionRow(HSSFSheet sheet, int cellRow,
			int cellCol) throws IOException {
		int retVal = 0;
		int sheetMergerCount = sheet.getNumMergedRegions();
		for (int i = 0; i < sheetMergerCount; i++) {
			CellRangeAddress cra = (CellRangeAddress) sheet.getMergedRegion(i);
			int firstRow = cra.getFirstRow();
			int firstCol = cra.getFirstColumn();
			int lastRow = cra.getLastRow();
			int lastCol = cra.getLastColumn();
			if (cellRow >= firstRow && cellRow <= lastRow) {
				if (cellCol >= firstCol && cellCol <= lastCol) {
					retVal = lastRow - firstRow + 1;
					break;
				}
			}
		}
		return 0;
	}

	/**
	 * ��Ԫ�񱳾�ɫת��
	 * 
	 * @param hc
	 * @return
	 */
	private String convertToStardColor(HSSFColor hc) {
		StringBuffer sb = new StringBuffer("");
		if (hc != null) {
			int a = HSSFColor.AUTOMATIC.index;
			int b = hc.getIndex();
			if (a == b) {
				return null;
			}
			sb.append("#");
			for (int i = 0; i < hc.getTriplet().length; i++) {
				String str;
				String str_tmp = Integer.toHexString(hc.getTriplet()[i]);
				if (str_tmp != null && str_tmp.length() < 2) {
					str = "0" + str_tmp;
				} else {
					str = str_tmp;
				}
				sb.append(str);
			}
		}
		return sb.toString();
	}

	/**
	 * ��Ԫ��Сƽ����
	 * 
	 * @param alignment
	 * @return
	 */
	private String convertAlignToHtml(short alignment) {
		String align = "left";
		switch (alignment) {
		case HSSFCellStyle.ALIGN_LEFT:
			align = "left";
			break;
		case HSSFCellStyle.ALIGN_CENTER:
			align = "center";
			break;
		case HSSFCellStyle.ALIGN_RIGHT:
			align = "right";
			break;
		default:
			break;
		}
		return align;
	}

	/**
	 * ��Ԫ��ֱ����
	 * 
	 * @param verticalAlignment
	 * @return
	 */
	private String convertVerticalAlignToHtml(short verticalAlignment) {
		String valign = "middle";
		switch (verticalAlignment) {
		case HSSFCellStyle.VERTICAL_BOTTOM:
			valign = "bottom";
			break;
		case HSSFCellStyle.VERTICAL_CENTER:
			valign = "center";
			break;
		case HSSFCellStyle.VERTICAL_TOP:
			valign = "top";
			break;
		default:
			break;
		}
		return valign;
	}

	public void makeFile() {
		String sdStateString = android.os.Environment.getExternalStorageState();
		if (sdStateString.equals(android.os.Environment.MEDIA_MOUNTED)) {
			try {
				File sdFile = android.os.Environment
						.getExternalStorageDirectory();
				String path = sdFile.getAbsolutePath() + File.separator
						+ "library";
				File dirFile = new File(path);
				if (!dirFile.exists()) {
					dirFile.mkdir();
				}
				File myFile = new File(path + File.separator +filename+ ".html");
				if (!myFile.exists()) {
					myFile.createNewFile();
				}
				this.htmlPath = myFile.getAbsolutePath();
			} catch (Exception e) {
			}
		}
	}

	/* ������sdcard�ϴ���ͼƬ */
	public void makePictureFile() {
		String sdString = android.os.Environment.getExternalStorageState();
		if (sdString.equals(android.os.Environment.MEDIA_MOUNTED))
			try {
				File picFile = android.os.Environment
						.getExternalStorageDirectory();
				String picPath = picFile.getAbsolutePath() + File.separator
						+ "library";
				File picDirFile = new File(picPath);
				if (!picDirFile.exists()) {
					picDirFile.mkdir();
				}
				File pictureFile = new File(picPath + File.separator
						+getFileName(nameStr)+ presentPicture + ".jpg");
				if (!pictureFile.exists()) {
					pictureFile.createNewFile();
				}
				this.picturePath = pictureFile.getAbsolutePath();
			} catch (Exception e) {
				System.out.println("PictureFile Catch Exception");
			}
	}
	
	public String getFileName(String pathandname){  
        
        int start=pathandname.lastIndexOf("/");  
        int end=pathandname.lastIndexOf(".");  
        if(start!=-1 && end!=-1){  
            return pathandname.substring(start+1,end);    
        }else{  
            return null;  
        }  
          
    }  

	public void writePicture() {
		Picture picture = (Picture) pictures.get(presentPicture);

		byte[] pictureBytes = picture.getContent();

		Bitmap bitmap = BitmapFactory.decodeByteArray(pictureBytes, 0,
				pictureBytes.length);

		makePictureFile();
		presentPicture++;

		File myPicture = new File(picturePath);

		try {

			FileOutputStream outputPicture = new FileOutputStream(myPicture);

			outputPicture.write(pictureBytes);

			outputPicture.close();
		} catch (Exception e) {
			System.out.println("outputPicture Exception");
		}

		String imageString = "<img src=\"" + picturePath + "\"";
		imageString = imageString + ">";

		try {
			output.write(imageString.getBytes());
		} catch (Exception e) {
			System.out.println("output Exception");
		}
	}

	public int decideSize(int size) {

		if (size >= 1 && size <= 8) {
			return 1;
		}
		if (size >= 9 && size <= 11) {
			return 2;
		}
		if (size >= 12 && size <= 14) {
			return 3;
		}
		if (size >= 15 && size <= 19) {
			return 4;
		}
		if (size >= 20 && size <= 29) {
			return 5;
		}
		if (size >= 30 && size <= 39) {
			return 6;
		}
		if (size >= 40) {
			return 7;
		}
		return 3;
	}

	private String decideColor(int a) {
		int color = a;
		switch (color) {
		case 1:
			return "#000000";
		case 2:
			return "#0000FF";
		case 3:
		case 4:
			return "#00FF00";
		case 5:
		case 6:
			return "#FF0000";
		case 7:
			return "#FFFF00";
		case 8:
			return "#FFFFFF";
		case 9:
			return "#CCCCCC";
		case 10:
		case 11:
			return "#00FF00";
		case 12:
			return "#080808";
		case 13:
		case 14:
			return "#FFFF00";
		case 15:
			return "#CCCCCC";
		case 16:
			return "#080808";
		default:
			return "#000000";
		}
	}

	private void getRange() {
		FileInputStream in = null;
		POIFSFileSystem pfs = null;

		try {
			in = new FileInputStream(nameStr);
			pfs = new POIFSFileSystem(in);
			hwpf = new HWPFDocument(pfs);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		range = hwpf.getRange();

		pictures = hwpf.getPicturesTable().getAllPictures();

		tableIterator = new TableIterator(range);

	}

	public void writeDOCXPicture(byte[] pictureBytes) {
		Bitmap bitmap = BitmapFactory.decodeByteArray(pictureBytes, 0,
				pictureBytes.length);
		makePictureFile();
		this.presentPicture++;
		File myPicture = new File(this.picturePath);
		try {
			FileOutputStream outputPicture = new FileOutputStream(myPicture);
			outputPicture.write(pictureBytes);
			outputPicture.close();
		} catch (Exception e) {
			System.out.println("outputPicture Exception");
		}
		String imageString = "<img src=\"" + this.picturePath + "\"";
		
		imageString = imageString + ">";
		try {
			this.output.write(imageString.getBytes());
		} catch (Exception e) {
			System.out.println("output Exception");
		}
	}

	public void writeParagraphContent(Paragraph paragraph) {
		Paragraph p = paragraph;
		int pnumCharacterRuns = p.numCharacterRuns();

		for (int j = 0; j < pnumCharacterRuns; j++) {

			CharacterRun run = p.getCharacterRun(j);

			if (run.getPicOffset() == 0 || run.getPicOffset() >= 1000) {
				if (presentPicture < pictures.size()) {
					writePicture();
				}
			} else {
				try {
					String text = run.text();
					if (text.length() >= 2 && pnumCharacterRuns < 2) {
						output.write(text.getBytes());
					} else {
						int size = run.getFontSize();
						int color = run.getColor();
						String fontSizeBegin = "<font size=\""
								+ decideSize(size) + "\">";
						String fontColorBegin = "<font color=\""
								+ decideColor(color) + "\">";
						String fontEnd = "</font>";
						String boldBegin = "<b>";
						String boldEnd = "</b>";
						String islaBegin = "<i>";
						String islaEnd = "</i>";

						output.write(fontSizeBegin.getBytes());
						output.write(fontColorBegin.getBytes());

						if (run.isBold()) {
							output.write(boldBegin.getBytes());
						}
						if (run.isItalic()) {
							output.write(islaBegin.getBytes());
						}

						output.write(text.getBytes());

						if (run.isBold()) {
							output.write(boldEnd.getBytes());
						}
						if (run.isItalic()) {
							output.write(islaEnd.getBytes());
						}
						output.write(fontEnd.getBytes());
						output.write(fontEnd.getBytes());
					}
				} catch (Exception e) {
					System.out.println("Write File Exception");
				}
			}
		}
	}
}
	


