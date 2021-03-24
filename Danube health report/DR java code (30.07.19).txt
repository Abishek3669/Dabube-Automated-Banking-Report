package dr_health_check;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.sql.Types;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.jar.Attributes.Name;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.imageio.stream.FileImageInputStream;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.LineSpacingRule;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.TextAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPicture;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.drawingml.x2006.main.CTGraphicalObject;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTAnchor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import static org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody.type;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHeight;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSpacing;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTString;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTabStop;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTrPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVerticalJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STLineSpacingRule;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STPageOrientation;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTabJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;

public class DR {
    
    private static String filePath = "DR_images";
    
    public static void main(String[] args) throws Exception {
        XWPFDocument document = new XWPFDocument();
        String ffPath = null;
        ffPath = "C:/DR_REPORT/DR_HEALTH_CHECK/" + filePath;
        File file = new File(ffPath + "/Capture1.png");
        File file1 = new File(ffPath + "/pro.png");
        File file2 = new File(ffPath + "/pro2.png");
        InputStream in = new FileInputStream(file);
        InputStream in1 = new FileInputStream(file1);
        InputStream in2 = new FileInputStream(file2);
        
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        
        CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();
        XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(document, sectPr);
        XWPFHeader header = headerFooterPolicy.createHeader(XWPFHeaderFooterPolicy.DEFAULT);
        paragraph = header.getParagraphArray(0);
        paragraph.setAlignment(ParagraphAlignment.LEFT);
        run.addTab();
        run = paragraph.createRun();
        String imgFile = "file";
        XWPFPicture picture = run.addPicture(in, XWPFDocument.PICTURE_TYPE_PNG, imgFile, Units.toEMU(500), Units.toEMU(70));
        String blipID = "";
        for (XWPFPictureData picturedata : header.getAllPackagePictures()) {
            blipID = header.getRelationId(picturedata);
        }
        picture.getCTPicture().getBlipFill().getBlip().setEmbed(blipID);
        run = paragraph.createRun();
        paragraph.setAlignment(ParagraphAlignment.RIGHT);
        run.addTab();
        CTTabStop tabStop = paragraph.getCTP().getPPr().addNewTabs().addNewTab();
        tabStop.setVal(STTabJc.RIGHT);
        in.close();
        XWPFFooter footer = headerFooterPolicy.createFooter(XWPFHeaderFooterPolicy.DEFAULT);
        paragraph = footer.getParagraphArray(0);
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        run = paragraph.createRun();
        run.setText("3i-Infotech Ltd.                                                               page |");
        paragraph.getCTP().addNewFldSimple().setInstr("PAGE +1");
        
        XWPFParagraph paragraph1 = document.createParagraph();
        XWPFRun run1 = paragraph1.createRun();
        run1.addBreak();
        run1.addBreak();
        run1.addBreak();
        run1.addBreak();
        run1.addBreak();
        run1.addBreak();
        run1.addBreak();
        run1.addBreak();
        run1.addBreak();
        run1.addBreak();
        run1.addBreak();
        run1.addPicture(in1, XWPFDocument.PICTURE_TYPE_JPEG, "file1", Units.toEMU(460), Units.toEMU(60)); // 200x200 pixels

        run1.addBreak();
        run1.addBreak();
        run1.addBreak();
        run1.addBreak();
        run1.addBreak();
        run1.addBreak();
        run1.addBreak();
        run1.addBreak();
        run1.addBreak();
        run1.addBreak();
        run1.addBreak();
        run1.addPicture(in2, XWPFDocument.PICTURE_TYPE_JPEG, "file2", Units.toEMU(460), Units.toEMU(60));
        in1.close();
        in2.close();
        
        String date1 = new SimpleDateFormat("ddMMyyyy").format(new Date());
        FileOutputStream out = new FileOutputStream("C:/DR_REPORT/DR_HEALTH_CHECK/DR_Health_Check_Report " + date1 + " .doc");
        
        XWPFParagraph para = document.createParagraph();
        XWPFRun ru = para.createRun();
        ru.addBreak(BreakType.PAGE);
        XWPFParagraph paragraph3 = document.createParagraph();
        XWPFRun run3 = paragraph3.createRun();
        run3.addBreak();
        run3.addBreak();
        run3.addBreak();
        run3.addBreak();
        run3.setFontSize(13);
        
        run3.setText("                       TABLE OF CONTENTS");
        run3.setBold(true);
        run3.addBreak();
        run3.addBreak();
        run3.addBreak();
        run3.addBreak();
        run3 = paragraph3.createRun();
        run3.setText("Report Details----------------------------------------------------------------------------------------------------- 3 ");
        run3.addBreak();
        run3.addBreak();
        run3 = paragraph3.createRun();
        run3.setText("Customer Details------------------------------------------------------------------------------------------------- 3 ");
        run3.addBreak();
        run3.addBreak();
        run3.setText("Report Overview------------------------------------------------------------------------------------------------- 4");
        run3.addBreak();
        run3.addBreak();
        run3.setText("Hardware and OS information----------------------------------------------------------------------------------4");
        run3.addBreak();
        run3.addBreak();
        run3.setText("Changes Observed------------------------------------------------------------------------------------------------4 ");
        run3.addBreak();
        run3.addBreak();
        run3.setText("Database Overview-----------------------------------------------------------------------------------------------4");
        run3.addBreak();
        run3.addBreak();
        run3.setText("Database Information ------------------------------------------------------------------------------------------- 4 ");
        run3.addBreak();
        run3.addBreak();
        run3.setText("Database parameters ---------------------------------------------------------------------------------------------5");
        run3.addBreak();
        run3.addBreak();
        run3.setText("DR sync information -------------------------------------------------------------------------------------------- 6");
        run3.addBreak();
        run3.addBreak();
        run3.setText("Archive log destination details ---------------------------------------------------------------------------------6 ");
        run3.addBreak();
        run3.addBreak();
        run3.setText("Alert log analysis------------------------------------------------------------------------------------------------- 6");
        run3.addBreak();
        run3.addBreak();
        run3.setText("Deleting/Moving of files and folder information-------------------------------------------------------------7");
        run3.addBreak();
        run3.addBreak();
        run3.setText("DR Maintenance Details---------------------------------------------------------------------------------------- 7 ");
        run3.addBreak();
        run3.addBreak();
        XWPFParagraph para1 = document.createParagraph();
        XWPFRun ru1 = para1.createRun();
        ru1.addBreak(BreakType.PAGE);
        ru1.addBreak();
        ru1.addBreak();
        ru1.addBreak();
        ru1.addBreak();
        ru1.addBreak();
        ru1.addBreak();
        
        int aRows = 0;
        int aCols = 0;
        XWPFTable tablea = document.createTable(aRows, aCols);
        tablea.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(9500));
        tablea.getCTTbl().getTblPr().unsetTblBorders();
        List<XWPFTableRow> rowsa = tablea.getRows();
        int rowCta = 0;
        int colCta = 0;
        for (XWPFTableRow rowa : rowsa) {
            CTTrPr trPra = rowa.getCtRow().addNewTrPr();
            CTHeight hta = trPra.addNewTrHeight();
            hta.setVal(BigInteger.valueOf(453));
            List<XWPFTableCell> cellsa = rowa.getTableCells();
            for (XWPFTableCell cella : cellsa) {
                CTTcPr tcpra = cella.getCTTc().addNewTcPr();
                CTVerticalJc vaa = tcpra.addNewVAlign();
                vaa.setVal(STVerticalJc.CENTER);
                CTShd ctshda = tcpra.addNewShd();
                ctshda.setVal(STShd.CLEAR);
                ctshda.setColor("FFFFFF");
                if (rowCta == 0) {
                    ctshda.setFill("FFFFFF");
                } else {
                    ctshda.setFill("FFFFFF");
                }
                XWPFParagraph paraa = cella.getParagraphs().get(0);
                XWPFRun rha = paraa.createRun();
                rha.setFontSize(10);
                rha.setBold(true);
                rha.setColor("99284C");
                rha.setText("Report Details");
                
                int bRows = 1;
                int bCols = 0;
                XWPFTable tableb = document.createTable(bRows, bCols);
                tableb.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(9500));
                tableb.getCTTbl().getTblPr().unsetTblBorders();
                List<XWPFTableRow> rowsb = tableb.getRows();
                int rowCtb = 0;
                for (XWPFTableRow rowb : rowsb) {
                    CTTrPr trPrb = rowb.getCtRow().addNewTrPr();
                    CTHeight htb = trPrb.addNewTrHeight();
                    htb.setVal(BigInteger.valueOf(300));
                    List<XWPFTableCell> cellsb = rowb.getTableCells();
                    for (XWPFTableCell cellb : cellsb) {
                        CTTcPr tcprb = cellb.getCTTc().addNewTcPr();
                        CTVerticalJc vab = tcprb.addNewVAlign();
                        vab.setVal(STVerticalJc.CENTER);
                        CTShd ctshdb = tcprb.addNewShd();
                        ctshdb.setColor("auto");
                        ctshdb.setVal(STShd.CLEAR);
                        if (rowCtb == 0) {
                            ctshdb.setFill("FFFFFF");
                            
                        } else {
                            ctshdb.setFill("FFFFFF");
                        }
                        tableb.getRow(0).getCell(0).setText("");
                    }
                }
                
                int cRows = 2;
                int cCols = 6;
                XWPFTable tablec = document.createTable(cRows, cCols);
                tablec.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(1500));
                tablec.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(2000));
                tablec.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(1500));
                tablec.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(1500));
                tablec.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(1500));
                tablec.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(1500));
                
                List<XWPFTableRow> rowsc = tablec.getRows();
                int rowCtc = 0;
                int colCtc = 0;
                for (XWPFTableRow rowc : rowsc) {
                    CTTrPr trPrc = rowc.getCtRow().addNewTrPr();
                    CTHeight htc = trPrc.addNewTrHeight();
                    htc.setVal(BigInteger.valueOf(453));
                    List<XWPFTableCell> cellsc = rowc.getTableCells();
                    for (XWPFTableCell cellc : cellsc) {
                        CTTcPr tcprc = cellc.getCTTc().addNewTcPr();
                        CTVerticalJc vac = tcprc.addNewVAlign();
                        vac.setVal(STVerticalJc.CENTER);
                        CTShd ctshdc = tcprc.addNewShd();
                        ctshdc.setColor("auto");
                        ctshdc.setVal(STShd.CLEAR);
                        if (rowCtc == 0) {
                            ctshdc.setFill("003366");
                        } else {
                            ctshdc.setFill("FFFFFF");
                        }
                        XWPFParagraph parac = cellc.getParagraphs().get(0);
                        XWPFRun rhc = parac.createRun();
                        if (rowCtc == 0) {
                            rhc.setText(" ");
                            rhc.setFontSize(10);
                            rhc.setBold(true);
                            parac.setAlignment(ParagraphAlignment.CENTER);
                        } else {
                            rhc.setText("");
                            parac.setAlignment(ParagraphAlignment.LEFT);
                        }
                        colCtc++;
                    }
                    colCtc = 0;
                    rowCtc++;
                }
                tablec.getRow(0).getCell(0).setText("Report Date ");
                
                tablec.getRow(0).getCell(1).setText("Report Name");
                tablec.getRow(0).getCell(2).setText("Report Version");
                tablec.getRow(0).getCell(3).setText("Report  Template Version ");
                tablec.getRow(0).getCell(4).setText("Released By");
                tablec.getRow(0).getCell(5).setText("Released To");
                tablec.getRow(1).getCell(0).setText(" ");
                tablec.getRow(1).getCell(1).setText(" ");
                tablec.getRow(1).getCell(2).setText(" ");
                tablec.getRow(1).getCell(3).setText(" ");
                tablec.getRow(1).getCell(4).setText(" ");
                tablec.getRow(1).getCell(5).setText(" ");
                
                XWPFParagraph para2 = document.createParagraph();
                XWPFRun ru2 = para2.createRun();
                ru2.addBreak();
                ru2.addBreak();
                ru2.addBreak();
                ru2.addBreak();
                ru2.addBreak();
                ru2.addBreak();
                int dRows = 1;
                int dCols = 0;
                XWPFTable tabled = document.createTable(dRows, dCols);
                tabled.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(9500));
                tabled.getCTTbl().getTblPr().unsetTblBorders();
                List<XWPFTableRow> rowsd = tabled.getRows();
                int rowCtd = 0;
                for (XWPFTableRow rowd : rowsd) {
                    CTTrPr trPrd = rowd.getCtRow().addNewTrPr();
                    CTHeight htd = trPrd.addNewTrHeight();
                    htd.setVal(BigInteger.valueOf(453));
                    List<XWPFTableCell> cellsd = rowd.getTableCells();
                    for (XWPFTableCell celld : cellsd) {
                        CTTcPr tcprd = celld.getCTTc().addNewTcPr();
                        CTVerticalJc vad = tcprd.addNewVAlign();
                        vad.setVal(STVerticalJc.CENTER);
                        CTShd ctshdd = tcprd.addNewShd();
                        ctshdd.setColor("auto");
                        ctshdd.setVal(STShd.CLEAR);
                        if (rowCtd == 0) {
                            ctshdd.setFill("FFFFFF");
                        } else {
                            ctshdd.setFill("FFFFFF");
                        }
                        XWPFParagraph parad = celld.getParagraphs().get(0);
                        XWPFRun rhd = parad.createRun();
                        rhd.setFontSize(10);
                        rhd.setBold(true);
                        rhd.setColor("99284C");
                        rhd.setText("Customer Details");
                        
                        int eRows = 1;
                        int eCols = 0;
                        XWPFTable tablee = document.createTable(eRows, eCols);
                        tablee.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(9500));
                        tablee.getCTTbl().getTblPr().unsetTblBorders();
                        List<XWPFTableRow> rowse = tablee.getRows();
                        int rowCte = 0;
                        for (XWPFTableRow rowe : rowse) {
                            CTTrPr trPre = rowe.getCtRow().addNewTrPr();
                            CTHeight hte = trPre.addNewTrHeight();
                            hte.setVal(BigInteger.valueOf(300));
                            List<XWPFTableCell> cellse = rowe.getTableCells();
                            for (XWPFTableCell celle : cellse) {
                                CTTcPr tcpre = celle.getCTTc().addNewTcPr();
                                CTVerticalJc vae = tcpre.addNewVAlign();
                                vae.setVal(STVerticalJc.CENTER);
                                CTShd ctshde = tcpre.addNewShd();
                                ctshde.setColor("auto");
                                ctshde.setVal(STShd.CLEAR);
                                if (rowCte == 0) {
                                    ctshde.setFill("FFFFFF");
                                    
                                } else {
                                    ctshde.setFill("FFFFFF");
                                }
                                tablee.getRow(0).getCell(0).setText("");
                                
                                int fRows = 2;
                                int fCols = 2;
                                XWPFTable tablef = document.createTable(fRows, fCols);
                                tablef.setWidth(2000);
                                tablef.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(4750));
                                tablef.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(4750));
                                
                                List<XWPFTableRow> rowsf = tablef.getRows();
                                int rowCtf = 0;
                                int colCtf = 0;
                                for (XWPFTableRow rowf : rowsf) {
                                    CTTrPr trPrf = rowf.getCtRow().addNewTrPr();
                                    CTHeight htf = trPrf.addNewTrHeight();
                                    htf.setVal(BigInteger.valueOf(330));
                                    List<XWPFTableCell> cellsf = rowf.getTableCells();
                                    for (XWPFTableCell cellf : cellsf) {
                                        CTTcPr tcprf = cellf.getCTTc().addNewTcPr();
                                        CTVerticalJc vaf = tcprf.addNewVAlign();
                                        vaf.setVal(STVerticalJc.CENTER);
                                        CTShd ctshdf = tcprf.addNewShd();
                                        ctshdf.setColor("auto");
                                        ctshdf.setVal(STShd.CLEAR);
                                        if (rowCtf == 0) {
                                            ctshdf.setFill("FFFFFF");
                                        } else {
                                            ctshdf.setFill("FFFFFF");
                                        }
                                        XWPFParagraph paraf = cellf.getParagraphs().get(0);
                                        XWPFRun rhf = paraf.createRun();
                                        if (rowCtf == 0) {
                                            rhf.setText(" ");
                                            rhf.setFontSize(10);
                                            rhf.setBold(true);
                                            paraf.setAlignment(ParagraphAlignment.CENTER);
                                        } else {
                                            rhf.setText("");
                                            paraf.setAlignment(ParagraphAlignment.CENTER);
                                        }
                                        colCtf++;
                                    }
                                    colCtf = 0;
                                    rowCtf++;
                                }
                                tablef.getRow(0).getCell(0).setText("Customer Name");
                                tablef.getRow(0).getCell(1).setText("");
                                tablef.getRow(1).getCell(0).setText("Geo's / Location ");
                                tablef.getRow(1).getCell(1).setText(" ");
                                
                                XWPFParagraph para3 = document.createParagraph();
                                XWPFRun ru3 = para3.createRun();
                                ru3.addBreak(BreakType.PAGE);
                                ru3.addBreak();
                                ru3.addBreak();
                                
                                int gRows = 1;
                                int gCols = 0;
                                XWPFTable tableg = document.createTable(gRows, gCols);
                                tableg.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(9500));
                                tableg.getCTTbl().getTblPr().unsetTblBorders();
                                List<XWPFTableRow> rowsg = tableg.getRows();
                                int rowCtg = 0;
                                for (XWPFTableRow rowg : rowsg) {
                                    CTTrPr trPrg = rowg.getCtRow().addNewTrPr();
                                    CTHeight htg = trPrg.addNewTrHeight();
                                    htg.setVal(BigInteger.valueOf(453));
                                    List<XWPFTableCell> cellsg = rowg.getTableCells();
                                    for (XWPFTableCell cellg : cellsg) {
                                        CTTcPr tcprg = cellg.getCTTc().addNewTcPr();
                                        CTVerticalJc vag = tcprg.addNewVAlign();
                                        vag.setVal(STVerticalJc.CENTER);
                                        CTShd ctshdg = tcprg.addNewShd();
                                        ctshdg.setColor("auto");
                                        ctshdg.setVal(STShd.CLEAR);
                                        if (rowCtg == 0) {
                                            ctshdg.setFill("FFFFFF");
                                        } else {
                                            ctshdg.setFill("FFFFFF");
                                        }
                                        XWPFParagraph parag = cellg.getParagraphs().get(0);
                                        XWPFRun rhg = parag.createRun();
                                        rhg.setFontSize(10);
                                        rhg.setBold(true);
                                        rhg.setColor("99284C");
                                        rhg.setText("Report Overview");
                                        
                                        int hRows = 1;
                                        int hCols = 0;
                                        XWPFTable tableh = document.createTable(hRows, hCols);
                                        tableh.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(9500));
                                        tableh.getCTTbl().getTblPr().unsetTblBorders();
                                        List<XWPFTableRow> rowsh = tableh.getRows();
                                        int rowCth = 0;
                                        for (XWPFTableRow rowh : rowsh) {
                                            CTTrPr trPrh = rowh.getCtRow().addNewTrPr();
                                            CTHeight hth = trPrh.addNewTrHeight();
                                            hth.setVal(BigInteger.valueOf(300));
                                            List<XWPFTableCell> cellsh = rowh.getTableCells();
                                            for (XWPFTableCell cellh : cellsh) {
                                                CTTcPr tcprh = cellh.getCTTc().addNewTcPr();
                                                CTVerticalJc vah = tcprh.addNewVAlign();
                                                vah.setVal(STVerticalJc.CENTER);
                                                CTShd ctshdh = tcprh.addNewShd();
                                                ctshdh.setColor("auto");
                                                ctshdh.setVal(STShd.CLEAR);
                                                if (rowCth == 0) {
                                                    ctshdh.setFill("FFFFFF");
                                                    
                                                } else {
                                                    ctshdh.setFill("FFFFFF");
                                                }
                                                tableh.getRow(0).getCell(0).setText("");
                                                
                                                int iRows = 7;
                                                int iCols = 5;
                                                XWPFTable tablei = document.createTable(iRows, iCols);
                                                tablei.setWidth(2000);
                                                tablei.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(2500));
                                                tablei.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(1750));
                                                tablei.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(1750));
                                                tablei.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(1750));
                                                tablei.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(1750));
                                                
                                                List<XWPFTableRow> rowsi = tablei.getRows();
                                                int rowCti = 0;
                                                int colCti = 0;
                                                for (XWPFTableRow rowi : rowsi) {
                                                    CTTrPr trPri = rowi.getCtRow().addNewTrPr();
                                                    CTHeight hti = trPri.addNewTrHeight();
                                                    hti.setVal(BigInteger.valueOf(330));
                                                    List<XWPFTableCell> cellsi = rowi.getTableCells();
                                                    for (XWPFTableCell celli : cellsi) {
                                                        CTTcPr tcpri = celli.getCTTc().addNewTcPr();
                                                        CTVerticalJc vai = tcpri.addNewVAlign();
                                                        vai.setVal(STVerticalJc.CENTER);
                                                        CTShd ctshdi = tcpri.addNewShd();
                                                        ctshdi.setColor("auto");
                                                        ctshdi.setVal(STShd.CLEAR);
                                                        if (rowCti == 0) {
                                                            ctshdi.setFill("003366");
                                                        } else {
                                                            ctshdi.setFill("FFFFFF");
                                                        }
                                                        XWPFParagraph parai = celli.getParagraphs().get(0);
                                                        XWPFRun rhi = parai.createRun();
                                                        if (rowCti == 0) {
                                                            rhi.setText(" ");
                                                            rhi.setFontSize(10);
                                                            rhi.setBold(true);
                                                            parai.setAlignment(ParagraphAlignment.CENTER);
                                                        } else {
                                                            rhi.setText("");
                                                            parai.setAlignment(ParagraphAlignment.CENTER);
                                                        }
                                                        colCti++;
                                                    }
                                                    colCti = 0;
                                                    rowCti++;
                                                }
                                                tablei.getRow(0).getCell(0).setText("Activity description");
                                                tablei.getRow(0).getCell(1).setText("Changes Observed");
                                                tablei.getRow(0).getCell(2).setText("Action Required");
                                                tablei.getRow(0).getCell(3).setText("Action Taken (Yes/No)");
                                                tablei.getRow(0).getCell(4).setText("Status (OK/ Warning/ Critical)");
                                                tablei.getRow(1).getCell(0).setText("Hardware information");
                                                tablei.getRow(1).getCell(1).setText("No");
                                                tablei.getRow(1).getCell(2).setText("");
                                                tablei.getRow(1).getCell(3).setText("");
                                                tablei.getRow(1).getCell(4).setText("");
                                                tablei.getRow(2).getCell(0).setText("Operating System");
                                                tablei.getRow(2).getCell(1).setText("No");
                                                tablei.getRow(2).getCell(2).setText("");
                                                tablei.getRow(2).getCell(3).setText("");
                                                tablei.getRow(2).getCell(4).setText("");
                                                tablei.getRow(3).getCell(0).setText("Database parameters");
                                                tablei.getRow(3).getCell(1).setText("No");
                                                tablei.getRow(3).getCell(2).setText("");
                                                tablei.getRow(3).getCell(3).setText("");
                                                tablei.getRow(3).getCell(4).setText("");
                                                tablei.getRow(4).getCell(0).setText("Archive log sequence difference between DC and DR");
                                                tablei.getRow(4).getCell(1).setText("0");
                                                tablei.getRow(4).getCell(2).setText("");
                                                tablei.getRow(4).getCell(3).setText("");
                                                tablei.getRow(4).getCell(4).setText("");
                                                tablei.getRow(5).getCell(0).setText("Archive log Switch");
                                                tablei.getRow(5).getCell(1).setText("No");
                                                tablei.getRow(5).getCell(2).setText("");
                                                tablei.getRow(5).getCell(3).setText("");
                                                tablei.getRow(5).getCell(4).setText("");
                                                tablei.getRow(6).getCell(0).setText("Alert log analysis");
                                                tablei.getRow(6).getCell(1).setText("No");
                                                tablei.getRow(6).getCell(2).setText("");
                                                tablei.getRow(6).getCell(3).setText("");
                                                tablei.getRow(6).getCell(4).setText("");
                                                XWPFParagraph para4 = document.createParagraph();
                                                XWPFRun ru4 = para4.createRun();
                                                ru4.addBreak();
                                                
                                                int jRows = 1;
                                                int jCols = 0;
                                                XWPFTable tablej = document.createTable(jRows, jCols);
                                                tablej.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(9500));
                                                tablej.getCTTbl().getTblPr().unsetTblBorders();
                                                List<XWPFTableRow> rowsj = tablej.getRows();
                                                int rowCtj = 0;
                                                for (XWPFTableRow rowj : rowsj) {
                                                    CTTrPr trPrj = rowj.getCtRow().addNewTrPr();
                                                    CTHeight htj = trPrj.addNewTrHeight();
                                                    htj.setVal(BigInteger.valueOf(453));
                                                    List<XWPFTableCell> cellsj = rowj.getTableCells();
                                                    for (XWPFTableCell cellj : cellsj) {
                                                        CTTcPr tcprj = cellj.getCTTc().addNewTcPr();
                                                        CTVerticalJc vaj = tcprj.addNewVAlign();
                                                        vaj.setVal(STVerticalJc.CENTER);
                                                        CTShd ctshdj = tcprj.addNewShd();
                                                        ctshdj.setColor("auto");
                                                        ctshdj.setVal(STShd.CLEAR);
                                                        if (rowCtj == 0) {
                                                            ctshdj.setFill("FFFFFF");
                                                        } else {
                                                            ctshdj.setFill("FFFFFF");
                                                        }
                                                        XWPFParagraph paraj = cellj.getParagraphs().get(0);
                                                        XWPFRun rhj = paraj.createRun();
                                                        rhj.setFontSize(10);
                                                        rhj.setBold(true);
                                                        rhj.setColor("99284C");
                                                        rhj.setText("Hardware and OS information");

                                                        /*int j1Rows = 1;
                                                         int j1Cols = 0;
                                                         XWPFTable tablej1 = document.createTable(j1Rows, j1Cols);
                                                         tablej1.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(9500));
                                                         tablej1.getCTTbl().getTblPr().unsetTblBorders();
                                                         List<XWPFTableRow> rowsj1 = tablej1.getRows();
                                                         int rowCtj1 = 0;
                                                         for (XWPFTableRow rowj1 : rowsj1) {
                                                         CTTrPr trPrj1 = rowj1.getCtRow().addNewTrPr();
                                                         CTHeight htj1 = trPrj1.addNewTrHeight();
                                                         htj1.setVal(BigInteger.valueOf(300));
                                                         List<XWPFTableCell> cellsj1 = rowj1.getTableCells();
                                                         for (XWPFTableCell cellj1 : cellsj1) {
                                                         CTTcPr tcprj1 = cellj1.getCTTc().addNewTcPr();
                                                         CTVerticalJc vaj1 = tcprj1.addNewVAlign();
                                                         vaj1.setVal(STVerticalJc.CENTER);
                                                         CTShd ctshdj1 = tcprj1.addNewShd();
                                                         ctshdj1.setColor("auto");
                                                         ctshdj1.setVal(STShd.CLEAR);
                                                         if (rowCtj1 == 0) {
                                                         ctshdj1.setFill("FFFFFF");

                                                         } else {
                                                         ctshdj1.setFill("FFFFFF");
                                                         }
                                                         tablej1.getRow(0).getCell(0).setText("");*/
                                                        int kRows = 1;
                                                        int kCols = 0;
                                                        XWPFTable tablek = document.createTable(kRows, kCols);
                                                        tablek.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(9500));
                                                        tablek.getCTTbl().getTblPr().unsetTblBorders();
                                                        List<XWPFTableRow> rowsk = tablek.getRows();
                                                        int rowCtk = 0;
                                                        for (XWPFTableRow rowk : rowsk) {
                                                            CTTrPr trPrk = rowk.getCtRow().addNewTrPr();
                                                            CTHeight htk = trPrk.addNewTrHeight();
                                                            htk.setVal(BigInteger.valueOf(453));
                                                            List<XWPFTableCell> cellsk = rowk.getTableCells();
                                                            for (XWPFTableCell cellk : cellsk) {
                                                                CTTcPr tcprk = cellk.getCTTc().addNewTcPr();
                                                                CTVerticalJc vak = tcprk.addNewVAlign();
                                                                vak.setVal(STVerticalJc.CENTER);
                                                                CTShd ctshdk = tcprk.addNewShd();
                                                                ctshdk.setColor("auto");
                                                                ctshdk.setVal(STShd.CLEAR);
                                                                if (rowCtk == 0) {
                                                                    ctshdk.setFill("FFFFFF");
                                                                } else {
                                                                    ctshdk.setFill("FFFFFF");
                                                                }
                                                                XWPFParagraph parak = cellk.getParagraphs().get(0);
                                                                XWPFRun rhk = parak.createRun();
                                                                rhk.setFontSize(10);
                                                                rhk.setColor("");
                                                                rhk.setText("The table below lists the server hardware / OS details and utilization information.");
                                                                
                                                                int k1Rows = 1;
                                                                int k1Cols = 0;
                                                                XWPFTable tablek1 = document.createTable(k1Rows, k1Cols);
                                                                tablek1.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(9500));
                                                                tablek1.getCTTbl().getTblPr().unsetTblBorders();
                                                                List<XWPFTableRow> rowsk1 = tablek1.getRows();
                                                                int rowCtk1 = 0;
                                                                for (XWPFTableRow rowk1 : rowsk1) {
                                                                    CTTrPr trPrk1 = rowk1.getCtRow().addNewTrPr();
                                                                    CTHeight htk1 = trPrk1.addNewTrHeight();
                                                                    htk1.setVal(BigInteger.valueOf(300));
                                                                    List<XWPFTableCell> cellsk1 = rowk1.getTableCells();
                                                                    for (XWPFTableCell cellk1 : cellsk1) {
                                                                        CTTcPr tcprk1 = cellk1.getCTTc().addNewTcPr();
                                                                        CTVerticalJc vak1 = tcprk1.addNewVAlign();
                                                                        vak1.setVal(STVerticalJc.CENTER);
                                                                        CTShd ctshdk1 = tcprk1.addNewShd();
                                                                        ctshdk1.setColor("auto");
                                                                        ctshdk1.setVal(STShd.CLEAR);
                                                                        if (rowCtk1 == 0) {
                                                                            ctshdk1.setFill("FFFFFF");
                                                                            
                                                                        } else {
                                                                            ctshdk1.setFill("FFFFFF");
                                                                        }
                                                                        tablek1.getRow(0).getCell(0).setText("");
                                                                        
                                                                        int lRows = 2;
                                                                        int lCols = 8;
                                                                        XWPFTable tablel = document.createTable(lRows, lCols);
                                                                        tablel.setWidth(2000);
                                                                        tablel.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(1250));
                                                                        tablel.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(1000));
                                                                        tablel.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(1800));
                                                                        tablel.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(900));
                                                                        tablel.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(950));
                                                                        tablel.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(1400));
                                                                        tablel.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(1100));
                                                                        tablel.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(1100));
                                                                        List<XWPFTableRow> rowsl = tablel.getRows();
                                                                        int rowCtl = 0;
                                                                        int colCtl = 0;
                                                                        for (XWPFTableRow rowl : rowsl) {
                                                                            CTTrPr trPrl = rowl.getCtRow().addNewTrPr();
                                                                            CTHeight htl = trPrl.addNewTrHeight();
                                                                            htl.setVal(BigInteger.valueOf(330));
                                                                            List<XWPFTableCell> cellsl = rowl.getTableCells();
                                                                            for (XWPFTableCell celll : cellsl) {
                                                                                CTTcPr tcprl = celll.getCTTc().addNewTcPr();
                                                                                CTVerticalJc val = tcprl.addNewVAlign();
                                                                                val.setVal(STVerticalJc.CENTER);
                                                                                CTShd ctshdl = tcprl.addNewShd();
                                                                                ctshdl.setColor("auto");
                                                                                ctshdl.setVal(STShd.CLEAR);
                                                                                if (rowCtl == 0) {
                                                                                    ctshdl.setFill("003366");
                                                                                } else {
                                                                                    ctshdl.setFill("FFFFFF");
                                                                                }
                                                                                XWPFParagraph paral = celll.getParagraphs().get(0);
                                                                                XWPFRun rhl = paral.createRun();
                                                                                if (rowCtl == 0) {
                                                                                    rhl.setText(" ");
                                                                                    rhl.setFontSize(10);
                                                                                    rhl.setBold(true);
                                                                                    paral.setAlignment(ParagraphAlignment.CENTER);
                                                                                } else {
                                                                                    rhl.setText("");
                                                                                    paral.setAlignment(ParagraphAlignment.CENTER);
                                                                                }
                                                                                colCtl++;
                                                                            }
                                                                            colCtl = 0;
                                                                            rowCtl++;
                                                                        }
                                                                        tablel.getRow(0).getCell(0).setText("Host Name");
                                                                        tablel.getRow(0).getCell(1).setText("IP");
                                                                        tablel.getRow(0).getCell(2).setText("OS");
                                                                        tablel.getRow(0).getCell(3).setText("Make");
                                                                        tablel.getRow(0).getCell(4).setText("Model");
                                                                        tablel.getRow(0).getCell(5).setText("CPU type");
                                                                        tablel.getRow(0).getCell(6).setText("CPU core count");
                                                                        tablel.getRow(0).getCell(7).setText("RAM (in GB)");
                                                                        
                                                                        tablel.getRow(1).getCell(0).setText("");
                                                                        tablel.getRow(1).getCell(1).setText("");
                                                                        tablel.getRow(1).getCell(2).setText("");
                                                                        tablel.getRow(1).getCell(3).setText("");
                                                                        tablel.getRow(1).getCell(4).setText("");
                                                                        tablel.getRow(1).getCell(5).setText(" ");
                                                                        tablel.getRow(1).getCell(6).setText("  ");
                                                                        tablel.getRow(1).getCell(7).setText(" ");
                                                                        
                                                                        XWPFParagraph para7 = document.createParagraph();
                                                                        XWPFRun ru7 = para7.createRun();
                                                                        ru7.addBreak();
                                                                        
                                                                        int mRows = 1;
                                                                        int mCols = 0;
                                                                        XWPFTable tablem = document.createTable(mRows, mCols);
                                                                        tablem.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(9500));
                                                                        tablem.getCTTbl().getTblPr().unsetTblBorders();
                                                                        List<XWPFTableRow> rowsm = tablem.getRows();
                                                                        int rowCtm = 0;
                                                                        for (XWPFTableRow rowm : rowsm) {
                                                                            CTTrPr trPrm = rowm.getCtRow().addNewTrPr();
                                                                            CTHeight htm = trPrm.addNewTrHeight();
                                                                            htm.setVal(BigInteger.valueOf(453));
                                                                            List<XWPFTableCell> cellsm = rowm.getTableCells();
                                                                            for (XWPFTableCell cellm : cellsm) {
                                                                                CTTcPr tcprm = cellm.getCTTc().addNewTcPr();
                                                                                CTVerticalJc vam = tcprm.addNewVAlign();
                                                                                vam.setVal(STVerticalJc.CENTER);
                                                                                CTShd ctshdm = tcprm.addNewShd();
                                                                                ctshdm.setColor("auto");
                                                                                ctshdm.setVal(STShd.CLEAR);
                                                                                if (rowCtm == 0) {
                                                                                    ctshdm.setFill("FFFFFF");
                                                                                } else {
                                                                                    ctshdm.setFill("FFFFFF");
                                                                                }
                                                                                XWPFParagraph param = cellm.getParagraphs().get(0);
                                                                                XWPFRun rhm = param.createRun();
                                                                                rhm.setFontSize(10);
                                                                                rhm.setBold(true);
                                                                                rhm.setColor("99284C");
                                                                                rhm.setText("Changes Observed");
                                                                                
                                                                                int m1Rows = 1;
                                                                                int m1Cols = 0;
                                                                                XWPFTable tablem1 = document.createTable(m1Rows, m1Cols);
                                                                                tablem1.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(9500));
                                                                                tablem1.getCTTbl().getTblPr().unsetTblBorders();
                                                                                List<XWPFTableRow> rowsm1 = tablem1.getRows();
                                                                                int rowCtm1 = 0;
                                                                                for (XWPFTableRow rowm1 : rowsm1) {
                                                                                    CTTrPr trPrm1 = rowm1.getCtRow().addNewTrPr();
                                                                                    CTHeight htm1 = trPrm1.addNewTrHeight();
                                                                                    htm1.setVal(BigInteger.valueOf(300));
                                                                                    List<XWPFTableCell> cellsm1 = rowm1.getTableCells();
                                                                                    for (XWPFTableCell cellm1 : cellsm1) {
                                                                                        CTTcPr tcprm1 = cellm1.getCTTc().addNewTcPr();
                                                                                        CTVerticalJc vam1 = tcprm1.addNewVAlign();
                                                                                        vam1.setVal(STVerticalJc.CENTER);
                                                                                        CTShd ctshdm1 = tcprm1.addNewShd();
                                                                                        ctshdm1.setColor("auto");
                                                                                        ctshdm1.setVal(STShd.CLEAR);
                                                                                        if (rowCtm1 == 0) {
                                                                                            ctshdm1.setFill("FFFFFF");
                                                                                            
                                                                                        } else {
                                                                                            ctshdm1.setFill("FFFFFF");
                                                                                        }
                                                                                        tablem1.getRow(0).getCell(0).setText("");
                                                                                        
                                                                                        int nRows = 1;
                                                                                        int nCols = 5;
                                                                                        XWPFTable tablen = document.createTable(nRows, nCols);
                                                                                        tablen.setWidth(2000);
                                                                                        tablen.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(1500));
                                                                                        tablen.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(2000));
                                                                                        tablen.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(2000));
                                                                                        tablen.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(2000));
                                                                                        tablen.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(2000));
                                                                                        
                                                                                        List<XWPFTableRow> rowsn = tablen.getRows();
                                                                                        int rowCtn = 0;
                                                                                        int colCtn = 0;
                                                                                        for (XWPFTableRow rown : rowsn) {
                                                                                            CTTrPr trPrn = rown.getCtRow().addNewTrPr();
                                                                                            CTHeight htn = trPrn.addNewTrHeight();
                                                                                            htn.setVal(BigInteger.valueOf(330));
                                                                                            List<XWPFTableCell> cellsn = rown.getTableCells();
                                                                                            for (XWPFTableCell celln : cellsn) {
                                                                                                CTTcPr tcprn = celln.getCTTc().addNewTcPr();
                                                                                                CTVerticalJc van = tcprn.addNewVAlign();
                                                                                                van.setVal(STVerticalJc.CENTER);
                                                                                                CTShd ctshdn = tcprn.addNewShd();
                                                                                                ctshdn.setColor("auto");
                                                                                                ctshdn.setVal(STShd.CLEAR);
                                                                                                if (rowCtn == 0) {
                                                                                                    ctshdn.setFill("003366");
                                                                                                } else {
                                                                                                    ctshdn.setFill("FFFFFF");
                                                                                                }
                                                                                                XWPFParagraph paran = celln.getParagraphs().get(0);
                                                                                                XWPFRun rhn = paran.createRun();
                                                                                                if (rowCtn == 0) {
                                                                                                    rhn.setText(" ");
                                                                                                    rhn.setFontSize(10);
                                                                                                    rhn.setBold(true);
                                                                                                    paran.setAlignment(ParagraphAlignment.CENTER);
                                                                                                } else {
                                                                                                    rhn.setText("");
                                                                                                    paran.setAlignment(ParagraphAlignment.CENTER);
                                                                                                }
                                                                                                colCtn++;
                                                                                            }
                                                                                            colCtn = 0;
                                                                                            rowCtn++;
                                                                                        }
                                                                                        tablen.getRow(0).getCell(0).setText(" S.No");
                                                                                        tablen.getRow(0).getCell(1).setText(" Item");
                                                                                        tablen.getRow(0).getCell(2).setText(" Old Value ");
                                                                                        tablen.getRow(0).getCell(3).setText(" New Value ");
                                                                                        tablen.getRow(0).getCell(4).setText("Remarks ");
                                                                                        int j1Rows = 1;
                                                                                        int j1Cols = 0;
                                                                                        XWPFTable tablej1 = document.createTable(j1Rows, j1Cols);
                                                                                        tablej1.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(9500));
                                                                                        List<XWPFTableRow> rowsj1 = tablej1.getRows();
                                                                                        int rowCtj1 = 0;
                                                                                        for (XWPFTableRow rowj1 : rowsj1) {
                                                                                            CTTrPr trPrj1 = rowj1.getCtRow().addNewTrPr();
                                                                                            CTHeight htj1 = trPrj1.addNewTrHeight();
                                                                                            htj1.setVal(BigInteger.valueOf(300));
                                                                                            List<XWPFTableCell> cellsj1 = rowj1.getTableCells();
                                                                                            for (XWPFTableCell cellj1 : cellsj1) {
                                                                                                CTTcPr tcprj1 = cellj1.getCTTc().addNewTcPr();
                                                                                                CTVerticalJc vaj1 = tcprj1.addNewVAlign();
                                                                                                vaj1.setVal(STVerticalJc.CENTER);
                                                                                                CTShd ctshdj1 = tcprj1.addNewShd();
                                                                                                ctshdj1.setColor("auto");
                                                                                                ctshdj1.setVal(STShd.CLEAR);
                                                                                                if (rowCtj1 == 0) {
                                                                                                    ctshdj1.setFill("FFFFFF");
                                                                                                    
                                                                                                } else {
                                                                                                    ctshdj1.setFill("FFFFFF");
                                                                                                }
                                                                                                tablej1.getRow(0).getCell(0).setText("                                      NO CHANGE");
                                                                                                
                                                                                                XWPFParagraph para8 = document.createParagraph();
                                                                                                XWPFRun ru8 = para8.createRun();
                                                                                                ru8.addBreak();
                                                                                                
                                                                                                int oRows = 1;
                                                                                                int oCols = 0;
                                                                                                XWPFTable tableo = document.createTable(oRows, oCols);
                                                                                                tableo.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(9500));
                                                                                                tableo.getCTTbl().getTblPr().unsetTblBorders();
                                                                                                List<XWPFTableRow> rowso = tableo.getRows();
                                                                                                int rowCto = 0;
                                                                                                for (XWPFTableRow rowo : rowso) {
                                                                                                    CTTrPr trPro = rowo.getCtRow().addNewTrPr();
                                                                                                    CTHeight hto = trPro.addNewTrHeight();
                                                                                                    hto.setVal(BigInteger.valueOf(453));
                                                                                                    List<XWPFTableCell> cellso = rowo.getTableCells();
                                                                                                    for (XWPFTableCell cello : cellso) {
                                                                                                        CTTcPr tcpro = cello.getCTTc().addNewTcPr();
                                                                                                        CTVerticalJc vao = tcpro.addNewVAlign();
                                                                                                        vao.setVal(STVerticalJc.CENTER);
                                                                                                        CTShd ctshdo = tcpro.addNewShd();
                                                                                                        ctshdo.setColor("auto");
                                                                                                        ctshdo.setVal(STShd.CLEAR);
                                                                                                        if (rowCto == 0) {
                                                                                                            ctshdo.setFill("FFFFFF");
                                                                                                        } else {
                                                                                                            ctshdo.setFill("FFFFFF");
                                                                                                        }
                                                                                                        XWPFParagraph parao = cello.getParagraphs().get(0);
                                                                                                        XWPFRun rho = parao.createRun();
                                                                                                        rho.setFontSize(10);
                                                                                                        rho.setBold(true);
                                                                                                        rho.setColor("99284C");
                                                                                                        rho.setText("Database Information");
                                                                                                        
                                                                                                        int o5Rows = 1;
                                                                                                        int o5Cols = 0;
                                                                                                        XWPFTable tableo5 = document.createTable(o5Rows, o5Cols);
                                                                                                        tableo5.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(9500));
                                                                                                        tableo5.getCTTbl().getTblPr().unsetTblBorders();
                                                                                                        List<XWPFTableRow> rowso5 = tableo5.getRows();
                                                                                                        int rowCto5 = 0;
                                                                                                        for (XWPFTableRow rowo5 : rowso5) {
                                                                                                            CTTrPr trPro5 = rowo5.getCtRow().addNewTrPr();
                                                                                                            CTHeight hto5 = trPro5.addNewTrHeight();
                                                                                                            hto5.setVal(BigInteger.valueOf(453));
                                                                                                            List<XWPFTableCell> cellso5 = rowo5.getTableCells();
                                                                                                            for (XWPFTableCell cello5 : cellso5) {
                                                                                                                CTTcPr tcpro5 = cello5.getCTTc().addNewTcPr();
                                                                                                                CTVerticalJc vao5 = tcpro5.addNewVAlign();
                                                                                                                vao5.setVal(STVerticalJc.CENTER);
                                                                                                                CTShd ctshdo5 = tcpro5.addNewShd();
                                                                                                                ctshdo5.setColor("auto");
                                                                                                                ctshdo5.setVal(STShd.CLEAR);
                                                                                                                if (rowCto5 == 0) {
                                                                                                                    ctshdo5.setFill("FFFFFF");
                                                                                                                } else {
                                                                                                                    ctshdo5.setFill("FFFFFF");
                                                                                                                }
                                                                                                                XWPFParagraph parao5 = cello5.getParagraphs().get(0);
                                                                                                                XWPFRun rho5 = parao5.createRun();
                                                                                                                rho5.setFontSize(10);
                                                                                                                rho5.setBold(true);
                                                                                                                rho5.setColor("99284C");
                                                                                                                rho5.setText("Database Overview");
                                                                                                                
                                                                                                                int o1Rows = 1;
                                                                                                                int o1Cols = 0;
                                                                                                                XWPFTable tableo1 = document.createTable(o1Rows, o1Cols);
                                                                                                                tableo1.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(9500));
                                                                                                                tableo1.getCTTbl().getTblPr().unsetTblBorders();
                                                                                                                List<XWPFTableRow> rowso1 = tableo1.getRows();
                                                                                                                int rowCto1 = 0;
                                                                                                                for (XWPFTableRow rowo1 : rowso1) {
                                                                                                                    CTTrPr trPro1 = rowo1.getCtRow().addNewTrPr();
                                                                                                                    CTHeight hto1 = trPro1.addNewTrHeight();
                                                                                                                    hto1.setVal(BigInteger.valueOf(300));
                                                                                                                    List<XWPFTableCell> cellso1 = rowo1.getTableCells();
                                                                                                                    for (XWPFTableCell cello1 : cellso1) {
                                                                                                                        CTTcPr tcpro1 = cello1.getCTTc().addNewTcPr();
                                                                                                                        CTVerticalJc vao1 = tcpro1.addNewVAlign();
                                                                                                                        vao1.setVal(STVerticalJc.CENTER);
                                                                                                                        CTShd ctshdo1 = tcpro1.addNewShd();
                                                                                                                        ctshdo1.setColor("auto");
                                                                                                                        ctshdo1.setVal(STShd.CLEAR);
                                                                                                                        if (rowCto1 == 0) {
                                                                                                                            ctshdo1.setFill("FFFFFF");
                                                                                                                            
                                                                                                                        } else {
                                                                                                                            ctshdo1.setFill("FFFFFF");
                                                                                                                        }
                                                                                                                        tableo1.getRow(0).getCell(0).setText("");
                                                                                                                        int pRows = 2;
                                                                                                                        int pCols = 3;
                                                                                                                        XWPFTable tablep = document.createTable(pRows, pCols);
                                                                                                                        tablep.setWidth(2000);
                                                                                                                        tablep.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(2500));
                                                                                                                        tablep.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(3000));
                                                                                                                        tablep.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(4000));
                                                                                                                        
                                                                                                                        List<XWPFTableRow> rowsp = tablep.getRows();
                                                                                                                        int rowCtp = 0;
                                                                                                                        int colCtp = 0;
                                                                                                                        for (XWPFTableRow rowp : rowsp) {
                                                                                                                            CTTrPr trPrp = rowp.getCtRow().addNewTrPr();
                                                                                                                            CTHeight htp = trPrp.addNewTrHeight();
                                                                                                                            htp.setVal(BigInteger.valueOf(330));
                                                                                                                            List<XWPFTableCell> cellsp = rowp.getTableCells();
                                                                                                                            for (XWPFTableCell cellp : cellsp) {
                                                                                                                                CTTcPr tcprp = cellp.getCTTc().addNewTcPr();
                                                                                                                                CTVerticalJc vap = tcprp.addNewVAlign();
                                                                                                                                vap.setVal(STVerticalJc.CENTER);
                                                                                                                                CTShd ctshdp = tcprp.addNewShd();
                                                                                                                                ctshdp.setColor("auto");
                                                                                                                                ctshdp.setVal(STShd.CLEAR);
                                                                                                                                if (rowCtp == 0) {
                                                                                                                                    ctshdp.setFill("003366");
                                                                                                                                } else {
                                                                                                                                    ctshdp.setFill("FFFFFF");
                                                                                                                                }
                                                                                                                                XWPFParagraph parap = cellp.getParagraphs().get(0);
                                                                                                                                XWPFRun rhp = parap.createRun();
                                                                                                                                if (rowCtp == 0) {
                                                                                                                                    rhp.setText(" ");
                                                                                                                                    rhp.setFontSize(10);
                                                                                                                                    rhp.setBold(true);
                                                                                                                                    parap.setAlignment(ParagraphAlignment.CENTER);
                                                                                                                                } else {
                                                                                                                                    rhp.setText("");
                                                                                                                                    parap.setAlignment(ParagraphAlignment.CENTER);
                                                                                                                                }
                                                                                                                                colCtp++;
                                                                                                                            }
                                                                                                                            colCtp = 0;
                                                                                                                            rowCtp++;
                                                                                                                        }
                                                                                                                        tablep.getRow(0).getCell(0).setText(" DB Name");
                                                                                                                        tablep.getRow(0).getCell(1).setText(" DB ID");
                                                                                                                        tablep.getRow(0).getCell(2).setText(" Database Version / Edition ");
                                                                                                                        tablep.getRow(1).getCell(0).setText("");
                                                                                                                        tablep.getRow(1).getCell(1).setText("");
                                                                                                                        tablep.getRow(1).getCell(2).setText("");
                                                                                                                        
                                                                                                                        XWPFParagraph para9 = document.createParagraph();
                                                                                                                        XWPFRun ru9 = para9.createRun();
                                                                                                                        ru9.addBreak(BreakType.PAGE);
                                                                                                                        int qRows = 1;
                                                                                                                        int qCols = 0;
                                                                                                                        XWPFTable tableq = document.createTable(qRows, qCols);
                                                                                                                        tableq.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(9500));
                                                                                                                        tableq.getCTTbl().getTblPr().unsetTblBorders();
                                                                                                                        List<XWPFTableRow> rowsq = tableq.getRows();
                                                                                                                        int rowCtq = 0;
                                                                                                                        for (XWPFTableRow rowq : rowsq) {
                                                                                                                            CTTrPr trPrq = rowq.getCtRow().addNewTrPr();
                                                                                                                            CTHeight htq = trPrq.addNewTrHeight();
                                                                                                                            htq.setVal(BigInteger.valueOf(453));
                                                                                                                            List<XWPFTableCell> cellsq = rowq.getTableCells();
                                                                                                                            for (XWPFTableCell cellq : cellsq) {
                                                                                                                                CTTcPr tcprq = cellq.getCTTc().addNewTcPr();
                                                                                                                                CTVerticalJc vaq = tcprq.addNewVAlign();
                                                                                                                                vaq.setVal(STVerticalJc.CENTER);
                                                                                                                                CTShd ctshdq = tcprq.addNewShd();
                                                                                                                                ctshdq.setColor("auto");
                                                                                                                                ctshdq.setVal(STShd.CLEAR);
                                                                                                                                if (rowCtq == 0) {
                                                                                                                                    ctshdq.setFill("FFFFFF");
                                                                                                                                } else {
                                                                                                                                    ctshdq.setFill("FFFFFF");
                                                                                                                                }
                                                                                                                                XWPFParagraph paraq = cellq.getParagraphs().get(0);
                                                                                                                                XWPFRun rhq = paraq.createRun();
                                                                                                                                rhq.setFontSize(10);
                                                                                                                                rhq.setBold(true);
                                                                                                                                rhq.setColor("99284C");
                                                                                                                                rhq.setText("Database parameters");

                                                                                                                                /*   int q1Rows = 1;
                                                                                                                                 int q1Cols = 0;
                                                                                                                                 XWPFTable tableq1 = document.createTable(q1Rows, q1Cols);
                                                                                                                                 tableq1.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(9500));
                                                                                                                                 tableq1.getCTTbl().getTblPr().unsetTblBorders();
                                                                                                                                 List<XWPFTableRow> rowsq1 = tableq1.getRows();
                                                                                                                                 int rowCtq1 = 0;
                                                                                                                                 for (XWPFTableRow rowq1 : rowsq1) {
                                                                                                                                 CTTrPr trPrq1 = rowq1.getCtRow().addNewTrPr();
                                                                                                                                 CTHeight htq1 = trPrq1.addNewTrHeight();
                                                                                                                                 htq1.setVal(BigInteger.valueOf(300));
                                                                                                                                 List<XWPFTableCell> cellsq1 = rowq1.getTableCells();
                                                                                                                                 for (XWPFTableCell cellq1 : cellsq1) {
                                                                                                                                 CTTcPr tcprq1 = cellq1.getCTTc().addNewTcPr();
                                                                                                                                 CTVerticalJc vaq1 = tcprq1.addNewVAlign();
                                                                                                                                 vaq1.setVal(STVerticalJc.CENTER);
                                                                                                                                 CTShd ctshdq1 = tcprq1.addNewShd();
                                                                                                                                 ctshdq1.setColor("auto");
                                                                                                                                 ctshdq1.setVal(STShd.CLEAR);
                                                                                                                                 if (rowCtq1 == 0) {
                                                                                                                                 ctshdq1.setFill("FFFFFF");

                                                                                                                                 } else {
                                                                                                                                 ctshdq1.setFill("FFFFFF");
                                                                                                                                 }
                                                                                                                                 tableq1.getRow(0).getCell(0).setText("");
                                                                                                                                 */
                                                                                                                                int rRows = 1;
                                                                                                                                int rCols = 0;
                                                                                                                                XWPFTable tabler = document.createTable(rRows, rCols);
                                                                                                                                tabler.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(9500));
                                                                                                                                tabler.getCTTbl().getTblPr().unsetTblBorders();
                                                                                                                                List<XWPFTableRow> rowsr = tabler.getRows();
                                                                                                                                int rowCtr = 0;
                                                                                                                                for (XWPFTableRow rowr : rowsr) {
                                                                                                                                    CTTrPr trPrr = rowr.getCtRow().addNewTrPr();
                                                                                                                                    CTHeight htr = trPrr.addNewTrHeight();
                                                                                                                                    htr.setVal(BigInteger.valueOf(453));
                                                                                                                                    List<XWPFTableCell> cellsr = rowr.getTableCells();
                                                                                                                                    for (XWPFTableCell cellr : cellsr) {
                                                                                                                                        CTTcPr tcprr = cellr.getCTTc().addNewTcPr();
                                                                                                                                        CTVerticalJc var = tcprr.addNewVAlign();
                                                                                                                                        var.setVal(STVerticalJc.CENTER);
                                                                                                                                        CTShd ctshdr = tcprr.addNewShd();
                                                                                                                                        ctshdr.setColor("auto");
                                                                                                                                        ctshdr.setVal(STShd.CLEAR);
                                                                                                                                        if (rowCtr == 0) {
                                                                                                                                            ctshdr.setFill("FFFFFF");
                                                                                                                                        } else {
                                                                                                                                            ctshdr.setFill("FFFFFF");
                                                                                                                                        }
                                                                                                                                        XWPFParagraph parar = cellr.getParagraphs().get(0);
                                                                                                                                        XWPFRun rhr = parar.createRun();
                                                                                                                                        rhr.setFontSize(10);
                                                                                                                                        rhr.setColor("");
                                                                                                                                        rhr.setText("  The table below lists the relevant parameters and their values.");
                                                                                                                                        
                                                                                                                                        int r1Rows = 1;
                                                                                                                                        int r1Cols = 0;
                                                                                                                                        XWPFTable tabler1 = document.createTable(r1Rows, r1Cols);
                                                                                                                                        tabler1.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(9500));
                                                                                                                                        tabler1.getCTTbl().getTblPr().unsetTblBorders();
                                                                                                                                        List<XWPFTableRow> rowsr1 = tabler1.getRows();
                                                                                                                                        int rowCtr1 = 0;
                                                                                                                                        for (XWPFTableRow rowr1 : rowsr1) {
                                                                                                                                            CTTrPr trPrr1 = rowr1.getCtRow().addNewTrPr();
                                                                                                                                            CTHeight htr1 = trPrr1.addNewTrHeight();
                                                                                                                                            htr1.setVal(BigInteger.valueOf(300));
                                                                                                                                            List<XWPFTableCell> cellsr1 = rowr1.getTableCells();
                                                                                                                                            for (XWPFTableCell cellr1 : cellsr1) {
                                                                                                                                                CTTcPr tcprr1 = cellr1.getCTTc().addNewTcPr();
                                                                                                                                                CTVerticalJc var1 = tcprr1.addNewVAlign();
                                                                                                                                                var1.setVal(STVerticalJc.CENTER);
                                                                                                                                                CTShd ctshdr1 = tcprr1.addNewShd();
                                                                                                                                                ctshdr1.setColor("auto");
                                                                                                                                                ctshdr1.setVal(STShd.CLEAR);
                                                                                                                                                if (rowCtr1 == 0) {
                                                                                                                                                    ctshdr1.setFill("FFFFFF");
                                                                                                                                                    
                                                                                                                                                } else {
                                                                                                                                                    ctshdr1.setFill("FFFFFF");
                                                                                                                                                }
                                                                                                                                                tabler1.getRow(0).getCell(0).setText("");
                                                                                                                                                
                                                                                                                                                int sRows = 24;
                                                                                                                                                int sCols = 3;
                                                                                                                                                XWPFTable tables = document.createTable(sRows, sCols);
                                                                                                                                                tables.setWidth(2000);
                                                                                                                                                tables.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(3000));
                                                                                                                                                tables.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(5000));
                                                                                                                                                tables.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(1500));
                                                                                                                                                
                                                                                                                                                List<XWPFTableRow> rowss = tables.getRows();
                                                                                                                                                int rowCts = 0;
                                                                                                                                                int colCts = 0;
                                                                                                                                                for (XWPFTableRow rows : rowss) {
                                                                                                                                                    CTTrPr trPrs = rows.getCtRow().addNewTrPr();
                                                                                                                                                    CTHeight hts = trPrs.addNewTrHeight();
                                                                                                                                                    hts.setVal(BigInteger.valueOf(330));
                                                                                                                                                    List<XWPFTableCell> cellss = rows.getTableCells();
                                                                                                                                                    for (XWPFTableCell cells : cellss) {
                                                                                                                                                        CTTcPr tcprs = cells.getCTTc().addNewTcPr();
                                                                                                                                                        CTVerticalJc vas = tcprs.addNewVAlign();
                                                                                                                                                        vas.setVal(STVerticalJc.CENTER);
                                                                                                                                                        CTShd ctshds = tcprs.addNewShd();
                                                                                                                                                        ctshds.setColor("auto");
                                                                                                                                                        ctshds.setVal(STShd.CLEAR);
                                                                                                                                                        if (rowCts == 0) {
                                                                                                                                                            ctshds.setFill("003366");
                                                                                                                                                        } else {
                                                                                                                                                            ctshds.setFill("FFFFFF");
                                                                                                                                                        }
                                                                                                                                                        XWPFParagraph paras = cells.getParagraphs().get(0);
                                                                                                                                                        XWPFRun rhs = paras.createRun();
                                                                                                                                                        if (rowCts == 0) {
                                                                                                                                                            rhs.setText(" ");
                                                                                                                                                            rhs.setFontSize(10);
                                                                                                                                                            rhs.setBold(true);
                                                                                                                                                            paras.setAlignment(ParagraphAlignment.CENTER);
                                                                                                                                                        } else {
                                                                                                                                                            rhs.setText("");
                                                                                                                                                            paras.setAlignment(ParagraphAlignment.CENTER);
                                                                                                                                                        }
                                                                                                                                                        colCts++;
                                                                                                                                                    }
                                                                                                                                                    colCts = 0;
                                                                                                                                                    rowCts++;
                                                                                                                                                }
                                                                                                                                                tables.getRow(0).getCell(0).setText(" Parameters");
                                                                                                                                                tables.getRow(0).getCell(1).setText(" Description");
                                                                                                                                                tables.getRow(0).getCell(2).setText(" Remarks(if different) ");
                                                                                                                                                tables.getRow(1).getCell(0).setText("");
                                                                                                                                                tables.getRow(1).getCell(1).setText("");
                                                                                                                                                tables.getRow(1).getCell(2).setText("NA");
                                                                                                                                                tables.getRow(2).getCell(0).setText("");
                                                                                                                                                tables.getRow(2).getCell(1).setText("");
                                                                                                                                                tables.getRow(2).getCell(2).setText("NA");
                                                                                                                                                tables.getRow(3).getCell(0).setText("");
                                                                                                                                                tables.getRow(3).getCell(1).setText("");
                                                                                                                                                tables.getRow(3).getCell(2).setText("NA");
                                                                                                                                                tables.getRow(4).getCell(0).setText("");
                                                                                                                                                tables.getRow(4).getCell(1).setText("");
                                                                                                                                                tables.getRow(4).getCell(2).setText("NA");
                                                                                                                                                tables.getRow(5).getCell(0).setText("");
                                                                                                                                                tables.getRow(5).getCell(1).setText("");
                                                                                                                                                tables.getRow(5).getCell(2).setText("NA");
                                                                                                                                                tables.getRow(6).getCell(0).setText("");
                                                                                                                                                tables.getRow(6).getCell(1).setText("");
                                                                                                                                                tables.getRow(6).getCell(2).setText("NA");
                                                                                                                                                tables.getRow(7).getCell(0).setText("");
                                                                                                                                                tables.getRow(7).getCell(1).setText("");
                                                                                                                                                tables.getRow(7).getCell(2).setText("NA");
                                                                                                                                                tables.getRow(8).getCell(0).setText("");
                                                                                                                                                tables.getRow(8).getCell(1).setText("");
                                                                                                                                                tables.getRow(8).getCell(2).setText("NA");
                                                                                                                                                tables.getRow(9).getCell(0).setText("");
                                                                                                                                                tables.getRow(9).getCell(1).setText("");
                                                                                                                                                tables.getRow(9).getCell(2).setText("NA");
                                                                                                                                                tables.getRow(10).getCell(0).setText("");
                                                                                                                                                tables.getRow(10).getCell(1).setText("");
                                                                                                                                                tables.getRow(10).getCell(2).setText("NA");
                                                                                                                                                tables.getRow(11).getCell(0).setText("");
                                                                                                                                                tables.getRow(11).getCell(1).setText("");
                                                                                                                                                tables.getRow(11).getCell(2).setText("NA");
                                                                                                                                                tables.getRow(12).getCell(0).setText("");
                                                                                                                                                tables.getRow(12).getCell(1).setText("");
                                                                                                                                                tables.getRow(12).getCell(2).setText("NA");
                                                                                                                                                tables.getRow(13).getCell(0).setText("");
                                                                                                                                                tables.getRow(13).getCell(1).setText("");
                                                                                                                                                tables.getRow(13).getCell(2).setText("NA");
                                                                                                                                                tables.getRow(14).getCell(0).setText("");
                                                                                                                                                tables.getRow(14).getCell(1).setText("");
                                                                                                                                                tables.getRow(14).getCell(2).setText("NA");
                                                                                                                                                tables.getRow(15).getCell(0).setText("");
                                                                                                                                                tables.getRow(15).getCell(1).setText("");
                                                                                                                                                tables.getRow(15).getCell(2).setText("NA");
                                                                                                                                                tables.getRow(16).getCell(0).setText("");
                                                                                                                                                tables.getRow(16).getCell(1).setText("");
                                                                                                                                                tables.getRow(16).getCell(2).setText("NA");
                                                                                                                                                tables.getRow(17).getCell(0).setText("");
                                                                                                                                                tables.getRow(17).getCell(1).setText("");
                                                                                                                                                tables.getRow(17).getCell(2).setText("NA");
                                                                                                                                                tables.getRow(18).getCell(0).setText("");
                                                                                                                                                tables.getRow(18).getCell(1).setText("");
                                                                                                                                                tables.getRow(18).getCell(2).setText("NA");
                                                                                                                                                tables.getRow(19).getCell(0).setText("");
                                                                                                                                                tables.getRow(19).getCell(1).setText("");
                                                                                                                                                tables.getRow(19).getCell(2).setText("NA");
                                                                                                                                                tables.getRow(20).getCell(0).setText("");
                                                                                                                                                tables.getRow(20).getCell(1).setText("");
                                                                                                                                                tables.getRow(20).getCell(2).setText("NA");
                                                                                                                                                tables.getRow(21).getCell(0).setText("");
                                                                                                                                                tables.getRow(21).getCell(1).setText("");
                                                                                                                                                tables.getRow(21).getCell(2).setText("NA");
                                                                                                                                                tables.getRow(22).getCell(0).setText("");
                                                                                                                                                tables.getRow(22).getCell(1).setText("");
                                                                                                                                                tables.getRow(22).getCell(2).setText("NA");
                                                                                                                                                tables.getRow(23).getCell(0).setText("");
                                                                                                                                                tables.getRow(23).getCell(1).setText("");
                                                                                                                                                tables.getRow(23).getCell(2).setText("NA");
                                                                                                                                                XWPFParagraph para10 = document.createParagraph();
                                                                                                                                                XWPFRun ru10 = para10.createRun();
                                                                                                                                                ru10.addBreak(BreakType.PAGE);
                                                                                                                                                
                                                                                                                                                int tRows = 1;
                                                                                                                                                int tCols = 0;
                                                                                                                                                XWPFTable tablet = document.createTable(tRows, tCols);
                                                                                                                                                tablet.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(9500));
                                                                                                                                                tablet.getCTTbl().getTblPr().unsetTblBorders();
                                                                                                                                                List<XWPFTableRow> rowst = tablet.getRows();
                                                                                                                                                int rowCtt = 0;
                                                                                                                                                for (XWPFTableRow rowt : rowst) {
                                                                                                                                                    CTTrPr trPrt = rowt.getCtRow().addNewTrPr();
                                                                                                                                                    CTHeight htt = trPrt.addNewTrHeight();
                                                                                                                                                    htt.setVal(BigInteger.valueOf(453));
                                                                                                                                                    List<XWPFTableCell> cellst = rowt.getTableCells();
                                                                                                                                                    for (XWPFTableCell cellt : cellst) {
                                                                                                                                                        CTTcPr tcprt = cellt.getCTTc().addNewTcPr();
                                                                                                                                                        CTVerticalJc vat = tcprt.addNewVAlign();
                                                                                                                                                        vat.setVal(STVerticalJc.CENTER);
                                                                                                                                                        CTShd ctshdt = tcprt.addNewShd();
                                                                                                                                                        ctshdt.setColor("auto");
                                                                                                                                                        ctshdt.setVal(STShd.CLEAR);
                                                                                                                                                        if (rowCtt == 0) {
                                                                                                                                                            ctshdt.setFill("FFFFFF");
                                                                                                                                                        } else {
                                                                                                                                                            ctshdt.setFill("FFFFFF");
                                                                                                                                                        }
                                                                                                                                                        XWPFParagraph parat = cellt.getParagraphs().get(0);
                                                                                                                                                        XWPFRun rht = parat.createRun();
                                                                                                                                                        rht.setFontSize(10);
                                                                                                                                                        rht.setBold(true);
                                                                                                                                                        rht.setColor("99284C");
                                                                                                                                                        rht.setText("DR sync information");
                                                                                                                                                        
                                                                                                                                                        int t2Rows = 1;
                                                                                                                                                        int t2Cols = 0;
                                                                                                                                                        XWPFTable tablet2 = document.createTable(t2Rows, t2Cols);
                                                                                                                                                        tablet2.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(9500));
                                                                                                                                                        tablet2.getCTTbl().getTblPr().unsetTblBorders();
                                                                                                                                                        List<XWPFTableRow> rowst2 = tablet2.getRows();
                                                                                                                                                        int rowCtt2 = 0;
                                                                                                                                                        for (XWPFTableRow rowt2 : rowst2) {
                                                                                                                                                            CTTrPr trPrt2 = rowt2.getCtRow().addNewTrPr();
                                                                                                                                                            CTHeight htt2 = trPrt2.addNewTrHeight();
                                                                                                                                                            htt2.setVal(BigInteger.valueOf(453));
                                                                                                                                                            List<XWPFTableCell> cellst2 = rowt2.getTableCells();
                                                                                                                                                            for (XWPFTableCell cellt2 : cellst2) {
                                                                                                                                                                CTTcPr tcprt2 = cellt2.getCTTc().addNewTcPr();
                                                                                                                                                                CTVerticalJc vat2 = tcprt2.addNewVAlign();
                                                                                                                                                                vat2.setVal(STVerticalJc.CENTER);
                                                                                                                                                                CTShd ctshdt2 = tcprt2.addNewShd();
                                                                                                                                                                ctshdt2.setColor("auto");
                                                                                                                                                                ctshdt2.setVal(STShd.CLEAR);
                                                                                                                                                                if (rowCtt2 == 0) {
                                                                                                                                                                    ctshdt2.setFill("FFFFFF");
                                                                                                                                                                } else {
                                                                                                                                                                    ctshdt2.setFill("FFFFFF");
                                                                                                                                                                }
                                                                                                                                                                XWPFParagraph parat2 = cellt2.getParagraphs().get(0);
                                                                                                                                                                XWPFRun rht2 = parat2.createRun();
                                                                                                                                                                rht2.setFontSize(10);
                                                                                                                                                                rht2.setText("The table below shows the DR & DC sync information.");
                                                                                                                                                                
                                                                                                                                                                int t1Rows = 1;
                                                                                                                                                                int t1Cols = 0;
                                                                                                                                                                XWPFTable tablet1 = document.createTable(t1Rows, t1Cols);
                                                                                                                                                                tablet1.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(9500));
                                                                                                                                                                tablet1.getCTTbl().getTblPr().unsetTblBorders();
                                                                                                                                                                List<XWPFTableRow> rowst1 = tablet1.getRows();
                                                                                                                                                                int rowCtt1 = 0;
                                                                                                                                                                for (XWPFTableRow rowt1 : rowst1) {
                                                                                                                                                                    CTTrPr trPrt1 = rowt1.getCtRow().addNewTrPr();
                                                                                                                                                                    CTHeight htt1 = trPrt1.addNewTrHeight();
                                                                                                                                                                    htt1.setVal(BigInteger.valueOf(300));
                                                                                                                                                                    List<XWPFTableCell> cellst1 = rowt1.getTableCells();
                                                                                                                                                                    for (XWPFTableCell cellt1 : cellst1) {
                                                                                                                                                                        CTTcPr tcprt1 = cellt1.getCTTc().addNewTcPr();
                                                                                                                                                                        CTVerticalJc vat1 = tcprt1.addNewVAlign();
                                                                                                                                                                        vat1.setVal(STVerticalJc.CENTER);
                                                                                                                                                                        CTShd ctshdt1 = tcprt1.addNewShd();
                                                                                                                                                                        ctshdt1.setColor("auto");
                                                                                                                                                                        ctshdt1.setVal(STShd.CLEAR);
                                                                                                                                                                        if (rowCtt1 == 0) {
                                                                                                                                                                            ctshdt1.setFill("FFFFFF");
                                                                                                                                                                            
                                                                                                                                                                        } else {
                                                                                                                                                                            ctshdt1.setFill("FFFFFF");
                                                                                                                                                                        }
                                                                                                                                                                        tablet1.getRow(0).getCell(0).setText("");
                                                                                                                                                                        
                                                                                                                                                                        int uRows = 2;
                                                                                                                                                                        int uCols = 3;
                                                                                                                                                                        XWPFTable tableu = document.createTable(uRows, uCols);
                                                                                                                                                                        tableu.setWidth(2000);
                                                                                                                                                                        tableu.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(3200));
                                                                                                                                                                        tableu.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(3200));
                                                                                                                                                                        tableu.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(3100));
                                                                                                                                                                        
                                                                                                                                                                        List<XWPFTableRow> rowsu = tableu.getRows();
                                                                                                                                                                        int rowCtu = 0;
                                                                                                                                                                        int colCtu = 0;
                                                                                                                                                                        for (XWPFTableRow rowu : rowsu) {
                                                                                                                                                                            CTTrPr trPru = rowu.getCtRow().addNewTrPr();
                                                                                                                                                                            CTHeight htu = trPru.addNewTrHeight();
                                                                                                                                                                            htu.setVal(BigInteger.valueOf(330));
                                                                                                                                                                            List<XWPFTableCell> cellsu = rowu.getTableCells();
                                                                                                                                                                            for (XWPFTableCell cellu : cellsu) {
                                                                                                                                                                                CTTcPr tcpru = cellu.getCTTc().addNewTcPr();
                                                                                                                                                                                CTVerticalJc vau = tcpru.addNewVAlign();
                                                                                                                                                                                vau.setVal(STVerticalJc.CENTER);
                                                                                                                                                                                CTShd ctshdu = tcpru.addNewShd();
                                                                                                                                                                                ctshdu.setColor("auto");
                                                                                                                                                                                ctshdu.setVal(STShd.CLEAR);
                                                                                                                                                                                if (rowCtu == 0) {
                                                                                                                                                                                    ctshdu.setFill("003366");
                                                                                                                                                                                } else {
                                                                                                                                                                                    ctshdu.setFill("FFFFFF");
                                                                                                                                                                                }
                                                                                                                                                                                XWPFParagraph parau = cellu.getParagraphs().get(0);
                                                                                                                                                                                XWPFRun rhu = parau.createRun();
                                                                                                                                                                                if (rowCtu == 0) {
                                                                                                                                                                                    rhu.setText(" ");
                                                                                                                                                                                    rhu.setFontSize(10);
                                                                                                                                                                                    rhu.setBold(true);
                                                                                                                                                                                    parau.setAlignment(ParagraphAlignment.CENTER);
                                                                                                                                                                                } else {
                                                                                                                                                                                    rhu.setText("");
                                                                                                                                                                                    parau.setAlignment(ParagraphAlignment.CENTER);
                                                                                                                                                                                }
                                                                                                                                                                                colCtu++;
                                                                                                                                                                            }
                                                                                                                                                                            colCtu = 0;
                                                                                                                                                                            rowCtu++;
                                                                                                                                                                        }
                                                                                                                                                                        tableu.getRow(0).getCell(0).setText(" Max sequence archived log in DC");
                                                                                                                                                                        tableu.getRow(0).getCell(1).setText("Max sequence archived log in DR");
                                                                                                                                                                        tableu.getRow(0).getCell(2).setText(" Remarks ");
                                                                                                                                                                        tableu.getRow(1).getCell(0).setText("");
                                                                                                                                                                        tableu.getRow(1).getCell(1).setText("");
                                                                                                                                                                        tableu.getRow(1).getCell(2).setText("Approximately 1-3 archive log sequence difference may occur between every 30 minutes which will be updated in every SYNC script.");
                                                                                                                                                                        
                                                                                                                                                                        XWPFParagraph para11 = document.createParagraph();
                                                                                                                                                                        XWPFRun ru11 = para11.createRun();
                                                                                                                                                                        ru11.addBreak();
                                                                                                                                                                        ru11.addBreak();
                                                                                                                                                                        
                                                                                                                                                                        int vRows = 1;
                                                                                                                                                                        int vCols = 0;
                                                                                                                                                                        XWPFTable tablev = document.createTable(vRows, vCols);
                                                                                                                                                                        tablev.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(9500));
                                                                                                                                                                        tablev.getCTTbl().getTblPr().unsetTblBorders();
                                                                                                                                                                        List<XWPFTableRow> rowsv = tablev.getRows();
                                                                                                                                                                        int rowCtv = 0;
                                                                                                                                                                        for (XWPFTableRow rowv : rowsv) {
                                                                                                                                                                            CTTrPr trPrv = rowv.getCtRow().addNewTrPr();
                                                                                                                                                                            CTHeight htv = trPrv.addNewTrHeight();
                                                                                                                                                                            htv.setVal(BigInteger.valueOf(453));
                                                                                                                                                                            List<XWPFTableCell> cellsv = rowv.getTableCells();
                                                                                                                                                                            for (XWPFTableCell cellv : cellsv) {
                                                                                                                                                                                CTTcPr tcprv = cellv.getCTTc().addNewTcPr();
                                                                                                                                                                                CTVerticalJc vav = tcprv.addNewVAlign();
                                                                                                                                                                                vav.setVal(STVerticalJc.CENTER);
                                                                                                                                                                                CTShd ctshdv = tcprv.addNewShd();
                                                                                                                                                                                ctshdv.setColor("auto");
                                                                                                                                                                                ctshdv.setVal(STShd.CLEAR);
                                                                                                                                                                                if (rowCtv == 0) {
                                                                                                                                                                                    ctshdv.setFill("FFFFFF");
                                                                                                                                                                                } else {
                                                                                                                                                                                    ctshdv.setFill("FFFFFF");
                                                                                                                                                                                }
                                                                                                                                                                                XWPFParagraph parav = cellv.getParagraphs().get(0);
                                                                                                                                                                                XWPFRun rhv = parav.createRun();
                                                                                                                                                                                rhv.setFontSize(10);
                                                                                                                                                                                rhv.setBold(true);
                                                                                                                                                                                rhv.setColor("99284C");
                                                                                                                                                                                rhv.setText("Archive log destination details");
                                                                                                                                                                                
                                                                                                                                                                                int v2Rows = 1;
                                                                                                                                                                                int v2Cols = 0;
                                                                                                                                                                                XWPFTable tablev2 = document.createTable(v2Rows, v2Cols);
                                                                                                                                                                                tablev2.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(9500));
                                                                                                                                                                                tablev2.getCTTbl().getTblPr().unsetTblBorders();
                                                                                                                                                                                List<XWPFTableRow> rowsv2 = tablev2.getRows();
                                                                                                                                                                                int rowCtv2 = 0;
                                                                                                                                                                                for (XWPFTableRow rowv2 : rowsv2) {
                                                                                                                                                                                    CTTrPr trPrv2 = rowv2.getCtRow().addNewTrPr();
                                                                                                                                                                                    CTHeight htv2 = trPrv2.addNewTrHeight();
                                                                                                                                                                                    htv2.setVal(BigInteger.valueOf(453));
                                                                                                                                                                                    List<XWPFTableCell> cellsv2 = rowv2.getTableCells();
                                                                                                                                                                                    for (XWPFTableCell cellv2 : cellsv2) {
                                                                                                                                                                                        CTTcPr tcprv2 = cellv2.getCTTc().addNewTcPr();
                                                                                                                                                                                        CTVerticalJc vav2 = tcprv2.addNewVAlign();
                                                                                                                                                                                        vav2.setVal(STVerticalJc.CENTER);
                                                                                                                                                                                        CTShd ctshdv2 = tcprv2.addNewShd();
                                                                                                                                                                                        ctshdv2.setColor("auto");
                                                                                                                                                                                        ctshdv2.setVal(STShd.CLEAR);
                                                                                                                                                                                        if (rowCtv2 == 0) {
                                                                                                                                                                                            ctshdv2.setFill("FFFFFF");
                                                                                                                                                                                        } else {
                                                                                                                                                                                            ctshdv2.setFill("FFFFFF");
                                                                                                                                                                                        }
                                                                                                                                                                                        XWPFParagraph parav2 = cellv2.getParagraphs().get(0);
                                                                                                                                                                                        XWPFRun rhv2 = parav2.createRun();
                                                                                                                                                                                        rhv2.setFontSize(10);
                                                                                                                                                                                        rhv2.setText("The below table provides the archive log destination details.");
                                                                                                                                                                                        
                                                                                                                                                                                        int v1Rows = 1;
                                                                                                                                                                                        int v1Cols = 0;
                                                                                                                                                                                        XWPFTable tablev1 = document.createTable(v1Rows, v1Cols);
                                                                                                                                                                                        tablev1.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(9500));
                                                                                                                                                                                        tablev1.getCTTbl().getTblPr().unsetTblBorders();
                                                                                                                                                                                        List<XWPFTableRow> rowsv1 = tablev1.getRows();
                                                                                                                                                                                        int rowCtv1 = 0;
                                                                                                                                                                                        for (XWPFTableRow rowv1 : rowsv1) {
                                                                                                                                                                                            CTTrPr trPrv1 = rowv1.getCtRow().addNewTrPr();
                                                                                                                                                                                            CTHeight htv1 = trPrv1.addNewTrHeight();
                                                                                                                                                                                            htv1.setVal(BigInteger.valueOf(300));
                                                                                                                                                                                            List<XWPFTableCell> cellsv1 = rowv1.getTableCells();
                                                                                                                                                                                            for (XWPFTableCell cellv1 : cellsv1) {
                                                                                                                                                                                                CTTcPr tcprv1 = cellv1.getCTTc().addNewTcPr();
                                                                                                                                                                                                CTVerticalJc vav1 = tcprv1.addNewVAlign();
                                                                                                                                                                                                vav1.setVal(STVerticalJc.CENTER);
                                                                                                                                                                                                CTShd ctshdv1 = tcprv1.addNewShd();
                                                                                                                                                                                                ctshdv1.setColor("auto");
                                                                                                                                                                                                ctshdv1.setVal(STShd.CLEAR);
                                                                                                                                                                                                if (rowCtv1 == 0) {
                                                                                                                                                                                                    ctshdv1.setFill("FFFFFF");
                                                                                                                                                                                                    
                                                                                                                                                                                                } else {
                                                                                                                                                                                                    ctshdv1.setFill("FFFFFF");
                                                                                                                                                                                                }
                                                                                                                                                                                                tablev1.getRow(0).getCell(0).setText("");
                                                                                                                                                                                                
                                                                                                                                                                                                int v3Rows = 2;
                                                                                                                                                                                                int v3Cols = 4;
                                                                                                                                                                                                XWPFTable tablev3 = document.createTable(v3Rows, v3Cols);
                                                                                                                                                                                                tablev3.setWidth(2000);
                                                                                                                                                                                                tablev3.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(2375));
                                                                                                                                                                                                tablev3.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(2375));
                                                                                                                                                                                                tablev3.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(2375));
                                                                                                                                                                                                tablev3.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(2375));
                                                                                                                                                                                                List<XWPFTableRow> rowsv3 = tablev3.getRows();
                                                                                                                                                                                                int rowCtv3 = 0;
                                                                                                                                                                                                int colCtv3 = 0;
                                                                                                                                                                                                for (XWPFTableRow rowv3 : rowsv3) {
                                                                                                                                                                                                    CTTrPr trPrv3 = rowv3.getCtRow().addNewTrPr();
                                                                                                                                                                                                    CTHeight htv3 = trPrv3.addNewTrHeight();
                                                                                                                                                                                                    htv3.setVal(BigInteger.valueOf(330));
                                                                                                                                                                                                    List<XWPFTableCell> cellsv3 = rowv3.getTableCells();
                                                                                                                                                                                                    for (XWPFTableCell cellv3 : cellsv3) {
                                                                                                                                                                                                        CTTcPr tcprv3 = cellv3.getCTTc().addNewTcPr();
                                                                                                                                                                                                        CTVerticalJc vav3 = tcprv3.addNewVAlign();
                                                                                                                                                                                                        vav3.setVal(STVerticalJc.CENTER);
                                                                                                                                                                                                        CTShd ctshdv3 = tcprv3.addNewShd();
                                                                                                                                                                                                        ctshdv3.setColor("auto");
                                                                                                                                                                                                        ctshdv3.setVal(STShd.CLEAR);
                                                                                                                                                                                                        if (rowCtv3 == 0) {
                                                                                                                                                                                                            ctshdv3.setFill("003366");
                                                                                                                                                                                                        } else {
                                                                                                                                                                                                            ctshdv3.setFill("FFFFFF");
                                                                                                                                                                                                        }
                                                                                                                                                                                                        XWPFParagraph parav3 = cellv3.getParagraphs().get(0);
                                                                                                                                                                                                        XWPFRun rhv3 = parav3.createRun();
                                                                                                                                                                                                        if (rowCtv3 == 0) {
                                                                                                                                                                                                            rhv3.setText(" ");
                                                                                                                                                                                                            rhv3.setFontSize(10);
                                                                                                                                                                                                            rhv3.setBold(true);
                                                                                                                                                                                                            parav3.setAlignment(ParagraphAlignment.CENTER);
                                                                                                                                                                                                        } else {
                                                                                                                                                                                                            rhv3.setText("");
                                                                                                                                                                                                            parav3.setAlignment(ParagraphAlignment.CENTER);
                                                                                                                                                                                                        }
                                                                                                                                                                                                        colCtv3++;
                                                                                                                                                                                                    }
                                                                                                                                                                                                    colCtv3 = 0;
                                                                                                                                                                                                    rowCtv3++;
                                                                                                                                                                                                }
                                                                                                                                                                                                tablev3.getRow(0).getCell(0).setText(" DC archive log destination");
                                                                                                                                                                                                tablev3.getRow(0).getCell(1).setText("Free space available in drive");
                                                                                                                                                                                                tablev3.getRow(0).getCell(2).setText(" DR archive log destination ");
                                                                                                                                                                                                tablev3.getRow(0).getCell(3).setText(" Free space available in drive ");
                                                                                                                                                                                                tablev3.getRow(1).getCell(0).setText("");
                                                                                                                                                                                                tablev3.getRow(1).getCell(1).setText("");
                                                                                                                                                                                                tablev3.getRow(1).getCell(2).setText("");
                                                                                                                                                                                                tablev3.getRow(1).getCell(3).setText("");
                                                                                                                                                                                                
                                                                                                                                                                                                XWPFParagraph para12 = document.createParagraph();
                                                                                                                                                                                                XWPFRun ru12 = para12.createRun();
                                                                                                                                                                                                ru12.addBreak();
                                                                                                                                                                                                ru12.addBreak();
                                                                                                                                                                                                int wRows = 1;
                                                                                                                                                                                                int wCols = 0;
                                                                                                                                                                                                XWPFTable tablew = document.createTable(wRows, wCols);
                                                                                                                                                                                                tablew.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(9500));
                                                                                                                                                                                                tablew.getCTTbl().getTblPr().unsetTblBorders();
                                                                                                                                                                                                List<XWPFTableRow> rowsw = tablew.getRows();
                                                                                                                                                                                                int rowCtw = 0;
                                                                                                                                                                                                for (XWPFTableRow roww : rowsw) {
                                                                                                                                                                                                    CTTrPr trPrw = roww.getCtRow().addNewTrPr();
                                                                                                                                                                                                    CTHeight htw = trPrw.addNewTrHeight();
                                                                                                                                                                                                    htw.setVal(BigInteger.valueOf(453));
                                                                                                                                                                                                    List<XWPFTableCell> cellsw = roww.getTableCells();
                                                                                                                                                                                                    for (XWPFTableCell cellw : cellsw) {
                                                                                                                                                                                                        CTTcPr tcprw = cellw.getCTTc().addNewTcPr();
                                                                                                                                                                                                        CTVerticalJc vaw = tcprw.addNewVAlign();
                                                                                                                                                                                                        vaw.setVal(STVerticalJc.CENTER);
                                                                                                                                                                                                        CTShd ctshdw = tcprw.addNewShd();
                                                                                                                                                                                                        ctshdw.setColor("auto");
                                                                                                                                                                                                        ctshdw.setVal(STShd.CLEAR);
                                                                                                                                                                                                        if (rowCtw == 0) {
                                                                                                                                                                                                            ctshdw.setFill("FFFFFF");
                                                                                                                                                                                                        } else {
                                                                                                                                                                                                            ctshdw.setFill("FFFFFF");
                                                                                                                                                                                                        }
                                                                                                                                                                                                        XWPFParagraph paraw = cellw.getParagraphs().get(0);
                                                                                                                                                                                                        XWPFRun rhw = paraw.createRun();
                                                                                                                                                                                                        rhw.setFontSize(10);
                                                                                                                                                                                                        rhw.setBold(true);
                                                                                                                                                                                                        rhw.setColor("99284C");
                                                                                                                                                                                                        rhw.setText("Alert log analysis");
                                                                                                                                                                                                        
                                                                                                                                                                                                        int w2Rows = 1;
                                                                                                                                                                                                        int w2Cols = 0;
                                                                                                                                                                                                        XWPFTable tablew2 = document.createTable(w2Rows, w2Cols);
                                                                                                                                                                                                        tablew2.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(9500));
                                                                                                                                                                                                        tablew2.getCTTbl().getTblPr().unsetTblBorders();
                                                                                                                                                                                                        List<XWPFTableRow> rowsw2 = tablew2.getRows();
                                                                                                                                                                                                        int rowCtw2 = 0;
                                                                                                                                                                                                        for (XWPFTableRow roww2 : rowsw2) {
                                                                                                                                                                                                            CTTrPr trPrw2 = roww2.getCtRow().addNewTrPr();
                                                                                                                                                                                                            CTHeight htw2 = trPrw2.addNewTrHeight();
                                                                                                                                                                                                            htw2.setVal(BigInteger.valueOf(453));
                                                                                                                                                                                                            List<XWPFTableCell> cellsw2 = roww2.getTableCells();
                                                                                                                                                                                                            for (XWPFTableCell cellw2 : cellsw2) {
                                                                                                                                                                                                                CTTcPr tcprw2 = cellw2.getCTTc().addNewTcPr();
                                                                                                                                                                                                                CTVerticalJc vaw2 = tcprw2.addNewVAlign();
                                                                                                                                                                                                                vaw2.setVal(STVerticalJc.CENTER);
                                                                                                                                                                                                                CTShd ctshdw2 = tcprw2.addNewShd();
                                                                                                                                                                                                                ctshdw2.setColor("auto");
                                                                                                                                                                                                                ctshdw2.setVal(STShd.CLEAR);
                                                                                                                                                                                                                if (rowCtw2 == 0) {
                                                                                                                                                                                                                    ctshdw2.setFill("FFFFFF");
                                                                                                                                                                                                                } else {
                                                                                                                                                                                                                    ctshdw2.setFill("FFFFFF");
                                                                                                                                                                                                                }
                                                                                                                                                                                                                XWPFParagraph paraw2 = cellw2.getParagraphs().get(0);
                                                                                                                                                                                                                XWPFRun rhw2 = paraw2.createRun();
                                                                                                                                                                                                                rhw2.setFontSize(10);
                                                                                                                                                                                                                rhw2.setText("The table below lists all the ORA-xxxxx errors presented in alert log, impact and action taken.");
                                                                                                                                                                                                                
                                                                                                                                                                                                                int w1Rows = 1;
                                                                                                                                                                                                                int w1Cols = 0;
                                                                                                                                                                                                                XWPFTable tablew1 = document.createTable(w1Rows, w1Cols);
                                                                                                                                                                                                                tablew1.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(9500));
                                                                                                                                                                                                                tablew1.getCTTbl().getTblPr().unsetTblBorders();
                                                                                                                                                                                                                List<XWPFTableRow> rowsw1 = tablew1.getRows();
                                                                                                                                                                                                                int rowCtw1 = 0;
                                                                                                                                                                                                                for (XWPFTableRow roww1 : rowsw1) {
                                                                                                                                                                                                                    CTTrPr trPrw1 = roww1.getCtRow().addNewTrPr();
                                                                                                                                                                                                                    CTHeight htw1 = trPrw1.addNewTrHeight();
                                                                                                                                                                                                                    htw1.setVal(BigInteger.valueOf(300));
                                                                                                                                                                                                                    List<XWPFTableCell> cellsw1 = roww1.getTableCells();
                                                                                                                                                                                                                    for (XWPFTableCell cellw1 : cellsw1) {
                                                                                                                                                                                                                        CTTcPr tcprw1 = cellw1.getCTTc().addNewTcPr();
                                                                                                                                                                                                                        CTVerticalJc vaw1 = tcprw1.addNewVAlign();
                                                                                                                                                                                                                        vaw1.setVal(STVerticalJc.CENTER);
                                                                                                                                                                                                                        CTShd ctshdw1 = tcprv1.addNewShd();
                                                                                                                                                                                                                        ctshdw1.setColor("auto");
                                                                                                                                                                                                                        ctshdw1.setVal(STShd.CLEAR);
                                                                                                                                                                                                                        if (rowCtw1 == 0) {
                                                                                                                                                                                                                            ctshdw1.setFill("FFFFFF");
                                                                                                                                                                                                                            
                                                                                                                                                                                                                        } else {
                                                                                                                                                                                                                            ctshdw1.setFill("FFFFFF");
                                                                                                                                                                                                                        }
                                                                                                                                                                                                                        tablew1.getRow(0).getCell(0).setText("");
                                                                                                                                                                                                                        
                                                                                                                                                                                                                        int w3Rows = 2;
                                                                                                                                                                                                                        int w3Cols = 5;
                                                                                                                                                                                                                        XWPFTable tablew3 = document.createTable(w3Rows, w3Cols);
                                                                                                                                                                                                                        tablew3.setWidth(2000);
                                                                                                                                                                                                                        tablew3.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(1900));
                                                                                                                                                                                                                        tablew3.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(1900));
                                                                                                                                                                                                                        tablew3.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(1900));
                                                                                                                                                                                                                        tablew3.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(1900));
                                                                                                                                                                                                                        tablew3.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(1900));
                                                                                                                                                                                                                        List<XWPFTableRow> rowsw3 = tablew3.getRows();
                                                                                                                                                                                                                        int rowCtw3 = 0;
                                                                                                                                                                                                                        int colCtw3 = 0;
                                                                                                                                                                                                                        for (XWPFTableRow roww3 : rowsw3) {
                                                                                                                                                                                                                            CTTrPr trPrw3 = roww3.getCtRow().addNewTrPr();
                                                                                                                                                                                                                            CTHeight htw3 = trPrw3.addNewTrHeight();
                                                                                                                                                                                                                            htw3.setVal(BigInteger.valueOf(330));
                                                                                                                                                                                                                            List<XWPFTableCell> cellsw3 = roww3.getTableCells();
                                                                                                                                                                                                                            for (XWPFTableCell cellw3 : cellsw3) {
                                                                                                                                                                                                                                CTTcPr tcprw3 = cellw3.getCTTc().addNewTcPr();
                                                                                                                                                                                                                                CTVerticalJc vaw3 = tcprw3.addNewVAlign();
                                                                                                                                                                                                                                vaw3.setVal(STVerticalJc.CENTER);
                                                                                                                                                                                                                                CTShd ctshdw3 = tcprw3.addNewShd();
                                                                                                                                                                                                                                ctshdw3.setColor("auto");
                                                                                                                                                                                                                                ctshdw3.setVal(STShd.CLEAR);
                                                                                                                                                                                                                                if (rowCtw3 == 0) {
                                                                                                                                                                                                                                    ctshdw3.setFill("003366");
                                                                                                                                                                                                                                } else {
                                                                                                                                                                                                                                    ctshdw3.setFill("FFFFFF");
                                                                                                                                                                                                                                }
                                                                                                                                                                                                                                XWPFParagraph paraw3 = cellw3.getParagraphs().get(0);
                                                                                                                                                                                                                                XWPFRun rhw3 = paraw3.createRun();
                                                                                                                                                                                                                                if (rowCtw3 == 0) {
                                                                                                                                                                                                                                    rhw3.setText(" ");
                                                                                                                                                                                                                                    rhw3.setFontSize(10);
                                                                                                                                                                                                                                    rhw3.setBold(true);
                                                                                                                                                                                                                                    paraw3.setAlignment(ParagraphAlignment.CENTER);
                                                                                                                                                                                                                                } else {
                                                                                                                                                                                                                                    rhw3.setText("");
                                                                                                                                                                                                                                    paraw3.setAlignment(ParagraphAlignment.CENTER);
                                                                                                                                                                                                                                }
                                                                                                                                                                                                                                colCtw3++;
                                                                                                                                                                                                                            }
                                                                                                                                                                                                                            colCtw3 = 0;
                                                                                                                                                                                                                            rowCtw3++;
                                                                                                                                                                                                                        }
                                                                                                                                                                                                                        tablew3.getRow(0).getCell(0).setText(" S.No");
                                                                                                                                                                                                                        tablew3.getRow(0).getCell(1).setText("ORA-xxxxx error");
                                                                                                                                                                                                                        tablew3.getRow(0).getCell(2).setText(" Error description ");
                                                                                                                                                                                                                        tablew3.getRow(0).getCell(3).setText(" Impact ");
                                                                                                                                                                                                                        tablew3.getRow(0).getCell(4).setText(" Action taken. ");
                                                                                                                                                                                                                        tablew3.getRow(1).getCell(0).setText("Nil");
                                                                                                                                                                                                                        tablew3.getRow(1).getCell(1).setText("Nil");
                                                                                                                                                                                                                        tablew3.getRow(1).getCell(2).setText("Nil");
                                                                                                                                                                                                                        tablew3.getRow(1).getCell(3).setText("Nil");
                                                                                                                                                                                                                        tablew3.getRow(1).getCell(4).setText("Nil");
                                                                                                                                                                                                                        
                                                                                                                                                                                                                        XWPFParagraph para13 = document.createParagraph();
                                                                                                                                                                                                                        XWPFRun ru13 = para13.createRun();
                                                                                                                                                                                                                        ru13.addBreak(BreakType.PAGE);
                                                                                                                                                                                                                        
                                                                                                                                                                                                                        int xRows = 1;
                                                                                                                                                                                                                        int xCols = 0;
                                                                                                                                                                                                                        XWPFTable tablex = document.createTable(xRows, xCols);
                                                                                                                                                                                                                        tablex.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(9500));
                                                                                                                                                                                                                        tablex.getCTTbl().getTblPr().unsetTblBorders();
                                                                                                                                                                                                                        List<XWPFTableRow> rowsx = tablex.getRows();
                                                                                                                                                                                                                        int rowCtx = 0;
                                                                                                                                                                                                                        for (XWPFTableRow rowx : rowsx) {
                                                                                                                                                                                                                            CTTrPr trPrx = rowx.getCtRow().addNewTrPr();
                                                                                                                                                                                                                            CTHeight htx = trPrx.addNewTrHeight();
                                                                                                                                                                                                                            htx.setVal(BigInteger.valueOf(453));
                                                                                                                                                                                                                            List<XWPFTableCell> cellsx = rowx.getTableCells();
                                                                                                                                                                                                                            for (XWPFTableCell cellx : cellsx) {
                                                                                                                                                                                                                                CTTcPr tcprx = cellx.getCTTc().addNewTcPr();
                                                                                                                                                                                                                                CTVerticalJc vax = tcprx.addNewVAlign();
                                                                                                                                                                                                                                vax.setVal(STVerticalJc.CENTER);
                                                                                                                                                                                                                                CTShd ctshdx = tcprx.addNewShd();
                                                                                                                                                                                                                                ctshdx.setColor("auto");
                                                                                                                                                                                                                                ctshdx.setVal(STShd.CLEAR);
                                                                                                                                                                                                                                if (rowCtx == 0) {
                                                                                                                                                                                                                                    ctshdx.setFill("FFFFFF");
                                                                                                                                                                                                                                } else {
                                                                                                                                                                                                                                    ctshdx.setFill("FFFFFF");
                                                                                                                                                                                                                                }
                                                                                                                                                                                                                                XWPFParagraph parax = cellx.getParagraphs().get(0);
                                                                                                                                                                                                                                XWPFRun rhx = parax.createRun();
                                                                                                                                                                                                                                rhx.setFontSize(10);
                                                                                                                                                                                                                                rhx.setBold(true);
                                                                                                                                                                                                                                rhx.setColor("99284C");
                                                                                                                                                                                                                                rhx.setText("Deleting/Moving of files and folder information");
                                                                                                                                                                                                                                
                                                                                                                                                                                                                                int x2Rows = 1;
                                                                                                                                                                                                                                int x2Cols = 0;
                                                                                                                                                                                                                                XWPFTable tablex2 = document.createTable(x2Rows, x2Cols);
                                                                                                                                                                                                                                tablex2.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(9500));
                                                                                                                                                                                                                                tablex2.getCTTbl().getTblPr().unsetTblBorders();
                                                                                                                                                                                                                                List<XWPFTableRow> rowsx2 = tablex2.getRows();
                                                                                                                                                                                                                                int rowCtx2 = 0;
                                                                                                                                                                                                                                for (XWPFTableRow rowx2 : rowsx2) {
                                                                                                                                                                                                                                    CTTrPr trPrx2 = rowx2.getCtRow().addNewTrPr();
                                                                                                                                                                                                                                    CTHeight htx2 = trPrx2.addNewTrHeight();
                                                                                                                                                                                                                                    htx2.setVal(BigInteger.valueOf(453));
                                                                                                                                                                                                                                    List<XWPFTableCell> cellsx2 = rowx2.getTableCells();
                                                                                                                                                                                                                                    for (XWPFTableCell cellx2 : cellsx2) {
                                                                                                                                                                                                                                        CTTcPr tcprx2 = cellx2.getCTTc().addNewTcPr();
                                                                                                                                                                                                                                        CTVerticalJc vax2 = tcprx2.addNewVAlign();
                                                                                                                                                                                                                                        vax2.setVal(STVerticalJc.CENTER);
                                                                                                                                                                                                                                        CTShd ctshdx2 = tcprx2.addNewShd();
                                                                                                                                                                                                                                        ctshdx2.setColor("auto");
                                                                                                                                                                                                                                        ctshdx2.setVal(STShd.CLEAR);
                                                                                                                                                                                                                                        if (rowCtx2 == 0) {
                                                                                                                                                                                                                                            ctshdx2.setFill("FFFFFF");
                                                                                                                                                                                                                                        } else {
                                                                                                                                                                                                                                            ctshdx2.setFill("FFFFFF");
                                                                                                                                                                                                                                        }
                                                                                                                                                                                                                                        XWPFParagraph parax2 = cellx2.getParagraphs().get(0);
                                                                                                                                                                                                                                        XWPFRun rhx2 = parax2.createRun();
                                                                                                                                                                                                                                        rhx2.setFontSize(10);
                                                                                                                                                                                                                                        rhx2.setText("The table below lists the deleted/moved files information.");
                                                                                                                                                                                                                                        
                                                                                                                                                                                                                                        int x1Rows = 1;
                                                                                                                                                                                                                                        int x1Cols = 0;
                                                                                                                                                                                                                                        XWPFTable tablex1 = document.createTable(x1Rows, x1Cols);
                                                                                                                                                                                                                                        tablex1.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(9500));
                                                                                                                                                                                                                                        tablex1.getCTTbl().getTblPr().unsetTblBorders();
                                                                                                                                                                                                                                        List<XWPFTableRow> rowsx1 = tablex1.getRows();
                                                                                                                                                                                                                                        int rowCtx1 = 0;
                                                                                                                                                                                                                                        for (XWPFTableRow rowx1 : rowsx1) {
                                                                                                                                                                                                                                            CTTrPr trPrx1 = rowx1.getCtRow().addNewTrPr();
                                                                                                                                                                                                                                            CTHeight htx1 = trPrx1.addNewTrHeight();
                                                                                                                                                                                                                                            htx1.setVal(BigInteger.valueOf(300));
                                                                                                                                                                                                                                            List<XWPFTableCell> cellsx1 = rowx1.getTableCells();
                                                                                                                                                                                                                                            for (XWPFTableCell cellx1 : cellsx1) {
                                                                                                                                                                                                                                                CTTcPr tcprx1 = cellx1.getCTTc().addNewTcPr();
                                                                                                                                                                                                                                                CTVerticalJc vax1 = tcprx1.addNewVAlign();
                                                                                                                                                                                                                                                vax1.setVal(STVerticalJc.CENTER);
                                                                                                                                                                                                                                                CTShd ctshdx1 = tcprx1.addNewShd();
                                                                                                                                                                                                                                                ctshdx1.setColor("auto");
                                                                                                                                                                                                                                                ctshdx1.setVal(STShd.CLEAR);
                                                                                                                                                                                                                                                if (rowCtx1 == 0) {
                                                                                                                                                                                                                                                    ctshdx1.setFill("FFFFFF");
                                                                                                                                                                                                                                                    
                                                                                                                                                                                                                                                } else {
                                                                                                                                                                                                                                                    ctshdx1.setFill("FFFFFF");
                                                                                                                                                                                                                                                }
                                                                                                                                                                                                                                                tablex1.getRow(0).getCell(0).setText("");
                                                                                                                                                                                                                                                
                                                                                                                                                                                                                                                int x3Rows = 3;
                                                                                                                                                                                                                                                int x3Cols = 5;
                                                                                                                                                                                                                                                XWPFTable tablex3 = document.createTable(x3Rows, x3Cols);
                                                                                                                                                                                                                                                tablex3.setWidth(2000);
                                                                                                                                                                                                                                                tablex3.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(1000));
                                                                                                                                                                                                                                                tablex3.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(4000));
                                                                                                                                                                                                                                                tablex3.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(1500));
                                                                                                                                                                                                                                                tablex3.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(1500));
                                                                                                                                                                                                                                                tablex3.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(1500));
                                                                                                                                                                                                                                                List<XWPFTableRow> rowsx3 = tablex3.getRows();
                                                                                                                                                                                                                                                int rowCtx3 = 0;
                                                                                                                                                                                                                                                int colCtx3 = 0;
                                                                                                                                                                                                                                                for (XWPFTableRow rowx3 : rowsx3) {
                                                                                                                                                                                                                                                    CTTrPr trPrx3 = rowx3.getCtRow().addNewTrPr();
                                                                                                                                                                                                                                                    CTHeight htx3 = trPrx3.addNewTrHeight();
                                                                                                                                                                                                                                                    htx3.setVal(BigInteger.valueOf(330));
                                                                                                                                                                                                                                                    List<XWPFTableCell> cellsx3 = rowx3.getTableCells();
                                                                                                                                                                                                                                                    for (XWPFTableCell cellx3 : cellsx3) {
                                                                                                                                                                                                                                                        CTTcPr tcprx3 = cellx3.getCTTc().addNewTcPr();
                                                                                                                                                                                                                                                        CTVerticalJc vax3 = tcprx3.addNewVAlign();
                                                                                                                                                                                                                                                        vax3.setVal(STVerticalJc.CENTER);
                                                                                                                                                                                                                                                        CTShd ctshdx3 = tcprx3.addNewShd();
                                                                                                                                                                                                                                                        ctshdx3.setColor("auto");
                                                                                                                                                                                                                                                        ctshdx3.setVal(STShd.CLEAR);
                                                                                                                                                                                                                                                        if (rowCtx3 == 0) {
                                                                                                                                                                                                                                                            ctshdx3.setFill("003366");
                                                                                                                                                                                                                                                        } else {
                                                                                                                                                                                                                                                            ctshdx3.setFill("FFFFFF");
                                                                                                                                                                                                                                                        }
                                                                                                                                                                                                                                                        XWPFParagraph parax3 = cellx3.getParagraphs().get(0);
                                                                                                                                                                                                                                                        XWPFRun rhx3 = parax3.createRun();
                                                                                                                                                                                                                                                        if (rowCtx3 == 0) {
                                                                                                                                                                                                                                                            rhx3.setText(" ");
                                                                                                                                                                                                                                                            rhx3.setFontSize(10);
                                                                                                                                                                                                                                                            rhx3.setBold(true);
                                                                                                                                                                                                                                                            parax3.setAlignment(ParagraphAlignment.CENTER);
                                                                                                                                                                                                                                                        } else {
                                                                                                                                                                                                                                                            rhx3.setText("");
                                                                                                                                                                                                                                                            parax3.setAlignment(ParagraphAlignment.CENTER);
                                                                                                                                                                                                                                                        }
                                                                                                                                                                                                                                                        colCtx3++;
                                                                                                                                                                                                                                                    }
                                                                                                                                                                                                                                                    colCtx3 = 0;
                                                                                                                                                                                                                                                    rowCtx3++;
                                                                                                                                                                                                                                                }
                                                                                                                                                                                                                                                tablex3.getRow(0).getCell(0).setText(" S.No");
                                                                                                                                                                                                                                                tablex3.getRow(0).getCell(1).setText("Folder name");
                                                                                                                                                                                                                                                tablex3.getRow(0).getCell(2).setText(" File name ");
                                                                                                                                                                                                                                                tablex3.getRow(0).getCell(3).setText(" Action ");
                                                                                                                                                                                                                                                tablex3.getRow(0).getCell(4).setText(" New location. ");
                                                                                                                                                                                                                                                tablex3.getRow(1).getCell(0).setText("1");
                                                                                                                                                                                                                                                tablex3.getRow(1).getCell(1).setText("");
                                                                                                                                                                                                                                                tablex3.getRow(1).getCell(2).setText("*.trc");
                                                                                                                                                                                                                                                tablex3.getRow(1).getCell(3).setText("Deleted the *.trc  files");
                                                                                                                                                                                                                                                tablex3.getRow(1).getCell(4).setText("NA");
                                                                                                                                                                                                                                                tablex3.getRow(2).getCell(0).setText("2");
                                                                                                                                                                                                                                                tablex3.getRow(2).getCell(1).setText("");
                                                                                                                                                                                                                                                tablex3.getRow(2).getCell(2).setText("*.trm");
                                                                                                                                                                                                                                                tablex3.getRow(2).getCell(3).setText("Deleted the *.trm  files");
                                                                                                                                                                                                                                                tablex3.getRow(2).getCell(4).setText("NA");
                                                                                                                                                                                                                                                XWPFParagraph para14 = document.createParagraph();
                                                                                                                                                                                                                                                XWPFRun ru14 = para14.createRun();
                                                                                                                                                                                                                                                ru14.addBreak();
                                                                                                                                                                                                                                                ru14.addBreak();
                                                                                                                                                                                                                                                int yRows = 1;
                                                                                                                                                                                                                                                int yCols = 0;
                                                                                                                                                                                                                                                XWPFTable tabley = document.createTable(yRows, yCols);
                                                                                                                                                                                                                                                tabley.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(9500));
                                                                                                                                                                                                                                                tabley.getCTTbl().getTblPr().unsetTblBorders();
                                                                                                                                                                                                                                                List<XWPFTableRow> rowsy = tabley.getRows();
                                                                                                                                                                                                                                                int rowCty = 0;
                                                                                                                                                                                                                                                for (XWPFTableRow rowy : rowsy) {
                                                                                                                                                                                                                                                    CTTrPr trPry = rowy.getCtRow().addNewTrPr();
                                                                                                                                                                                                                                                    CTHeight hty = trPry.addNewTrHeight();
                                                                                                                                                                                                                                                    hty.setVal(BigInteger.valueOf(453));
                                                                                                                                                                                                                                                    List<XWPFTableCell> cellsy = rowy.getTableCells();
                                                                                                                                                                                                                                                    for (XWPFTableCell celly : cellsy) {
                                                                                                                                                                                                                                                        CTTcPr tcpry = celly.getCTTc().addNewTcPr();
                                                                                                                                                                                                                                                        CTVerticalJc vay = tcpry.addNewVAlign();
                                                                                                                                                                                                                                                        vay.setVal(STVerticalJc.CENTER);
                                                                                                                                                                                                                                                        CTShd ctshdy = tcpry.addNewShd();
                                                                                                                                                                                                                                                        ctshdy.setColor("auto");
                                                                                                                                                                                                                                                        ctshdy.setVal(STShd.CLEAR);
                                                                                                                                                                                                                                                        if (rowCty == 0) {
                                                                                                                                                                                                                                                            ctshdy.setFill("FFFFFF");
                                                                                                                                                                                                                                                        } else {
                                                                                                                                                                                                                                                            ctshdy.setFill("FFFFFF");
                                                                                                                                                                                                                                                        }
                                                                                                                                                                                                                                                        XWPFParagraph paray = celly.getParagraphs().get(0);
                                                                                                                                                                                                                                                        XWPFRun rhy = paray.createRun();
                                                                                                                                                                                                                                                        rhy.setFontSize(10);
                                                                                                                                                                                                                                                        rhy.setBold(true);
                                                                                                                                                                                                                                                        rhy.setColor("99284C");
                                                                                                                                                                                                                                                        rhy.setText("DR Maintenance Details");
                                                                                                                                                                                                                                                        XWPFParagraph para15 = document.createParagraph();
                                                                                                                                                                                                                                                        XWPFRun ru15 = para15.createRun();
                                                                                                                                                                                                                                                        ru15.addBreak();
                                                                                                                                                                                                                                                        XWPFParagraph para16 = document.createParagraph();
                                                                                                                                                                                                                                                        XWPFRun ru16 = para16.createRun();
                                                                                                                                                                                                                                                        ru16.setFontSize(10);
                                                                                                                                                                                                                                                        para16.setAlignment(ParagraphAlignment.BOTH);
                                                                                                                                                                                                                                                        ru16.setText(" 1. Kindly ensure the free space availability of archive log destination on both DC and DR server.");
                                                                                                                                                                                                                                                        
                                                                                                                                                                                                                                                        ru16.addBreak();
                                                                                                                                                                                                                                                        ru16.setText(" 2. Check the archive synch between DC and DR using below query.");
                                                                                                                                                                                                                                                        ru16.addBreak();
                                                                                                                                                                                                                                                        ru16.addBreak();
                                                                                                                                                                                                                                                        ru16.setText("                  Select max(sequence#) from v$log_history ");
                                                                                                                                                                                                                                                        ru16.addBreak();
                                                                                                                                                                                                                                                        ru16.addBreak();
                                                                                                                                                                                                                                                        ru16.setText("Note: ");
                                                                                                                                                                                                                                                        ru16.addBreak();
                                                                                                                                                                                                                                                        ru16.setText(" It is not mandatory to have exact archive log sequence number matching for DR server.");
                                                                                                                                                                                                                                                        ru16.setText("It is not mandatory to have exact archive log sequence number matching for DR server. ");
                                                                                                                                                                                                                                                        ru16.addBreak();
                                                                                                                                                                                                                                                        ru16.addBreak();
                                                                                                                                                                                                                                                        ru16.setText("Approximately 5-20 archive log sequence difference may occur between every 30 minutes which will be updated in every SYNC script.");
                                                                                                                                                                                                                                                    }
                                                                                                                                                                                                                                                }
                                                                                                                                                                                                                                            }
                                                                                                                                                                                                                                        }
                                                                                                                                                                                                                                    }
                                                                                                                                                                                                                                }
                                                                                                                                                                                                                            }
                                                                                                                                                                                                                        }
                                                                                                                                                                                                                    }
                                                                                                                                                                                                                }
                                                                                                                                                                                                            }
                                                                                                                                                                                                        }
                                                                                                                                                                                                    }
                                                                                                                                                                                                }
                                                                                                                                                                                            }
                                                                                                                                                                                        }
                                                                                                                                                                                    }
                                                                                                                                                                                }
                                                                                                                                                                            }
                                                                                                                                                                            
                                                                                                                                                                        }
                                                                                                                                                                    }
                                                                                                                                                                }
                                                                                                                                                            }
                                                                                                                                                        }
                                                                                                                                                    }
                                                                                                                                                }
                                                                                                                                            }
                                                                                                                                        }
                                                                                                                                    }
                                                                                                                                }
                                                                                                                            }
                                                                                                                        }
                                                                                                                    }
                                                                                                                }
                                                                                                                
                                                                                                            }
                                                                                                        }
                                                                                                    }
                                                                                                }
                                                                                            }
                                                                                        }
                                                                                    }
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        DR d = new DR();
        d.pagesix(document);
        d.getQuery();
        document.write(out);
    }

    private void pagesix(XWPFDocument document) throws IOException, SQLException, FileNotFoundException, ClassNotFoundException {
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        ArrayList<String> query = getQuery();
        System.out.println("Values  : " + query);
        System.out.println("Size : " + query.size());
        
        for (int i = 0; i < query.size(); i++) {
            System.out.print("i value : " + i + "--- " + query.get(i));
            //  value = aa@ @bb @ @33@ @44@@    
            String arr[] = query.get(i).split("@@");
            String DB_REP_QRY = arr[0];
            
        }
    }
    
    private ArrayList<String> getQuery() throws SQLException, FileNotFoundException, IOException {
        ArrayList<String> query = new ArrayList();
        InputStream input = null;
        Connection con = null;
        try {
            Class.forName("oracle.jdbc.driver.OracleDriver");
        } catch (ClassNotFoundException ex) {
        }
        Properties props = new Properties();
        
        try (FileInputStream in = new FileInputStream("C:/DR_REPORT/DR_HEALTH_CHECK/DR_db.properties")) {
            props.load(in);
            String driver = props.getProperty("jdbc.driver");
            //   System.out.println("driver " + driver);
            if (driver != null) {
                Class.forName(driver);
            }
            String url = props.getProperty("jdbc.url");
            // System.out.println("url" + url);
            String username = props.getProperty("jdbc.username");
            //  System.out.println("username" + username);
            String password = props.getProperty("jdbc.password");
            // System.out.println("password" + password);
            con = DriverManager.getConnection(url, username, password);
            Statement stmt = con.createStatement();
            boolean status = stmt.execute("select DR_REP_QRY from dr_rep_dba");
            if (status) {
                ResultSet rs = stmt.getResultSet();
                while (rs.next()) {
                    query.add(rs.getString(1));
                }
            } else {
                int count = stmt.getUpdateCount();
                //  System.out.println("Total records updated : " + count);
            }
        } catch (ClassNotFoundException | SQLException e) {
            //  System.out.println(e);
        } finally {
            try {
                if (con != null) {
                    con.close();
                }
            } catch (SQLException ex) {
                //    System.out.println(ex);
            }
        }
        return query;
    }
    
}
