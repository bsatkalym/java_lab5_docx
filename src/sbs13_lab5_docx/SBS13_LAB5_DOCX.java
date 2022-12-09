package sbs13_lab5_docx;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class SBS13_LAB5_DOCX {

    public static void main(String[] args) throws InvalidFormatException, IOException {
        String dir = new File(".").getAbsoluteFile().getParentFile().getAbsolutePath()
                + System.getProperty("file.separator");
        XWPFDocument doc = new XWPFDocument(OPCPackage.open(dir + "input.docx"));
        for (XWPFParagraph p : doc.getParagraphs()) {
            List<XWPFRun> runs = p.getRuns();
            if (runs != null) {
                for (XWPFRun r : runs) {
                    String text = r.getText(0);
                    if (text != null && text.toLowerCase().contains("$one")) {
                        text = text.replace(text, "Пример");
                        r.setText(text, 0);
                    }
                    if (text != null && text.toLowerCase().contains("$two")) {
                        text = text.replace(text, " работы");
                        r.setText(text, 0);
                    }
                    if (text != null && text.toLowerCase().contains("$three")) {
                        text = text.replace(text, "С");
                        r.setText(text, 0);
                    }
                    if (text != null && text.toLowerCase().contains("$four")) {
                        text = text.replace(text, "DOCX");
                        r.setText(text, 0);
                    }
                    if (text != null && text.toLowerCase().contains("$five")) {
                        text = text.replace(text, "через");
                        r.setText(text, 0);
                    }
                    if (text != null && text.toLowerCase().contains("$six")) {
                        text = text.replace(text, "Apache poi");
                        r.setText(text, 0);
                    }
                }
            }
        }
        for (XWPFTable tbl : doc.getTables()) {
            for (XWPFTableRow row : tbl.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph p : cell.getParagraphs()) {
                        for (XWPFRun r : p.getRuns()) {
                            String text = r.getText(0);
                            if (text != null && text.toLowerCase().contains("$one")) {
                                text = text.replace(text, "Пример");
                                r.setText(text, 0);
                            }
                            if (text != null && text.toLowerCase().contains("$two")) {
                                text = text.replace(text, "работы");
                                r.setText(text, 0);
                            }
                            if (text != null && text.toLowerCase().contains("$three")) {
                                text = text.replace(text, "С");
                                r.setText(text, 0);
                            }
                            if (text != null && text.toLowerCase().contains("$four")) {
                                text = text.replace(text, "DOCX");
                                r.setText(text, 0);
                            }
                            if (text != null && text.toLowerCase().contains("$five")) {
                                text = text.replace(text, "через");
                                r.setText(text, 0);
                            }
                            if (text != null && text.toLowerCase().contains("$six")) {
                                text = text.replace(text, "Apache poi");
                                r.setText(text, 0);
                            }
                        }
                    }
                }
            }
        }
        FileOutputStream fos = new FileOutputStream(dir + "output.docx");
        doc.write(fos);
        fos.close();
    }

}
