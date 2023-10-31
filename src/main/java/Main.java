

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Table;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class Main {

    public static String output = "F:\\ProbaJava.docx";

    public static void main(String[] args) throws IOException {

        //XWPFDocument — целостное представление Word документа.
        XWPFDocument document = new XWPFDocument();

        //создание параграфа
        //всё, что будет создано до создания нового параграфа, будет в этом параграфе
        XWPFParagraph title = document.createParagraph();
        title.setAlignment(ParagraphAlignment.LEFT);

        //document.createTable().createRow().createCell().addParagraph();
        XWPFTable table = document.createTable(5,7);
        //table.setWidth("20.00%");

        CTTblWidth widthRepr = table.getCTTbl().getTblPr().addNewTblW();
        widthRepr.setType(STTblWidth.AUTO);
        widthRepr.setW("20.00%");
        table.getRows().get(3).getCell(4).setText("3/4");
        //table.getRow(2).getCell(1).setWidth("30");

        //table.getRow(0).setHeight(20);
        //table.set
        //Cell cell = document.createTable()

        //XWPFRun - набор данных о выводе текста внутри параграфа.
        // Находится может только внутри параграфа, создается через вызов метода параграфа-родителя
        // для каждого параграфа своё оформление через XWPFRun
        XWPFRun paraRun = title.createRun();
        paraRun.setText("Уиииииии!");
        paraRun.setColor("009933");
        paraRun.setBold(true);
        paraRun.setFontFamily("Courier");
        paraRun.setFontSize(20);

        //следующий параграф
        //т.к. XWPFRun не прописан, оформление будет по умолчанию
        document.createParagraph().createRun().setText("Второй параграф");
        document.createParagraph().createRun().setText("Третий параграф");


        FileOutputStream out = new FileOutputStream(output);
        document.write(out);
        out.close();
        document.close();


//        document.saveToFile(
//                "output/CreateAWordDocument.docx",
//                FileFormat.Docx);
    }
}
