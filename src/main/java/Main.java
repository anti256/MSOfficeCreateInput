

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Table;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
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




        //XWPFRun - набор данных о выводе текста внутри параграфа.
        // Находится может только внутри параграфа, создается через вызов метода параграфа-родителя
        // для каждого параграфа своё оформление через XWPFRun
        XWPFRun paraRun = title.createRun();
        paraRun.setText("Уиииииии!");
        paraRun.setColor("009933");
        paraRun.setBold(true);
        paraRun.setFontFamily("Courier");
        paraRun.setFontSize(20);
        //новая строка без смены параграфа
        paraRun.addBreak();
        paraRun.setText("ещё строка");

        //следующий параграф
        //т.к. XWPFRun не прописан, оформление будет по умолчанию
        document.createParagraph().createRun().setText("Второй параграф");
        document.createParagraph().createRun().setText("Третий параграф");

        //document.createTable().createRow().createCell().addParagraph();
        XWPFTable table = document.createTable(7,5);
        table.setWidthType(TableWidthType.DXA);
        table.setWidth("100.00%");
        int wid = table.getRow(1).getCell(1).getWidth();
        table.setWidth(wid);


        //делает тоже самое по поводу ширины таблицы
        //CTTblWidth widthRepr = table.getCTTbl().getTblPr().addNewTblW();
        //widthRepr.setType(STTblWidth.PCT);
        //widthRepr.setW("100.00%");//ширина таблицы во всю возможную ширину на странице
        //table.getRows().get(3).getCell(4).setText("3/4");
        //table.getRow(1).getCell(0).setWidth("20.0%");

        FileOutputStream out = new FileOutputStream(output);
        document.write(out);
        out.close();
        document.close();


//        document.saveToFile(
//                "output/CreateAWordDocument.docx",
//                FileFormat.Docx);
    }
}
