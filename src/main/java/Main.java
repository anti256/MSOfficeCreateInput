

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
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

    //установка границ для таблицы - можно и без этого
    static void setAllBorders(XWPFTable table, XWPFTable.XWPFBorderType borderType, int size, int space, java.lang.String rgbColor) {
        table.setTopBorder(borderType, size, space, rgbColor);
        table.setRightBorder(borderType, size, space, rgbColor);
        table.setBottomBorder(borderType, size, space, rgbColor);
        table.setLeftBorder(borderType, size, space, rgbColor);
        table.setInsideHBorder(borderType, size, space, rgbColor);
        table.setInsideVBorder(borderType, size, space, rgbColor);
    }

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
        setAllBorders(table, XWPFTable.XWPFBorderType.DOUBLE, 4, 0, "FF0000");
        table.setWidthType(TableWidthType.PCT);//dxa - установка в процентах, nil - устанавливает ширину в ноль
        table.setWidth("100.00%");

        int w = table.getWidth()/5;
        XWPFTableRow row = table.getRow(1);
        for (XWPFTableCell cell: row.getTableCells()
             ) {
            cell.setWidthType(TableWidthType.DXA);
            cell.setWidth("" + w);
        }


        //делает тоже самое по поводу ширины таблицы
//        CTTblWidth widthRepr = table.getCTTbl().getTblPr().addNewTblW();
//        widthRepr.setType(STTblWidth.PCT);
//        widthRepr.setW("100.00%");//ширина таблицы во всю возможную ширину на странице
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
//К сожалению, так и не удалось найти каким образом зафиксировать ширины столбцов. Сделать таблицу с определенными ширинами столбцов не проблема,
//проблема когда открываешь созданный файл с таблицей в WORD, при редактировании ширины столбцов сдвигаются.