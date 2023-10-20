
import com.spire.doc.Document;
import com.spire.doc.FileFormat;
import com.spire.doc.Section;
import com.spire.doc.documents.BuiltinStyle;
import com.spire.doc.documents.Paragraph;
import com.spire.doc.documents.ParagraphStyle;

public class Main {
    public static void main(String[] args) {
        Document document = new Document();

        document.saveToFile(
                "output/CreateAWordDocument.docx",
                FileFormat.Docx);
    }
}
