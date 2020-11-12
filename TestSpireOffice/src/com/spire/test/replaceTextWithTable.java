package com.spire.test;

import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

/**
 * 测试读取Word的文字，转换为table。
* @author WuBin
*
*/
public class replaceTextWithTable {
    public static void main(String[] args) {
        //Create word document.
        Document document = new Document();

        // Load the file from disk.
        //document.loadFromFile("data/Template_Docx_1.docx");
        document.loadFromFile("output/createTableOfContentByDefault.docx");

        //Return TextSection by finding the key text string "Christmas Day, December 25".
        Section section = document.getSections().get(0);
        String replaceText = "2020.09.11";
        TextSelection selection = document.findString(replaceText, true, true);

        //Return TextRange from TextSection, then get OwnerParagraph through TextRange.
        TextRange range = selection.getAsOneRange();
        Paragraph paragraph = range.getOwnerParagraph();

        //Return the zero-based index of the specified paragraph.
        Body body = paragraph.ownerTextBody();
        int index = body.getChildObjects().indexOf(paragraph);

        //Create a new table.        
        Table table = section.addTable(true);
        /**
        table.resetCells(3, 3);
s		**/
        
        table = createTable.addTableAndContent(table);

        // Remove the paragraph and insert table into the collection at the specified index.
        body.getChildObjects().remove(paragraph);
        body.getChildObjects().insert(index, table);

        String result = "output/replaceTextWithTable.docx";

        // Save to file.
        document.saveToFile(result, FileFormat.Docx_2013);
        
        System.out.println(">>> replace word is over.");
    }
}
