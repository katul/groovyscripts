 import org.apache.poi.ss.usermodel.*
import org.apache.poi.hssf.usermodel.*
import org.apache.poi.xssf.usermodel.*
import org.apache.poi.ss.util.*
import org.apache.poi.ss.usermodel.*
import java.io.*
import com.day.cq.commons.jcr.*;

main();
  //http://poi.apache.org/spreadsheet/quick-guide.html#Iterator
class GroovyExcelParser {
  def parse(path) {
    InputStream inp = new FileInputStream(path)
    Workbook wb = WorkbookFactory.create(inp);
    Sheet sheet = wb.getSheetAt(0);

    Iterator<Row> rowIt = sheet.rowIterator()
    Row row = rowIt.next()
    def headers = getRowData(row)

    def rows = []
    while(rowIt.hasNext()) {
      row = rowIt.next()
      rows << getRowData(row)
    }
    [headers, rows]
  }

  def getRowData(Row row) {
    def data = []
    for (Cell cell : row) {
      getValue(row, cell, data)
    }
    data
  }

  def getRowReference(Row row, Cell cell) {
    def rowIndex = row.getRowNum()
    def colIndex = cell.getColumnIndex()
    CellReference ref = new CellReference(rowIndex, colIndex)
    ref.getRichStringCellValue().getString()
  }
 
  def getValue(Row row, Cell cell, List data) {
    def rowIndex = row.getRowNum()
    def colIndex = cell.getColumnIndex()
    def value = ""
    switch (cell.getCellType()) {
      case CellType.BLANK:
        value = "atul";
        break;
      case CellType.STRING:
        value = cell.getRichStringCellValue().getString();
        break;
      case CellType.NUMERIC:
        if (DateUtil.isCellDateFormatted(cell)) {
            value = cell.getDateCellValue();
        } else {
            value = cell.getNumericCellValue();
        }
        break;
      case CellType.BOOLEAN:
        value = cell.getBooleanCellValue();
        break;
      case CellType.FORMULA:
        value = cell.getCellFormula();
        break;
      default:
        value = ""
    }
    data[colIndex] = value
    data
  }

  def toXml(header, row) {
    def obj = "<object>\n"
    row.eachWithIndex { datum, i -> 
      def headerName = header[i]
      obj += "\t<$headerName>$datum</$headerName>\n" 
    } 
    obj += "</object>"
  }
}
  def main() {
        def filename = '/Users/archanagupta/Desktop/announcements-data.xlsx'
        GroovyExcelParser parser = new GroovyExcelParser()
        def (headers, rows) = parser.parse(filename)
        def rootPagePath = "/content/we-retail/en/announcements/";
        println 'Headers'
        println '------------------'
        headers.each { header -> 
          println header
        }
        println "\n"
        println 'Rows'
        println '------------------'
        def  count = 0;
        rows.each { row ->
          String date =  row[0];
          String time =  row[1];
          String title =  row[2];
          String tags =  row[3];
          String content =  row[4];
          if(!tags.isEmpty()){
              if(!tags.contains(",")){
                  def tagVal = tags.split("#")[1];
                  def tag = tagVal.replaceAll(" ","-").toLowerCase();
                  def rootPage = rootPagePath + tag + "/";
                  if(getPage(rootPage)==null){
                      pageManager.create(rootPagePath,JcrUtil.createValidName(tag),"we-retail/templates/content-page",tagVal);
                  }
                 Page page = pageManager.create(rootPage,JcrUtil.createValidName(title),"we-retail/templates/content-page",title);
                 //println page.path;
                 Node pageNode = page.adaptTo(Node.class);
                 Node contentNode = pageNode.getNode("jcr:content");
                 Node contentNode = pageNode.getorAddNode("jcr:content");
                 contentNode.setProperty("announcementTitle",title);
                 contentNode.setProperty("announcementDate",date);
                 contentNode.setProperty("announcementTime",time);
                 contentNode.setProperty("announcementContent",content);
                 contentNode.setProperty("announcementTags",tags);
                }
            }
        }
    }
