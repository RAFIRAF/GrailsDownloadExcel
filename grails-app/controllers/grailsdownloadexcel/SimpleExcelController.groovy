package grailsdownloadexcel

import jxl.Workbook
import jxl.write.Label
import jxl.write.WritableSheet
import jxl.write.WritableWorkbook

class SimpleExcelController {

    def index() { }
    def downloadSampleExcel(String s){
        response.setContentType('application/vnd.ms-excel')
        response.setHeader('Content-Disposition','Attachment;Filename="example.xls"')
        WritableWorkbook writableWorkbook = Workbook.createWorkbook(response.outputStream)
        WritableSheet writableSheet = writableWorkbook.createSheet("Students",0)
        writableSheet.addCell(new Label(0,0, "First Name"))
        writableSheet.addCell(new Label(1,0, "Last Name"))
        writableSheet.addCell(new Label(2,0, "Age"))
        writableSheet.addCell(new Label(0,1, "John"))
        writableSheet.addCell(new Label(1,1, "Doe"))
        writableSheet.addCell(new Label(2,1, "20"))
        writableSheet.addCell(new Label(0,2, "Jane"))
        writableSheet.addCell(new Label(1,2, "Smith"))
        writableSheet.addCell(new Label(2,2, "18"))
        WritableSheet sheet2 = writableWorkbook.createSheet("Courses", 1)
        sheet2.addCell(new Label(0,0, "Course Name"))
        sheet2.addCell(new Label(1,0, "Number of units"))
        sheet2.addCell(new Label(0,1, "Algebra"))
        sheet2.addCell(new Label(1,1, "3"))
        sheet2.addCell(new Label(0,2, "English Grammar"))
        sheet2.addCell(new Label(1,2, s))
        writableWorkbook.write()
        writableWorkbook.close()
    }
}
