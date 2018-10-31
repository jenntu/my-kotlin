package test.java

import org.apache.poi.hssf.usermodel.HSSFClientAnchor
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFClientAnchor
import org.apache.poi.xssf.usermodel.XSSFPicture
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.ByteArrayOutputStream
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.util.ArrayList
import javax.imageio.ImageIO

object TestAllPic {

    @JvmStatic
    fun main(args: Array<String>) {

        // 创建文件
        val file = File("/Users/duzhen/Documents/testImg.xlsx")

        // 创建流
        val input = FileInputStream(file)

        // 获取文件后缀名
        val fileExt = file.name.substring(file.name.lastIndexOf(".") + 1)

        // 创建Workbook
        var wb: Workbook? = null

        // 创建sheet
        var sheet: Sheet? = null

        //根据后缀判断excel 2003 or 2007+
        if (fileExt == "xls") {
            wb = WorkbookFactory.create(input) as HSSFWorkbook
        } else {
            wb = XSSFWorkbook(input)
        }

        //获取excel sheet总数
        val sheetNumbers = wb.numberOfSheets

        // sheet list
        val sheetList = ArrayList<Map<String, PictureData>>()

        // 循环sheet
        for (i in 0 until sheetNumbers) {

            sheet = wb.getSheetAt(i)
            println("this is sheet $i")
            getPictureNameOfCell(sheet, wb)

        }



    }


    /**
     * 获取单元格中的图片文件名，如果单元格内容不是图片返回null。
     *
     * @param sheet
     *            工作表
     * @param cell
     *            单元格
     * @return 图片文件名
     */
    fun getPictureNameOfCell(sheet: Sheet, wb: Workbook): String? {
        if (sheet is XSSFSheet) {
            var drawing = sheet.drawingPatriarch
            var xShapeList = drawing.shapes
            xShapeList.forEach {
                if (it is XSSFPicture) {
                    var xAnchor = it.preferredSize
                    var picName = it.pictureData.packagePart.partName.name
                    print("row1 and col1 -> rowNum: ${xAnchor.row1}, colNum: ${xAnchor.col1}, picName: $picName")
                    println("  ||  row2 and col2 ->  rowNum: ${xAnchor.row2}, colNum: ${xAnchor.col2}, picName: $picName")

                    var patriarch = sheet.createDrawingPatriarch()
                    //anchor主要用于设置图片的属性
                    var anchor = XSSFClientAnchor(0, 0, 255, 255, 1, 1, 5, 8)
                    anchor.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
                    //插入图片
                    patriarch.createPicture(anchor, wb.addPicture(it.pictureData.data, HSSFWorkbook.PICTURE_TYPE_JPEG));

                    var fileOut = FileOutputStream("/Users/duzhen/test.xlsx");
                    // 写入excel文件
                    wb.write(fileOut);
                    System.out.println("----Excle文件已生成------");

                }
            }
        }
        return null
    }
}