package test.java

import org.apache.poi.hssf.usermodel.HSSFSheet
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFPicture
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream
import java.util.ArrayList

object GetImg {

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
            for (rowNum in 0..sheet.lastRowNum) {
                var row = sheet.getRow(rowNum)
                for (colNum in 0..row.lastCellNum) {
                    var picName = getPictureNameOfCell(sheet, rowNum, colNum)
                    println("sheet : $i ,rowNum: $rowNum, colNum: $colNum, value : $picName)")
                }
            }
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
    fun getPictureNameOfCell(sheet: Sheet, cellRowIndex: Int, cellColIndex: Int): String? {
        if (sheet is XSSFSheet) {
            var drawing = sheet.drawingPatriarch
            var xShapeList = drawing.shapes
            xShapeList.forEach {
                if (it is XSSFPicture) {
                    var xAnchor = it.preferredSize
                    if (xAnchor.row1 == cellRowIndex && xAnchor.col1.toInt() == cellColIndex) {
                        return it.pictureData.packagePart.partName.name
                    }
                }
            }
        }
        return null
    }

}