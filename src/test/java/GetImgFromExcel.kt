package test.java


import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.io.IOException
import java.io.InputStream
import java.util.ArrayList
import java.util.HashMap

import org.apache.poi.POIXMLDocumentPart
import org.apache.poi.hssf.usermodel.HSSFClientAnchor
import org.apache.poi.hssf.usermodel.HSSFPicture
import org.apache.poi.hssf.usermodel.HSSFPictureData
import org.apache.poi.hssf.usermodel.HSSFShape
import org.apache.poi.hssf.usermodel.HSSFSheet
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.openxml4j.exceptions.InvalidFormatException
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFClientAnchor
import org.apache.poi.xssf.usermodel.XSSFDrawing
import org.apache.poi.xssf.usermodel.XSSFPicture
import org.apache.poi.xssf.usermodel.XSSFShape
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker

/**
 *
 * 获取excel中 图片，并得到图片位置，支持03 07 多sheet
 *
 */
object GetImgFromExcel {

    /**
     * @param args
     * @throws IOException
     * @throws InvalidFormatException
     */
    @Throws(InvalidFormatException::class, IOException::class)
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
            // map等待存储excel图片
            val sheetIndexPicMap: Map<String, PictureData>?

            // 判断用07还是03的方法获取图片
            if (fileExt == "xls") {
                sheetIndexPicMap = getSheetPictrues03(i, sheet as HSSFSheet, (wb as HSSFWorkbook?)!!)
            } else {
                sheetIndexPicMap = getSheetPictrues07(i, (sheet as XSSFSheet?)!!, wb as XSSFWorkbook)
            }
            // 将当前sheet图片map存入list
            sheetList.add(sheetIndexPicMap!!)
        }

        printImg(sheetList)

    }

    /**
     * 获取Excel2003图片
     * @param sheetNum 当前sheet编号
     * @param sheet 当前sheet对象
     * @param workbook 工作簿对象
     * @return Map key:图片单元格索引（0_1_1）String，value:图片流PictureData
     * @throws IOException
     */
    fun getSheetPictrues03(sheetNum: Int,
                           sheet: HSSFSheet, workbook: HSSFWorkbook): Map<String, PictureData>? {

        val sheetIndexPicMap = HashMap<String, PictureData>()
        val pictures = workbook.allPictures
        if (pictures.size != 0) {
            for (shape in sheet.drawingPatriarch.children) {
                val anchor = shape.anchor as HSSFClientAnchor
                if (shape is HSSFPicture) {
                    val pictureIndex = shape.pictureIndex - 1
                    val picData = pictures[pictureIndex]
                    val picIndex = (sheetNum.toString() + "_"
                            + anchor.row1.toString() + "_"
                            + anchor.col1.toString())
                    sheetIndexPicMap[picIndex] = picData
                }
            }
            return sheetIndexPicMap
        } else {
            return null
        }
    }

    /**
     * 获取Excel2007图片
     * @param sheetNum 当前sheet编号
     * @param sheet 当前sheet对象
     * @param workbook 工作簿对象
     * @return Map key:图片单元格索引（0_1_1）String，value:图片流PictureData
     */
    fun getSheetPictrues07(sheetNum: Int,
                           sheet: XSSFSheet, workbook: XSSFWorkbook): Map<String, PictureData> {
        val sheetIndexPicMap = HashMap<String, PictureData>()

        for (dr in sheet.relations) {
            if (dr is XSSFDrawing) {
                val shapes = dr.shapes
                for (shape in shapes) {
                    val pic = shape as XSSFPicture
                    val anchor = pic.preferredSize
                    val ctMarker = anchor.from
                    val picIndex = (sheetNum.toString() + "_"
                            + ctMarker.row + "_" + ctMarker.col)
                    sheetIndexPicMap[picIndex] = pic.pictureData
                }
            }
        }

        return sheetIndexPicMap
    }

    @Throws(IOException::class)
    fun printImg(sheetList: List<Map<String, PictureData>>) {

        for (map in sheetList) {
            val key = map.keys.toTypedArray()
            for (i in 0 until map.size) {
                // 获取图片流
                val pic = map[key[i]]
                // 获取图片索引
                val picName = key[i]
                // 获取图片格式
                val ext = pic!!.suggestFileExtension()

                val data = pic.getData()

                val out = FileOutputStream("D:\\pic$picName.$ext")
                out.write(data)
                out.close()
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
    fun getPictureNameOfCell(sheet: Sheet, cell: Cell): String? {
        var cellRowIndex = cell.rowIndex
        var cellColIndex = cell.columnIndex
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