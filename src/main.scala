import java.io._
import java.awt.Color

import org.apache.poi.xssf.usermodel.{XSSFColor, XSSFWorkbook}
import javax.imageio.ImageIO

import org.apache.poi.ss.usermodel.CellStyle

object main {
  val PixelHeight:Short = 25
  val PixelWidth:Short = 50
  val FilePrefix:String = "sasaki_nozomi"
  def main(args: Array[String]): Unit = {
    val image = ImageIO.read(new File(".\\" + FilePrefix + ".jpg"))

    val workbook = new XSSFWorkbook()
    workbook.createSheet(FilePrefix)
    val sheet = workbook.getSheetAt(0)
    val width = image.getWidth()
    val height = image.getHeight()
    for( j <- 0 until height-1 ) {
      sheet.createRow(j)
      val row = sheet.getRow(j)
      row.setHeight(PixelHeight)
      for( i <- 0 until width-1 ) {
        row.createCell(i)
        val cel = row.getCell(i)
        val pixel = image.getRGB(i, j)
        val color = new XSSFColor(this.getRGBArray(pixel))
        val style = workbook.createCellStyle()
        style.setFillPattern(CellStyle.SOLID_FOREGROUND)
        style.setFillForegroundColor(color)
        cel.setCellStyle(style)
        print(i + ",")
      }
      println()
      println(j + "/" + height)
    }
    for( i <- 0 until width-1 ) {
      sheet.setColumnWidth(i, PixelWidth)
    }

    val outfile = new java.io.FileOutputStream(FilePrefix + "_scala.xlsx")
    workbook.write(outfile)
    outfile.close()
  }
  def getRGBArray(pixel:Int):Color = {
    val r:Int = (pixel >> 16) & 0xff
    val g:Int = (pixel >> 8) & 0xff
    val b:Int = pixel & 0xff
    new Color(r,g,b)
  }
}
