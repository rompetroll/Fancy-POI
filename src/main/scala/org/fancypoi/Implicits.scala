package org.fancypoi

import excel.FancyExcelUtils.AddrRangeStart
import excel.{ FancyWorkbook, FancySheet, FancyRow, FancyCell }
import org.apache.poi.ss.usermodel._

/**
 * @author ishiiyoshinori
 * @since 2011-05-04
 */
object Implicits {

  implicit def workbook2fancy(w: Workbook): org.fancypoi.excel.FancyWorkbook = new FancyWorkbook(w)

  implicit def sheet2fancy(s: Sheet): org.fancypoi.excel.FancySheet = new FancySheet(s)

  implicit def row2fancy(r: Row): org.fancypoi.excel.FancyRow = new FancyRow(r)

  implicit def cell2fancy(c: Cell): org.fancypoi.excel.FancyCell = new FancyCell(c)

  implicit def workbook2plain(w: FancyWorkbook): org.apache.poi.ss.usermodel.Workbook = w.workbook

  implicit def sheet2plain(s: FancySheet): org.apache.poi.ss.usermodel.Sheet = s._sheet

  implicit def row2plain(r: FancyRow): org.apache.poi.ss.usermodel.Row = r._row

  implicit def cell2plain(c: FancyCell): org.apache.poi.ss.usermodel.Cell = c._cell

  implicit def indexedColors2Int(indexedColor: IndexedColors): Short = indexedColor.getIndex

  implicit def str2Addr(addr: String): org.fancypoi.excel.FancyExcelUtils.AddrRangeStart = new AddrRangeStart(addr)

}
