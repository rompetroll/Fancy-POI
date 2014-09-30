package org.fancypoi.excel

/**
 * Scala-themed support for POI cell types.
 * Usage example:
 * {{{
 *   import org.fancypoi.Implicits._
 *   import org.fancypoi.FancyCellType
 *
 *   val cell: Cell = ???
 *   cell.cellType match {
 *     case CellTypeNumeric => println(cell.numericValue)
 *     case CellTypeString  => println(cell.stringValue)
 *   }
 * }}}
 *
 * @author FS
 */
object FancyCellType {
  sealed abstract class CellType(val cellTypeIndex: Int, val cellTypeString: String) {
    override def toString = cellTypeString
  }
  case object CellTypeNumeric extends CellType(0, "CELL_TYPE_NUMERIC")
  case object CellTypeString extends CellType(1, "CELL_TYPE_STRING")
  case object CellTypeFormula extends CellType(2, "CELL_TYPE_FORMULA")
  case object CellTypeBlank extends CellType(3, "CELL_TYPE_BLANK")
  case object CellTypeBoolean extends CellType(4, "CELL_TYPE_BOOLEAN")
  case object CellTypeError extends CellType(5, "CELL_TYPE_ERROR")

  def fromPoiCellType(i: Int): CellType = i match {
    case 0 => CellTypeNumeric
    case 1 => CellTypeString
    case 2 => CellTypeFormula
    case 3 => CellTypeBlank
    case 4 => CellTypeBoolean
    case 5 => CellTypeError
    case x => throw new IllegalArgumentException(s"Unknown cell type: $x")
  }
}
