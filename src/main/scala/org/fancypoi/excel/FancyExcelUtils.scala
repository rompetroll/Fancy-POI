package org.fancypoi.excel

import org.apache.poi.ss.usermodel._
import org.apache.poi.ss.util.CellReference.convertColStringToIndex
import org.apache.poi.ss.util.CellReference.convertNumToColString
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy
import org.apache.poi.hssf.usermodel.HSSFRichTextString
import org.apache.poi.xssf.usermodel.XSSFRichTextString
import org.fancypoi.Implicits._
import org.fancypoi.MonadicConversions._

object FancyExcelUtils {

  def addrToIndexes(address: String) = {
    val m = "([A-Z]+)(\\d+)".r.findAllIn(address).matchData.toList(0)
    val colAddr = m.group(1)
    val rowAddr = m.group(2)
    (convertColStringToIndex(colAddr), rowAddr.toInt - 1)
  }

  def copyRowStyle(from: FancyRow, to: FancyRow) {
    for {
      colIndex <- (from.firstColIndex to from.lastColIndex).toList
      cell = from.getCell(colIndex, MissingCellPolicy.RETURN_NULL_AND_BLANK)
      if cell != null
      cellStyle = cell.getCellStyle
    } yield {
      to.cellAt(colIndex).replaceStyle(cellStyle)
    }
  }

  def searchFont(workbook: FancyWorkbook, font: FancyFont) = {
    val indexes = (0 to workbook.getNumberOfFonts).toList.filter(_ != 4).map(_.toShort)
    indexes.find {
      index =>
        val registered = workbook.workbook.getFontAt(index)
        equalFont(font, registered)
    }.map(workbook.workbook.getFontAt)
  }

  def searchStyle(workbook: FancyWorkbook, style: FancyCellStyle): Option[CellStyle] = {
    (0 to workbook.workbook.getNumCellStyles - 1).map(_.toInt).find {
      index =>
        val registered = workbook.workbook.getCellStyleAt(index toShort)
        equalStyleWithoutFont(style, registered) && equalFont(style.getFont, workbook.getFontAt(registered.getFontIndex))
    }.map(workbook.workbook.getCellStyleAt)
  }

  /**
   * 目に見えるセルかどうかを判定します。
   *
   * 以下の条件を満たすと目に見えないと判定します。
   *     ・値がブランク
   *     ・背景色、前面色がない、または、自動設定
   *     ・罫線もない
   */
  def isViewableCell(cell: FancyCell) = {
    !(cell.cellType == FancyCellType.CellTypeBlank &&
      List(0, 64).contains(cell.style.getFillBackgroundColor) &&
      List(0, 64).contains(cell.style.getFillForegroundColor) &&
      cell.style.getBorderBottomEnum == BorderStyle.NONE &&
      cell.style.getBorderLeftEnum == BorderStyle.NONE &&
      cell.style.getBorderRightEnum == BorderStyle.NONE &&
      cell.style.getBorderTopEnum == BorderStyle.NONE)
  }

  def copyStyleWithoutFont(from: CellStyle, to: CellStyle) = {
    to.setAlignment(from.getAlignmentEnum)
    to.setBorderBottom(from.getBorderBottomEnum)
    to.setBorderLeft(from.getBorderLeftEnum)
    to.setBorderRight(from.getBorderRightEnum)
    to.setBorderTop(from.getBorderTopEnum)
    to.setBottomBorderColor(from.getBottomBorderColor)
    to.setDataFormat(from.getDataFormat)
    to.setFillForegroundColor(from.getFillForegroundColor) // setFillBackgroundColor の前にsetFillForegroundColorをセットしないとgetFillBackgroundColorの値が変わってしまう。
    to.setFillBackgroundColor(from.getFillBackgroundColor)
    to.setFillPattern(from.getFillPatternEnum)
    to.setHidden(from.getHidden)
    to.setIndention(from.getIndention)
    to.setLeftBorderColor(from.getLeftBorderColor)
    to.setLocked(from.getLocked)
    to.setRightBorderColor(from.getRightBorderColor)
    to.setRotation(from.getRotation)
    to.setTopBorderColor(from.getTopBorderColor)
    to.setVerticalAlignment(from.getVerticalAlignmentEnum)
    to.setWrapText(from.getWrapText)
    to
  }

  def copyFont(from: Font, to: Font) = {
    to.setBold(from.getBold)
    to.setCharSet(from.getCharSet)
    to.setColor(from.getColor)
    to.setFontHeight(from.getFontHeight)
    to.setFontName(from.getFontName)
    to.setItalic(from.getItalic)
    to.setStrikeout(from.getStrikeout)
    to.setTypeOffset(from.getTypeOffset)
    to.setUnderline(from.getUnderline)
    to
  }

  protected def diff(expected: Any, actual: Any, name: String) = (expected == actual) ~ ("[" + name + "] expected=" + expected + ",actual=" + actual)

  def equalFont(font1: Font, font2: Font) = {
    diff(font1.getBold, font2.getBold, "bold") &&
      diff(font1.getColor, font2.getColor, "color") &&
      diff(font1.getFontHeight, font2.getFontHeight, "fontHeight") &&
      diff(font1.getFontName, font2.getFontName, "fontName") &&
      diff(font1.getItalic, font2.getItalic, "italic") &&
      diff(font1.getStrikeout, font2.getStrikeout, "strikeout") &&
      diff(font1.getTypeOffset, font2.getTypeOffset, "typeOffset") &&
      diff(font1.getUnderline, font2.getUnderline, "underline")
  }

  def equalStyleWithoutFont(style1: CellStyle, style2: CellStyle) = {
    diff(style1.getAlignmentEnum, style2.getAlignmentEnum, "alignment") &&
      diff(style1.getBorderBottomEnum, style2.getBorderBottomEnum, "borderBottom") &&
      diff(style1.getBorderLeftEnum, style2.getBorderLeftEnum, "borderLeft") &&
      diff(style1.getBorderRightEnum, style2.getBorderRightEnum, "borderRight") &&
      diff(style1.getBorderTopEnum, style2.getBorderTopEnum, "borderTop") &&
      diff(style1.getBottomBorderColor, style2.getBottomBorderColor, "bottomBorderColor") &&
      diff(style1.getDataFormat, style2.getDataFormat, "dateFormat") &&
      diff(style1.getFillBackgroundColor, style2.getFillBackgroundColor, "fillBackgroundColor") &&
      diff(style1.getFillForegroundColor, style2.getFillForegroundColor, "fillForegroundColor") &&
      diff(style1.getFillPatternEnum, style2.getFillPatternEnum, "fillPattern") &&
      diff(style1.getHidden, style2.getHidden, "hidden") &&
      diff(style1.getIndention, style2.getIndention, "indention") &&
      diff(style1.getLeftBorderColor, style2.getLeftBorderColor, "leftBorderColor") &&
      diff(style1.getLocked, style2.getLocked, "locked") &&
      diff(style1.getRightBorderColor, style2.getRightBorderColor, "rightBorderColor") &&
      diff(style1.getRotation, style2.getRotation, "rotation") &&
      diff(style1.getTopBorderColor, style2.getTopBorderColor, "borderColor") &&
      diff(style1.getVerticalAlignmentEnum, style2.getVerticalAlignmentEnum, "verticalAlignment") &&
      diff(style1.getWrapText, style2.getWrapText, "wrapText")
  }

  def equalHyperlink(link1: Hyperlink, link2: Hyperlink) = {
    (link1, link2) match {
      case (null, null) => true
      case (null, _)    => false
      case (_, null)    => false
      case _ =>
        diff(link1.getAddress, link2.getAddress, "address") &&
          diff(link1.getLabel, link2.getLabel, "label") &&
          diff(link1.getTypeEnum, link2.getTypeEnum, "type")
    }
  }

  def equalComment(cm1: Comment, cm2: Comment, w1: Workbook, w2: Workbook) = {
    (cm1, cm2) match {
      case (null, null) => true
      case (null, _)    => false
      case (_, null)    => false
      case _ =>
        diff(cm1.getAuthor, cm2.getAuthor, "author") &&
          diff(cm1.isVisible, cm2.isVisible, "isVisible") &&
          (equalRichTextStyring(cm1.getString, cm2.getString, w1, w2) ~ "string")
    }
  }

  def equalRichTextStyring(str1: RichTextString, str2: RichTextString, w1: Workbook, w2: Workbook): Boolean = {
    (str1, str2) match {
      case (s1: HSSFRichTextString, s2: HSSFRichTextString) =>
        if (s1.getString != s2.getString) return false
        if (s1.numFormattingRuns != s2.numFormattingRuns) return false
        val s1FontIndexes = (0 to s1.numFormattingRuns - 1).map(s1.getFontOfFormattingRun).toList
        val s2FontIndexes = (0 to s2.numFormattingRuns - 1).map(s2.getFontOfFormattingRun).toList
        !s1FontIndexes.zip(s2FontIndexes).find {
          case (f1, f2) => !FancyExcelUtils.equalFont(w1.getFontAt(f1), w2.getFontAt(f2))
        }.isDefined
      case (x1: XSSFRichTextString, x2: XSSFRichTextString) =>
        throw new RuntimeException
      //        if (x1.numFormattingRuns != x2.numFormattingRuns) return false
      //        val x1FontIndexes = (0 to x1.numFormattingRuns - 1).map(x1.getFontOfFormattingRun).toList
      //        val x2FontIndexes = (0 to x2.numFormattingRuns - 1).map(x2.getFontOfFormattingRun).toList
      //        !x1FontIndexes.zip(x2FontIndexes).find {case (f1, f2) => !FancyPOIUtil.equalFont(workbook.getFontAt(f1), workbook.getFontAt(f2))}.isDefined
      case _ => false
    }
  }

  def toStringFont(font: Font): String = {
    List("index" -> font.getIndex,
      "fontName" -> font.getFontName,
      "fontHeight" -> font.getFontHeight,
      "italic" -> font.getItalic,
      "strikout" -> font.getStrikeout,
      "bold" -> font.getBold,
      "underline" -> font.getUnderline,
      "typeOffset" -> font.getTypeOffset,
      "charset" -> font.getCharSet,
      "color" -> font.getColor).map {
        case (k, v) => k + "=" + v
      }.mkString(",")
  }

  def toStringStyle(style: CellStyle): String = {
    List(
      "index" -> style.getIndex,
      "fontIndex" -> style.getFontIndex,
      "alignment" -> style.getAlignmentEnum,
      "borderBottom" -> style.getBorderBottomEnum,
      "borderLeft" -> style.getBorderLeftEnum,
      "borderRight" -> style.getBorderRightEnum,
      "borderTop" -> style.getBorderTopEnum,
      "bottomBorderColor" -> style.getBottomBorderColor,
      "dataFormat" -> style.getDataFormat,
      "fillBackgroundColor" -> style.getFillBackgroundColor,
      "fillForegroundColor" -> style.getFillForegroundColor,
      "fillPattern" -> style.getFillPatternEnum,
      "fontIndex" -> style.getFontIndex,
      "hidden" -> style.getHidden,
      "indention" -> style.getIndention,
      "leftBorderColor" -> style.getLeftBorderColor,
      "locked" -> style.getLocked,
      "rightBorderColor" -> style.getRightBorderColor,
      "rotation" -> style.getRotation,
      "topBorderColor" -> style.getTopBorderColor,
      "verticalAlignment" -> style.getVerticalAlignmentEnum,
      "wrapText" -> style.getWrapText).map {
        case (k, v) => k + "=" + v
      }.mkString(",")
  }

  def !![T](any: T) = any match {
    case null => None
    case _    => Some(any)
  }

  class AddrRangeStart(startAddr: String) {

    val numericReg = "^[0-9]+$".r
    val alphabeticalReg = "^[A-Z]+$".r

    def isRowAddr(addr: String) = numericReg.findFirstIn(addr).isDefined

    def isColAddr(addr: String) = alphabeticalReg.findFirstIn(addr).isDefined

    def ~(endAddr: String): List[String] = if (isRowAddr(startAddr) && isRowAddr(endAddr)) {
      (startAddr.toInt to endAddr.toInt).map(_.toString).toList
    } else if (isColAddr(startAddr) && isColAddr(endAddr)) {
      (convertColStringToIndex(startAddr) to convertColStringToIndex(endAddr))
        .map(convertNumToColString).toList
    } else throw new IllegalArgumentException

  }

}
