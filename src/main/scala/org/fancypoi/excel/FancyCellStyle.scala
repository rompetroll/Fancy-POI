package org.fancypoi.excel

import org.apache.poi.ss.usermodel._

/**
 * CellStyleのバリューオブジェクト
 */
object FancyCellStyle {
  val DEFAULT_CELL_STYLE_INDEX = -1 toShort
}

class FancyCellStyle extends CellStyle {
  private var _align: HorizontalAlignment = HorizontalAlignment.GENERAL
  private var _backgroundColor: Short = IndexedColors.WHITE.getIndex
  private var _borderBottom: BorderStyle = BorderStyle.NONE
  private var _borderLeft: BorderStyle = BorderStyle.NONE
  private var _borderRight: BorderStyle = BorderStyle.NONE
  private var _borderTop: BorderStyle = BorderStyle.NONE
  private var _bottomBorderColor: Short = IndexedColors.BLACK.getIndex
  private var _dataFormat: Short = 0
  private var _fillPattern: FillPatternType = FillPatternType.NO_FILL
  private var _font: Font = _
  private var _foregroundColor: Short = IndexedColors.WHITE.getIndex
  private var _hidden: Boolean = false
  private var _indent: Short = 0
  private var _leftBorderColor: Short = IndexedColors.BLACK.getIndex
  private var _locked: Boolean = true
  private var _rightBorderColor: Short = IndexedColors.BLACK.getIndex
  private var _rotation: Short = 0
  private var _topBorderColor: Short = IndexedColors.BLACK.getIndex
  private var _verticalAlign: VerticalAlignment = VerticalAlignment.BOTTOM
  private var _wrapped: Boolean = false
  private var _shrinkToFit: Boolean = false

  def getIndex: Short = FancyCellStyle.DEFAULT_CELL_STYLE_INDEX

  def getAlignment: Short = _align.getCode

  def setAlignment(align: Short): Unit = _align = HorizontalAlignment.forInt(align)

  def getAlignmentEnum: HorizontalAlignment = _align

  def setAlignment(align: HorizontalAlignment): Unit = _align = align

  def getBorderBottom: Short = _borderBottom.getCode

  def setBorderBottom(border: Short): Unit = _borderBottom = BorderStyle.valueOf(border)

  def getBorderBottomEnum(): BorderStyle = _borderBottom

  def setBorderBottom(border: BorderStyle): Unit = _borderBottom = border

  def getBorderLeft: Short = _borderLeft.getCode

  def setBorderLeft(border: Short): Unit = _borderLeft = BorderStyle.valueOf(border)

  def getBorderLeftEnum(): BorderStyle = _borderLeft

  def setBorderLeft(border: BorderStyle): Unit = _borderLeft = border

  def getBorderRight: Short = _borderRight.getCode

  def setBorderRight(border: Short): Unit = _borderRight = BorderStyle.valueOf(border)

  def getBorderRightEnum(): BorderStyle = _borderRight

  def setBorderRight(border: BorderStyle): Unit = _borderRight = border

  def getBorderTop: Short = _borderTop.getCode

  def setBorderTop(border: Short): Unit = _borderTop = BorderStyle.valueOf(border)

  def getBorderTopEnum(): BorderStyle = _borderTop

  def setBorderTop(border: BorderStyle): Unit = _borderTop = border

  def getBottomBorderColor: Short = _bottomBorderColor

  def setBottomBorderColor(color: Short): Unit = _bottomBorderColor = color

  def getDataFormat: Short = _dataFormat

  def getDataFormatString: String = throw new RuntimeException("Can't stringize data format.")

  def setDataFormat(fmt: Short): Unit = _dataFormat = fmt

  def getFillBackgroundColor: Short = _backgroundColor

  def getFillBackgroundColorColor: Color = throw new RuntimeException("Can't convert color index to color object.")

  def setFillBackgroundColor(bg: Short): Unit = _backgroundColor = bg

  def getFillForegroundColor: Short = _foregroundColor

  def getFillForegroundColorColor: Color = throw new RuntimeException("Can't convert color index to color object.")

  def setFillForegroundColor(fg: Short): Unit = _foregroundColor = fg

  def getFillPattern: Short = _fillPattern.getCode

  def setFillPattern(fp: Short): Unit = _fillPattern = FillPatternType.forInt(fp)

  def getFillPatternEnum: FillPatternType = _fillPattern

  def setFillPattern(fp: FillPatternType): Unit = _fillPattern = fp

  def getFontIndex: Short = _font.getIndex

  def setFont(font: Font): Unit = _font = font

  // CellStyleVOの固有メソッド
  def getFont = _font

  def getHidden: Boolean = _hidden

  def setHidden(hidden: Boolean): Unit = _hidden = hidden

  def getIndention: Short = _indent

  def setIndention(indent: Short): Unit = _indent = indent

  def getLeftBorderColor: Short = _leftBorderColor

  def setLeftBorderColor(color: Short): Unit = _leftBorderColor = color

  def getLocked: Boolean = _locked

  def setLocked(locked: Boolean): Unit = _locked = locked

  def getRightBorderColor: Short = _rightBorderColor

  def setRightBorderColor(color: Short): Unit = _rightBorderColor = color

  def getRotation: Short = _rotation

  def setRotation(rotation: Short): Unit = _rotation = rotation

  def getTopBorderColor: Short = _topBorderColor

  def setTopBorderColor(color: Short): Unit = _topBorderColor = color

  def getVerticalAlignment: Short = _verticalAlign.getCode

  def setVerticalAlignment(align: Short): Unit = _verticalAlign = VerticalAlignment.forInt(align)

  def getVerticalAlignmentEnum(): VerticalAlignment = _verticalAlign

  def setVerticalAlignment(align: VerticalAlignment): Unit = _verticalAlign = align

  def getWrapText: Boolean = _wrapped

  def setWrapText(wrapped: Boolean): Unit = _wrapped = wrapped

  def cloneStyleFrom(source: CellStyle): Unit = throw new UnsupportedOperationException("Can't clone style.")

  def getShrinkToFit: Boolean = _shrinkToFit

  def setShrinkToFit(shrinkToFit: Boolean): Unit = _shrinkToFit = shrinkToFit
}
