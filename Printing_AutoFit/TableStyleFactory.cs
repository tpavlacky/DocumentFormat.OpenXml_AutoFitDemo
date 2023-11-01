using DocumentFormat.OpenXml.Wordprocessing;
using Sitraffic.Common.PrintingX.Document.DocumentStyle;

namespace Printing_AutoFit
{
  internal class TableStyleFactory
  {
    internal static Style GenerateStyle()
    {
      var style = new Style() { Type = StyleValues.Table, Default = true, CustomStyle = true };
      var styleName = new StyleName() { Val = "MyGrid" };

      var tableProperties = new StyleTableProperties();
      var tableStyleRowBandSize = new TableStyleRowBandSize() { Val = 1 };
      var indentation = new TableIndentation() { Width = 0, Type = TableWidthUnitValues.Dxa };
      var tableCellMargin = new TableCellMarginDefault(
        new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa },
        new TableCellLeftMargin() { Width = 35, Type = TableWidthValues.Dxa },
        new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa },
        new TableCellRightMargin() { Width = 35, Type = TableWidthValues.Dxa });
      var tableColumnBandSize = new TableStyleColumnBandSize() { Val = 1 };
      var tableBorders = new TableBorders(
        new TopBorder() { Val = BorderValues.Single, Color = SitrafficPrintingStyle.TABLE_BORDER_COLOR, ThemeColor = ThemeColorValues.Accent5, ThemeTint = "99", Size = 4U, Space = 0U },
        new LeftBorder() { Val = BorderValues.Single, Color = SitrafficPrintingStyle.TABLE_BORDER_COLOR, ThemeColor = ThemeColorValues.Accent5, ThemeTint = "99", Size = 4U, Space = 0U },
        new BottomBorder() { Val = BorderValues.Single, Color = SitrafficPrintingStyle.TABLE_BORDER_COLOR, ThemeColor = ThemeColorValues.Accent5, ThemeTint = "99", Size = 4U, Space = 0U },
        new RightBorder() { Val = BorderValues.Single, Color = SitrafficPrintingStyle.TABLE_BORDER_COLOR, ThemeColor = ThemeColorValues.Accent5, ThemeTint = "99", Size = 4U, Space = 0U },
        new InsideHorizontalBorder() { Val = BorderValues.Single, Color = SitrafficPrintingStyle.TABLE_BORDER_COLOR, ThemeColor = ThemeColorValues.Accent5, ThemeTint = "99", Size = 4U, Space = 0U },
        new InsideVerticalBorder() { Val = BorderValues.Single, Color = SitrafficPrintingStyle.TABLE_BORDER_COLOR, ThemeColor = ThemeColorValues.Accent5, ThemeTint = "99", Size = 4U, Space = 0U });
      tableProperties.Append(tableStyleRowBandSize, tableColumnBandSize, indentation, tableBorders, tableCellMargin);

      var paragraphProperties = new StyleParagraphProperties(new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto });
      var runProperties = new StyleRunProperties(new Color() { Val = "2F5496", ThemeColor = ThemeColorValues.Accent5, ThemeShade = "BF" });
      var firstRowStyle = CreateFirstRowStyle();
      var lastRowStyle = CreateLastRowStyleTable();
      var firstColumnStyle = CreateFirstColumnStyle();
      var lastColumnStyle = CreateLastColumnStyle();

      style.Append(
        styleName,
        paragraphProperties,
        runProperties,
        tableProperties,
        firstRowStyle,
        lastRowStyle,
        firstColumnStyle,
        lastColumnStyle
        );

      return style;
    }

    private static TableStyleProperties CreateLastColumnStyle()
    {
      return new TableStyleProperties(CreateBasePropertyStyle()) { Type = TableStyleOverrideValues.LastColumn };
    }

    private static TableStyleProperties CreateFirstColumnStyle()
    {
      var firstColumnStyle = new TableStyleProperties() { Type = TableStyleOverrideValues.FirstColumn };
      //if (_firstColumnProperties != null)
      //{
      //  var runPropertiesBaseStyle = CreateBasePropertyStyle();
      //  firstColumnStyle.Append(runPropertiesBaseStyle, new TableStyleConditionalFormattingTableProperties(), CreateFirstRowStyleForVerticalTable());
      //}
      return firstColumnStyle;
    }

    private static TableStyleProperties CreateFirstRowStyle()
    {
      var firstRowStyle = new TableStyleProperties() { Type = TableStyleOverrideValues.FirstRow };
      //if (_firstRowProperties != null)
      //{
      //  firstRowStyle.Append(CreateBasePropertyStyle(), new TableStyleConditionalFormattingTableProperties(), CreateFirstRowStyleForVerticalTable());
      //}

      return firstRowStyle;
    }

    private static TableStyleProperties CreateLastRowStyleTable()
    {
      var lastRowStyle = new TableStyleProperties() { Type = TableStyleOverrideValues.LastRow };
      var runPropertiesBaseStyle = CreateBasePropertyStyle();
      var conditionalFormattingLastRow = new TableStyleConditionalFormattingTableCellProperties(
        new TableCellBorders(
          new TopBorder() { Val = BorderValues.Double, Color = SitrafficPrintingStyle.TABLE_BORDER_COLOR, ThemeColor = ThemeColorValues.Accent5, ThemeTint = "99", Size = 4U, Space = 0U })
      );
      lastRowStyle.Append(runPropertiesBaseStyle, new TableStyleConditionalFormattingTableProperties(), conditionalFormattingLastRow);
      return lastRowStyle;
    }

    private static RunPropertiesBaseStyle CreateBasePropertyStyle()
    {
      return new RunPropertiesBaseStyle(
        new Bold(),
        new BoldComplexScript());
    }

    private static TableStyleConditionalFormattingTableCellProperties CreateFirstRowStyleForVerticalTable()
    {
      return new TableStyleConditionalFormattingTableCellProperties(
        new TableCellBorders(
          new BottomBorder() { Val = BorderValues.Single, Color = SitrafficPrintingStyle.TABLE_BORDER_COLOR, ThemeColor = ThemeColorValues.Accent5, ThemeTint = "99", Size = 12U, Space = 0U }
        )
      );
    }
  }
}
