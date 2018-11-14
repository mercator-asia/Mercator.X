using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Mercator.OfficeX
{
    public class SheetX
    {
        /// <summary>
        /// 获取工作簿
        /// </summary>
        /// <param name="fileName">文件名</param>
        /// <returns></returns>
        public static IWorkbook GetWorkbook(string fileName)
        {
            IWorkbook workbook = null;

            if (File.Exists(fileName))
            {
                FileStream fileStream = null;
                try
                {
                    fileStream = File.Open(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                    var extName = Path.GetExtension(fileName).ToLower();
                    if (fileStream != null)
                    {
                        if (extName == ".xls" || extName == ".xlt")
                        {
                            workbook = new HSSFWorkbook(fileStream);
                        }
                        else if (extName == ".xlsx" || extName == ".xltx")
                        {
                            workbook = new XSSFWorkbook(fileStream);
                        }
                    }
                }
                finally
                {
                    fileStream?.Close();
                }
            }          
            
            return workbook;
        }

        /// <summary>
        /// 获取公式计算器
        /// </summary>
        /// <param name="workbook">工作簿</param>
        /// <returns></returns>
        public static IFormulaEvaluator GetFormulaEvaluator(IWorkbook workbook)
        {
            if (workbook is HSSFWorkbook)
            {
                return new HSSFFormulaEvaluator(workbook);
            }
            else
            {
                return new XSSFFormulaEvaluator(workbook);
            }
        }

        /// <summary>
        /// 设置单元格的数据
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <param name="cellData">数据</param>
        public static void SetCellData(ICell cell, CellData cellData)
        {
            if (cell == null || cellData == null) { return; }

            switch(cellData.DataType)
            {
                case CellDataType.Text:
                    cell.SetCellValue(cellData.TextValue);
                    cell.SetCellType(CellType.String);
                    break;
                case CellDataType.Numeric:
                    cell.SetCellValue(cellData.NumericalValue);
                    cell.SetCellType(CellType.Numeric);
                    break;
                case CellDataType.DateTime:
                    cell.SetCellValue(cellData.DateTimeValue);
                    cell.SetCellType(CellType.Numeric);
                    break;
            }

            IWorkbook workbook = cell.Sheet.Workbook;
            IFont font = GetExpectedCellFont(cell, cellData.CellFont);
            short dataFormat = GetExpectedCellDataFormat(cell, cellData.DataFormat);

            SetCellStyle(cell, font, dataFormat);
        }

        /// <summary>
        /// 设置单元格的样式
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <param name="font">字体</param>
        /// <param name="dataFormat">数据格式</param>
        private static short SetCellStyle(ICell cell, IFont font, short dataFormat)
        {
            short cellStyleIndex = -1;

            IWorkbook workbook = cell.Sheet.Workbook;
            for (short iCellStyle = 0; iCellStyle < workbook.NumCellStyles; iCellStyle++)
            {
                var cellStyle = workbook.GetCellStyleAt(iCellStyle);
                if (cellStyle.Alignment == cell.CellStyle.Alignment
                    && cellStyle.BorderBottom == cell.CellStyle.BorderBottom
                    && cellStyle.BorderDiagonal == cell.CellStyle.BorderDiagonal
                    && cellStyle.BorderDiagonalColor == cell.CellStyle.BorderDiagonalColor
                    && cellStyle.BorderDiagonalLineStyle == cell.CellStyle.BorderDiagonalLineStyle
                    && cellStyle.BorderLeft == cell.CellStyle.BorderLeft
                    && cellStyle.BorderRight == cell.CellStyle.BorderRight
                    && cellStyle.BorderTop == cell.CellStyle.BorderTop
                    && cellStyle.BottomBorderColor == cell.CellStyle.BottomBorderColor
                    && cellStyle.FillBackgroundColor == cell.CellStyle.FillBackgroundColor
                    && cellStyle.FillBackgroundColorColor == cell.CellStyle.FillBackgroundColorColor
                    && cellStyle.FillForegroundColor == cell.CellStyle.FillForegroundColor
                    && cellStyle.FillForegroundColorColor == cell.CellStyle.FillForegroundColorColor
                    && cellStyle.FillPattern == cell.CellStyle.FillPattern
                    && cellStyle.Indention == cell.CellStyle.Indention
                    && cellStyle.IsHidden == cell.CellStyle.IsHidden
                    && cellStyle.IsLocked == cell.CellStyle.IsLocked
                    && cellStyle.LeftBorderColor == cell.CellStyle.LeftBorderColor
                    && cellStyle.RightBorderColor == cell.CellStyle.RightBorderColor
                    && cellStyle.Rotation == cell.CellStyle.Rotation
                    && cellStyle.ShrinkToFit == cell.CellStyle.ShrinkToFit
                    && cellStyle.TopBorderColor == cell.CellStyle.TopBorderColor
                    && cellStyle.VerticalAlignment == cell.CellStyle.VerticalAlignment
                    && cellStyle.WrapText == cell.CellStyle.WrapText
                    && cellStyle.DataFormat == dataFormat
                    && cellStyle.FontIndex == font.Index)
                {
                    cell.CellStyle = cellStyle;
                    cellStyleIndex = iCellStyle;
                    break;
                }
            }

            if (cellStyleIndex < 0)
            {
                ICellStyle newCellStyle = workbook.CreateCellStyle();
                newCellStyle.CloneStyleFrom(cell.CellStyle);
                newCellStyle.DataFormat = dataFormat;
                newCellStyle.SetFont(font);
                cell.CellStyle = newCellStyle;
                cellStyleIndex = newCellStyle.Index;
            }

            return cellStyleIndex;
        }

        /// <summary>
        /// 获取单元格的期望数据格式（指定数据格式存在则返回已有值，否则创建新数据格式）
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <param name="dataFormatBySet">指定格式</param>
        /// <returns></returns>
        private static short GetExpectedCellDataFormat(ICell cell, string dataFormat)
        {
            short expectedDataFormat = -1;

            IWorkbook workbook = cell.Sheet.Workbook;

            if (!string.IsNullOrEmpty(dataFormat))
            {
                short builtinFormat = HSSFDataFormat.GetBuiltinFormat(dataFormat);
                if (builtinFormat < 0)
                {
                    expectedDataFormat = workbook.CreateDataFormat().GetFormat(dataFormat);
                }
                else
                {
                    expectedDataFormat = builtinFormat;
                }
            }
            else
            {
                expectedDataFormat = cell.CellStyle.DataFormat;
            }

            return expectedDataFormat;
        }

        /// <summary>
        /// 获取单元格的期望字体（指定字体存在则返回已有值，否则创建新字体）
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <param name="fontBySet">指定字体</param>
        /// <returns></returns>
        private static IFont GetExpectedCellFont(ICell cell, CellFont font)
        {
            IFont expectedFont = null;

            IWorkbook workbook = cell.Sheet.Workbook;
            IFont fontInCell = cell.CellStyle.GetFont(workbook);

            short boldWeight = fontInCell.Boldweight;
            short color = fontInCell.Color;
            short fontHeight = fontInCell.FontHeightInPoints;
            string name = fontInCell.FontName;
            bool italic = fontInCell.IsItalic;
            bool strikeout = fontInCell.IsStrikeout;
            FontSuperScript typeOffset = fontInCell.TypeOffset;
            FontUnderlineType underline = fontInCell.Underline;

            if (font.ColorChanged) { color = font.Color; }
            if (font.IsBoldChanged) { boldWeight = (short)FontBoldWeight.Bold; }
            if (font.FontNameChanged) { name = font.FontName; }
            if (font.IsFontHeightChanged) { fontHeight = font.FontHeight; }
            if (font.IsItalicChanged) { italic = font.IsItalic; }
            if (font.IsStrikeoutChanged) { strikeout = font.IsStrikeout; }
            if (font.IsUnderlineTypeChanged) { underline = font.UnderlineType; }
            if (font.IsTypeOffsetChanged) { typeOffset = font.TypeOffset; }

            expectedFont = workbook.FindFont(boldWeight, color, (short)(fontHeight*20), name, italic, strikeout, typeOffset, underline);

            if (expectedFont == null)
            {
                expectedFont = workbook.CreateFont();
                expectedFont.Boldweight = boldWeight;
                expectedFont.Color = color;
                expectedFont.FontHeightInPoints = fontHeight;
                expectedFont.FontName = name;
                expectedFont.IsItalic = italic;
                expectedFont.IsStrikeout = strikeout;
                expectedFont.TypeOffset = typeOffset;
                expectedFont.Underline = underline;
            }

            return expectedFont;
        }

        /// <summary>
        /// 判断单元格是否为合并单元格
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="rowIndex">行号</param>
        /// <param name="columnIndex">列号</param>
        /// <param name="dimension">维度</param>
        /// <returns></returns>
        public static bool IsMergeCell(ISheet sheet, int rowIndex, int columnIndex, out Dimension dimension)
        {
            dimension = new Dimension
            {
                FirstRowIndex = rowIndex,
                LastRowIndex = rowIndex,
                FirstColumnIndex = columnIndex,
                LastColumnIndex = columnIndex
            };

            bool result = false;

            for (int i = 0; i < sheet.NumMergedRegions; i++)
            {
                CellRangeAddress range = sheet.GetMergedRegion(i);
                sheet.IsMergedRegion(range);

                if ((rowIndex >= range.FirstRow && range.LastRow >= rowIndex) && (columnIndex >= range.FirstColumn && range.LastColumn >= columnIndex))
                {
                    dimension.FirstRowIndex = range.FirstRow;
                    dimension.LastRowIndex = range.LastRow;
                    dimension.FirstColumnIndex = range.FirstColumn;
                    dimension.LastColumnIndex = range.LastColumn;

                    result = true;
                    break;
                }
            }

            return result;
        }

        /// <summary>
        /// 查找关键字所在的行
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="keyword">关键字</param>
        /// <param name="columnIndex">列号</param>
        /// <returns></returns>
        public static int FindRow(ISheet sheet, string keyword, int columnIndex = 0)
        {
            for (int i = 0; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                ICell cell = row.GetCell(columnIndex);
                if (cell.StringCellValue.Equals(keyword))
                {
                    return i;
                }
            }
            return -1;
        }

        /// <summary>
        /// 查找关键字所在的列
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="keyword">关键字</param>
        /// <param name="rowIndex">行号</param>
        /// <returns></returns>
        public static int FindColumn(ISheet sheet, string keyword, int rowIndex = 0)
        {
            IRow row = sheet.GetRow(rowIndex);

            for (int i = 0; i < row.LastCellNum; i++)
            {
                ICell cell = row.GetCell(i);
                if (cell.StringCellValue.Equals(keyword))
                {
                    return i;
                }
            }
            return -1;
        }

        /// <summary>
        /// 插入行
        /// </summary>
        /// <param name="sheet">工作簿</param>
        /// <param name="starRow">起始行</param>
        /// <param name="rows">行数</param>
        private static void InsertRow(ISheet sheet, int starRow, int rows)
        {
            /*
             * ShiftRows(int startRow, int endRow, int n, bool copyRowHeight, bool resetOriginalRowHeight); 
             * 
             * startRow 开始行
             * endRow 结束行
             * n 移动行数
             * copyRowHeight 复制的行是否高度在移
             * resetOriginalRowHeight 是否设置为默认的原始行的高度
             * 
             */
            var startRowIndex = sheet.GetRow(starRow + 1) != null ? starRow + 1 : starRow;
            sheet.ShiftRows(startRowIndex, sheet.LastRowNum, rows, true, true); // 从下一行开始复制，避免第一行的样式被覆盖

            var sourceRow = sheet.GetRow(starRow);

            for (int i = 0; i < rows; i++)
            {
                IRow targetRow = null;
                ICell sourceCell = null;
                ICell targetCell = null;

                short m;

                targetRow = sheet.CreateRow(starRow + i + 1);
                targetRow.HeightInPoints = sourceRow.HeightInPoints;

                for (m = (short)sourceRow.FirstCellNum; m < sourceRow.LastCellNum; m++)
                {

                    sourceCell = sourceRow.GetCell(m);
                    targetCell = targetRow.CreateCell(m);

                    targetCell.CellStyle = sourceCell.CellStyle;
                    targetCell.SetCellType(sourceCell.CellType);

                }
            }
        }

        /// <summary>
        /// 获取端锚
        /// </summary>
        /// <param name="workbook">工作簿</param>
        /// <param name="dx1">左上角单元格的Left值</param>
        /// <param name="dy1">左上角单元格的Top值</param>
        /// <param name="dx2">右下角单元格的Left值</param>
        /// <param name="dy2">右下角单元格的Top值</param>
        /// <param name="col1">左上角单元格的列</param>
        /// <param name="row1">左上角单元格的行</param>
        /// <param name="col2">右下角单元格的列</param>
        /// <param name="row2">右下角单元格的行</param>
        /// <returns></returns>
        private static IClientAnchor GetClientAnchor(IWorkbook workbook, int dx1, int dy1, int dx2, int dy2, int col1, int row1, int col2, int row2)
        {
            // 定义插入图片锚点
            IClientAnchor anchor = null;
            var workbookTypeName = workbook.GetType().Name;
            if (workbookTypeName.Equals("HSSFWorkbook"))
            {
                anchor = new HSSFClientAnchor(dx1, dy1, dx2, dy2, col1, row1, col2, row2);
            }
            else if (workbookTypeName.Equals("XSSFWorkbook"))
            {
                anchor = new XSSFClientAnchor(dx1, dy1, dx2, dy2, col1, row1, col2, row2);
            }
            return anchor;
        }

        /// <summary>
        /// 设置单元格的图片
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <param name="imageFileName">图片文件名</param>
        public static void SetCellPicture(ICell cell, string pictureFileName)
        {
            if (string.IsNullOrEmpty(pictureFileName) || !File.Exists(pictureFileName)) { return; }

            byte[] pictureData = null;
            using (FileStream file = new FileStream(pictureFileName, FileMode.Open))
            {
                pictureData = new byte[file.Length];
                file.Read(pictureData, 0, pictureData.Length);
            }
            if (pictureData == null || pictureData.Length <= 0) { return; }

            // 图片格式
            var format = PictureType.None;
            // 获取文件扩展名
            string extension = Path.GetExtension(pictureFileName).ToLower();
            switch (extension)
            {
                case ".bmp":
                    format = PictureType.BMP;
                    break;
                case ".jpg":
                    format = PictureType.JPEG;
                    break;
                case ".png":
                    format = PictureType.PNG;
                    break;
                case ".gif":
                    format = PictureType.GIF;
                    break;
                case ".tif":
                    format = PictureType.TIFF;
                    break;
                case ".emf":
                    format = PictureType.EMF;
                    break;
            }

            IWorkbook workbook = cell.Sheet.Workbook;

            int dx1 = 10, dy1 = 0;
            int col1 = cell.ColumnIndex, row1 = cell.RowIndex;
            int dx2 = 1023, dy2 = 255;
            int col2 = col1, row2 = row1;
            var dimension = new Dimension();
            if (IsMergeCell(cell.Sheet, cell.RowIndex, cell.ColumnIndex, out dimension))
            {
                col2 = dimension.LastColumnIndex;
                row2 = dimension.LastRowIndex;
            }

            IDrawing drawingPatriarch = cell.Sheet.DrawingPatriarch != null ? cell.Sheet.DrawingPatriarch : cell.Sheet.CreateDrawingPatriarch();
            IClientAnchor clientAnchor = GetClientAnchor(workbook, dx1, dy1, dx2, dy2, col1, row1, col2, row2);
            int pictureIndex = workbook.AddPicture(pictureData, format);
            IPicture picture = drawingPatriarch.CreatePicture(clientAnchor, pictureIndex);
            picture.Resize();
        }

        /// <summary>
        /// 替换占位符
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="hashTable">包含值的占位符哈希表</param>
        public static void ReplacePlaceholders(ISheet sheet, Hashtable hashTable)
        {
            IWorkbook workbook = sheet.Workbook;

            if (workbook == null || sheet == null || hashTable == null) { return; }

            // 遍历模板表的所有行
            for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                // 获取当前行
                IRow row = sheet.GetRow(rowIndex);
                if (row == null) { continue; }
                // 遍历当前行的所有列
                for (int columnIndex = 0; columnIndex <= row.LastCellNum; columnIndex++)
                {
                    // 当前列
                    ICell cell = row.GetCell(columnIndex);
                    // 如果当前单元格是合并单元格则将列号设置为合并列的最后一列
                    var dimension = new Dimension();
                    if (IsMergeCell(sheet, rowIndex, columnIndex, out dimension)) { columnIndex = dimension.LastColumnIndex; }
                    var cellValue = Regex.Replace(GetStringValue(cell), @"\s", "");
                    // 当前单元格不是字符串类型时转到下一列
                    if (string.IsNullOrEmpty(cellValue)) { continue; }
                    // 匹配占位符字符串
                    if (!Placeholder.IsMatch(cellValue)) { continue; }

                    // 从包含值的占位符哈希表查找与当前单元格匹配的元素执行替换
                    foreach (DictionaryEntry entry in hashTable)
                    {
                        // 如果DictionaryEntry的值为空则转到下一次循环
                        if (entry.Value == null) { continue; }

                        // 取得占位符对象
                        var placeholder = (Placeholder)entry.Key;
                        // 获取单元格中定义的占位符对象
                        var placeholderFromCell = new Placeholder(Regex.Replace(GetStringValue(cell), @"\s", ""));
                        // 判断占位符是否与单元格中定义的占位符对象匹配
                        if (!Equals(placeholderFromCell.Key, placeholder.Key) || !Equals(placeholderFromCell.Type, placeholder.Type)) { continue; }
                        // 清空占位符
                        cell.SetCellValue("");

                        // 根据占位符类型确定替换方式
                        switch (placeholder.Type)
                        {
                            case PlaceholderType.Text:
                                SetCellData(cell, (CellData)entry.Value);
                                break;
                            case PlaceholderType.DateTime:
                                SetCellData(cell, (CellData)entry.Value);
                                break;
                            case PlaceholderType.Numeric:
                                SetCellData(cell, (CellData)entry.Value);
                                break;
                            case PlaceholderType.Picture:
                                SetCellPicture(cell, ((CellData)entry.Value).TextValue);
                                break;
                            case PlaceholderType.Table:
                                rowIndex = SetCellTable(cell, (CellData[,])entry.Value, rowIndex);
                                break;
                        }
                        break;
                    }
                }
            }
        }

        /// <summary>
        /// 填充表格
        /// </summary>
        /// <param name="cell">占位符所在的单元格</param>
        /// <param name="values">占位符的值</param>
        /// <param name="startRow">起始行（起始行小于0则表示不插入行）</param>
        public static int SetCellTable(ICell cell, CellData[,] values, int startRow = -1)
        {
            ISheet sheet = cell.Sheet;
            IWorkbook workbook = sheet.Workbook;

            // 获取占位符所在单元格的行号
            var rowIndex = cell.RowIndex;
            // 获取占位符所在单元格的列号
            var columnIndex = cell.ColumnIndex;
            // 要素二维数组的行数
            var elementRowCount = values.GetLength(0);
            // 要素二维数组的列数
            var elementColumnCount = values.GetLength(1);
            // 如果起始行参数大于0则插入新行
            if (startRow >= 0) { InsertRow(sheet, startRow, elementRowCount - 1); }
            // 遍历新增加行的每一行（将要素中的每个元素值插入到对应新增的单元格中）
            for (int i = 0; i < elementRowCount; i++)
            {
                // 获取当前行
                var tableRow = sheet.GetRow(i + rowIndex);
                if (tableRow == null) { tableRow = sheet.CreateRow(i + rowIndex); }
                // 遍历当前行中的每一列
                for (int j = 0; j < elementColumnCount; j++)
                {
                    // 获取当前单元格
                    var tableCell = tableRow.GetCell(j + columnIndex);
                    if (tableCell == null) { tableCell = tableRow.CreateCell(j + columnIndex); }
                    // 获取当前要素
                    var value = values[i, j];
                    // 设置单元格的值和样式
                    SetCellData(tableCell, value);
                }
            }

            // 返回插入表后的最后一行
            return rowIndex + elementRowCount - 1;
        }

        /// <summary>
        /// 获取单元格的值（期待数字）
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <returns></returns>
        public static double GetDoubleValue(ICell cell)
        {
            var value = 0d;

            if (cell != null)
            {
                switch (cell.CellType)
                {
                    case CellType.Blank:
                        break;
                    case CellType.Boolean:
                        break;
                    case CellType.Error:
                        break;
                    case CellType.Formula:
                        var formulaEvaluator = GetFormulaEvaluator(cell.Sheet.Workbook);
                        var cellValue = formulaEvaluator.Evaluate(cell);
                        switch (cellValue.CellType)
                        {
                            case CellType.Blank:
                                break;
                            case CellType.Boolean:
                                if (cellValue.BooleanValue) { value = 1; } else { value = 0; }
                                break;
                            case CellType.Error:
                                break;
                            case CellType.Numeric:
                                value = cellValue.NumberValue;
                                break;
                            case CellType.String:
                                try
                                {
                                    if (!string.IsNullOrEmpty(cell.StringCellValue.Trim()))
                                    {
                                        value = Convert.ToDouble(cellValue.StringValue);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    throw (ex);
                                }
                                break;
                            case CellType.Unknown:
                                break;
                            case CellType.Formula: // CellType.Formula will never happen
                                break;
                        }
                        break;
                    case CellType.Numeric:
                        value = cell.NumericCellValue;
                        break;
                    case CellType.String:
                        try
                        {
                            if (!string.IsNullOrEmpty(cell.StringCellValue.Trim()))
                            {
                                value = Convert.ToDouble(cell.StringCellValue);
                            }
                        }
                        catch (Exception ex)
                        {
                            throw (ex);
                        }
                        break;
                    case CellType.Unknown:
                        break;
                }

            }

            return value;
        }

        /// <summary>
        /// 获取单元格的值（期待日期）
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <returns></returns>
        public static DateTime GetDateValue(ICell cell)
        {
            var value = DateTime.MaxValue;

            if (cell != null)
            {
                switch (cell.CellType)
                {
                    case CellType.Blank:
                        break;
                    case CellType.Boolean:
                        break;
                    case CellType.Error:
                        break;
                    case CellType.Formula:
                        var formulaEvaluator = GetFormulaEvaluator(cell.Sheet.Workbook);
                        var cellValue = formulaEvaluator.Evaluate(cell);
                        switch (cellValue.CellType)
                        {
                            case CellType.Blank:
                                break;
                            case CellType.Boolean:
                                break;
                            case CellType.Error:
                                break;
                            case CellType.Numeric:
                                try
                                {
                                    if (DateUtil.IsValidExcelDate(cellValue.NumberValue))
                                        value = cell.DateCellValue;
                                }
                                catch (Exception ex)
                                {
                                    throw (ex);
                                }
                                break;
                            case CellType.String:
                                try
                                {
                                    value = Convert.ToDateTime(cellValue.StringValue);
                                }
                                catch (Exception ex)
                                {
                                    throw (ex);
                                }
                                break;
                            case CellType.Unknown:
                                break;
                            case CellType.Formula: // CellType.Formula will never happen
                                break;
                        }
                        break;
                    case CellType.Numeric:
                        try
                        {
                            value = cell.DateCellValue;
                        }
                        catch (Exception ex)
                        {
                            throw (ex);
                        }
                        break;
                    case CellType.String:
                        try
                        {
                            if (!string.IsNullOrEmpty(cell.StringCellValue.Trim()))
                            {
                                value = Convert.ToDateTime(cell.StringCellValue);
                            }
                        }
                        catch (Exception ex)
                        {
                            throw (ex);
                        }
                        break;
                    case CellType.Unknown:
                        break;
                }

            }

            return value;
        }

        /// <summary>
        /// 获取单元格的值（期待字符串）
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static string GetStringValue(ICell cell)
        {
            var value = string.Empty;

            if (cell != null)
            {
                switch (cell.CellType)
                {
                    case CellType.Blank:
                        break;
                    case CellType.Boolean:
                        value = cell.BooleanCellValue ? "是" : "否";
                        break;
                    case CellType.Error:
                        value = cell.ErrorCellValue.ToString();
                        break;
                    case CellType.Formula:
                        var formulaEvaluator = GetFormulaEvaluator(cell.Sheet.Workbook);
                        var cellValue = formulaEvaluator.Evaluate(cell);
                        switch (cellValue.CellType)
                        {
                            case CellType.Blank:
                                break;
                            case CellType.Boolean:
                                value = cellValue.BooleanValue ? "是" : "否";
                                break;
                            case CellType.Error:
                                value = cellValue.ErrorValue.ToString();
                                break;
                            case CellType.Numeric:
                                value = cellValue.NumberValue.ToString();
                                break;
                            case CellType.String:
                                value = cellValue.StringValue;
                                break;
                            case CellType.Unknown:
                                break;
                            case CellType.Formula: // CellType.Formula will never happen
                                break;
                        }
                        break;
                    case CellType.Numeric:
                        value = cell.NumericCellValue.ToString();
                        break;
                    case CellType.String:
                        value = cell.StringCellValue;
                        break;
                    case CellType.Unknown:
                        break;
                }

            }

            return value;
        }

        /// <summary>
        /// 获取单元格的值（期待布尔值）
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static bool GetBooleanValue(ICell cell)
        {
            var value = false;

            if (cell != null)
            {
                switch (cell.CellType)
                {
                    case CellType.Blank:
                        break;
                    case CellType.Boolean:
                        value = cell.BooleanCellValue;
                        break;
                    case CellType.Error:
                        break;
                    case CellType.Formula:
                        var formulaEvaluator = GetFormulaEvaluator(cell.Sheet.Workbook);
                        var cellValue = formulaEvaluator.Evaluate(cell);
                        switch (cellValue.CellType)
                        {
                            case CellType.Blank:
                                break;
                            case CellType.Boolean:
                                value = cellValue.BooleanValue;
                                break;
                            case CellType.Error:
                                break;
                            case CellType.Numeric:
                                if (cellValue.NumberValue > 0)
                                    value = true;
                                break;
                            case CellType.String:
                                switch (cellValue.StringValue)
                                {
                                    case "1":
                                        value = true;
                                        break;
                                    case "是":
                                        value = true;
                                        break;
                                    case "对":
                                        value = true;
                                        break;
                                    case "男":
                                        value = true;
                                        break;
                                }
                                break;
                            case CellType.Unknown:
                                break;
                            case CellType.Formula: // CellType.Formula will never happen
                                break;
                        }
                        break;
                    case CellType.Numeric:
                        if (cell.NumericCellValue > 0)
                            value = true;
                        break;
                    case CellType.String:
                        switch (cell.StringCellValue)
                        {
                            case "1":
                                value = true;
                                break;
                            case "是":
                                value = true;
                                break;
                            case "对":
                                value = true;
                                break;
                            case "男":
                                value = true;
                                break;
                        }
                        break;
                    case CellType.Unknown:
                        break;
                }

            }

            return value;
        }
    }

    public class CellData
    {
        /// <summary>
        /// 数值
        /// </summary>
        public double NumericalValue;

        /// <summary>
        /// 文本值
        /// </summary>
        public string TextValue;

        /// <summary>
        /// 时间日期值
        /// </summary>
        public DateTime DateTimeValue;

        /// <summary>
        /// 数据格式
        /// </summary>
        public string DataFormat = string.Empty;

        /// <summary>
        /// 数据类型
        /// </summary>
        public CellDataType DataType
        {
            get
            {
                return _DataType;
            }
        }
        private CellDataType _DataType;

        /// <summary>
        /// 字体
        /// </summary>
        public CellFont CellFont;

        /// <summary>
        /// 单元格数据
        /// </summary>
        public CellData()
        {
            NumericalValue = double.MaxValue;
            DateTimeValue = DateTime.MaxValue;
            TextValue = string.Empty;

            CellFont = new CellFont();
        }

        /// <summary>
        /// 数值型单元格数据
        /// </summary>
        /// <param name="numericalValue">数值</param>
        public CellData(double numericalValue) : this()
        {
            NumericalValue = numericalValue;
            _DataType = CellDataType.Numeric;
        }

        /// <summary>
        /// 数值型单元格数据
        /// </summary>
        /// <param name="numericalValue">数值</param>
        /// <param name="dataFormat">数据格式</param>
        public CellData(double numericalValue, string dataFormat) : this(numericalValue)
        {
            DataFormat = dataFormat;
        }

        /// <summary>
        /// 数值型单元格数据
        /// </summary>
        /// <param name="numericalValue">数值</param>
        /// <param name="dataFormat">数据格式</param>
        /// <param name="cellFont">字体</param>
        public CellData(double numericalValue, string dataFormat, CellFont cellFont) : this(numericalValue, dataFormat)
        {
            CellFont = cellFont;
        }

        /// <summary>
        /// 时间日期型单元格数据
        /// </summary>
        /// <param name="dateTimeValue">时间日期值</param>
        public CellData(DateTime dateTimeValue) : this()
        {
            DateTimeValue = dateTimeValue;
            _DataType = CellDataType.DateTime;
        }

        /// <summary>
        /// 时间日期型单元格数据
        /// </summary>
        /// <param name="dateTimeValue">时间日期值</param>
        /// <param name="dataFormat">数据格式</param>
        public CellData(DateTime dateTimeValue, string dataFormat) : this(dateTimeValue)
        {
            DataFormat = dataFormat;
        }

        /// <summary>
        /// 时间日期型单元格数据
        /// </summary>
        /// <param name="dateTimeValue">时间日期值</param>
        /// <param name="dataFormat">数据格式</param>
        /// <param name="cellFont">字体</param>
        public CellData(DateTime dateTimeValue, string dataFormat, CellFont cellFont) : this(dateTimeValue, dataFormat)
        {
            CellFont = cellFont;
        }

        /// <summary>
        /// 文本型单元格数据
        /// </summary>
        /// <param name="textValue">文本值</param>
        public CellData(string textValue) : this()
        {
            TextValue = textValue;
            _DataType = CellDataType.Text;
        }

        /// <summary>
        /// 文本型单元格数据
        /// </summary>
        /// <param name="textValue">文本值</param>
        /// <param name="dataFormat">数据格式</param>
        public CellData(string textValue, string dataFormat) : this(textValue)
        {
            DataFormat = dataFormat;
        }

        /// <summary>
        /// 文本型单元格数据
        /// </summary>
        /// <param name="textValue">文本值</param>
        /// <param name="dataFormat">数据格式</param>
        /// <param name="cellFont">字体</param>
        public CellData(string textValue, string dataFormat, CellFont cellFont) : this(textValue, dataFormat)
        {
            CellFont = cellFont;
        }
    }

    public enum CellDataType
    {
        /// <summary>
        /// 数字型
        /// </summary>
        Numeric,
        /// <summary>
        /// 文本型
        /// </summary>
        Text,
        /// <summary>
        /// 日期型
        /// </summary>
        DateTime
    }

    public class CellFont
    {
        /// <summary>
        /// 字体
        /// </summary>
        public string FontName
        {
            get
            {
                return _FontName;
            }
            set
            {
                _FontName = value;
                _FontNameChanged = true;
            }
        }
        private string _FontName = "Times New Roman";

        /// <summary>
        /// 字体赋值标志
        /// </summary>
        public bool FontNameChanged
        {
            get
            {
                return _FontNameChanged;
            }
        }
        private bool _FontNameChanged = false;

        /// <summary>
        /// 颜色
        /// </summary>
        public short Color
        {
            get
            {
                return _Color;
            }
            set
            {
                _Color = value;
                _ColorChanged = true;

            }
        }
        private short _Color = IndexedColors.Black.Index;

        /// <summary>
        /// 颜色赋值标志
        /// </summary>
        public bool ColorChanged
        {
            get
            {
                return _ColorChanged;
            }
        }
        private bool _ColorChanged = false;

        /// <summary>
        /// 粗体
        /// </summary>
        public bool IsBold
        {
            get
            {
                return _IsBold;
            }
            set
            {
                _IsBold = value;
                _IsBoldChanged = true;
            }
        }
        private bool _IsBold = false;

        /// <summary>
        /// 粗体赋值标志
        /// </summary>
        public bool IsBoldChanged
        {
            get
            {
                return _IsBoldChanged;
            }
        }
        private bool _IsBoldChanged = false;

        /// <summary>
        /// 斜体
        /// </summary>
        public bool IsItalic
        {
            get
            {
                return _IsItalic;
            }
            set
            {
                _IsItalic = value;
                _IsItalicChanged = true;
            }
        }
        private bool _IsItalic = false;

        /// <summary>
        /// 斜体赋值标志
        /// </summary>
        public bool IsItalicChanged
        {
            get
            {
                return _IsItalicChanged;
            }
        }
        private bool _IsItalicChanged = false;

        /// <summary>
        /// 字号
        /// </summary>
        public short FontHeight
        {
            get
            {
                return _FontHeight;
            }
            set
            {
                _FontHeight = value;
                _IsFontHeightChanged = true;
            }
        }
        private short _FontHeight = 10;

        /// <summary>
        /// 字号赋值标志
        /// </summary>
        public bool IsFontHeightChanged
        {
            get
            {
                return _IsFontHeightChanged;
            }
        }
        private bool _IsFontHeightChanged = false;

        /// <summary>
        /// 删除线
        /// </summary>
        public bool IsStrikeout
        {
            get
            {
                return _IsStrikeout;
            }
            set
            {
                _IsStrikeout = value;
                _IsStrikeoutChanged = true;
            }
        }
        private bool _IsStrikeout = false;

        /// <summary>
        /// 删除线赋值标志
        /// </summary>
        public bool IsStrikeoutChanged
        {
            get
            {
                return _IsStrikeoutChanged;
            }
        }
        private bool _IsStrikeoutChanged = false;

        /// <summary>
        /// 下划线
        /// </summary>
        public FontUnderlineType UnderlineType
        {
            get
            {
                return _UnderlineType;
            }
            set
            {
                _UnderlineType = value;
                _IsStrikeoutChanged = true;
            }
        }
        private FontUnderlineType _UnderlineType = FontUnderlineType.None;

        /// <summary>
        /// 下划线赋值标志
        /// </summary>
        public bool IsUnderlineTypeChanged
        {
            get
            {
                return _IsUnderlineTypeChanged;
            }
        }
        private bool _IsUnderlineTypeChanged = false;

        /// <summary>
        /// 上下标
        /// </summary>
        public FontSuperScript TypeOffset
        {
            get
            {
                return _TypeOffset;
            }
            set
            {
                _TypeOffset = value;
                _IsTypeOffsetChanged = true;
            }
        }
        private FontSuperScript _TypeOffset = FontSuperScript.None;

        /// <summary>
        /// 上下标赋值标志
        /// </summary>
        public bool IsTypeOffsetChanged
        {
            get
            {
                return _IsTypeOffsetChanged;
            }
        }
        private bool _IsTypeOffsetChanged = false;

        public static short GetColorIndex(FontColor fontColor)
        {
            short color = 0;
            switch(fontColor)
            {
                case FontColor.Black:
                    color = IndexedColors.Black.Index;
                    break;
                case FontColor.Red:
                    color = IndexedColors.Red.Index;
                    break;
                case FontColor.Orange:
                    color = IndexedColors.Orange.Index;
                    break;
                case FontColor.White:
                    color = IndexedColors.White.Index;
                    break;
                case FontColor.Blue:
                    color = IndexedColors.Blue.Index;
                    break;
                case FontColor.Yellow:
                    color = IndexedColors.Yellow.Index;
                    break;
                case FontColor.Green:
                    color = IndexedColors.Green.Index;
                    break;
                case FontColor.Pink:
                    color = IndexedColors.Pink.Index;
                    break;
            }
            return color;
        }
    }

    public enum FontColor
    {
        Black,
        Red,
        Orange,
        White,
        Blue,
        Yellow,
        Green,
        Pink
    }

    /// <summary>
    /// 合并单元格维度
    /// </summary>
    public struct Dimension
    {
        /// <summary>
        /// 合并单元格的起始行索引
        /// </summary>
        public int FirstRowIndex;
        /// <summary>
        /// 合并单元格的结束行索引
        /// </summary>
        public int LastRowIndex;
        /// <summary>
        /// 合并单元格的起始列索引
        /// </summary>
        public int FirstColumnIndex;
        /// <summary>
        /// 合并单元格的结束列索引
        /// </summary>
        public int LastColumnIndex;
    }

    /// <summary>
    /// 占位符类型
    /// </summary>
    public enum PlaceholderType
    {
        /// <summary>
        /// 文本
        /// </summary>
        Text,
        /// <summary>
        /// 日期
        /// </summary>
        DateTime,
        /// <summary>
        /// 图片
        /// </summary>
        Picture,
        /// <summary>
        /// 数值
        /// </summary>
        Numeric,
        /// <summary>
        /// 表格
        /// </summary>
        Table
    }

    /// <summary>
    /// 占位符
    /// </summary>
    public class Placeholder
    {
        /// <summary>
        /// 关键字
        /// </summary>
        public string Key
        {
            get
            {
                return _Key;
            }
            set
            {
                _Key = Regex.Replace(value, @"\s", "");
            }
        }
        private string _Key;
        /// <summary>
        /// 占位符类型
        /// </summary>
        public PlaceholderType Type;

        /// <summary>
        /// 占位符
        /// </summary>
        public Placeholder()
        {

        }

        /// <summary>
        /// 占位符
        /// </summary>
        /// <param name="key">关键字</param>
        /// <param name="type">占位符类型</param>
        public Placeholder(string key, PlaceholderType type)
        {
            Key = key;
            Type = type;
        }

        /// <summary>
        /// 占位符
        /// </summary>
        /// <param name="placeholderTag">占位符标签</param>
        public Placeholder(string placeholderTag)
        {
            placeholderTag = Regex.Replace(placeholderTag, @"\s", "");
            string[] tags = placeholderTag.Split('.');
            Key = tags[0].ToLower();
            var type = tags[1].Substring(1, tags[1].Length - 2);
            switch (type)
            {
                case "文本":
                    Type = PlaceholderType.Text;
                    break;
                case "日期":
                    Type = PlaceholderType.DateTime;
                    break;
                case "数字":
                    Type = PlaceholderType.Numeric;
                    break;
                case "图片":
                    Type = PlaceholderType.Picture;
                    break;
                case "表格":
                    Type = PlaceholderType.Table;
                    break;
            }
        }

        /// <summary>
        /// 匹配字符串（格式为：{关键字}.[类型]）
        /// </summary>
        /// <param name="input">输入字符串</param>
        /// <returns></returns>
        public static bool IsMatch(string input)
        {
            input = Regex.Replace(input, @"\s", "");
            var result = false;
            var patterns = new List<string>();
            patterns.Add(@"^[\{]\S+[\}]\.[\[](文本)[\]]");
            patterns.Add(@"^[\{]\S+[\}]\.[\[](日期)[\]]");
            patterns.Add(@"^[\{]\S+[\}]\.[\[](数字)[\]]");
            patterns.Add(@"^[\{]\S+[\}]\.[\[](图片)[\]]");
            patterns.Add(@"^[\{]\S+[\}]\.[\[](表格)[\]]");

            foreach (var pattern in patterns)
            {
                if (Regex.IsMatch(input, pattern))
                {
                    result = true;
                    break;
                }
            }

            return result;
        }

        /// <summary>
        /// 判断是否存在指定名称的占位符
        /// </summary>
        /// <param name="table"></param>
        /// <param name="placeholderName"></param>
        /// <returns></returns>
        public static bool ContainsKey(Hashtable table, string placeholderName)
        {
            placeholderName = Regex.Replace(placeholderName, @"\s", "");
            foreach (DictionaryEntry de in table)
            {
                var palceholder = (Placeholder)de.Key;
                if (palceholder.Key.Equals(placeholderName))
                {
                    return true;
                }
            }

            return false;
        }
    }
}
