using Mercator.OfficeX;
using NPOI.SS.UserModel;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Mercator.ExcelX
{
    public class TemplateHandler
    {
        /// <summary>
        /// 工作簿对象
        /// </summary>
        private IWorkbook _Workbook;

        /// <summary>
        /// 模板文件名
        /// </summary>
        private string _TemplateFileName;

        /// <summary>
        /// 表名
        /// </summary>
        public Dictionary<string, string> Sheets
        {
            get
            {
                Dictionary<string, string> sheets = new Dictionary<string, string>();
                foreach(var sheetName in _SheetNames)
                    sheets.Add(sheetName, "");
                return sheets;
            }
        }
        private List<string> _SheetNames;

        /// <summary>
        /// 占位符集合
        /// </summary>
        public IReadOnlyList<Placeholder> Placeholders
        {
            get
            {
                return _Placeholders;
            }
        }
        private List<Placeholder> _Placeholders;

        /// <summary>
        /// 模板类型
        /// </summary>
        private TemplateType _TemplateType;

        /// <summary>
        /// Excel模板处理器
        /// </summary>
        /// <param name="sheetName">表名</param>
        public TemplateHandler(string templateFileName, TemplateType templateType = TemplateType.Normal)
        {
            _TemplateFileName = templateFileName;
            _TemplateType = templateType;
            _SheetNames = new List<string>();
            _Placeholders = new List<Placeholder>();
            // 获取工作簿对象
            _Workbook = SheetX.GetWorkbook(_TemplateFileName);
            if (_Workbook == null) { return; }

            for (int i = 0; i < _Workbook.NumberOfSheets; i++)
            {
                // 当前表
                var sheet = _Workbook.GetSheetAt(i);
                // 支持公式
                sheet.ForceFormulaRecalculation = true;
                _SheetNames.Add(_Workbook.GetSheetName(i));
                for (int iRow = 0; iRow <= sheet.LastRowNum; iRow++)
                {
                    // 当前行
                    var row = sheet.GetRow(iRow);
                    // 行为空则转到下一行
                    if (row == null) { continue; }
                    for (int iColumn = 0; iColumn <= row.LastCellNum; iColumn++)
                    {
                        // 当前单元格
                        var cell = row.GetCell(iColumn);
                        // 单元格为空则转到下一列
                        if (cell == null) { continue; }
                        // 单元格是合并单元格则将列号设置为合并列的最后一列
                        var dimension = new Dimension();
                        if (SheetX.IsMergeCell(sheet, iRow, iColumn, out dimension)) { iColumn = dimension.LastColumnIndex; }
                        // 获取单元格的字符串值
                        var cellValue = SheetX.GetStringValue(cell);
                        // 单元格的字符串值为空则转到下一列
                        if (string.IsNullOrEmpty(cellValue)) { continue; }
                        // 单元格的字符串值不是占位符则转到下一列
                        if (!Placeholder.IsMatch(cellValue)) { continue; }
                        // 添加占位符对象
                        _Placeholders.Add(new Placeholder(cellValue));
                    }
                }
            }
        }

        /// <summary>
        /// 替换占位符
        /// </summary>
        /// <param name="values">数据</param>
        /// <param name="sheets">要修改的新旧表名字典</param>
        private void ReplacePlaceholders(Hashtable values, Dictionary<string, string> sheets)
        {
            if (_TemplateType != TemplateType.Normal)
            {
                foreach (KeyValuePair<string, string> sheet in sheets)
                {
                    // 复制表并重命名
                    if (!string.IsNullOrEmpty(sheet.Key) && !string.IsNullOrEmpty(sheet.Value) && sheet.Value != sheet.Key && !_SheetNames.Contains(sheet.Key))
                        _Workbook.GetSheet(sheet.Value).CopySheet(sheet.Key);
                }

                // 重新获取全部表名
                var sheetNames = new List<string>();
                for (int i = 0; i < _Workbook.NumberOfSheets; i++)
                    sheetNames.Add(_Workbook.GetSheetName(i));

                foreach (var sheetName in sheetNames)
                {
                    // 删除不需要的表
                    if (!sheets.Keys.Contains(sheetName))
                        _Workbook.RemoveSheetAt(_Workbook.GetSheetIndex(sheetName));
                }
            }

            for (int i=0;i<_Workbook.NumberOfSheets;i++)
            {
                var sheet = _Workbook.GetSheetAt(i);
                SheetX.ReplacePlaceholders(sheet, values);
                sheet.ForceFormulaRecalculation = true;
            }
        }

        /// <summary>
        /// 生成表格
        /// </summary>
        /// <param name="sheetFileName">表格文件名</param>
        /// <param name="values">数据</param>
        /// <param name="sheets">要修改的新旧表名字典</param>
        public void ToSheet(string sheetFileName, Hashtable values, Dictionary<string, string> sheets)
        {
            if (_Workbook.NumberOfSheets <= 0) { return; }
            ReplacePlaceholders(values, sheets);
            using (FileStream fileStream = new FileStream(sheetFileName, FileMode.Create)) { _Workbook.Write(fileStream); }
        }
    }

    public enum TemplateType
    {
        /// <summary>
        /// 普通模板
        /// </summary>
        Normal,
        /// <summary>
        /// 包含抽象表的模板
        /// </summary>
        Abstract,
        /// <summary>
        /// 包含冗余表的模板
        /// </summary>
        Extra
    }
}
