using Mercator.ShapeX;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml;

namespace Mercator.Evaluate.Assistant
{
    public partial class MainForm : Form
    {
        private string _SHPFileName;

        private List<Patch> _patches;

        public MainForm()
        {
            InitializeComponent();

            _patches = new List<Patch>();
        }

        public delegate void AppendTextCallback(string text);
        public void AppendText(string text)
        {
            if (this.textBox.InvokeRequired)
            {
                AppendTextCallback d = new AppendTextCallback(AppendText);
                this.Invoke(d, new object[] { text });
            }
            else
            {
                this.textBox.AppendText(text + Environment.NewLine);
            }
        }

        public delegate void SetDataCallback(DataTable table);
        public void SetData(DataTable table)
        {
            if (this.dataGridView.InvokeRequired)
            {
                SetDataCallback d = new SetDataCallback(SetData);
                this.Invoke(d, new object[] { table });
            }
            else
            {
                this.dataGridView.DataSource = table;
                dataGridView.RowsDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                for (int i = 0; i < this.dataGridView.Columns.Count; i++)
                {
                    dataGridView.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                }
            }
        }

        /// <summary>
        /// 从Configuration.xml中加载县级分等单元图层字段定义
        /// </summary>
        /// <returns>字段定义集合</returns>
        private List<Shapelib.DBFField> LoadXJFDDYFields()
        {
            List<Shapelib.DBFField> DBFFields = new List<Shapelib.DBFField>();

            var _ConfigurationFileName = AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "Configuration.xml";
            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.Load(_ConfigurationFileName);
            XmlElement root = xmlDocument.DocumentElement;
            XmlNode XJFDDYNode = root.SelectSingleNode("./XJFDDY");

            foreach (XmlNode fieldNode in XJFDDYNode.ChildNodes)
            {
                var name = fieldNode.Attributes["Name"].InnerText.Trim();
                var type = fieldNode.Attributes["Type"].InnerText.Trim();
                Shapelib.DBFFieldType fieldType = Shapelib.DBFFieldType.FTInvalid;
                switch (type)
                {
                    case "FTString":
                        fieldType = Shapelib.DBFFieldType.FTString;
                        break;
                    case "FTDouble":
                        fieldType = Shapelib.DBFFieldType.FTDouble;
                        break;
                    case "FTInteger":
                        fieldType = Shapelib.DBFFieldType.FTInteger;
                        break;
                    case "FTLogical":
                        fieldType = Shapelib.DBFFieldType.FTLogical;
                        break;
                    case "FTDate":
                        fieldType = Shapelib.DBFFieldType.FTDate;
                        break;
                }
                var width = Convert.ToInt32(fieldNode.Attributes["Width"].InnerText.Trim());
                var decimals = Convert.ToInt32(fieldNode.Attributes["Decimals"].InnerText.Trim());
                var description = fieldNode.Attributes["Description"].InnerText.Trim();

                var field = new Shapelib.DBFField(name, fieldType, width, decimals, description);

                DBFFields.Add(field);
            }

            return DBFFields;

        }

        private DataTable GetLPolygonsDataTable()
        {
            var dataTable = new DataTable();

            dataTable.Columns.Add(new DataColumn("编号"));
            dataTable.Columns.Add(new DataColumn("地类"));
            dataTable.Columns.Add(new DataColumn("有效土层厚度"));
            dataTable.Columns.Add(new DataColumn("表层土壤质地"));
            dataTable.Columns.Add(new DataColumn("剖面构型"));
            dataTable.Columns.Add(new DataColumn("土壤有机质含量"));
            dataTable.Columns.Add(new DataColumn("土壤PH值"));
            dataTable.Columns.Add(new DataColumn("障碍层距地表深度"));
            dataTable.Columns.Add(new DataColumn("排水条件"));
            dataTable.Columns.Add(new DataColumn("地形坡度"));
            dataTable.Columns.Add(new DataColumn("灌溉保证率"));
            dataTable.Columns.Add(new DataColumn("地表岩石出露度"));

            foreach (var patch in _patches)
            {
                var dataRow = dataTable.NewRow();
                dataRow["编号"] = patch.Name;
                dataRow["地类"] = patch.LandClass;
                dataRow["有效土层厚度"] = GetFactorValue(patch, "有效土层厚度");
                dataRow["表层土壤质地"] = GetFactorValue(patch, "表层土壤质地");
                dataRow["剖面构型"] = GetFactorValue(patch, "剖面构型");
                dataRow["土壤有机质含量"] = GetFactorValue(patch, "土壤有机质含量");
                dataRow["土壤PH值"] = GetFactorValue(patch, "土壤PH值");
                dataRow["障碍层距地表深度"] = GetFactorValue(patch, "障碍层距地表深度");
                dataRow["排水条件"] = GetFactorValue(patch, "排水条件");
                dataRow["地形坡度"] = GetFactorValue(patch, "地形坡度");
                dataRow["灌溉保证率"] = GetFactorValue(patch, "灌溉保证率");
                dataRow["地表岩石出露度"] = GetFactorValue(patch, "地表岩石出露度");
                dataTable.Rows.Add(dataRow);
            }

            return dataTable;
        }

        private string GetFactorValue(Patch patch,string factorName)
        {
            var factorValue = string.Empty;
            foreach(var factor in patch.Factors)
            {
                if(factor.Name == factorName)
                {
                    factorValue = factor.Value;
                    break;
                }
            }

            return factorValue;
        }

        private void ClearDataGridView()
        {
            dataGridView.DataSource = null;
        }

        private void GetPatchs(string shpFileName)
        {
            _patches = new List<Patch>();

            var pathName = Path.GetDirectoryName(shpFileName);
            var fileName = Path.GetFileNameWithoutExtension(shpFileName);

            var dbf = string.Format(@"{0}\{1}.dbf", pathName, fileName);
            var hDBF = Shapelib.DBFOpen(dbf, "rb");

            var recordCount = Shapelib.DBFGetRecordCount(hDBF);

            for (int iShape = 0; iShape < recordCount; iShape++)
            {
                var patch = new Patch();

                // 图斑编号/名称
                patch.Name = Shapelib.DBFReadAnsiStringAttribute(hDBF, iShape, Shapelib.DBFGetFieldIndex(hDBF, "XJDYBH"));

                // 指标区
                patch.ThirdIndexRegion = Shapelib.DBFReadAnsiStringAttribute(hDBF, iShape, Shapelib.DBFGetFieldIndex(hDBF, "SJQMC"));

                // 县
                patch.County = Shapelib.DBFReadAnsiStringAttribute(hDBF, iShape, Shapelib.DBFGetFieldIndex(hDBF, "X"));

                // 地类
                patch.ClassCode = Shapelib.DBFReadStringAttribute(hDBF, iShape, Shapelib.DBFGetFieldIndex(hDBF, "DLDM"));

                // 面积
                patch.Area = Shapelib.DBFReadDoubleAttribute(hDBF, iShape, Shapelib.DBFGetFieldIndex(hDBF, "ZRDMJ"));

                // 土地利用系数
                patch.UtilizationCoefficient = Shapelib.DBFReadDoubleAttribute(hDBF, iShape, Shapelib.DBFGetFieldIndex(hDBF, "TDLYXS"));

                // 土地经济系数
                patch.EconomicalCoefficient = Shapelib.DBFReadDoubleAttribute(hDBF, iShape, Shapelib.DBFGetFieldIndex(hDBF, "TDJJXS"));

                // 是否为新增耕地
                var iField = Shapelib.DBFGetFieldIndex(hDBF, "SFXZGD");
                if (iField >= 0) { patch.IsNew = Convert.ToBoolean(Shapelib.DBFReadIntegerAttribute(hDBF, iShape, iField)); }

                patch.IsCPPC = IndexTypeMenuItem.Checked;

                if (patch.LandClass == "水田")
                {
                    patch.Crops[0].GrainOutput = Shapelib.DBFReadDoubleAttribute(hDBF, iShape, Shapelib.DBFGetFieldIndex(hDBF, "JZZWCL"));
                    patch.Crops[0].Cost = Shapelib.DBFReadDoubleAttribute(hDBF, iShape, Shapelib.DBFGetFieldIndex(hDBF, "JZZWCB"));

                    patch.Crops[1].GrainOutput = Shapelib.DBFReadDoubleAttribute(hDBF, iShape, Shapelib.DBFGetFieldIndex(hDBF, "ZDZW1CL"));
                    patch.Crops[1].Cost = Shapelib.DBFReadDoubleAttribute(hDBF, iShape, Shapelib.DBFGetFieldIndex(hDBF, "ZDZW1CB"));
                }
                else
                {
                    patch.Crops[0].GrainOutput = Shapelib.DBFReadDoubleAttribute(hDBF, iShape, Shapelib.DBFGetFieldIndex(hDBF, "ZDZW1CL"));
                    patch.Crops[0].Cost = Shapelib.DBFReadDoubleAttribute(hDBF, iShape, Shapelib.DBFGetFieldIndex(hDBF, "ZDZW1CB"));

                    patch.Crops[1].GrainOutput = Shapelib.DBFReadDoubleAttribute(hDBF, iShape, Shapelib.DBFGetFieldIndex(hDBF, "ZDZW2CL"));
                    patch.Crops[1].Cost = Shapelib.DBFReadDoubleAttribute(hDBF, iShape, Shapelib.DBFGetFieldIndex(hDBF, "ZDZW2CB"));
                }

                // 分等因素
                for (int i = 0; i < 7; i++)
                {
                    switch(patch.Factors[i].Name)
                    {
                        case "有效土层厚度":
                            patch.Factors[i].Value = Shapelib.DBFReadIntegerAttribute(hDBF, iShape, Shapelib.DBFGetFieldIndex(hDBF, "YXTCHD")).ToString();
                            break;
                        case "表层土壤质地":
                            patch.Factors[i].Value = Shapelib.DBFReadAnsiStringAttribute(hDBF, iShape, Shapelib.DBFGetFieldIndex(hDBF, "BCTRZD"));
                            break;
                        case "剖面构型":
                            patch.Factors[i].Value = Shapelib.DBFReadAnsiStringAttribute(hDBF, iShape, Shapelib.DBFGetFieldIndex(hDBF, "PMGX"));
                            break;
                        case "土壤有机质含量":
                            patch.Factors[i].Value = Shapelib.DBFReadIntegerAttribute(hDBF, iShape, Shapelib.DBFGetFieldIndex(hDBF, "TRYJZHL")).ToString();
                            break;
                        case "土壤PH值":
                            patch.Factors[i].Value = Shapelib.DBFReadIntegerAttribute(hDBF, iShape, Shapelib.DBFGetFieldIndex(hDBF, "TRSJD")).ToString();
                            break;
                        case "障碍层距地表深度":
                            patch.Factors[i].Value = Shapelib.DBFReadDoubleAttribute(hDBF, iShape, Shapelib.DBFGetFieldIndex(hDBF, "ZACJDBSD")).ToString();
                            break;
                        case "排水条件":
                            patch.Factors[i].Value = Shapelib.DBFReadIntegerAttribute(hDBF, iShape, Shapelib.DBFGetFieldIndex(hDBF, "PSTJ")).ToString();
                            break;
                        case "地形坡度":
                            patch.Factors[i].Value = Shapelib.DBFReadDoubleAttribute(hDBF, iShape, Shapelib.DBFGetFieldIndex(hDBF, "DXPD")).ToString();
                            break;
                        case "灌溉保证率":
                            patch.Factors[i].Value = Shapelib.DBFReadAnsiStringAttribute(hDBF, iShape, Shapelib.DBFGetFieldIndex(hDBF, "GGBZL"));
                            break;
                        case "地表岩石出露度":
                            patch.Factors[i].Value = Shapelib.DBFReadIntegerAttribute(hDBF, iShape, Shapelib.DBFGetFieldIndex(hDBF, "DBYSLTD")).ToString();
                            break;
                    }
                }

                // 整治前评价指标
                for (int i = 0; i < 5; i++)
                {
                    switch (patch.BeforeUtilizationFactors[i].Type)
                    {
                        case UtilizationFactorType.Water:
                            patch.BeforeUtilizationFactors[i].Value = Shapelib.DBFReadIntegerAttribute(hDBF, iShape, Shapelib.DBFGetFieldIndex(hDBF, "SYTJ1"));
                            break;
                        case UtilizationFactorType.Irrigation:
                            patch.BeforeUtilizationFactors[i].Value = Shapelib.DBFReadIntegerAttribute(hDBF, iShape, Shapelib.DBFGetFieldIndex(hDBF, "GGTJ1"));
                            break;
                        case UtilizationFactorType.Drainage:
                            patch.BeforeUtilizationFactors[i].Value = Shapelib.DBFReadIntegerAttribute(hDBF, iShape, Shapelib.DBFGetFieldIndex(hDBF, "PSTJ1"));
                            break;
                        case UtilizationFactorType.Road:
                            patch.BeforeUtilizationFactors[i].Value = Shapelib.DBFReadIntegerAttribute(hDBF, iShape, Shapelib.DBFGetFieldIndex(hDBF, "DLTDD1"));
                            break;
                        case UtilizationFactorType.Flatness:
                            patch.BeforeUtilizationFactors[i].Value = Shapelib.DBFReadIntegerAttribute(hDBF, iShape, Shapelib.DBFGetFieldIndex(hDBF, "TKPZD1"));
                            break;
                    }
                }

                // 整治后评价指标
                for (int i = 0; i < 5; i++)
                {
                    switch (patch.AfterUtilizationFactors[i].Type)
                    {
                        case UtilizationFactorType.Water:
                            patch.AfterUtilizationFactors[i].Value = Shapelib.DBFReadIntegerAttribute(hDBF, iShape, Shapelib.DBFGetFieldIndex(hDBF, "SYTJ2"));
                            break;
                        case UtilizationFactorType.Irrigation:
                            patch.AfterUtilizationFactors[i].Value = Shapelib.DBFReadIntegerAttribute(hDBF, iShape, Shapelib.DBFGetFieldIndex(hDBF, "GGTJ2"));
                            break;
                        case UtilizationFactorType.Drainage:
                            patch.AfterUtilizationFactors[i].Value = Shapelib.DBFReadIntegerAttribute(hDBF, iShape, Shapelib.DBFGetFieldIndex(hDBF, "PSTJ2"));
                            break;
                        case UtilizationFactorType.Road:
                            patch.AfterUtilizationFactors[i].Value = Shapelib.DBFReadIntegerAttribute(hDBF, iShape, Shapelib.DBFGetFieldIndex(hDBF, "DLTDD2"));
                            break;
                        case UtilizationFactorType.Flatness:
                            patch.AfterUtilizationFactors[i].Value = Shapelib.DBFReadIntegerAttribute(hDBF, iShape, Shapelib.DBFGetFieldIndex(hDBF, "TKPZD2"));
                            break;
                    }
                }

                _patches.Add(patch);
            }

            Shapelib.DBFClose(hDBF);
        }

        private void OpenShpMenuItem_Click(object sender, EventArgs e)
        {
            var dialog = new OpenFileDialog();
            dialog.Filter = "Shp文件(*.shp)|*.shp";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                if (!OpenShpBGWorker.IsBusy)
                {
                    _SHPFileName = dialog.FileName;
                    OpenShpBGWorker.RunWorkerAsync(_SHPFileName);
                }
                else
                    MessageBox.Show(string.Format("正在打开{0}，请稍后再试。", dialog.FileName), "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CloseShpMenuItem_Click(object sender, EventArgs e)
        {
            ClearDataGridView();
            _patches.Clear();
        }

        private void SaveXlsMenuItem_Click(object sender, EventArgs e)
        {
            if (_patches.Count <= 0) { return; }

            var dialog = new SaveFileDialog();
            dialog.Filter = "Excel文件(*.xls)|*.xls";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                if (!SaveXlsBGWorker.IsBusy)
                {
                    SaveXlsBGWorker.RunWorkerAsync(dialog.FileName);
                }
                else
                    MessageBox.Show("正在生成上一次报表，请稍后再试。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OpenShpBGWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            var shpFileName = e.Argument.ToString();
            GetPatchs(shpFileName);
            SetData(GetLPolygonsDataTable());
        }

        private void SaveXlsBGWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            var xlsFileName = e.Argument.ToString();
            var EvaluateTemplateHandler = new EvaluateTemplateHandler(AppDomain.CurrentDomain.SetupInformation.ApplicationBase + @"Templates\Evaluate.xlt");
            EvaluateTemplateHandler.ToSheet(xlsFileName, _patches);
            e.Result = xlsFileName;
        }

        private void SaveXlsBGWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            MessageBox.Show(string.Format("成功生成{0}。",e.Result), "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void SetAttributeMenuItem_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_SHPFileName)) { return; }

            var form = new ResearchForm();
            if(form.ShowDialog(this)== DialogResult.OK)
            {
                var pathName = Path.GetDirectoryName(_SHPFileName);
                var fileName = Path.GetFileNameWithoutExtension(_SHPFileName);

                var dbf = string.Format(@"{0}\{1}.dbf", pathName, fileName);
                var hDBF = Shapelib.DBFOpen(dbf, "rb+");

                var recordCount = Shapelib.DBFGetRecordCount(hDBF);
                for(int i=0;i< recordCount;i++)
                {
                    var iField = Shapelib.DBFGetFieldIndex(hDBF, "JZZWCL");
                    if (iField >= 0) { Shapelib.DBFWriteDoubleAttribute(hDBF, i, iField, form.JZZWCL); }
                    iField = Shapelib.DBFGetFieldIndex(hDBF, "JZZWCB");
                    if (iField >= 0) { Shapelib.DBFWriteDoubleAttribute(hDBF, i, iField, form.JZZWCB); }
                    iField = Shapelib.DBFGetFieldIndex(hDBF, "ZDZW1CL");
                    if (iField >= 0) { Shapelib.DBFWriteDoubleAttribute(hDBF, i, iField, form.ZDZW1CL); }
                    iField = Shapelib.DBFGetFieldIndex(hDBF, "ZDZW1CB");
                    if (iField >= 0) { Shapelib.DBFWriteDoubleAttribute(hDBF, i, iField, form.ZDZW1CB); }
                    iField = Shapelib.DBFGetFieldIndex(hDBF, "ZDZW2CL");
                    if (iField >= 0) { Shapelib.DBFWriteDoubleAttribute(hDBF, i, iField, form.ZDZW2CL); }
                    iField = Shapelib.DBFGetFieldIndex(hDBF, "ZDZW2CB");
                    if (iField >= 0) { Shapelib.DBFWriteDoubleAttribute(hDBF, i, iField, form.ZDZW2CB); }
                }

                Shapelib.DBFClose(hDBF);

                MessageBox.Show("成功添加 产量/成本 字段。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void AddCustomFieldsMenuItem_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_SHPFileName)) { return; }

            var pathName = Path.GetDirectoryName(_SHPFileName);
            var fileName = Path.GetFileNameWithoutExtension(_SHPFileName);

            var dbf = string.Format(@"{0}\{1}.dbf", pathName, fileName);
            var hDBF = Shapelib.DBFOpen(dbf, "rb+");

            if (Shapelib.DBFGetFieldIndex(hDBF, "JZZWCL") < 0)
                Shapelib.DBFAddField(hDBF, "JZZWCL", Shapelib.DBFFieldType.FTDouble, 7, 2);
            if (Shapelib.DBFGetFieldIndex(hDBF, "JZZWCB") < 0)
                Shapelib.DBFAddField(hDBF, "JZZWCB", Shapelib.DBFFieldType.FTDouble, 7, 2);
            if (Shapelib.DBFGetFieldIndex(hDBF, "ZDZW1CL") < 0)
                Shapelib.DBFAddField(hDBF, "ZDZW1CL", Shapelib.DBFFieldType.FTDouble, 7, 2);
            if (Shapelib.DBFGetFieldIndex(hDBF, "ZDZW1CB") < 0)
                Shapelib.DBFAddField(hDBF, "ZDZW1CB", Shapelib.DBFFieldType.FTDouble, 7, 2);
            if (Shapelib.DBFGetFieldIndex(hDBF, "ZDZW2CL") < 0)
                Shapelib.DBFAddField(hDBF, "ZDZW2CL", Shapelib.DBFFieldType.FTDouble, 7, 2);
            if (Shapelib.DBFGetFieldIndex(hDBF, "ZDZW2CB") < 0)
                Shapelib.DBFAddField(hDBF, "ZDZW2CB", Shapelib.DBFFieldType.FTDouble, 7, 2);

            if (Shapelib.DBFGetFieldIndex(hDBF, "SYTJ1") < 0)
                Shapelib.DBFAddField(hDBF, "SYTJ1", Shapelib.DBFFieldType.FTInteger, 1, 0);
            if (Shapelib.DBFGetFieldIndex(hDBF, "GGTJ1") < 0)
                Shapelib.DBFAddField(hDBF, "GGTJ1", Shapelib.DBFFieldType.FTInteger, 1, 0);
            if (Shapelib.DBFGetFieldIndex(hDBF, "PSTJ1") < 0)
                Shapelib.DBFAddField(hDBF, "PSTJ1", Shapelib.DBFFieldType.FTInteger, 1, 0);
            if (Shapelib.DBFGetFieldIndex(hDBF, "DLTDD1") < 0)
                Shapelib.DBFAddField(hDBF, "DLTDD1", Shapelib.DBFFieldType.FTInteger, 1, 0);
            if (Shapelib.DBFGetFieldIndex(hDBF, "TKPZD1") < 0)
                Shapelib.DBFAddField(hDBF, "TKPZD1", Shapelib.DBFFieldType.FTInteger, 1, 0);

            if (Shapelib.DBFGetFieldIndex(hDBF, "SYTJ2") < 0)
                Shapelib.DBFAddField(hDBF, "SYTJ2", Shapelib.DBFFieldType.FTInteger, 1, 0);
            if (Shapelib.DBFGetFieldIndex(hDBF, "GGTJ2") < 0)
                Shapelib.DBFAddField(hDBF, "GGTJ2", Shapelib.DBFFieldType.FTInteger, 1, 0);
            if (Shapelib.DBFGetFieldIndex(hDBF, "PSTJ2") < 0)
                Shapelib.DBFAddField(hDBF, "PSTJ2", Shapelib.DBFFieldType.FTInteger, 1, 0);
            if (Shapelib.DBFGetFieldIndex(hDBF, "DLTDD2") < 0)
                Shapelib.DBFAddField(hDBF, "DLTDD2", Shapelib.DBFFieldType.FTInteger, 1, 0);
            if (Shapelib.DBFGetFieldIndex(hDBF, "TKPZD2") < 0)
                Shapelib.DBFAddField(hDBF, "TKPZD2", Shapelib.DBFFieldType.FTInteger, 1, 0);

            if (Shapelib.DBFGetFieldIndex(hDBF, "SFXZGD") < 0)
                Shapelib.DBFAddField(hDBF, "SFXZGD", Shapelib.DBFFieldType.FTInteger, 1, 0);

            Shapelib.DBFClose(hDBF);

            MessageBox.Show("成功添加 评定辅助 字段。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void SaveLayerAsMenuItem_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_SHPFileName)) { return; }

            if(MessageBox.Show("是否重新计算当前图斑的质量等别？","提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question)==DialogResult.Yes)
            {
                var dbf = string.Format(@"{0}\{1}.dbf", Path.GetDirectoryName(_SHPFileName), Path.GetFileNameWithoutExtension(_SHPFileName));
                var hDBF = Shapelib.DBFOpen(dbf, "rb+");
                var hSHP = Shapelib.SHPOpen(_SHPFileName, "rb");

                var recordCount = Shapelib.DBFGetRecordCount(hDBF);

                var formula = "";
                for (int i = 0; i < recordCount; i++)
                {
                    var iField = Shapelib.DBFGetFieldIndex(hDBF, "XJDYBH");
                    if (iField >= 0) { Shapelib.DBFWriteStringAttribute(hDBF, i, iField, _patches[i].Name); }
                    if (_patches[i].LandClass == "水田")
                    {
                        iField = Shapelib.DBFGetFieldIndex(hDBF, "JZZWMC");
                        if (iField >= 0) { Shapelib.DBFWriteStringAttribute(hDBF, i, iField, _patches[i].Crops[0].Name); }
                        iField = Shapelib.DBFGetFieldIndex(hDBF, "JZGWZS");
                        if (iField >= 0) { Shapelib.DBFWriteIntegerAttribute(hDBF, i, iField, (int)SQLiteHelper.GetCropPTP(_patches[i].County, _patches[i].Crops[0].Name)); }
                        iField = Shapelib.DBFGetFieldIndex(hDBF, "JZQHZS");
                        if (iField >= 0) { Shapelib.DBFWriteIntegerAttribute(hDBF, i, iField, 0); }
                        iField = Shapelib.DBFGetFieldIndex(hDBF, "JZZWF");
                        if (iField >= 0) { Shapelib.DBFWriteDoubleAttribute(hDBF, i, iField, SQLiteHelper.CalculateCropNaturalQualityScore(_patches[i], _patches[i].Crops[0].Name, out formula)); }

                        iField = Shapelib.DBFGetFieldIndex(hDBF, "ZDZW1MC");
                        if (iField >= 0) { Shapelib.DBFWriteStringAttribute(hDBF, i, iField, _patches[i].Crops[1].Name); }
                        iField = Shapelib.DBFGetFieldIndex(hDBF, "ZDZW1GW");
                        if (iField >= 0) { Shapelib.DBFWriteIntegerAttribute(hDBF, i, iField, (int)SQLiteHelper.GetCropPTP(_patches[i].County, _patches[i].Crops[1].Name)); }
                        iField = Shapelib.DBFGetFieldIndex(hDBF, "ZDZW1QH");
                        if (iField >= 0) { Shapelib.DBFWriteIntegerAttribute(hDBF, i, iField, 0); }
                        iField = Shapelib.DBFGetFieldIndex(hDBF, "ZDZW1F");
                        if (iField >= 0) { Shapelib.DBFWriteDoubleAttribute(hDBF, i, iField, SQLiteHelper.CalculateCropNaturalQualityScore(_patches[i], _patches[i].Crops[1].Name, out formula)); }
                        iField = Shapelib.DBFGetFieldIndex(hDBF, "ZDZW1B");
                        if (iField >= 0) { Shapelib.DBFWriteDoubleAttribute(hDBF, i, iField, _patches[i].Crops[1].GOCCoefficient); }
                    }
                    else
                    {
                        iField = Shapelib.DBFGetFieldIndex(hDBF, "ZDZW1MC");
                        if (iField >= 0) { Shapelib.DBFWriteStringAttribute(hDBF, i, iField, _patches[i].Crops[0].Name); }
                        iField = Shapelib.DBFGetFieldIndex(hDBF, "ZDZW1GW");
                        if (iField >= 0) { Shapelib.DBFWriteIntegerAttribute(hDBF, i, iField, 0); }
                        iField = Shapelib.DBFGetFieldIndex(hDBF, "ZDZW1QH");
                        if (iField >= 0) { Shapelib.DBFWriteIntegerAttribute(hDBF, i, iField, (int)SQLiteHelper.GetCropCPPC(_patches[i].County, _patches[i].Crops[0].Name)); }
                        iField = Shapelib.DBFGetFieldIndex(hDBF, "ZDZW1F");
                        if (iField >= 0) { Shapelib.DBFWriteDoubleAttribute(hDBF, i, iField, SQLiteHelper.CalculateCropNaturalQualityScore(_patches[i], _patches[i].Crops[0].Name, out formula)); }
                        iField = Shapelib.DBFGetFieldIndex(hDBF, "ZDZW1B");
                        if (iField >= 0) { Shapelib.DBFWriteDoubleAttribute(hDBF, i, iField, _patches[i].Crops[0].GOCCoefficient); }

                        iField = Shapelib.DBFGetFieldIndex(hDBF, "ZDZW2MC");
                        if (iField >= 0) { Shapelib.DBFWriteStringAttribute(hDBF, i, iField, _patches[i].Crops[1].Name); }
                        iField = Shapelib.DBFGetFieldIndex(hDBF, "ZDZW2GW");
                        if (iField >= 0) { Shapelib.DBFWriteIntegerAttribute(hDBF, i, iField, 0); }
                        iField = Shapelib.DBFGetFieldIndex(hDBF, "ZDZW2QH");
                        if (iField >= 0) { Shapelib.DBFWriteIntegerAttribute(hDBF, i, iField, (int)SQLiteHelper.GetCropCPPC(_patches[i].County, _patches[i].Crops[1].Name)); }
                        iField = Shapelib.DBFGetFieldIndex(hDBF, "ZDZW2F");
                        if (iField >= 0) { Shapelib.DBFWriteDoubleAttribute(hDBF, i, iField, SQLiteHelper.CalculateCropNaturalQualityScore(_patches[i], _patches[i].Crops[1].Name, out formula)); }
                        iField = Shapelib.DBFGetFieldIndex(hDBF, "ZDZW2B");
                        if (iField >= 0) { Shapelib.DBFWriteDoubleAttribute(hDBF, i, iField, _patches[i].Crops[1].GOCCoefficient); }
                    }

                    // 计算
                    var naturalQualityScore = SQLiteHelper.CalculateNaturalQualityScore(_patches[i]);
                    var naturalQualityGrade = SQLiteHelper.CalculateNaturalQualityGrade(_patches[i].NaturalQualityGradeIndex);

                    SQLiteHelper.CalculateUtilizationScore(_patches[i], out double score1, out string formula1, out double score2, out string formula2);
                    double utilizationCoefficient = _patches[i].UtilizationCoefficient;
                    if (!_patches[i].IsNew)
                    {
                        utilizationCoefficient = SQLiteHelper.CalculateUtilizationCoefficient(_patches[i].UtilizationCoefficient, score1, score2);
                    }
                    var utilizationGrade = SQLiteHelper.CalculateUtilizationGrade(_patches[i].UtilizationGradeIndex);

                    var economicalGrade = SQLiteHelper.CalculateEconomicalGrade(_patches[i].EconomicalGradeIndex);

                    iField = Shapelib.DBFGetFieldIndex(hDBF, "ZHZLF");
                    if (iField >= 0) { Shapelib.DBFWriteDoubleAttribute(hDBF, i, iField, naturalQualityScore); }
                    iField = Shapelib.DBFGetFieldIndex(hDBF, "ZRDZS");
                    if (iField >= 0) { Shapelib.DBFWriteIntegerAttribute(hDBF, i, iField, (int)_patches[i].NaturalQualityGradeIndex); }
                    iField = Shapelib.DBFGetFieldIndex(hDBF, "ZRDB");
                    if (iField >= 0) { Shapelib.DBFWriteIntegerAttribute(hDBF, i, iField, naturalQualityGrade); }
                    iField = Shapelib.DBFGetFieldIndex(hDBF, "TDLYXS");
                    if (iField >= 0) { Shapelib.DBFWriteDoubleAttribute(hDBF, i, iField, utilizationCoefficient); }
                    iField = Shapelib.DBFGetFieldIndex(hDBF, "LYDZS");
                    if (iField >= 0) { Shapelib.DBFWriteIntegerAttribute(hDBF, i, iField, (int)_patches[i].UtilizationGradeIndex); }
                    iField = Shapelib.DBFGetFieldIndex(hDBF, "LYD");
                    if (iField >= 0) { Shapelib.DBFWriteIntegerAttribute(hDBF, i, iField, utilizationGrade); }
                    iField = Shapelib.DBFGetFieldIndex(hDBF, "TDJJXS");
                    if (iField >= 0) { Shapelib.DBFWriteDoubleAttribute(hDBF, i, iField, _patches[i].EconomicalCoefficient); }
                    iField = Shapelib.DBFGetFieldIndex(hDBF, "DBZ");
                    if (iField >= 0) { Shapelib.DBFWriteIntegerAttribute(hDBF, i, iField, (int)_patches[i].EconomicalGradeIndex); }
                    iField = Shapelib.DBFGetFieldIndex(hDBF, "DB");
                    if (iField >= 0) { Shapelib.DBFWriteIntegerAttribute(hDBF, i, iField, economicalGrade); }
                    iField = Shapelib.DBFGetFieldIndex(hDBF, "GJZRDZS");
                    if (iField >= 0) { Shapelib.DBFWriteIntegerAttribute(hDBF, i, iField, (int)_patches[i].StateNaturalQualityGradeIndex); }
                    iField = Shapelib.DBFGetFieldIndex(hDBF, "GJZRDB");
                    if (iField >= 0) { Shapelib.DBFWriteIntegerAttribute(hDBF, i, iField, _patches[i].StateNaturalQualityGrade); }
                    iField = Shapelib.DBFGetFieldIndex(hDBF, "GJLYDZS");
                    if (iField >= 0) { Shapelib.DBFWriteIntegerAttribute(hDBF, i, iField, (int)_patches[i].StateUtilizationGradeIndex); }
                    iField = Shapelib.DBFGetFieldIndex(hDBF, "GJLYDB");
                    if (iField >= 0) { Shapelib.DBFWriteIntegerAttribute(hDBF, i, iField, _patches[i].StateUtilizationGrade); }
                    iField = Shapelib.DBFGetFieldIndex(hDBF, "GJDBZS");
                    if (iField >= 0) { Shapelib.DBFWriteIntegerAttribute(hDBF, i, iField, (int)_patches[i].StateEconomicalGradeIndex); }
                    iField = Shapelib.DBFGetFieldIndex(hDBF, "GJDB");
                    if (iField >= 0) { Shapelib.DBFWriteIntegerAttribute(hDBF, i, iField, _patches[i].StateEconomicalGrade); }

                    // 计算图斑面积
                    // iField = Shapelib.DBFGetFieldIndex(hDBF, "ZRDMJ");
                    // if (iField >= 0) { Shapelib.DBFWriteDoubleAttribute(hDBF, i, iField, Shapelib.SHPCalculateArea(hSHP, i)/10000); }
                }

                Shapelib.DBFClose(hDBF);
                Shapelib.SHPClose(hSHP);
            }

            var dialog = new SaveFileDialog();
            dialog.Filter = "Shp文件(*.shp)|*.shp";
            dialog.AddExtension = true;
            if (dialog.ShowDialog()== DialogResult.OK)
            {
                // 当前SHP文件句柄
                var hShp = Shapelib.SHPOpen(_SHPFileName, "rb");
                // 当前DBF文件句柄
                var dbf = string.Format(@"{0}\{1}.dbf", Path.GetDirectoryName(_SHPFileName), Path.GetFileNameWithoutExtension(_SHPFileName));
                var hDbf = Shapelib.DBFOpen(dbf, "rb");
                // 实体数
                int pnEntities = 0;
                // 形状类型
                Shapelib.ShapeType pshpType = Shapelib.ShapeType.NullShape;
                // 界限坐标数组
                double[] adfMinBound = new double[4], adfMaxBound = new double[4];
                // 获取实体数、形状类型、界限坐标等信息
                Shapelib.SHPGetInfo(hShp, ref pnEntities, ref pshpType, adfMinBound, adfMaxBound);

                // 分等单元SHP文件名
                var shpFileName = dialog.FileName;
                // 分等单元DBF文件名
                var dbfFileName = shpFileName.Replace(".shp", ".dbf");
                // 创建分等单元SHP文件
                var hNewShp = Shapelib.SHPCreate(shpFileName, pshpType);
                // 创建分等单元DBF文件
                var hNewDbf = Shapelib.DBFCreate(dbfFileName);
                // 创建字段
                var fields = LoadXJFDDYFields();
                foreach(var field in fields)
                    Shapelib.DBFAddField(hNewDbf, field.Name, field.Type, field.Width, field.Decimals);
                
                // 复制图元到分等单元
                for (int iShape = 0; iShape < pnEntities; iShape++)
                {
                    // SHPObject对象
                    Shapelib.SHPObject shpObject = new Shapelib.SHPObject();
                    // 读取SHPObject对象指针
                    var shpObjectPtr = Shapelib.SHPReadObject(hShp, iShape);
                    // 忽略可能存在问题的实体
                    if (shpObjectPtr == IntPtr.Zero) { continue; }
                    // 指针转换为SHPObject对象
                    Marshal.PtrToStructure(shpObjectPtr, shpObject);
                    // 写入SHPObject到分等单元
                    Shapelib.SHPWriteObject(hNewShp, -1, shpObjectPtr);
                }

                // 复制属性到分等单元
                for (int iShape = 0; iShape < pnEntities; iShape++)
                {
                    for (int iField = 0; iField < fields.Count; iField++)
                    {
                        var field = fields[iField];
                        switch(field.Type)
                        {
                            case Shapelib.DBFFieldType.FTDate:
                                var dateAttribute = Shapelib.DBFReadDateAttribute(hDbf, iShape, Shapelib.DBFGetFieldIndex(hDbf, field.Name));
                                Shapelib.DBFWriteDateAttribute(hNewDbf, iShape, iField, dateAttribute);
                                break;
                            case Shapelib.DBFFieldType.FTDouble:
                                var doubleAttribute = Shapelib.DBFReadDoubleAttribute(hDbf, iShape, Shapelib.DBFGetFieldIndex(hDbf, field.Name));
                                Shapelib.DBFWriteDoubleAttribute(hNewDbf, iShape, iField, doubleAttribute);
                                break;
                            case Shapelib.DBFFieldType.FTInteger:
                                var integerAttribute = Shapelib.DBFReadIntegerAttribute(hDbf, iShape, Shapelib.DBFGetFieldIndex(hDbf, field.Name));
                                Shapelib.DBFWriteIntegerAttribute(hNewDbf, iShape, iField, integerAttribute);
                                break;
                            case Shapelib.DBFFieldType.FTLogical:
                                var logicalAttribute = Shapelib.DBFReadLogicalAttribute(hDbf, iShape, Shapelib.DBFGetFieldIndex(hDbf, field.Name));
                                Shapelib.DBFWriteLogicalAttribute(hNewDbf, iShape, iField, logicalAttribute);
                                break;
                            case Shapelib.DBFFieldType.FTString:
                                var stringAttribute = Shapelib.DBFReadAnsiStringAttribute(hDbf, iShape, Shapelib.DBFGetFieldIndex(hDbf, field.Name));
                                Shapelib.DBFWriteStringAttribute(hNewDbf, iShape, iField, stringAttribute);
                                break;
                        }
                    }
                }

                // 关闭分等单元DBF文件句柄
                Shapelib.DBFClose(hNewDbf);
                // 关闭分等单元SHP文件句柄
                Shapelib.SHPClose(hNewShp);

                // 关闭DBF文件句柄
                Shapelib.DBFClose(hDbf);
                // 关闭SHP文件句柄
                Shapelib.SHPClose(hShp);

                MessageBox.Show(string.Format("成功创建{0}。", shpFileName), "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void IndexTypeMenuItem_CheckedChanged(object sender, EventArgs e)
        {
            if(_patches.Count>0)
            {
                for(int i=0;i< _patches.Count;i++)
                {
                    if(_patches[i].LandClass=="水浇地")
                    {
                        _patches[i].IsCPPC = IndexTypeMenuItem.Checked;
                    }
                }
            }
        }
    }
}
