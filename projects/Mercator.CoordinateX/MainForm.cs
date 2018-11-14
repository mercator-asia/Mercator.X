using Mercator.ShapeX;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Mercator.CoordinateX
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void OpenShpFileButton_Click(object sender, EventArgs e)
        {
            var dialog = new OpenFileDialog();
            dialog.Filter = "shp文件(*.shp)|*.shp";
            if(dialog.ShowDialog()==DialogResult.OK)
            {
                ShpFileNameTextBox.Text = dialog.FileName;
            }
        }

        private void CreateCoordinateFileButton_Click(object sender, EventArgs e)
        {
            var shpFileName = ShpFileNameTextBox.Text.Trim();
            if (!string.IsNullOrEmpty(shpFileName) && File.Exists(shpFileName))
            {
                var dialog = new SaveFileDialog();
                dialog.Filter = "报备坐标文件(*.txt)|*.txt";
                dialog.AddExtension = true;
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    var coordinateProperty = new CoordinateProperty();
                    coordinateProperty.CoordinateSystem = CoordinateSystemTextBox.Text.Trim();
                    coordinateProperty.ZoneType = Convert.ToInt16(ZoneTypeTextBox.Text.Trim());
                    coordinateProperty.ProjectionType = ProjectionTypeTextBox.Text.Trim();
                    coordinateProperty.Unit = UnitTextBox.Text.Trim();
                    coordinateProperty.Zone = Convert.ToInt16(ZoneTextBox.Text.Trim());
                    coordinateProperty.Decimals = Convert.ToDouble(DecimalsTextBox.Text.Trim());
                    coordinateProperty.Parameters = ParametersTextBox.Text.Trim();

                    var identifierFieldName = IdentifierFieldNameTextBox.Text.Trim();

                    var coordinateFileName = dialog.FileName;
                    Shapelib.SHPCreateCoordinateFile(coordinateFileName, coordinateProperty, shpFileName, identifierFieldName);

                    MessageBox.Show(string.Format("生成{0}成功。", coordinateFileName), "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            } 
            else
            {
                MessageBox.Show("不可读取的SHP文件。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OpenCoordinateFileButton_Click(object sender, EventArgs e)
        {
            var dialog = new OpenFileDialog();
            dialog.Filter = "报备坐标文件(*.txt)|*.txt";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                CoordinateFileNameTextBox.Text = dialog.FileName;
            }
        }

        private void CreateShpFileButton_Click(object sender, EventArgs e)
        {
            var coordinateFileName = CoordinateFileNameTextBox.Text.Trim();
            if (!string.IsNullOrEmpty(coordinateFileName) && File.Exists(coordinateFileName))
            {
                var dialog = new SaveFileDialog();
                dialog.Filter = "shp文件(*.shp)|*.shp";
                dialog.AddExtension = true;
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    var identifierFieldName = IdentifierFieldNameTextBox.Text.Trim();
                    var shpFileName = dialog.FileName;

                    Shapelib.SHPCreateFromCoordinateFile(coordinateFileName, shpFileName, identifierFieldName);
                    MessageBox.Show(string.Format("生成{0}成功。", shpFileName), "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }  
            }
            else
            {
                MessageBox.Show("不可读取的报备坐标文件。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
