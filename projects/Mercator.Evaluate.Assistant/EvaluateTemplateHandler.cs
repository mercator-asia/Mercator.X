using Mercator.ExcelX;
using Mercator.OfficeX;
using System.Collections;
using System.Collections.Generic;

namespace Mercator.Evaluate.Assistant
{
    public class EvaluateTemplateHandler : TemplateHandler
    {
        public EvaluateTemplateHandler(string xltFileName) : base(xltFileName)
        {

        }

        /// <summary>
        /// 耕地质量等别评定
        /// </summary>
        /// <param name="lPolygons">评价单元集合</param>
        /// <returns></returns>
        private CellData[,] Evaluate(List<Patch> patches)
        {
            var data = new CellData[patches.Count + 2, 53];

            var zarea = 0d;
            var karea = 0d;

            int i = 0;

            foreach (var patch in patches)
            {
                data[i, 0] = new CellData(patch.Name);
                data[i, 1] = new CellData(patch.LandClass);
                data[i, 2] = new CellData(patch.IsNew?"是":"否");
                data[i, 3] = new CellData(GetFactorValue(patch, "有效土层厚度"));
                data[i, 4] = new CellData(GetFactorValue(patch, "表层土壤质地"));
                data[i, 5] = new CellData(GetFactorValue(patch, "剖面构型"));
                data[i, 6] = new CellData(GetFactorValue(patch, "土壤有机质含量"));
                data[i, 7] = new CellData(GetFactorValue(patch, "土壤PH值"));
                data[i, 8] = new CellData(GetFactorValue(patch, "障碍层距地表深度"));
                data[i, 9] = new CellData(GetFactorValue(patch, "排水条件"));
                data[i, 10] = new CellData(GetFactorValue(patch, "地形坡度"));
                data[i, 11] = new CellData(GetFactorValue(patch, "灌溉保证率"));
                data[i, 12] = new CellData(GetFactorValue(patch, "地表岩石出露度"));

                var PP = 0d;
                var score = SQLiteHelper.CalculateCropNaturalQualityScore(patch, patch.Crops[0].Name,out string formula);
                var index = 0d;
                var ratio = patch.Crops[0].GOCCoefficient;

                switch (patch.LandClass)
                {
                    case "水田":
                        PP = SQLiteHelper.GetCropPTP(patch.County, patch.Crops[0].Name);
                        break;
                    case "水浇地":
                        PP = patch.IsCPPC ? SQLiteHelper.GetCropCPPC(patch.County, patch.Crops[0].Name) : SQLiteHelper.GetCropPTP(patch.County, patch.Crops[0].Name);
                        break;
                    case "旱地":
                        PP = SQLiteHelper.GetCropCPPC(patch.County, patch.Crops[0].Name);
                        break;
                }

                index = score * PP * ratio;

                data[i, 13] = new CellData(patch.Crops[0].Name);
                data[i, 14] = new CellData(formula);
                data[i, 15] = new CellData(score,"0.000");
                data[i, 16] = new CellData(PP,"0");
                data[i, 17] = new CellData(ratio, "0.0");
                data[i, 18] = new CellData(index, "0");

                score = SQLiteHelper.CalculateCropNaturalQualityScore(patch, patch.Crops[1].Name, out formula);
                ratio = patch.Crops[1].GOCCoefficient;

                switch(patch.LandClass)
                {
                    case "水田":
                        PP = SQLiteHelper.GetCropPTP(patch.County, patch.Crops[1].Name);
                        break;
                    case "水浇地":
                        PP = patch.IsCPPC ? SQLiteHelper.GetCropCPPC(patch.County, patch.Crops[1].Name) : SQLiteHelper.GetCropPTP(patch.County, patch.Crops[1].Name);
                        break;
                    case "旱地":
                        PP = SQLiteHelper.GetCropCPPC(patch.County, patch.Crops[1].Name);
                        break;
                }

                index = score * PP * ratio;

                data[i, 19] = new CellData(patch.Crops[1].Name);
                data[i, 20] = new CellData(formula);
                data[i, 21] = new CellData(score, "0.000");
                data[i, 22] = new CellData(PP, "0");
                data[i, 23] = new CellData(ratio, "0.0");
                data[i, 24] = new CellData(index, "0");

                double utilizationCoefficient = patch.UtilizationCoefficient;

                if (!patch.IsNew)
                {
                    SQLiteHelper.CalculateUtilizationScore(patch, out double score1, out string formula1, out double score2, out string formula2);

                    data[i, 25] = new CellData(patch.UtilizationCoefficient, "0.000");
                    data[i, 26] = new CellData(patch.BeforeUtilizationFactors[0].Value, "0");
                    data[i, 27] = new CellData(patch.BeforeUtilizationFactors[1].Value, "0");
                    data[i, 28] = new CellData(patch.BeforeUtilizationFactors[2].Value, "0");
                    data[i, 29] = new CellData(patch.BeforeUtilizationFactors[3].Value, "0");
                    data[i, 30] = new CellData(patch.BeforeUtilizationFactors[4].Value, "0");
                    data[i, 31] = new CellData(formula1);
                    data[i, 32] = new CellData(score1, "0.0");

                    data[i, 33] = new CellData(patch.AfterUtilizationFactors[0].Value, "0");
                    data[i, 34] = new CellData(patch.AfterUtilizationFactors[1].Value, "0");
                    data[i, 35] = new CellData(patch.AfterUtilizationFactors[2].Value, "0");
                    data[i, 36] = new CellData(patch.AfterUtilizationFactors[3].Value, "0");
                    data[i, 37] = new CellData(patch.AfterUtilizationFactors[4].Value, "0");
                    data[i, 38] = new CellData(formula2);
                    data[i, 39] = new CellData(score2, "0.0");

                    utilizationCoefficient = SQLiteHelper.CalculateUtilizationCoefficient(patch.UtilizationCoefficient, score1, score2);

                    data[i, 40] = new CellData(utilizationCoefficient, "0.000");

                    zarea += patch.Area;
                }
                else
                {
                    karea += patch.Area;
                }

                if (double.IsInfinity(utilizationCoefficient)) { continue; }

                data[i, 41] = new CellData(SQLiteHelper.GetUtilizationCoefficient(patch.County, utilizationCoefficient), "0.000");
                data[i, 42] = new CellData(SQLiteHelper.GetEconomicalCoefficient(patch.County, patch.EconomicalCoefficient), "0.000");

                data[i, 43] = new CellData(patch.NaturalQualityGradeIndex, "0");
                data[i, 44] = new CellData(patch.UtilizationGradeIndex, "0");
                data[i, 45] = new CellData(patch.EconomicalGradeIndex, "0");
                data[i, 46] = new CellData(patch.StateNaturalQualityGradeIndex, "0");
                data[i, 47] = new CellData(patch.StateUtilizationGradeIndex, "0");
                data[i, 48] = new CellData(patch.StateEconomicalGradeIndex, "0");
                data[i, 49] = new CellData(patch.StateNaturalQualityGrade, "0");
                data[i, 50] = new CellData(patch.StateUtilizationGrade, "0");
                data[i, 51] = new CellData(patch.StateEconomicalGrade, "0");

                data[i, 52] = new CellData(patch.Area, "0.0000");

                var cellFontBlue = new CellFont()
                {
                    Color = CellFont.GetColorIndex(FontColor.Blue),
                    IsBold = true
                };
                data[i, 49] = new CellData(patch.StateNaturalQualityGrade, "0", cellFontBlue);

                var cellFontRed = new CellFont()
                {
                    Color = CellFont.GetColorIndex(FontColor.Red),
                    IsBold = true
                };
                data[i, 50] = new CellData(patch.StateUtilizationGrade, "0", cellFontRed);

                var cellFontGreen = new CellFont()
                {
                    Color = CellFont.GetColorIndex(FontColor.Green),
                    IsBold = true
                };
                data[i, 51] = new CellData(patch.StateEconomicalGrade, "0", cellFontGreen);

                i++;
            }


            double zrd = 0d, lyd = 0d, jjd = 0d;
            double zrd1 = 0d, lyd1 = 0d, jjd1 = 0d;
            for (int j = 0; j < i; j++)
            {
                if(!patches[j].IsNew)
                {
                    zrd += data[j, 49].NumericalValue * data[j, 52].NumericalValue / zarea;
                    lyd += data[j, 50].NumericalValue * data[j, 52].NumericalValue / zarea;
                    jjd += data[j, 51].NumericalValue * data[j, 52].NumericalValue / zarea;
                }
                else
                {
                    zrd1 += data[j, 49].NumericalValue * data[j, 52].NumericalValue / karea;
                    lyd1 += data[j, 50].NumericalValue * data[j, 52].NumericalValue / karea;
                    jjd1 += data[j, 51].NumericalValue * data[j, 52].NumericalValue / karea;
                }
            }

            data[i, 48] = new CellData("整理");
            data[i, 49] = new CellData(zrd, "0.0");
            data[i, 50] = new CellData(lyd, "0.0");
            data[i, 51] = new CellData(jjd, "0.0");
            data[i, 52] = new CellData(zarea, "0.0000");

            data[i + 1, 48] = new CellData("开发");
            data[i + 1, 49] = new CellData(zrd1, "0.0");
            data[i + 1, 50] = new CellData(lyd1, "0.0");
            data[i + 1, 51] = new CellData(jjd1, "0.0");
            data[i + 1, 52] = new CellData(karea, "0.0000");

            return data;
        }

        /// <summary>
        /// 获取评价因子的分值
        /// </summary>
        /// <param name="patch">评价单元</param>
        /// <param name="factorName">评价因子</param>
        /// <returns></returns>
        private string GetFactorValue(Patch patch, string factorName)
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

        /// <summary>
        /// 获取一般占位符数据表
        /// </summary>
        /// <param name="patches">评价单元集合</param>
        /// <returns></returns>
        public Hashtable GetPlaceholdersWithData(List<Patch> patches)
        {
            var table = new Hashtable();
            foreach (var placeholder in Placeholders)
            {
                switch (placeholder.Key)
                {
                    case "{质量等别评定}":
                        if (!Placeholder.ContainsKey(table, placeholder.Key))
                            table.Add(placeholder, Evaluate(patches));
                        break;
                }
            }
            return table;
        }

        /// <summary>
        /// 另存文件
        /// </summary>
        /// <param name="xlsFileName">评定结果文件（.xls）</param>
        /// <param name="patches">评价单元集合</param>
        public void ToSheet(string xlsFileName, List<Patch> patches)
        {
            // 获得填充了数据的占位符集合
            var table = GetPlaceholdersWithData(patches);
            // 保存文件
            base.ToSheet(xlsFileName, table, Sheets);
        }
    }
}
