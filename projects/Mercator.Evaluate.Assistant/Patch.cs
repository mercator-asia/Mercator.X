using System;
using System.ComponentModel;
using System.Data.SQLite;

namespace Mercator.Evaluate.Assistant
{
    public class Patch
    {
        /// <summary>
        /// 名称
        /// </summary>
        public string Name;

        /// <summary>
        /// 面积
        /// </summary>
        public double Area;

        /// <summary>
        /// 县、区
        /// </summary>
        public string County;

        /// <summary>
        /// 地类
        /// </summary>
        public string LandClass
        {
            get
            {
                switch(ClassCode)
                {
                    case "011":
                        return "水田";
                    case "012":
                        return "水浇地";
                    case "013":
                        return "旱地";
                    default:
                        return "";
                }
            }
        }

        /// <summary>
        /// 地类编码
        /// </summary>
        public string ClassCode
        {
            get
            {
                return _ClassCode;
            }
            set
            {
                _ClassCode = value;
                if (_ClassCode== "011")
                {
                    Factors[0] = new EvaluationFactor("表层土壤质地");
                    Factors[1] = new EvaluationFactor("剖面构型");
                    Factors[2] = new EvaluationFactor("土壤有机质含量");
                    Factors[3] = new EvaluationFactor("土壤PH值");
                    Factors[4] = new EvaluationFactor("障碍层距地表深度", EvaluationFactorValueType.Interval);
                    Factors[5] = new EvaluationFactor("排水条件");
                    Factors[6] = new EvaluationFactor("灌溉保证率");

                    Crops[0] = new Crop("水稻");
                    Crops[1] = new Crop("小麦");
                }
                else
                {
                    Factors[0] = new EvaluationFactor("有效土层厚度", EvaluationFactorValueType.Interval);
                    Factors[1] = new EvaluationFactor("表层土壤质地");
                    Factors[2] = new EvaluationFactor("土壤有机质含量");
                    Factors[3] = new EvaluationFactor("土壤PH值");
                    Factors[4] = new EvaluationFactor("地形坡度", EvaluationFactorValueType.Interval);
                    Factors[5] = new EvaluationFactor("灌溉保证率");
                    Factors[6] = new EvaluationFactor("地表岩石出露度");

                    Crops[0] = new Crop("小麦");
                    Crops[1] = new Crop("玉米");
                }
            }
        }
        public string _ClassCode;

        /// <summary>
        /// 三级指标区
        /// </summary>
        public string ThirdIndexRegion;

        /// <summary>
        /// 分等因素
        /// </summary>
        public EvaluationFactor[] Factors;

        /// <summary>
        /// 整治前土地利用水平评价指标
        /// </summary>
        public UtilizationFactor[] BeforeUtilizationFactors;

        /// <summary>
        /// 整治后土地利用水平评价指标
        /// </summary>
        public UtilizationFactor[] AfterUtilizationFactors;

        /// <summary>
        /// 是否是新增耕地
        /// </summary>
        public bool IsNew = false;

        /// <summary>
        /// 指定作物
        /// </summary>
        public Crop[] Crops;

        /// <summary>
        /// 土地经济系数
        /// </summary>
        public double EconomicalCoefficient;

        /// <summary>
        /// 整治前土地利用系数
        /// </summary>
        public double UtilizationCoefficient;

        /// <summary>
        /// 利用等
        /// </summary>
        public int StateUtilizationGrade
        {
            get
            {
                return SQLiteHelper.CalculateStateUtilizationGrade(StateUtilizationGradeIndex);
            }
        }

        /// <summary>
        /// 利用等指数
        /// </summary>
        public double UtilizationGradeIndex
        {
            get
            {
                if(!IsNew)
                {
                    SQLiteHelper.CalculateUtilizationScore(this, out double score1, out string formula1, out double score2, out string formula2);
                    var uc = SQLiteHelper.CalculateUtilizationCoefficient(UtilizationCoefficient, score1, score2);
                    return SQLiteHelper.CalculateUtilizationGradeIndex(NaturalQualityGradeIndex, uc);
                }
                else
                {
                    return SQLiteHelper.CalculateUtilizationGradeIndex(NaturalQualityGradeIndex, UtilizationCoefficient);
                }
            }
        }

        /// <summary>
        /// 国家利用等指数
        /// </summary>
        public double StateUtilizationGradeIndex
        {
            get
            {
                return SQLiteHelper.CalculateStateUtilizationGradeIndex(UtilizationGradeIndex);
            }
        }

        /// <summary>
        /// 经济等
        /// </summary>
        public int StateEconomicalGrade
        {
            get
            {
                return SQLiteHelper.CalculateStateEconomicalGrade(StateEconomicalGradeIndex);
            }
        }

        /// <summary>
        /// 经济等指数
        /// </summary>
        public double EconomicalGradeIndex
        {
            get
            {
                return SQLiteHelper.CalculateEconomicalGradeIndex(UtilizationGradeIndex, EconomicalCoefficient);
            }
        }

        /// <summary>
        /// 国家经济等指数
        /// </summary>
        public double StateEconomicalGradeIndex
        {
            get
            {
                return SQLiteHelper.CalculateStateEconomicalGradeIndex(EconomicalGradeIndex);
            }
        }

        /// <summary>
        /// 自然质量等
        /// </summary>
        public int StateNaturalQualityGrade
        {
            get
            {
                return SQLiteHelper.CalculateStateNaturalQualityGrade(StateNaturalQualityGradeIndex);
            }
        }

        /// <summary>
        /// 自然质量等指数
        /// </summary>
        public double NaturalQualityGradeIndex
        {
            get
            {
                return SQLiteHelper.CalculateNaturalQualityGradeIndex(this);
            }
        }

        /// <summary>
        /// 国家自然质量等指数
        /// </summary>
        public double StateNaturalQualityGradeIndex
        {
            get
            {
                return SQLiteHelper.CalculateStateNaturalQualityGradeIndex(NaturalQualityGradeIndex);
            }
        }

        /// <summary>
        /// 综合自然质量分
        /// </summary>
        public double NaturalQualityScore
        {
            get
            {
                return SQLiteHelper.CalculateNaturalQualityScore(this);
            }
        }

        /// <summary>
        /// 评价单元
        /// </summary>
        public Patch()
        {
            Factors = new EvaluationFactor[7];
            Crops = new Crop[2];
            BeforeUtilizationFactors = new UtilizationFactor[5];
            for(int i=0;i<5;i++)
            {
                BeforeUtilizationFactors[i] = new UtilizationFactor();
            }
            BeforeUtilizationFactors[0].Type = UtilizationFactorType.Water;
            BeforeUtilizationFactors[1].Type = UtilizationFactorType.Irrigation;
            BeforeUtilizationFactors[2].Type = UtilizationFactorType.Drainage;
            BeforeUtilizationFactors[3].Type = UtilizationFactorType.Road;
            BeforeUtilizationFactors[4].Type = UtilizationFactorType.Flatness;

            AfterUtilizationFactors = new UtilizationFactor[5];
            for (int i = 0; i < 5; i++)
            {
                AfterUtilizationFactors[i] = new UtilizationFactor();
            }
            AfterUtilizationFactors[0].Type = UtilizationFactorType.Water;
            AfterUtilizationFactors[1].Type = UtilizationFactorType.Irrigation;
            AfterUtilizationFactors[2].Type = UtilizationFactorType.Drainage;
            AfterUtilizationFactors[3].Type = UtilizationFactorType.Road;
            AfterUtilizationFactors[4].Type = UtilizationFactorType.Flatness;
        }
    }

    public class Crop
    {
        /// <summary>
        /// 作物名称
        /// </summary>
        public string Name;

        /// <summary>
        /// 光温生产潜力
        /// </summary>
        public double PTP;

        /// <summary>
        /// 气候生产潜力
        /// </summary>
        public double CPPC;

        /// <summary>
        /// 实际产量（千克/亩）
        /// </summary>
        public double GrainOutput;

        /// <summary>
        /// 实际成本（元/亩）
        /// </summary>
        public double Cost;

        /// <summary>
        /// 产量比系数
        /// </summary>
        public double GOCCoefficient
        {
            get
            {
                var ratio = 0d;
                switch(Name)
                {
                    case "水稻":
                        ratio = 1.0d;
                        break;
                    case "玉米":
                        ratio = 0.8d;
                        break;
                    case "小麦":
                        ratio =1.3d;
                        break;
                }
                return ratio;
            }
        }

        /// <summary>
        /// 作物
        /// </summary>
        /// <param name="name">作物名称</param>
        public Crop(string name)
        {
            Name = name;
        }
    }

    public class SQLiteHelper
    {
        /// <summary>
        /// SQLite数据库连接字符串
        /// </summary>
        private static string _ConnectionString = "Data Source=" + System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase + @"Database\Evaluate.db";

        /// <summary>
        /// 计算指定作物的农用地自然质量分
        /// </summary>
        /// <param name="patch">地块</param>
        /// <param name="cropType">指定作物</param>
        /// <returns></returns>
        public static double CalculateCropNaturalQualityScore(Patch patch, string cropName, out string formula)
        {
            var score = 0d;
            formula = "（";

            SQLiteConnection connection = new SQLiteConnection(_ConnectionString);
            SQLiteCommand command = new SQLiteCommand();
            command.Connection = connection;
            connection.Open();

            var scores = new int[] { 0, 0, 0, 0, 0, 0, 0 };
            var powers = new double[] { 0, 0, 0, 0, 0, 0, 0 };

            var landClass = patch.LandClass;

            if (landClass == "水浇地") { landClass = "旱地"; }

            for (int i = 0; i < 7; i++)
            {
                if (patch.Factors[i].ValueType == EvaluationFactorValueType.Fixed)
                {
                    command.CommandText = string.Format("SELECT 分值 FROM 自然质量分记分规则表 WHERE 分等因素='{0}' AND 等值='{1}' AND 指定作物='{2}' AND 农用地类型='{3}' AND 指标区='{4}'", patch.Factors[i].Name, patch.Factors[i].Value, cropName, landClass, patch.ThirdIndexRegion);
                    SQLiteDataReader reader = command.ExecuteReader();
                    if (reader.Read()) { scores[i] = reader.GetInt32(0); }
                    reader.Close();
                }
                else
                {
                    command.CommandText = string.Format("SELECT 分值 FROM 自然质量分记分规则表 WHERE 分等因素='{0}' AND 最大值>{1} AND 最小值<={1} AND 指定作物='{2}' AND 农用地类型='{3}' AND 指标区='{4}'", patch.Factors[i].Name, patch.Factors[i].Value, cropName, landClass, patch.ThirdIndexRegion);
                    SQLiteDataReader reader = command.ExecuteReader();
                    if (reader.Read()) { scores[i] = reader.GetInt32(0); }
                    reader.Close();
                }

                command.CommandText = string.Format("SELECT 权重 FROM 分等因素及权重表 WHERE 指标区='{0}' AND 分等因素='{1}' AND 农用地类型='{2}'", patch.ThirdIndexRegion, patch.Factors[i].Name, landClass);
                SQLiteDataReader powerReader = command.ExecuteReader();
                if (powerReader.Read()) { powers[i] = powerReader.GetDouble(0); }
                powerReader.Close();

                score += scores[i] * powers[i];

                formula += string.Format("{0}×{1}+", scores[i], powers[i]);
            }

            formula = formula.Substring(0, formula.Length - 1) + "）÷100";

            score = score / 100;

            connection.Close();

            return score;
        }

        /// <summary>
        /// 获取指定作物的光温潜力指数
        /// </summary>
        /// <param name="patch"></param>
        /// <param name="cropName"></param>
        /// <returns></returns>
        public static double GetCropPTP(string county, string cropName)
        {
            var ptp = 0d;
            SQLiteConnection connection = new SQLiteConnection(_ConnectionString);
            SQLiteCommand command = new SQLiteCommand();
            command.Connection = connection;
            connection.Open();
            command.CommandText = string.Format("SELECT 光温潜力指数 FROM 作物生产潜力指数 WHERE 县='{0}' AND 指定作物='{1}'", county, cropName);
            SQLiteDataReader reader = command.ExecuteReader();
            if (reader.Read())
            {
                ptp = reader.GetDouble(0);
            }
            reader.Close();
            connection.Close();

            return ptp;
        }

        /// <summary>
        /// 获取指定作物的气候潜力指数
        /// </summary>
        /// <param name="patch"></param>
        /// <param name="cropName"></param>
        /// <returns></returns>
        public static double GetCropCPPC(string county, string cropName)
        {
            var ptp = 0d;
            SQLiteConnection connection = new SQLiteConnection(_ConnectionString);
            SQLiteCommand command = new SQLiteCommand();
            command.Connection = connection;
            connection.Open();
            command.CommandText = string.Format("SELECT 气候潜力指数 FROM 作物生产潜力指数 WHERE 县='{0}' AND 指定作物='{1}'", county, cropName);
            SQLiteDataReader reader = command.ExecuteReader();
            if (reader.Read())
            {
                ptp = reader.GetDouble(0);
            }
            reader.Close();
            connection.Close();

            return ptp;
        }

        /// <summary>
        /// 计算评价单元的综合自然质量分
        /// </summary>
        public static double CalculateNaturalQualityScore(Patch patch)
        {
            var totalScore = 0d;

            var valueUp = 0d;
            var valueDown = 0d;

            for (int i = 0; i < 2; i++)
            {
                string formula;
                var score = CalculateCropNaturalQualityScore(patch, patch.Crops[i].Name,out formula);
                if(patch.LandClass=="水田")
                {
                    var ptp = GetCropPTP(patch.County, patch.Crops[i].Name);
                    valueUp += score * ptp * patch.Crops[i].GOCCoefficient;
                    valueDown += ptp * patch.Crops[i].GOCCoefficient;
                }
                else
                {                   
                    var cppc = GetCropCPPC(patch.County, patch.Crops[i].Name);
                    valueUp += score * cppc * patch.Crops[i].GOCCoefficient;
                    valueDown += cppc * patch.Crops[i].GOCCoefficient;
                }
            }

            totalScore = valueUp / valueDown;

            return totalScore;
        }

        /// <summary>
        /// 计算评价单元的自然质量等指数
        /// </summary>
        /// <param name="patch"></param>
        /// <returns></returns>
        public static double CalculateNaturalQualityGradeIndex(Patch patch)
        {
            var totalScore = 0d;

            if (patch.LandClass == "水田")
            {
                for (int i = 0; i < 2; i++)
                {
                    string formula;
                    var score = CalculateCropNaturalQualityScore(patch, patch.Crops[i].Name, out formula);
                    var ptp = GetCropPTP(patch.County, patch.Crops[i].Name);
                    totalScore += score * ptp * patch.Crops[i].GOCCoefficient;
                }
            }
            else
            {
                for (int i = 0; i < 2; i++)
                {
                    string formula;
                    var score = CalculateCropNaturalQualityScore(patch, patch.Crops[i].Name, out formula);
                    var cppc = GetCropCPPC(patch.County, patch.Crops[i].Name);
                    totalScore += score * cppc * patch.Crops[i].GOCCoefficient;
                }
            }

            return totalScore;
        }

        /// <summary>
        /// 计算评价单元的自然质量等（省等）
        /// </summary>
        /// <param name="patch"></param>
        /// <returns></returns>
        public static int CalculateNaturalQualityGrade(double naturalQualityGradeIndex)
        {
            var grade = 0;

            var index = naturalQualityGradeIndex;

            // 向上进位取整
            grade = (int)Math.Ceiling(index / 200);

            return grade;
        }

        /// <summary>
        /// 计算评价单元的实际粮食产量（Y = ∑指定作物实际产量×指定作物产量比系数）
        /// </summary>
        /// <param name="patch"></param>
        /// <returns></returns>
        private static double CalculateGrainOutput(Patch patch)
        {
            var output = 0d;

            for (int i = 0; i < 2; i++)
            {
                output += patch.Crops[i].GOCCoefficient * patch.Crops[i].GrainOutput;
            }

            return output;
        }

        /// <summary>
        /// 计算评价单元的最大粮食产量（Ymax = ∑指定作物最大产量×指定作物产量比系数）
        /// </summary>
        /// <param name="patch"></param>
        /// <returns></returns>
        private static double CalculateMaxGrainOutput(Patch patch)
        {
            var output = 0d;

            for (int i = 0; i < 2; i++)
            {
                output += patch.Crops[i].GOCCoefficient * SelectMaxGrainOutput(patch.County, patch.Crops[i].Name);
            }

            return output;
        }

        /// <summary>
        /// 计算评价单元的土地利用系数（Ki = Y / Ymax）
        /// </summary>
        /// <param name="patch"></param>
        /// <returns></returns>
        private static double CalculateUtilizationCoefficient(Patch patch)
        {
            var utilizationCoefficient = 0d;

            utilizationCoefficient = CalculateGrainOutput(patch) / CalculateMaxGrainOutput(patch);

            if(utilizationCoefficient > 1) { utilizationCoefficient = 1; }

            return utilizationCoefficient;
        }

        /// <summary>
        /// 计算土地综合利用分值
        /// </summary>
        /// <param name="patch">地块</param>
        /// <returns></returns>
        public static void CalculateUtilizationScore(Patch patch, out double score1, out string formula1, out double score2, out string formula2)
        {
            formula1 = formula2 = "";
            score1 = score2 = 0d;
            for (int i = 0; i < 5; i++)
            {
                score1 += patch.BeforeUtilizationFactors[i].Score * patch.BeforeUtilizationFactors[i].Power;
                score2 += patch.AfterUtilizationFactors[i].Score * patch.AfterUtilizationFactors[i].Power;
                formula1 += string.Format("{0}×{1}+", patch.BeforeUtilizationFactors[i].Score, patch.BeforeUtilizationFactors[i].Power);
                formula2 += string.Format("{0}×{1}+", patch.AfterUtilizationFactors[i].Score, patch.AfterUtilizationFactors[i].Power);
            }
            formula1 = formula1.Substring(0, formula1.Length - 1);
            formula2 = formula2.Substring(0, formula2.Length - 1);
        }

        /// <summary>
        /// 根据修正系数计算评价单元的土地利用系数
        /// </summary>
        /// <param name="beforeCoefficient">整治前土地利用系数(K)</param>
        /// <param name="beforeScore">整治前土地综合利用分值</param>
        /// <param name="afterScore">整治后土地综合利用分值</param>
        /// <returns></returns>
        public static double CalculateUtilizationCoefficient(double beforeCoefficient, double beforeScore, double afterScore)
        {
            return beforeCoefficient * afterScore / beforeScore;
        }

        /// <summary>
        /// 获取三级区的最大粮食产量（计算Ymax）
        /// </summary>
        /// <param name="county"></param>
        /// <param name="cropName"></param>
        /// <returns></returns>
        private static int SelectMaxGrainOutput(string county,string cropName)
        {
            var maxValue = 0;
            SQLiteConnection connection = new SQLiteConnection(_ConnectionString);
            SQLiteCommand command = new SQLiteCommand();
            command.Connection = connection;
            connection.Open();
            command.CommandText = string.Format("SELECT 最大单产 FROM 最大单产 WHERE 县='{0}' AND 指定作物='{1}'", county, cropName);
            SQLiteDataReader reader = command.ExecuteReader();
            if (reader.Read())
            {
                maxValue = reader.GetInt32(0) / 15;     // 数据库中的单位为：千克/公顷
            }
            reader.Close();
            connection.Close();
            return maxValue;
        }

        /// <summary>
        /// 计算评价单元的“产量-成本”指数（ai = Yi / C）
        /// </summary>
        /// <param name="patch"></param>
        /// <returns></returns>
        private static double CalculateGOCIndex(Patch patch)
        {
            var grainOutput = CalculateGrainOutput(patch);
            var cost = 0d;
            for (int i = 0; i < 2; i++)
            {
                cost += patch.Crops[i].Cost;
            }

            return grainOutput / cost;
        }

        /// <summary>
        /// 计算评价单元的土地经济系数（Kc = ai / A）
        /// </summary>
        /// <param name="patch"></param>
        /// <returns></returns>
        private static double CalculateEconomicalCoefficient(Patch patch)
        {
            var economicalCoefficient = 0d;

            economicalCoefficient = CalculateGOCIndex(patch) / SelectMaxGOCIndex(patch.County);

            if (economicalCoefficient > 1) { economicalCoefficient = 1; }

            return economicalCoefficient;
        }

        /// <summary>
        /// 获取三级区的最大产量-成本指数（A）
        /// </summary>
        /// <param name="county"></param>
        /// <returns></returns>
        private static double SelectMaxGOCIndex(string county)
        {
            var maxValue = 0d;
            SQLiteConnection connection = new SQLiteConnection(_ConnectionString);
            SQLiteCommand command = new SQLiteCommand();
            command.Connection = connection;
            connection.Open();
            command.CommandText = string.Format("SELECT 最大产量成本指数 FROM 最大产量成本指数 WHERE 县='{0}'", county);
            SQLiteDataReader reader = command.ExecuteReader();
            if (reader.Read())
            {
                maxValue = reader.GetDouble(0);
            }
            reader.Close();
            connection.Close();
            return maxValue;
        }

        /// <summary>
        /// 根据等值区获取评价单元的综合土地利用系数
        /// </summary>
        /// <param name="patch"></param>
        /// <returns></returns>
        public static double GetUtilizationCoefficient(string county, double utilizationCoefficient)
        {
            var coefficient = 0d;

            SQLiteConnection connection = new SQLiteConnection(_ConnectionString);
            SQLiteCommand command = new SQLiteCommand();
            command.Connection = connection;
            connection.Open();
            command.CommandText = string.Format("SELECT 综合土地利用系数 FROM 综合土地利用系数 WHERE 县='{0}' AND 下限值<{1} AND 上限值>={1}", county, utilizationCoefficient);
            SQLiteDataReader reader = command.ExecuteReader();
            if (reader.Read())
            {
                coefficient = reader.GetDouble(0);
            }
            reader.Close();
            connection.Close();

            return coefficient;
        }

        /// <summary>
        /// 计算评价单元的综合土地利用等指数
        /// </summary>
        /// <param name="patch"></param>
        /// <returns></returns>
        public static double CalculateUtilizationGradeIndex(double naturalQualityGradeIndex, double utilizationCoefficient)
        {
            var index = 0d;

            index = naturalQualityGradeIndex * utilizationCoefficient;

            return index;
        }

        /// <summary>
        /// 计算评价单元的土地利用等（省等）
        /// </summary>
        /// <param name="patch"></param>
        /// <returns></returns>
        public static int CalculateUtilizationGrade(double utilizationGradeIndex)
        {
            var grade = 0;

            // 向上进位取整
            grade = (int)Math.Ceiling(utilizationGradeIndex / 200);

            return grade;
        }

        /// <summary>
        /// 获取评价单元的综合经济系数
        /// </summary>
        /// <param name="patch"></param>
        /// <returns></returns>
        public static double GetEconomicalCoefficient(string county, double economicalCoefficient)
        {
            var coefficient = 0d;

            SQLiteConnection connection = new SQLiteConnection(_ConnectionString);
            SQLiteCommand command = new SQLiteCommand();
            command.Connection = connection;
            connection.Open();
            command.CommandText = string.Format("SELECT 综合土地经济系数 FROM 综合土地经济系数 WHERE 县='{0}' AND 下限值<{1} AND 上限值>={1}", county, economicalCoefficient);
            SQLiteDataReader reader = command.ExecuteReader();
            if (reader.Read())
            {
                coefficient = reader.GetDouble(0);
            }
            reader.Close();
            connection.Close();

            return coefficient;
        }

        /// <summary>
        /// 计算评价单元的经济等指数
        /// </summary>
        /// <param name="patch"></param>
        /// <returns></returns>
        public static double CalculateEconomicalGradeIndex(double utilizationGradeIndex, double economicalCoefficient)
        {
            var index = 0d;

            index = utilizationGradeIndex * economicalCoefficient;

            return index;
        }

        /// <summary>
        /// 计算评价单元的经济等（省等）
        /// </summary>
        /// <param name="patch"></param>
        /// <returns></returns>
        public static int CalculateEconomicalGrade(double economicalGradeIndex)
        {
            var grade = 0;

            // 向上进位取整
            grade = (int)Math.Ceiling(economicalGradeIndex / 200);

            return grade;
        }

        /// <summary>
        /// 计算评价单元的国家自然质量等指数
        /// </summary>
        /// <param name="patch"></param>
        /// <returns></returns>
        public static double CalculateStateNaturalQualityGradeIndex(double naturalQualityGradeIndex)
        {
            return naturalQualityGradeIndex * 0.5148 + 1020.28;
        }

        /// <summary>
        /// 计算评价单元的国家利用等指数
        /// </summary>
        /// <param name="patch"></param>
        /// <returns></returns>
        public static double CalculateStateUtilizationGradeIndex(double utilizationGradeIndex)
        {
            return utilizationGradeIndex * 0.5598 + 539.70;
        }

        /// <summary>
        /// 计算评价单元的国家经济等指数
        /// </summary>
        /// <param name="patch"></param>
        /// <returns></returns>
        public static double CalculateStateEconomicalGradeIndex(double economicalGradeIndex)
        {
            return economicalGradeIndex * 0.6998 + 676.04;
        }

        /// <summary>
        /// 计算评价单元的国家自然质量等
        /// </summary>
        /// <param name="patch"></param>
        /// <returns></returns>
        public static int CalculateStateNaturalQualityGrade(double stateNaturalQualityGradeIndex)
        {
            var index = stateNaturalQualityGradeIndex;

            if (index > 0 && index <= 400) { return 15; }
            if (index > 400 && index <= 800) { return 14; }
            if (index > 800 && index <= 1200) { return 13; }
            if (index > 1200 && index <= 1600) { return 12; }
            if (index > 1600 && index <= 2000) { return 11; }
            if (index > 2000 && index <= 2400) { return 10; }
            if (index > 2400 && index <= 2800) { return 9; }
            if (index > 2800 && index <= 3200) { return 8; }
            if (index > 3200 && index <= 3600) { return 7; }
            if (index > 3600 && index <= 4000) { return 6; }
            if (index > 4000 && index <= 4400) { return 5; }
            if (index > 4400 && index <= 4800) { return 4; }
            if (index > 4800 && index <= 5200) { return 3; }
            if (index > 5200 && index <= 5600) { return 2; }
            if (index > 5600 && index <= 6000) { return 1; }

            return 0;
        }

        /// <summary>
        /// 计算评价单元的国家利用等
        /// </summary>
        /// <param name="patch"></param>
        /// <returns></returns>
        public static int CalculateStateUtilizationGrade(double stateUtilizationGradeIndex)
        {
            var index = stateUtilizationGradeIndex;

            if (index > 0 && index <= 200) { return 15; }
            if (index > 200 && index <= 400) { return 14; }
            if (index > 400 && index <= 600) { return 13; }
            if (index > 600 && index <= 800) { return 12; }
            if (index > 800 && index <= 1000) { return 11; }
            if (index > 1000 && index <= 1200) { return 10; }
            if (index > 1200 && index <= 1400) { return 9; }
            if (index > 1400 && index <= 1600) { return 8; }
            if (index > 1600 && index <= 1800) { return 7; }
            if (index > 1800 && index <= 2000) { return 6; }
            if (index > 2000 && index <= 2200) { return 5; }
            if (index > 2200 && index <= 2400) { return 4; }
            if (index > 2400 && index <= 2600) { return 3; }
            if (index > 2600 && index <= 2800) { return 2; }
            if (index > 2800 && index <= 3000) { return 1; }

            return 0;
        }

        /// <summary>
        /// 计算评价单元的国家经济等
        /// </summary>
        /// <param name="patch"></param>
        /// <returns></returns>
        public static int CalculateStateEconomicalGrade(double stateEconomicalGradeIndex)
        {
            var index = stateEconomicalGradeIndex;

            if (index > 0 && index <= 200) { return 15; }
            if (index > 200 && index <= 400) { return 14; }
            if (index > 400 && index <= 600) { return 13; }
            if (index > 600 && index <= 800) { return 12; }
            if (index > 800 && index <= 1000) { return 11; }
            if (index > 1000 && index <= 1200) { return 10; }
            if (index > 1200 && index <= 1400) { return 9; }
            if (index > 1400 && index <= 1600) { return 8; }
            if (index > 1600 && index <= 1800) { return 7; }
            if (index > 1800 && index <= 2000) { return 6; }
            if (index > 2000 && index <= 2200) { return 5; }
            if (index > 2200 && index <= 2400) { return 4; }
            if (index > 2400 && index <= 2600) { return 3; }
            if (index > 2600 && index <= 2800) { return 2; }
            if (index > 2800 && index <= 3000) { return 1; }

            return 0;
        }
    }

    /// <summary>
    /// 分等因素
    /// </summary>
    public class EvaluationFactor
    {
        public string Name;
        public string Value;
        public EvaluationFactorValueType ValueType;

        public EvaluationFactor(string name, EvaluationFactorValueType valueType=EvaluationFactorValueType.Fixed)
        {
            Name = name;
            ValueType = valueType;
        }
    }

    /// <summary>
    /// 分等因素值类型
    /// </summary>
    public enum EvaluationFactorValueType
    {
        [Description("固定值")]
        Fixed,
        [Description("区间值")]
        Interval
    }

    public class UtilizationFactor
    {
        /// <summary>
        /// 指标名称
        /// </summary>
        public string Name
        {
            get
            {
                return _Name;
            }
        }
        public string _Name;

        /// <summary>
        /// 指标级别
        /// </summary>
        public int Value;

        /// <summary>
        /// 指标分值
        /// </summary>
        public int Score
        {
            get
            {
                short score = 0;
                switch (Type)
                {
                    case UtilizationFactorType.Water:
                        _Name = "水源类型";
                        _Power = 0.1;
                        _Description = "反映耕地的生产成本及土地整理项目的工程投入";
                        switch (Value)
                        {
                            case 1:
                                _Standard = "地表水（河流、水库）";
                                score = 100;
                                break;
                            case 2:
                                _Standard = "地下水";
                                score = 80;
                                break;
                            case 3:
                                _Standard = "天然降水";
                                score = 50;
                                break;
                        }
                        break;
                    case UtilizationFactorType.Irrigation:
                        _Name = "灌溉方式";
                        _Power = 0.25;
                        _Description = "可描述的灌溉条件，反映资金投资及灌溉的效果";
                        switch (Value)
                        {
                            case 1:
                                _Standard = "管灌或管灌为主";
                                score = 100;
                                break;
                            case 2:
                                _Standard = "渠灌或渠灌为主";
                                score = 80;
                                break;
                            case 3:
                                _Standard = "集雨灌或水窖方式灌溉";
                                score = 50;
                                break;
                            case 4:
                                _Standard = "天然无灌溉";
                                score = 30;
                                break;
                        }
                        break;
                    case UtilizationFactorType.Drainage:
                        _Name = "排水方式";
                        _Power = 0.2;
                        _Description = "反映生产成本的高低及排水的条件的优劣";
                        switch (Value)
                        {
                            case 1:
                                _Standard = "自排";
                                score = 100;
                                break;
                            case 2:
                                _Standard = "人工排水";
                                score = 80;
                                break;
                        }
                        break;
                    case UtilizationFactorType.Road:
                        _Name = "道路通达性";
                        _Power = 0.25;
                        _Description = "描述耕地所处位置的交通便利程度，一个地方的道路密度越大，生产就越方便，反之越不方便。反映土地整理项目的工程投入";
                        switch (Value)
                        {
                            case 1:
                                _Standard = "达到工程建设标准，有完善的田间道路系统，生产便捷";
                                score = 100;
                                break;
                            case 2:
                                _Standard = "达到工程建设标准，但还未形成健全的田间道路体系";
                                score = 80;
                                break;
                            case 3:
                                _Standard = "没有达到工程建设标准，并不能满足生产要求";
                                score = 50;
                                break;
                        }
                        break;
                    case UtilizationFactorType.Flatness:
                        _Name = "田块平整度";
                        _Power = 0.2;
                        _Description = "可描述的田块平整情况，平整度高的田块利于农作物的灌溉和机械操作";
                        switch (Value)
                        {
                            case 1:
                                _Standard = "田块平整规则，便于机械耕作";
                                score = 100;
                                break;
                            case 2:
                                _Standard = "田块比较平整规则，不影响机械耕作";
                                score = 80;
                                break;
                            case 3:
                                _Standard = "田块平整不太规则，对机械耕作影响不大";
                                score = 50;
                                break;
                            case 4:
                                _Standard = "田块既不平整也不规则，机械难以耕作";
                                score = 30;
                                break;
                        }
                        break;
                }

                return score;
            }
        }

        /// <summary>
        /// 指标类型
        /// </summary>
        public UtilizationFactorType Type;

        /// <summary>
        /// 权重
        /// </summary>
        public double Power
        {
            get
            {
                return _Power;
            }
        }
        public double _Power;

        /// <summary>
        /// 指标描述
        /// </summary>
        public string Description
        {
            get
            {
                return _Description;
            }
        }
        public string _Description;

        /// <summary>
        /// 分级标准
        /// </summary>
        public string Standard
        {
            get
            {
                return _Standard;
            }
        }
        private string _Standard;
    }

    public enum UtilizationFactorType
    {
        [Description("水源类型")]
        Water,
        [Description("灌溉方式")]
        Irrigation,
        [Description("排水方式")]
        Drainage,
        [Description("道路通达性")]
        Road,
        [Description("田块平整度")]
        Flatness
    }
}
