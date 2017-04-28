using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using ClosedXML.Excel;


namespace HNCDIExternalProjectManage
{
    /// <summary>
    /// PrizesDetailYear.xaml 的交互逻辑
    /// </summary>
    public partial class PrizesDetailYear : Window
    {
        private DataClassesProjectClassifyDataContext dataContext;
        private int _totalPrizeClassify = 0; //奖金总类别数
        private List<string> _prizeClasifies; //奖金类别列表 
        private List<Prizes> _prizeses;
        private List<string> _departmets; //部门列表
        private int _totalDepartment; //总部门数
        private List<string> _declareDepartments; //申报部门列表
        private int _totalDeclareDepartment; //申报部门数
        private List<Employee> _employees; //员工列表
        private int _totalEmployee; //总员工数
        private decimal _totalMoney; //总奖金
        private string _year; //年度
        private List<string> _projects; //项目列表
        private int _totalProjects; //项目总数


        public PrizesDetailYear()
        {
            InitializeComponent();
        }

        // Creates a SpreadsheetDocument.
        private void CreatePackage(string filePath)
        {
            using (XLWorkbook package = new XLWorkbook())
            {
                CreateParts(package);
                package.SaveAs(filePath);
            }
        }

        // Adds child parts and generates content of the specified part.
        private void CreateParts(XLWorkbook document)
        {
            //生成“汇总”sheet
            GenerateTotalSheet(document);
            //生成部门汇总
            GenerateDepartmentTotalSheet(document);
            //生成个人汇总
            GenaratePersonalTotalSheet(document);
            //生成各部门汇总
            GenaratePerDepartmentTotalSheet(document);
        }

        /// <summary>
        /// 生成“汇总”sheet
        /// </summary>
        /// <param name="document"></param>
        private void GenerateTotalSheet(XLWorkbook document)
        {
            try
            {
                var ws = document.Worksheets.Add("汇总");
                ws.Cell(1, 1).Value = "HNCDI(" + _year + "年度)获奖成果与发表论著奖励明细(单位：万元)";
                var title = ws.Range("A1:F1");
                title.Row(1).Merge(); //合并标题行
                title.Style.Font.FontSize = 14;
                title.Style.Font.FontName = "宋体";
                title.Style.Font.Bold = true;
                title.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                ws.Cell(2, 1).Value = "序号";
                ws.Cell(2, 2).Value = "奖励对象";
                ws.Cell(2, 3).Value = "奖别";
                ws.Cell(2, 4).Value = "单位";
                ws.Cell(2, 5).Value = "奖金";
                ws.Cell(2, 6).Value = "备注";

                var colA = ws.Column("A");
                colA.Width = 3;
                ws.Column("B").Width = 29;
                ws.Column("C").Width = 20.88;
                ws.Column("D").Width = 12.35;
                ws.Column("E").Width = 9.5;
                ws.Column("F").Width = 4.38;
                ws.Row(1).AdjustToContents(1);


                var secondRow = ws.Range("A2:F2");
                secondRow.Style.Font.FontSize = 10;
                secondRow.Style.Font.FontName = "宋体";
                secondRow.Style.Font.Bold = true;
                secondRow.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Row(2).AdjustToContents();

                int currentRow = 3;
                int currentClassify = 1;
                DigitToChnText digitToChnText = new DigitToChnText();
                foreach (string prizeClassify in _prizeClasifies)
                {
                    //计算该类总奖金
                    decimal classifyMoney =
                        _prizeses.Where(p => p.PrizeClassify.Equals(prizeClassify)).Sum(p => p.Prize) ?? 0;
                    ws.Cell(currentRow, 1).Value = digitToChnText.Convert(currentClassify.ToString(), false);
                    ws.Cell(currentRow, 2).Value = prizeClassify;
                    ws.Cell(currentRow, 5).Value = (classifyMoney / 10000).ToString("N4");

                    var classifyRow = ws.Range("A" + currentRow.ToString() + ":F" + currentRow.ToString());
                    classifyRow.Style.Font.Bold = true;
                    classifyRow.Style.Font.FontSize = 10;
                    classifyRow.Style.Font.FontName = "宋体";
                    classifyRow.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    classifyRow.Style.Fill.BackgroundColor = XLColor.Gainsboro;
                    ws.Row(currentRow).AdjustToContents();

                    currentRow++;
                    //统计该类各项目
                    int projectNo = 1;
                    List<Prizes> classifyPrizes = _prizeses.Where(p => p.PrizeClassify.Equals(prizeClassify)).ToList();
                    var prizeses =
                        classifyPrizes.Select(
                            p =>
                                new
                                {
                                    PrizeClassify = p.PrizeClassify,
                                    Project = p.Project,
                                    AwardName = p.AwardName,
                                    DeclareDepartment = p.DeclareDepartment
                                })
                            .Distinct()
                            .ToList();

                    foreach (var prizes in prizeses)
                    {
                        ws.Cell(currentRow, 1).Value = projectNo;
                        ws.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        ws.Cell(currentRow, 2).Value = prizes.Project;
                        ws.Cell(currentRow, 2).Style.Alignment.WrapText = true;
                        ws.Cell(currentRow, 3).Value = prizes.AwardName;
                        ws.Cell(currentRow, 4).Value = prizes.DeclareDepartment;
                        //计算该项目奖金
                        decimal projectMoney =
                            _prizeses.Where(
                                p =>
                                    p.Project.Equals(prizes.Project) && p.AwardName.Equals(prizes.AwardName) &&
                                    p.DeclareDepartment.Equals(prizes.DeclareDepartment)).Sum(p => p.Prize) ?? 0;
                        ws.Cell(currentRow, 5).Value = (projectMoney / 10000).ToString("N4");

                        var projectRow = ws.Range("A" + currentRow.ToString() + ":F" + currentRow.ToString());
                        projectRow.Style.Font.FontSize = 9;
                        projectRow.Style.Font.FontName = "宋体";
                        ws.Range("C" + currentRow.ToString() + ":F" + currentRow.ToString()).Style.Alignment.Horizontal
                            =
                            XLAlignmentHorizontalValues.Center;
                        ws.Row(currentRow).AdjustToContents();
                        currentRow++;
                        projectNo++;
                    }
                    currentClassify++;
                }

                //添加总计行
                ws.Cell(currentRow, 2).Value = "总  计";
                ws.Cell(currentRow, 5).Value = (_totalMoney / 10000).ToString("N4");
                var totalRow = ws.Range("A" + currentRow.ToString() + ":F" + currentRow.ToString());
                totalRow.Style.Font.FontSize = 10;
                totalRow.Style.Font.FontName = "宋体";
                totalRow.Style.Font.Bold = true;
                totalRow.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                totalRow.Style.Fill.BackgroundColor = XLColor.DarkOrange;

                var mainBody = ws.Range("A2:F" + currentRow.ToString());
                mainBody.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                mainBody.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
                mainBody.Style.Border.OutsideBorderColor = XLColor.Black;
                mainBody.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                mainBody.Style.Border.InsideBorderColor = XLColor.Black;
                currentRow++;
                ws.Cell(currentRow, 1).Value = "编制";
                ws.Cell(currentRow, 3).Value = "审核：";
                ws.Cell(currentRow, 4).Value = "批准：";
                ws.Row(currentRow).AdjustToContents();
                ws.Columns().AdjustToContents();
                ws.Rows().AdjustToContents();
                ws.Rows().Height = 15;
                ws.Row(1).Height = 20;
                //MessageBox.Show("生成汇总表成功！");
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
        }

        /// <summary>
        /// 生成“部门汇总”sheet
        /// </summary>
        /// <param name="document"></param>
        private void GenerateDepartmentTotalSheet(XLWorkbook document)
        {
            try
            {
                var ws = document.Worksheets.Add("部门汇总");

                //标题行
                ws.Cell(1, 1).Value = _year + "年度年终奖励汇总表";
                var titleRow = ws.Range(1, 1, 1, _totalPrizeClassify + 3);
                titleRow.Merge();
                titleRow.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                titleRow.Style.Font.Bold = true;
                titleRow.Style.Font.FontName = "宋体";
                titleRow.Style.Font.FontSize = 14;
                ws.Row(1).AdjustToContents();

                //第二行
                var cellRow2EndCell = ws.Cell(2, _totalPrizeClassify + 3);
                cellRow2EndCell.Value = "单位：元";
                cellRow2EndCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                cellRow2EndCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                cellRow2EndCell.Style.Font.FontName = "宋体";
                cellRow2EndCell.Style.Font.FontSize = 10;
                ws.Row(2).AdjustToContents();

                //第三行
                ws.Cell(3, 1).Value = "序号";
                ws.Cell(3, 2).Value = "部门";
                Dictionary<int, string> dictionarylassify = new Dictionary<int, string>();
                int currentCol = 3;
                foreach (string classify in _prizeClasifies)
                {
                    ws.Cell(3, currentCol).Value = classify;
                    dictionarylassify.Add(currentCol, classify);
                    currentCol++;
                }
                ws.Cell(3, currentCol).Value = "合计";
                var row3Rang = ws.Range(3, 1, 3, currentCol);
                row3Rang.Style.Font.FontSize = 10;
                row3Rang.Style.Font.FontName = "宋体";
                row3Rang.Style.Fill.BackgroundColor = XLColor.CadetBlue;

                //填写部门数据
                int departmentNo = 1;
                int currentRow = 4;
                string currentClassify = "";
                Dictionary<string, decimal> dictionaryClassifyMoney =
                    _prizeClasifies.ToDictionary<string, string, decimal>(classify => classify, classify => 0);
                foreach (string department in _departmets)
                {
                    decimal totalDepartmentMoney = 0;
                    ws.Cell(currentRow, 1).Value = departmentNo.ToString();
                    ws.Cell(currentRow, 2).Value = department;
                    for (currentCol = 3; currentCol < _totalPrizeClassify + 3; currentCol++)
                    {
                        currentClassify = dictionarylassify[currentCol];
                        var currentMoney = _prizeses.Where(
                            p => p.Department.Equals(department) && p.PrizeClassify.Equals(currentClassify))
                            .Sum(p => p.Prize) ?? 0;
                        totalDepartmentMoney += currentMoney;
                        dictionaryClassifyMoney[currentClassify] += currentMoney;
                        ws.Cell(currentRow, currentCol).Value = currentMoney == 0 ? "0" : currentMoney.ToString("N0");
                    }
                    ws.Cell(currentRow, currentCol).Value = totalDepartmentMoney == 0
                        ? "0"
                        : totalDepartmentMoney.ToString("N0");
                    ws.Cell(currentRow, currentCol).Style.Font.Bold = true;
                    departmentNo++;
                    currentRow++;
                }
                //合计行
                ws.Cell(currentRow, 2).Value = "合计";
                for (currentCol = 3; currentCol < _totalPrizeClassify + 3; currentCol++)
                {
                    ws.Cell(currentRow, currentCol).Value = dictionaryClassifyMoney[dictionarylassify[currentCol]] == 0
                        ? "0"
                        : dictionaryClassifyMoney[dictionarylassify[currentCol]].ToString("N0");
                }
                ws.Cell(currentRow, currentCol).Value = _totalMoney == 0 ? "0" : _totalMoney.ToString("N0");
                var rowLast = ws.Range(currentRow, 1, currentRow, currentCol);
                rowLast.Style.Fill.BackgroundColor = XLColor.Orange;
                rowLast.Style.Font.Bold = true;
                var dataBody = ws.Range(3, 1, currentRow, currentCol);
                dataBody.Style.Font.FontName = "宋体";
                dataBody.Style.Font.FontSize = 10;
                dataBody.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
                dataBody.Style.Border.OutsideBorderColor = XLColor.Black;
                dataBody.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                dataBody.Style.Border.InsideBorderColor = XLColor.Black;
                dataBody.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                dataBody.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                dataBody.Style.Alignment.WrapText = false;
                ws.Columns().AdjustToContents();
                ws.Rows().AdjustToContents();
                ws.Rows().Height = 15;
                ws.Row(1).Height = 20;
                //MessageBox.Show("生成部门汇总表成功");
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
        }

        /// <summary>
        /// 生成“个人汇总”sheet
        /// </summary>
        /// <param name="document"></param>
        private void GenaratePersonalTotalSheet(XLWorkbook document)
        {
            try
            {
                var ws = document.Worksheets.Add("个人汇总");
                //标题行
                ws.Cell(1, 1).Value = _year + "年度科技成果年终奖励个人汇总表（单位：元）";
                var title = ws.Range(1, 1, 1, _totalPrizeClassify + 4);
                title.Merge();
                title.Style.Font.Bold = true;
                title.Style.Font.FontName = "宋体";
                title.Style.Font.FontSize = 14;
                title.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                title.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                //第二行
                ws.Cell(2, 1).Value = "序号";
                ws.Cell(2, 2).Value = "部门";
                ws.Cell(2, 3).Value = "姓名";
                Dictionary<string, decimal> dictionaryClassifyMoney =
                    _prizeClasifies.ToDictionary<string, string, decimal>(classify => classify, classify => 0);
                Dictionary<int, string> dictionaryClassify = new Dictionary<int, string>();
                int currentCol = 4;
                foreach (string classify in _prizeClasifies)
                {
                    ws.Cell(2, currentCol).Value = classify;
                    dictionaryClassify.Add(currentCol, classify);
                    currentCol++;
                }
                ws.Cell(2, currentCol).Value = "合计";
                ws.Cell(2, currentCol).Style.Font.Bold = true;

                //填写个人数据
                DigitToChnText dtc = new DigitToChnText();
                int departmentNo = 1;
                int currentRow = 3;
                foreach (string department in _departmets)
                {
                    var departmentStartrow = currentRow;
                    int currentEmployeeNo = 1;
                    string currentClassify = "";
                    decimal totalDepartmentMoney = 0;
                    ws.Cell(currentRow, 1).Value = dtc.Convert(departmentNo.ToString(), false);
                    ws.Cell(currentRow, 1).Style.Fill.BackgroundColor = XLColor.Orange;
                    ws.Cell(currentRow, 2).Value = department;
                    ws.Cell(currentRow, 2).Style.Alignment.WrapText = true;
                    ws.Cell(currentRow, 3).Value = "小计";
                    ws.Cell(currentRow, 3).Style.Fill.BackgroundColor = XLColor.Orange;
                    for (currentCol = 4; currentCol < _totalPrizeClassify + 4; currentCol++)
                    {
                        currentClassify = dictionaryClassify[currentCol];
                        var departmentMoney =
                            _prizeses.Where(
                                p => p.Department.Equals(department) && p.PrizeClassify.Equals(currentClassify))
                                .Sum(p => p.Prize) ?? 0;
                        totalDepartmentMoney += departmentMoney;
                        dictionaryClassifyMoney[dictionaryClassify[currentCol]] += departmentMoney;
                        ws.Cell(currentRow, currentCol).Value = departmentMoney == 0
                            ? "0"
                            : departmentMoney.ToString("N0");
                        ws.Cell(currentRow, currentCol).Style.Fill.BackgroundColor = XLColor.Orange;
                    }
                    //部门合计
                    ws.Cell(currentRow, currentCol).Value = totalDepartmentMoney == 0
                        ? "0"
                        : totalDepartmentMoney.ToString("N0");
                    ws.Cell(currentRow, currentCol).Style.Fill.BackgroundColor = XLColor.Orange;
                    currentRow++;
                    ws.Cell(currentRow, currentCol).Style.Font.Bold = true;

                    //部门员工汇总
                    var employees =
                        _prizeses.Where(p => p.Department.Equals(department))
                            .Select(
                                p => new { Name = p.Name, AccountName = p.AccountName, Department = p.Department })
                            .Distinct()
                            .ToList();
                    Dictionary<string, decimal> dictionaryEmSum = new Dictionary<string, decimal>();

                    foreach (var E in employees)
                    {
                        var thisMoney = _prizeses.Where(p => p.AccountName.Equals(E.AccountName) && p.Name.Equals(E.Name) && p.Department.Equals(E.Department)).Sum(p => p.Prize) ?? 0;
                        if (!string.IsNullOrWhiteSpace(E.AccountName))
                        {
                            dictionaryEmSum.Add(E.AccountName, thisMoney);
                        }
                        else
                        {
                            dictionaryEmSum.Add(E.Name, thisMoney);
                        }
                    }

                    dictionaryEmSum = dictionaryEmSum.OrderByDescending(E => E.Value).ToDictionary(E => E.Key, p => p.Value);

                    foreach (KeyValuePair<string, decimal> em in dictionaryEmSum)
                    {
                        ws.Cell(currentRow, 1).Value = currentEmployeeNo.ToString();
                        var employee = employees.FirstOrDefault(e => e.AccountName.Equals(em.Key) || e.Name.Equals(em.Key));
                        ws.Cell(currentRow, 3).Value = employee == null ? "" : employee.Name;
                        for (currentCol = 4; currentCol < _totalPrizeClassify + 4; currentCol++)
                        {
                            var employeeClassifyMoney = _prizeses.Where(
                                p =>
                                    p.AccountName.Equals(em.Key) &&
                                    p.PrizeClassify.Equals(dictionaryClassify[currentCol])).Sum(p => p.Prize) ?? 0;

                            ws.Cell(currentRow, currentCol).Value = employeeClassifyMoney == 0
                                ? "0"
                                : employeeClassifyMoney.ToString("N0");
                        }
                        //合计栏
                        ws.Cell(currentRow, currentCol).Value = em.Value == 0
                            ? "0"
                            : em.Value.ToString("N0");
                        ws.Cell(currentRow, currentCol).Style.Font.Bold = true;
                        currentRow++;
                        currentEmployeeNo++;
                    }


                    //合并部门名称列
                    var departmentRows = ws.Range(departmentStartrow, 2, currentRow - 1, 2);
                    departmentRows.Merge();
                    departmentNo++;
                }
                //总计行
                ws.Cell(currentRow, 2).Value = "合计";
                ws.Range(currentRow, 2, currentRow, 3).Merge();
                for (currentCol = 4; currentCol < _totalPrizeClassify + 4; currentCol++)
                {
                    ws.Cell(currentRow, currentCol).Value = dictionaryClassifyMoney[dictionaryClassify[currentCol]] == 0
                        ? "0"
                        : dictionaryClassifyMoney[dictionaryClassify[currentCol]].ToString("N0");
                }
                ws.Cell(currentRow, currentCol).Value = _totalMoney == 0 ? "0" : _totalMoney.ToString("N0");
                ws.Range(currentRow, 1, currentRow, currentCol).Style.Fill.BackgroundColor = XLColor.DarkOrange;
                ws.Range(currentRow, 1, currentRow, currentCol).Style.Font.Bold = true;
                var mainBody = ws.Range(2, 1, currentRow, currentCol);
                mainBody.Style.Font.FontSize = 10;
                mainBody.Style.Font.FontName = "宋体";
                mainBody.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                mainBody.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                mainBody.Style.Alignment.WrapText = false;
                mainBody.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
                mainBody.Style.Border.OutsideBorderColor = XLColor.Black;
                mainBody.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                mainBody.Style.Border.InsideBorderColor = XLColor.Black;
                currentRow++;
                ws.Cell(currentRow, 1).Value = "编制：";
                ws.Cell(currentRow, 4).Value = "复核：";
                ws.Cell(currentRow, 7).Value = "审核：";
                ws.Cell(currentRow, currentCol).Value = DateTime.Now.ToString("0:yyyy.MM.dd",
                    DateTimeFormatInfo.InvariantInfo);
                var rowLast = ws.Range(currentRow, 1, currentRow, currentCol);
                rowLast.Style.Font.FontName = "宋体";
                rowLast.Style.Font.FontSize = 10;
                rowLast.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                rowLast.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                ws.Columns().AdjustToContents();
                ws.Column(2).Width = 30;
                ws.Rows().AdjustToContents();
                ws.Rows().Height = 15;
                ws.Row(1).Height = 20;
                //MessageBox.Show("生成个人汇总表成功！");
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
        }

        /// <summary>
        /// 生成各部门汇总表
        /// </summary>
        /// <param name="document"></param>
        private void GenaratePerDepartmentTotalSheet(XLWorkbook document)
        {
            foreach (string department in _departmets)
            {
                CreateDepartmentSheet(document, department);
            }
        }

        private void CreateDepartmentSheet(XLWorkbook document, string department)
        {
            try
            {
                var ws = document.Worksheets.Add(department);
                ws.Cell(1, 1).Value = "HNCDI(" + _year + "年度)项目奖励明细";
                //奖项列别列表
                List<string> departmentClassifies =
                    _prizeses.Where(p => p.Department.Equals(department))
                        .Select(p => p.PrizeClassify)
                        .Distinct()
                        .ToList();
                //奖项列表
                var awards =
                    _prizeses.Where(p => p.Department.Equals(department))
                        .Select(
                            p => new { Classify = p.PrizeClassify, Project = p.Project, AwardName = p.AwardName })
                        .Distinct()
                        .ToList();
                var title = ws.Range(1, 1, 1, 2 + awards.Count() + departmentClassifies.Count());
                int lastestCol = 2 + awards.Count() + departmentClassifies.Count();
                for (int i = 1; i <= lastestCol; i++)
                {
                    ws.Column(i).Width = 9;
                }
                title.Merge();
                title.Style.Font.FontSize = 14;
                title.Style.Font.FontName = "宋体";
                title.Style.Font.Bold = true;
                title.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                title.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                Dictionary<int, string> dictionaryClassify = new Dictionary<int, string>();
                Dictionary<int, string> dictionaryProject = new Dictionary<int, string>();
                Dictionary<int, string> dictionaryAward = new Dictionary<int, string>();
                //员工列表
                var employees =
                    _prizeses.Where(p => p.Department.Equals(department))
                        .Select(
                            p => new { Name = p.Name, AccountName = p.AccountName, Department = p.Department })
                        .Distinct()
                        .ToList();
                //空一行
                //奖项类别行
                ws.Cell(3, 1).Value = "姓名";
                ws.Range(3, 1, 4, 1).Merge();
                int currentCol = 3;
                foreach (string classify in departmentClassifies)
                {
                    ws.Cell(3, currentCol).Value = classify;
                    int currentClassifyCols = awards.Count(p => p.Classify.Equals(classify));
                    int lastCol = currentCol + currentClassifyCols;
                    ws.Range(3, currentCol, 3, currentCol + currentClassifyCols).Merge();
                    for (; currentCol <= lastCol; currentCol++)
                    {
                        dictionaryClassify.Add(currentCol, classify);
                    }
                }
                //项目行
                ws.Cell(4, 2).Value = "总计";
                ws.Range(4, 2, 5, 2).Merge();
                ws.Range(4, 2, 5, 2).Style.Fill.BackgroundColor = XLColor.ForestGreen;

                for (currentCol = 3; currentCol <= lastestCol; currentCol++)
                {
                    if (currentCol == 3)
                    {
                        ws.Cell(4, currentCol).Value = "小计";
                        ws.Range(4, currentCol, 5, currentCol).Merge();
                        ws.Range(4, currentCol, 5, currentCol).Style.Fill.BackgroundColor = XLColor.Yellow;
                        continue;
                    }
                    else
                    {
                        if (dictionaryClassify[currentCol] != dictionaryClassify[currentCol - 1])
                        {
                            ws.Cell(4, currentCol).Value = "小计";
                            ws.Range(4, currentCol, 5, currentCol).Merge();
                            ws.Range(4, currentCol, 5, currentCol).Style.Fill.BackgroundColor = XLColor.Yellow;
                            continue;
                        }
                    }
                    if (dictionaryClassify[currentCol] == "国家、省部级优秀勘察、设计、咨询奖")
                    {
                        string test = "";
                    }
                    List<string> projects =
                        awards.Where(p => p.Classify.Equals(dictionaryClassify[currentCol]))
                            .Select(p => p.Project)
                            .Distinct()
                            .ToList();
                    foreach (string project in projects)
                    {
                        ws.Cell(4, currentCol).Value = project;
                        ws.Cell(4, currentCol).Style.Alignment.WrapText = true;
                        if (awards.Count(p => p.Classify.Equals(dictionaryClassify[currentCol]) && p.Project.Equals(project)) > 1)
                        {
                            int projectCols = awards.Count(p => p.Classify.Equals(dictionaryClassify[currentCol]) && p.Project.Equals(project));
                            ws.Range(4, currentCol, 4, currentCol + projectCols - 1).Merge();
                            for (int i = 1; i <= projectCols; i++)
                            {
                                dictionaryProject.Add(currentCol, project);
                                currentCol++;
                            }
                        }
                        else
                        {
                            dictionaryProject.Add(currentCol, project);
                            currentCol++;
                        }
                    }
                    currentCol--;
                }
                //奖项行
                for (currentCol = 4; currentCol <= lastestCol; currentCol++)
                {
                    if (currentCol > 4 && dictionaryClassify[currentCol] != dictionaryClassify[currentCol - 1])
                    {
                        continue;
                    }
                    var col = currentCol;
                    var projectAwards =
                        awards.Where(
                            p =>
                                p.Classify.Equals(dictionaryClassify[col]) &&
                                p.Project.Equals(dictionaryProject[col]));
                    foreach (var award in projectAwards)
                    {
                        ws.Cell(5, currentCol).Value = award.AwardName;
                        ws.Cell(5, currentCol).Style.Alignment.WrapText = true;
                        dictionaryAward.Add(currentCol, award.AwardName);
                        currentCol++;
                    }
                    currentCol--;
                }

                //填写个人数据
                int currentRow = 6;

                foreach (var employee in employees)
                {
                    ws.Cell(currentRow, 1).Value = employee.Name;
                    //总计
                    List<Prizes> employeePrizeses =
                        _prizeses.Where(p => p.AccountName.Equals(employee.AccountName) && p.Name.Equals(employee.Name)).ToList();
                    decimal employeeMoney = employeePrizeses.Sum(p => p.Prize) ?? 0;
                    ws.Cell(currentRow, 2).Value = employeeMoney.ToString("N0");
                    for (currentCol = 3; currentCol <= lastestCol; currentCol++)
                    {
                        if (currentCol == 3)
                        {
                            employeeMoney =
                                employeePrizeses.Where(p => p.PrizeClassify.Equals(dictionaryClassify[currentCol]))
                                    .Sum(p => p.Prize) ?? 0;
                            ws.Cell(currentRow, currentCol).Value = employeeMoney == 0
                                ? "0"
                                : employeeMoney.ToString("N0");
                            ws.Cell(currentRow, currentCol).Style.Fill.BackgroundColor = XLColor.Yellow;
                            continue;
                        }
                        else
                        {
                            if (dictionaryClassify[currentCol] != dictionaryClassify[currentCol - 1])
                            {
                                employeeMoney =
                                    employeePrizeses.Where(p => p.PrizeClassify.Equals(dictionaryClassify[currentCol]))
                                        .Sum(p => p.Prize) ?? 0;
                                ws.Cell(currentRow, currentCol).Value = employeeMoney == 0
                                    ? "0"
                                    : employeeMoney.ToString("N0");
                                ws.Cell(currentRow, currentCol).Style.Fill.BackgroundColor = XLColor.Yellow;
                                continue;
                            }
                        }
                        Prizes prize =
                            employeePrizeses.FirstOrDefault(
                                p =>
                                    p.PrizeClassify.Equals(dictionaryClassify[currentCol]) &&
                                    p.Project.Equals(dictionaryProject[currentCol]) &&
                                    p.AwardName.Equals(dictionaryAward[currentCol]));
                        if (prize?.Prize == null) continue;
                        ws.Cell(currentRow, currentCol).Value = (Convert.ToDecimal(prize.Prize)).ToString("N0");
                    }
                    currentRow++;
                }
                //合计行
                ws.Cell(currentRow, 1).Value = "合计";
                decimal departmentMoney = _prizeses.Where(p => p.Department.Equals(department)).Sum(p => p.Prize) ?? 0;
                ws.Cell(currentRow, 2).Value = departmentMoney == 0 ? "0" : departmentMoney.ToString("N0");
                for (currentCol = 3; currentCol <= lastestCol; currentCol++)
                {
                    if (currentCol == 3)
                    {
                        departmentMoney =
                            _prizeses.Where(
                                p =>
                                    p.Department.Equals(department) &&
                                    p.PrizeClassify.Equals(dictionaryClassify[currentCol])).Sum(p => p.Prize) ?? 0;
                        ws.Cell(currentRow, currentCol).Value = departmentMoney == 0
                            ? "0"
                            : departmentMoney.ToString("N0");
                        continue;
                    }
                    else
                    {
                        if (dictionaryClassify[currentCol] != dictionaryClassify[currentCol - 1])
                        {
                            departmentMoney =
                                _prizeses.Where(
                                    p =>
                                        p.Department.Equals(department) &&
                                        p.PrizeClassify.Equals(dictionaryClassify[currentCol])).Sum(p => p.Prize) ?? 0;
                            ws.Cell(currentRow, currentCol).Value = departmentMoney == 0
                                ? "0"
                                : departmentMoney.ToString("N0");
                            continue;
                        }
                    }
                    departmentMoney =
                        _prizeses.Where(
                            p =>
                                p.Department.Equals(department) &&
                                p.PrizeClassify.Equals(dictionaryClassify[currentCol]) &&
                                p.Project.Equals(dictionaryProject[currentCol]) &&
                                p.AwardName.Equals(dictionaryAward[currentCol])).Sum(p => p.Prize) ?? 0;
                    ws.Cell(currentRow, currentCol).Value = departmentMoney == 0 ? "0" : departmentMoney.ToString("N0");
                }
                var lastestRow = ws.Range(currentRow, 2, currentRow, lastestCol);
                lastestRow.Style.Fill.BackgroundColor = XLColor.ForestGreen;
                ws.Range(4, 2, currentRow, 2).Style.Fill.BackgroundColor = XLColor.ForestGreen;
                ws.Range(currentRow, 3, currentRow, lastestCol).Style.Fill.BackgroundColor = XLColor.ForestGreen;
                var mainbody = ws.Range(3, 1, currentRow, lastestCol);
                mainbody.Style.Font.FontSize = 10;
                mainbody.Style.Font.FontName = "宋体";
                mainbody.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                mainbody.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                mainbody.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
                mainbody.Style.Border.OutsideBorderColor = XLColor.Black;
                mainbody.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                mainbody.Style.Border.InsideBorderColor = XLColor.Black;
                ws.Columns().Width = 10;
                ws.Rows().AdjustToContents();
                ws.Rows().Height = 15;
                ws.Row(1).Height = 20;
                ws.Row(4).Height = 80;
                ws.Row(5).Height = 80;
                //MessageBox.Show("导出部门 " + department + " 数据成功！");
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
        }

        private void ButtonRun_GotFocus(object sender, RoutedEventArgs e)
        {
            ButtonRun.MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
        }

        private void ButtonCancel_GotFocus(object sender, RoutedEventArgs e)
        {
            ButtonCancel.MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
        }

        private void ButtonRun_Click(object sender, RoutedEventArgs e)
        {
            if (dataContext == null) dataContext = new DataClassesProjectClassifyDataContext();

            _year = TextBoxYear.Text.Trim();
            //计算总金额
            _totalMoney = dataContext.Prizes.Where(p => p.PayYear.Equals(_year)).Sum(p => p.Prize) ?? 0;
            //计算部门数
            _prizeses = new List<Prizes>();
            _departmets = new List<string>();
            _prizeses = dataContext.Prizes.Where(P => P.PayYear.Equals(_year)).ToList();
            _departmets = _prizeses.Select(p => p.Department).Distinct().ToList();
            _totalDepartment = _departmets.Count();

            //获取申报部门列表及总数
            _declareDepartments = new List<string>();
            _declareDepartments = _prizeses.Select(p => p.DeclareDepartment).Distinct().ToList();
            _totalDeclareDepartment = _declareDepartments.Count();

            //获取员工列表
            _employees = new List<Employee>();
            _employees = _prizeses.Select(p => new Employee { AccountName = p.AccountName, Name = p.Name, Department = p.Department }).Distinct().ToList();
            _totalEmployee = _employees.Count();

            //获取奖金类别
            _prizeClasifies = new List<string>();
            _prizeClasifies = _prizeses.Select(p => p.PrizeClassify).Distinct().ToList();
            _totalPrizeClassify = _prizeClasifies.Count();

            //获取项目列表
            _projects = new List<string>();
            _projects = _prizeses.Select(p => p.Project).Distinct().ToList();
            _totalProjects = _projects.Count();

            Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog();
            saveFileDialog.Filter = "Execl files (*.xlsx)|*.xlsx";
            saveFileDialog.FilterIndex = 0;
            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.FileName = _year + "年度获奖成果及论著奖励汇总表";
            saveFileDialog.CreatePrompt = true;
            saveFileDialog.Title = "要创建Excel文件";
            saveFileDialog.ShowDialog();
            if (saveFileDialog.FileName == "")
            {
                MessageBox.Show("错误", "请选择文件或输入文件名", MessageBoxButton.OK);
                return;
            }
            FileInfo fileToCreate = new FileInfo(saveFileDialog.FileName);
            if (fileToCreate.Exists)
            {
                try
                {
                    fileToCreate.Delete();
                }
                catch (Exception error)
                {
                    MessageBox.Show(error.Message, "删除失败", MessageBoxButton.OK);
                    return;
                }
            }
            try
            {
                CreatePackage(saveFileDialog.FileName);
                MessageBox.Show("导出成功！");
                this.Close();
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message, "导出失败", MessageBoxButton.OK);
                this.Close();
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            dataContext = new DataClassesProjectClassifyDataContext();
        }

        private void CheckEmpty()
        {
            ButtonRun.IsEnabled = false;
            if (string.IsNullOrWhiteSpace(TextBoxYear.Text)) return;
            if (TextBoxYear.Text.Trim().Length != 4) return;
            Regex reg1 = new Regex(@"^[0-9]\d*$");
            if (!reg1.IsMatch(TextBoxYear.Text.Trim())) return;
            ButtonRun.IsEnabled = true;
        }

        private void TextBoxYear_LostFocus(object sender, RoutedEventArgs e)
        {
            CheckEmpty();
        }

        private void TextBoxYear_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Enter) return;
            TextBoxYear.MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
            CheckEmpty();
        }

        private void TextBoxYear_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBoxYear.MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
            if (!String.IsNullOrWhiteSpace(TextBoxYear.Text))
            {
                if (TextBoxYear.Text.Trim().Length == 4)
                {
                    CheckEmpty();
                }
            }
        }
    }
}
