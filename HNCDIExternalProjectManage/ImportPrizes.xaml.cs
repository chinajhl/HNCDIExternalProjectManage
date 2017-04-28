using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
using Microsoft.Win32;

namespace HNCDIExternalProjectManage
{
    /// <summary>
    /// 导入部门年度奖金明细表
    /// ImportPrizes.xaml 的交互逻辑
    /// </summary>
    public partial class ImportPrizes : Window
    {
        private DataClassesProjectClassifyDataContext dataContext;
        private XLWorkbook _wb;
        private string _filename;
        private string _year;
        private string _declareDepartment; //申报部门
        private int _lastClassifyColNo, _lastProjectColNo, _lastCategoryColNo;
        private List<string> _listPrizeClassify;
        private Dictionary<int, string> _dictionaryPrizeClassify;
        private List<string> _listProject;
        private Dictionary<int, string> _dictionaryProject;
        private Dictionary<int, string> _dictionaryCategory;
        private int _totalRows;
        private DomainOperate _doo;
        public Employee CurrentEmployee { private get; set; }
        public string CurrentEmployeeName { get; set; }
        private string CurrentEmployeeAccountName { get; set; }
        private string CurrentEmployeeDepartment { get; set; }
        private List<Prizes> _listPrizeses;

        public ImportPrizes()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {

        }

        private void ButtonBrowser_GotFocus(object sender, RoutedEventArgs e)
        {
            //ButtonBrowser.MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
        }

        private void ButtonImport_GotFocus(object sender, RoutedEventArgs e)
        {
            ButtonImport.MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
        }

        private void ButtonCancel_GotFocus(object sender, RoutedEventArgs e)
        {
            ButtonCancel.MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
        }

        private void TextBoxSourceFile_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TextBoxSourceFile.Text))
            {
                _filename = "";
                ButtonImport.IsEnabled = false;
            }
            else
            {
                _filename = TextBoxSourceFile.Text.Trim();
                ButtonImport.IsEnabled = true;
            }
        }

        private void ButtonBrowser_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel文件 |*.xls;*.xlsx";
            openFileDialog.Multiselect = false;
            openFileDialog.Title = "请选择源文件";
            var result = openFileDialog.ShowDialog() ?? false;
            if (result)
            {
                TextBoxSourceFile.Text = openFileDialog.FileName;
            }
        }

        private void ButtonImport_Click(object sender, RoutedEventArgs e)
        {
            dataContext = new DataClassesProjectClassifyDataContext();
            try
            {
                _wb = new XLWorkbook(_filename);
                if (!GetYear())
                {
                    MessageBox.Show("标题行中没有包含 年度 描述，或者报表为空！", "错误");
                    return;
                }
                if (!GetDeclareDepartment())
                {
                    MessageBox.Show("标题行中没有包含 申报部门 描述，或者报表为空！", "错误");
                    return;
                }
                GetPrizeClassifies();

                GetProjects();

                GetPrizeCategories();

                ReadData();

                MessageBox.Show("导入数据成功！");
            }
            catch (Exception error)
            {
                MessageBox.Show("导入数据失败！\n" + error.Message, "错误");
            }
        }

        /// <summary>
        /// 获取奖金年度
        /// </summary>
        private bool GetYear()
        {
            var ws = _wb.Worksheets.FirstOrDefault();
            if (ws != null)
            {
                var firstRow = ws.FirstRowUsed();
                var title = firstRow.RowUsed();
                int colNo = 1;
                if (title.Cell(colNo).IsEmpty())
                {
                    colNo++;
                }
                string caption = title.Cell(colNo).GetString();
                if (caption != null && caption.IndexOf("年度", StringComparison.Ordinal) > 4)
                {
                    _year = caption.Substring(caption.IndexOf("年度", StringComparison.Ordinal) - 4, 4);
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// 获取申报部门
        /// </summary>
        /// <returns></returns>
        private bool GetDeclareDepartment()
        {
            var ws = _wb.Worksheets.FirstOrDefault();
            if (ws != null)
            {
                var firstRow = ws.FirstRowUsed();
                var title = firstRow.RowUsed();
                int colNo = 1;
                if (title.Cell(colNo).IsEmpty())
                {
                    colNo++;
                }
                string caption = title.Cell(colNo).GetString();
                if (caption != null && caption.IndexOf("申报部门", StringComparison.Ordinal) > 0)
                {
                    _declareDepartment = caption.Substring(caption.IndexOf("申报部门", StringComparison.Ordinal));
                    _declareDepartment = _declareDepartment.Replace("申报部门", "");
                    _declareDepartment = _declareDepartment.Replace("：", "");
                    _declareDepartment = _declareDepartment.Replace(":", "");
                    _declareDepartment = _declareDepartment.Replace("）", "");
                    _declareDepartment = _declareDepartment.Replace(")", "");
                    if (!string.IsNullOrWhiteSpace(_declareDepartment))
                    {
                        return true;
                    }
                }
                return false;
            }
            return false;
        }

        /// <summary>
        /// 获取奖金类别列表
        /// </summary>
        private void GetPrizeClassifies()
        {
            var ws = _wb.Worksheets.FirstOrDefault();
            if (ws != null)
            {
                var firstRow = ws.FirstRowUsed();
                firstRow = firstRow.RowBelow();
                var prizeClassifies = firstRow.RowUsed();
                int maxColNo = prizeClassifies.CellCount();
                _lastClassifyColNo = maxColNo;
                int perClassifyCol = 1; //每个奖项类别所占列数
                int currentCol = 4;
                string prizeClassify = "";
                _dictionaryPrizeClassify = new Dictionary<int, string>();
                _listPrizeClassify = new List<string>();
                while (true)
                {
                    if (prizeClassifies.Cell(currentCol).IsEmpty())
                    {
                        if (!string.IsNullOrWhiteSpace(prizeClassify))
                        {
                            _dictionaryPrizeClassify.Add(currentCol, prizeClassify);
                        }
                        currentCol++;
                        perClassifyCol++;
                        continue;
                    }
                    else
                    {
                        prizeClassify = prizeClassifies.Cell(currentCol).GetString();
                        _dictionaryPrizeClassify.Add(currentCol, prizeClassify);
                        if (!_listPrizeClassify.Contains(prizeClassify))
                        {
                            _listPrizeClassify.Add(prizeClassify);
                        }
                        perClassifyCol = 1;
                        if (currentCol == maxColNo)
                        {
                            break;
                        }
                        currentCol++;
                    }
                }
            }
        }

        /// <summary>
        /// 获取项目列表
        /// </summary>
        private void GetProjects()
        {
            var ws = _wb.Worksheets.FirstOrDefault();
            if (ws != null)
            {
                var firstRow = ws.FirstRowUsed();
                firstRow = firstRow.RowBelow();
                firstRow = firstRow.RowBelow();
                var projectRow = firstRow.RowUsed();
                int maxColNo = projectRow.CellCount();
                _lastProjectColNo = maxColNo;

                int perProjectCols = 1;
                int currentCol = 4;
                string project = "";
                _listProject = new List<string>();
                _dictionaryProject = new Dictionary<int, string>();

                while (currentCol <= maxColNo)
                {
                    if (projectRow.Cell(currentCol).IsEmpty())
                    {
                        if (currentCol <= _lastClassifyColNo)
                        {
                            if (currentCol == 4)
                            {
                                _dictionaryProject.Add(4, "");
                                currentCol++;
                                continue;
                            }
                            if (_dictionaryPrizeClassify[currentCol] == _dictionaryPrizeClassify[currentCol - 1])
                            {
                                if (!string.IsNullOrWhiteSpace(project))
                                {
                                    _dictionaryProject.Add(currentCol, project);
                                    perProjectCols++;
                                }
                            }
                            if (string.IsNullOrWhiteSpace(project))
                            {
                                _dictionaryProject.Add(currentCol, "");
                            }
                        }
                        else
                        {
                            _dictionaryPrizeClassify.Add(currentCol, _dictionaryPrizeClassify[_lastClassifyColNo]);
                            _lastClassifyColNo++;
                            _dictionaryProject.Add(currentCol, project);
                            perProjectCols++;
                        }
                        currentCol++;
                    }
                    else
                    {
                        string temp = projectRow.Cell(currentCol).GetString();
                        temp = temp.Replace(" ", "");
                        if (temp.Contains("小计"))
                        {
                            project = "";
                            _dictionaryProject.Add(currentCol, "");
                            currentCol++;
                        }
                        else
                        {
                            project = temp;
                            _dictionaryProject.Add(currentCol, project);
                            if (currentCol > _lastClassifyColNo)
                            {
                                _dictionaryPrizeClassify.Add(currentCol, _dictionaryPrizeClassify[_lastClassifyColNo]);
                                _lastClassifyColNo++;
                            }
                            perProjectCols = 1;
                            currentCol++;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 读取奖项
        /// </summary>
        private void GetPrizeCategories()
        {
            var ws = _wb.Worksheets.FirstOrDefault();
            if (ws != null)
            {
                var firstRow = ws.FirstRowUsed();
                firstRow = firstRow.RowBelow();
                firstRow = firstRow.RowBelow();
                firstRow = firstRow.RowBelow();
                var categoryRow = firstRow.RowUsed();
                int maxColNo = categoryRow.CellCount();
                _lastCategoryColNo = maxColNo;

                int currentCol = 4;
                string category = "";
                //int perCategoryCols = 1;
                _dictionaryCategory = new Dictionary<int, string>();

                while (currentCol <= maxColNo)
                {
                    if (categoryRow.Cell(currentCol).IsEmpty())
                    {
                        if (currentCol <= _lastProjectColNo)
                        {
                            _dictionaryCategory.Add(currentCol, "");
                        }
                        else
                        {
                            _dictionaryPrizeClassify.Add(currentCol, _dictionaryPrizeClassify[_lastClassifyColNo]);
                            _lastClassifyColNo++;
                            _dictionaryProject.Add(currentCol, _dictionaryProject[_lastProjectColNo]);
                            _lastProjectColNo++;
                            _dictionaryCategory.Add(currentCol, "");
                        }
                        currentCol++;
                    }
                    else
                    {
                        string temp = categoryRow.Cell(currentCol).GetString();
                        temp = temp.Replace(" ", "");
                        if (temp.Contains("小计"))
                        {
                            _dictionaryCategory.Add(currentCol, "");
                            currentCol++;
                        }
                        else
                        {
                            category = temp;
                            _dictionaryCategory.Add(currentCol, category);
                            if (currentCol > _lastProjectColNo)
                            {
                                _dictionaryPrizeClassify.Add(currentCol, _dictionaryPrizeClassify[_lastClassifyColNo]);
                                _lastClassifyColNo++;
                                _dictionaryProject.Add(currentCol, _dictionaryProject[_lastProjectColNo]);
                                _lastProjectColNo++;
                            }
                            //perCategoryCols = 1;
                            currentCol++;
                        }
                    }
                }
                int lastCol = _dictionaryPrizeClassify.Count() > _dictionaryProject.Count()
                    ? 3 + _dictionaryPrizeClassify.Count()
                    : 3 + _dictionaryProject.Count();

                if (maxColNo < lastCol)
                {
                    maxColNo++;
                    while (maxColNo <= lastCol)
                    {
                        _dictionaryCategory.Add(maxColNo, "");
                        maxColNo++;
                    }
                    _lastCategoryColNo = lastCol;
                }
            }
        }

        private bool IsNumber(string str)
        {
            if (string.IsNullOrEmpty(str))    //验证这个参数是否为空  
                return false;                           //是，就返回False  
            ASCIIEncoding ascii = new ASCIIEncoding();//new ASCIIEncoding 的实例  
            byte[] bytestr = ascii.GetBytes(str);         //把string类型的参数保存到数组里  

            foreach (byte c in bytestr)                   //遍历这个数组里的内容  
            {
                if (c < 48 || c > 57)                          //判断是否为数字  
                {
                    return false;                              //不是，就返回False  
                }
            }
            return true;                                        //是，就返回True  
        }

        /// <summary>
        /// 读取奖金数据
        /// </summary>
        private void ReadData()
        {
            var ws = _wb.Worksheets.FirstOrDefault();
            if (ws != null)
            {
                List<string> logins = new List<string>();
                _doo = new DomainOperate("hncdi");
                _totalRows = ws.Rows().Count();
                int currentRow = 5;
                _listPrizeses = new List<Prizes>();
                while (currentRow <= _totalRows)
                {
                    var row = ws.Row(currentRow);
                    if (row.Cell(1).IsEmpty())
                    {
                        currentRow++;
                        continue;
                    }
                    string temp = row.Cell(1).GetString();
                    temp = temp.Replace(" ", "").Trim();
                    if (temp == "合计" || temp == "小计" || temp == "总计" || temp == "共计") break;
                    //if (row.Cell(2).IsEmpty() && row.Cell(1).GetString().Contains("计")) break;
                    var dataRow = row.RowUsed();
                    int totalCol = dataRow.CellCount();
                    int currentCol = 1;
                    if (dataRow.Cell(currentCol).IsEmpty())
                    {
                        currentRow++;
                        continue;
                    }
                    string userName = dataRow.Cell(currentCol).GetString();
                    userName = userName.Replace(" ", "").Trim();
                    if (IsNumber(userName))
                    {
                        currentRow++;
                        continue;
                    }
                    GetUserAccount(userName, out logins);

                    if (!logins.Any())
                    {
                        CurrentEmployee = new Employee
                        {
                            Name = userName,
                            AccountName = "",
                            Department = _declareDepartment
                        };
                    }
                    else
                    {
                        if (logins.Count() > 1)
                        {
                            SelectPrizeEmployee selectPrizeEmployee = new SelectPrizeEmployee();
                            selectPrizeEmployee.EmployeeName = userName;
                            selectPrizeEmployee.ShowDialog();
                            CurrentEmployeeName = selectPrizeEmployee.EmployeeName;
                            CurrentEmployeeAccountName = selectPrizeEmployee.SelectedAccountName;
                            CurrentEmployeeDepartment = selectPrizeEmployee.SelectedDepartment;
                            if (string.IsNullOrWhiteSpace(CurrentEmployeeName))
                            {
                                CurrentEmployee = null;
                            }
                            else
                            {
                                CurrentEmployee = new Employee
                                {
                                    Name = CurrentEmployeeName,
                                    AccountName = CurrentEmployeeAccountName,
                                    Department = CurrentEmployeeDepartment
                                };
                            }
                        }
                        else
                        {
                            CurrentEmployee = new Employee
                            {
                                Name = userName,
                                AccountName = logins[0],
                                Department = _doo.GetOuByLoginID(logins[0])
                            };
                        }
                    }
                    if (CurrentEmployee == null)
                    {
                        currentRow++;
                        continue;
                    }
                    currentCol = 4;
                    while (currentCol <= totalCol)
                    {
                        if (string.IsNullOrWhiteSpace(_dictionaryProject[currentCol]) || dataRow.Cell(currentCol).IsEmpty())
                        {
                            currentCol++;
                            continue;
                        }
                        try
                        {
                            decimal prize = Convert.ToDecimal(dataRow.Cell(currentCol).Value);
                            if (prize <= 0.0M)
                            {
                                currentCol++;
                                continue;
                            }
                            if (currentCol > _lastCategoryColNo)
                            {
                                _dictionaryCategory.Add(currentCol, "");
                                _lastCategoryColNo++;
                            }
                            if (currentCol > _lastProjectColNo)
                            {
                                _dictionaryProject.Add(currentCol, _dictionaryProject[_lastProjectColNo]);
                                _lastProjectColNo++;
                            }
                            if (currentCol > _lastClassifyColNo)
                            {
                                _dictionaryPrizeClassify.Add(currentCol, _dictionaryPrizeClassify[_lastClassifyColNo]);
                                _lastClassifyColNo++;
                            }

                            Prizes prizes = new Prizes
                            {
                                Department = CurrentEmployee.Department,
                                DeclareDepartment = _declareDepartment,
                                Name = CurrentEmployee.Name,
                                AccountName = CurrentEmployee.AccountName,
                                Project = _dictionaryProject[currentCol],
                                PrizeClassify = _dictionaryPrizeClassify[currentCol],
                                AwardName = _dictionaryCategory[currentCol],
                                Prize = prize,
                                PayYear = _year
                            };
                            //检查重复项
                            var pr =
                                dataContext.Prizes.FirstOrDefault(
                                    p =>
                                        p.Department.Equals(prizes.Department) &&
                                        //p.DeclareDepartment.Equals(prizes.DeclareDepartment) &&
                                        p.Name.Equals(prizes.Name) && p.AccountName.Equals(prizes.AccountName) &&
                                        p.Project.Equals(prizes.Project) && p.PrizeClassify.Equals(prizes.PrizeClassify) &&
                                        p.PayYear.Equals(prizes.PayYear) &&
                                        p.AwardName.Equals(prizes.AwardName));

                            if (pr != null)
                            {
                                if (pr.Prize != prizes.Prize)
                                {
                                    pr.Prize = prizes.Prize;
                                    dataContext.SubmitChanges();
                                }
                            }
                            else
                            {
                                dataContext.Prizes.InsertOnSubmit(prizes);
                                dataContext.SubmitChanges();
                            }
                            currentCol++;
                        }
                        catch (Exception error)
                        {
                            MessageBox.Show(error.Message + "当前行" + currentRow.ToString());
                            currentCol++;
                        }
                    }
                    currentRow++;
                }
            }
        }

        /// <summary>
        /// 获取用户账号
        /// </summary>
        /// <param name="username">用户姓名</param>
        /// <param name="useraccount">账号</param>
        /// <returns></returns>
        private bool GetUserAccount(string username, out List<string> useraccount)
        {
            useraccount = _doo.GetLoginIDByUserName(username);
            if (useraccount.Count > 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
