using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace HNCDIExternalProjectManage
{
    /// <summary>
    /// ContractManage.xaml 的交互逻辑
    /// </summary>
    public partial class ContractManage : Window
    {
        private DataClassesProjectClassifyDataContext dataContext;
        private int projectID;

        public int ProjectID
        {
            get { return projectID; }
            set { projectID = value; }
        }

        private string projectName;

        private int contractID;

        public ContractManage()
        {
            InitializeComponent();
        }

        private void buttonPreYear_Click(object sender, RoutedEventArgs e)
        {
            ContractDate.DisplayDate = ContractDate.DisplayDate.AddYears(-1);
            if (ContractDate.SelectedDate != null)
            {
                ContractDate.SelectedDate = ((DateTime)(ContractDate.SelectedDate)).AddYears(-1);
            }
        }

        private void buttonNextYear_Click(object sender, RoutedEventArgs e)
        {
            ContractDate.DisplayDate = ContractDate.DisplayDate.AddYears(1);
            if (ContractDate.SelectedDate != null)
            {
                ContractDate.SelectedDate = ((DateTime)(ContractDate.SelectedDate)).AddYears(1);
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            dataContext = new DataClassesProjectClassifyDataContext();
            listboxContractType.DisplayMemberPath = "ContractType";
            listboxContractType.SelectedValuePath = "TypeID";
            listboxContractType.DataContext = dataContext.ContractTypes;
            listboxContractType.SelectedIndex = 0;
            var p = dataContext.ProjectBase.Single(pb => pb.ProjectId.Equals(projectID));
            this.Title = p.ProjectName + "——合同管理";
            projectName = p.ProjectName;
            textboxProjectName.Text = projectName;
            textboxSecondParty.Text = "湖南省交通规划勘察设计院";
            datagridContracts.DataContext = dataContext.ProjectContracts.Where(pc => pc.ProjectID.Equals(projectID));
        }

        private void listboxContractType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (listboxContractType.SelectedIndex == 0)
            {
                textboxFirstParty.Text = "";
                textboxSecondParty.Text = "湖南省交通规划勘察设计院";
            }
            else
            {
                textboxFirstParty.Text = "湖南省交通规划勘察设计院";
                textboxSecondParty.Text = "";
            }
        }

        private void buttonSumbit_Click(object sender, RoutedEventArgs e)
        {
            if (listboxContractType.SelectedItem == null)
            {
                MessageBox.Show("请选择合同类型！", "错误");
                return;
            }
            if (ContractDate.SelectedDate == null)
            {
                MessageBox.Show("请选择签订日期！", "错误");
                return;
            }
            if (dataContext == null)
            {
                dataContext = new DataClassesProjectClassifyDataContext();
            }
            var pn = dataContext.ProjectContracts.Where(p => p.ContractNo.Trim().Equals(textboxContractNo.Text.Trim()) && p.ContractNo.Trim() != "");
            if (pn.Count() > 0)
            {
                MessageBox.Show("合同编号重复，已经录入该合同？", "错误");
                return;
            }
            ProjectContracts projectContract = new ProjectContracts();
            projectContract.ProjectID = projectID;
            projectContract.TypeID = ((ContractTypes)(listboxContractType.SelectedItem)).TypeID;
            projectContract.ContractNo = textboxContractNo.Text;
            try
            {
                projectContract.SumMoney = (Decimal)Double.Parse(textboxSumMoney.Text);
            }
            catch (FormatException)
            {
                MessageBox.Show("金额格式不对！", "错误");
                return;
            }
            projectContract.ProjectName = textboxProjectName.Text.Trim();
            projectContract.FirstParty = textboxFirstParty.Text.Trim();
            projectContract.SecondParty = textboxSecondParty.Text.Trim();
            projectContract.ContractPeriod = textboxContractPeriod.Text.Trim();
            projectContract.Principal = textboxPrincipal.Text.Trim();
            projectContract.ContractDate = ContractDate.SelectedDate;
            projectContract.Note = textboxNote.Text;

            dataContext.ProjectContracts.InsertOnSubmit(projectContract);
            dataContext.SubmitChanges();
            datagridContracts.DataContext = dataContext.ProjectContracts.Where(pc => pc.ProjectID.Equals(projectID));
            listboxContractType.DisplayMemberPath = "ContractType";
            listboxContractType.SelectedValuePath = "TypeID";
            listboxContractType.DataContext = dataContext.ContractTypes;
            listboxContractType.SelectedIndex = 0;
            ((MainWindow)(this.Owner)).DialogR = true;
        }

        private void datagridContracts_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (datagridContracts.SelectedItem != null)
            {
                ProjectContracts projectContract = (ProjectContracts)datagridContracts.SelectedItem;
                contractID = projectContract.ID;
                buttonUpdate.IsEnabled = true;
                buttonDelete.IsEnabled = true;                        
                if (dataContext == null)
                {
                    dataContext = new DataClassesProjectClassifyDataContext();
                }
                listboxContractType.SelectedItem = dataContext.ContractTypes.Single(ct => ct.TypeID.Equals(projectContract.TypeID));
                textboxContractNo.Text = projectContract.ContractNo;
                textboxProjectName.Text = projectContract.ProjectName;
                textboxFirstParty.Text = projectContract.FirstParty;
                textboxSecondParty.Text = projectContract.SecondParty;
                textboxContractPeriod.Text = projectContract.ContractPeriod;
                textboxPrincipal.Text = projectContract.Principal;
                textboxSumMoney.Text = projectContract.SumMoney.ToString();
                try
                {
                    ContractDate.DisplayDate = (DateTime)projectContract.ContractDate;
                    ContractDate.SelectedDate = projectContract.ContractDate;
                }
                catch (Exception)
                { }
                textboxNote.Text = projectContract.Note;
            }
        }

        private void buttonDelete_Click(object sender, RoutedEventArgs e)
        {
            if (contractID == 0)
            {
                MessageBox.Show("请选择合同项！", "错误");
                return;
            }
            if (MessageBox.Show("该项合同将被删除！确认要删除该项合同信息？", "警告", MessageBoxButton.YesNo) == MessageBoxResult.No)
            {
                return;
            }
            dataContext = new DataClassesProjectClassifyDataContext();
            var pc = dataContext.ProjectContracts.Single(p => p.ID.Equals(contractID));
            dataContext.ProjectContracts.DeleteOnSubmit(pc);
            dataContext.SubmitChanges();
            dataContext = new DataClassesProjectClassifyDataContext();
            datagridContracts.DataContext = dataContext.ProjectContracts.Where(p => p.ProjectID.Equals(projectID));
            ((MainWindow)(this.Owner)).DialogR = true;
            Clear();
        }

        private void Clear()
        {
            contractID = 0;
            listboxContractType.SelectedIndex = 0;
            textboxContractNo.Text = "";
            textboxFirstParty.Text = "";
            textboxSecondParty.Text = "湖南省交通规划勘察设计院";
            textboxProjectName.Text = projectName;
            textboxContractPeriod.Text = "";
            textboxPrincipal.Text = "";
            textboxSumMoney.Text = "";
            ContractDate.DisplayDate = DateTime.Now;
            ContractDate.SelectedDate = DateTime.Now;
            textboxNote.Text = "";
            buttonUpdate.IsEnabled = false;
            buttonDelete.IsEnabled = false;
        }

        private void SetBlackOutDate()
        {
            try
            {
                DateTime dt = ContractDate.DisplayDate;
                ContractDate.BlackoutDates.Clear();
                if (dt.Month == 1)
                {
                    ContractDate.BlackoutDates.Add(new CalendarDateRange(new DateTime(dt.Year - 1, 12, 15), new DateTime(dt.Year - 1, 12, 31)));
                }
                else
                {
                    DateTime s = new DateTime(dt.Year, dt.Month - 1, 15);
                    DateTime e = new DateTime(dt.Year, dt.Month - 1, DateTime.DaysInMonth(dt.Year, dt.Month - 1));
                    ContractDate.BlackoutDates.Add(new CalendarDateRange(s, e));
                }
                if (dt.Month == 12)
                {
                    ContractDate.BlackoutDates.Add(new CalendarDateRange(new DateTime(dt.Year + 1, 1, 1), new DateTime(dt.Year + 1, 1, 15)));
                }
                else
                {
                    DateTime s = new DateTime(dt.Year, dt.Month + 1, 1);
                    DateTime e = new DateTime(dt.Year, dt.Month + 1, 15);
                    ContractDate.BlackoutDates.Add(new CalendarDateRange(s, e));
                }
            }
            catch (Exception)
            {
            }
        }

        private void ContractDate_DisplayDateChanged(object sender, CalendarDateChangedEventArgs e)
        {
            SetBlackOutDate();
        }

        private void datagridContracts_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = e.Row.GetIndex() + 1;
        }

        private void ContractDate_Loaded(object sender, RoutedEventArgs e)
        {
            SetBlackOutDate();
        }

        private void buttonUpdate_Click(object sender, RoutedEventArgs e)
        {
            if (contractID == 0)
            {
                MessageBox.Show("请选择合同项！", "错误");
                return;
            }
            if (listboxContractType.SelectedItem == null)
            {
                MessageBox.Show("请选择合同类型！", "错误");
                return;
            }
            if (ContractDate.SelectedDate == null)
            {
                MessageBox.Show("请选择签订日期！", "错误");
                return;
            }
            if (dataContext == null)
            {
                dataContext = new DataClassesProjectClassifyDataContext();
            }
            int cid = ((ProjectContracts)(datagridContracts.SelectedItem)).ID;
            var pn = dataContext.ProjectContracts.Where(p => p.ContractNo.Trim().Equals(textboxContractNo.Text.Trim()) && p.ContractNo.Trim() != "" && p.ID != cid);
            if (pn.Any())
            {
                MessageBox.Show("合同编号重复！", "错误");
                return;
            }
            var projectContract = dataContext.ProjectContracts.Single(pc => pc.ID.Equals(contractID));
            projectContract.ProjectID = projectID;
            projectContract.TypeID = ((ContractTypes)(listboxContractType.SelectedItem)).TypeID;
            projectContract.ContractNo = textboxContractNo.Text;
            try
            {
                projectContract.SumMoney = (Decimal)Double.Parse(textboxSumMoney.Text);
            }
            catch (FormatException)
            {
                MessageBox.Show("金额格式不对！", "错误");
                return;
            }
            projectContract.ProjectName = textboxProjectName.Text.Trim();
            projectContract.FirstParty = textboxFirstParty.Text.Trim();
            projectContract.SecondParty = textboxSecondParty.Text.Trim();
            projectContract.ContractPeriod = textboxContractPeriod.Text.Trim();
            projectContract.Principal = textboxPrincipal.Text.Trim();
            projectContract.ContractDate = ContractDate.SelectedDate;
            projectContract.Note = textboxNote.Text;
            dataContext.SubmitChanges();
            datagridContracts.DataContext = dataContext.ProjectContracts.Where(pc => pc.ProjectID.Equals(projectID));
            listboxContractType.DisplayMemberPath = "ContractType";
            listboxContractType.SelectedValuePath = "TypeID";
            listboxContractType.DataContext = dataContext.ContractTypes;
            listboxContractType.SelectedIndex = 0;
            ((MainWindow)(this.Owner)).DialogR = true;
        }
    }
}