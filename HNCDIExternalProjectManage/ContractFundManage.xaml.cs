﻿using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace HNCDIExternalProjectManage
{
    /// <summary>
    ///     ContractFundManage.xaml 的交互逻辑
    /// </summary>
    public partial class ContractFundManage : Window
    {
        private int contractID;
        private DataClassesProjectClassifyDataContext dataContent;

        private int fundID;

        private decimal totalInComing, totalPayfor;

        public ContractFundManage()
        {
            InitializeComponent();
        }

        public int ProjectID { get; set; }

        private void datagridContractIn_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = e.Row.GetIndex() + 1;
        }

        private void datagridContractPay_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = e.Row.GetIndex() + 1;
        }

        private void dataGridFund_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = e.Row.GetIndex() + 1;
        }

        private void dataGridFund_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var funds = (ContractFunds) dataGridFund.SelectedItem;
            if (funds != null)
            {
                fundID = funds.Id;
                if (funds.ContractID != null) contractID = (int) funds.ContractID;
                textBoxContractNo.Text = funds.ContractNo;
                FundSource.Text = funds.Source;
                FundClassifys.SelectedItem =
                    dataContent.FundClassify.Single(fc => fc.FandClassifyId.Equals(funds.FundClassifyID));
                Money.Text = funds.Money.ToString();
                FundDate.DisplayDateStart = DateTime.MinValue;
                FundDate.DisplayDateEnd = DateTime.MaxValue;
                try
                {
                    if (funds.Date != null)
                    {
                        FundDate.SelectedDate = (DateTime) funds.Date;
                        FundDate.DisplayDate = (DateTime) funds.Date;
                    }
                }
                catch (Exception)
                {
                }
                Handled.Text = funds.Handled;
                SubPrincipal.Text = funds.SubPrincipal;
            }
        }

        private void buttonPreYear_Click(object sender, RoutedEventArgs e)
        {
            FundDate.DisplayDate = FundDate.DisplayDate.AddYears(-1);
            if (FundDate.SelectedDate != null)
                FundDate.SelectedDate = ((DateTime) FundDate.SelectedDate).AddYears(-1);
        }

        private void FundDate_DisplayDateChanged(object sender, CalendarDateChangedEventArgs e)
        {
            SetBlackOutDate();
        }

        private void SetBlackOutDate()
        {
            try
            {
                var dt = FundDate.DisplayDate;
                FundDate.BlackoutDates.Clear();
                if (dt.Month == 1)
                {
                    FundDate.BlackoutDates.Add(new CalendarDateRange(new DateTime(dt.Year - 1, 12, 15),
                        new DateTime(dt.Year - 1, 12, 31)));
                }
                else
                {
                    var s = new DateTime(dt.Year, dt.Month - 1, 15);
                    var e = new DateTime(dt.Year, dt.Month - 1, DateTime.DaysInMonth(dt.Year, dt.Month - 1));
                    FundDate.BlackoutDates.Add(new CalendarDateRange(s, e));
                }
                if (dt.Month == 12)
                {
                    FundDate.BlackoutDates.Add(new CalendarDateRange(new DateTime(dt.Year + 1, 1, 1),
                        new DateTime(dt.Year + 1, 1, 15)));
                }
                else
                {
                    var s = new DateTime(dt.Year, dt.Month + 1, 1);
                    var e = new DateTime(dt.Year, dt.Month + 1, 15);
                    FundDate.BlackoutDates.Add(new CalendarDateRange(s, e));
                }
            }
            catch (Exception)
            {
            }
        }

        private void FundDate_Loaded(object sender, RoutedEventArgs e)
        {
            SetBlackOutDate();
        }

        private void buttonNextYear_Click(object sender, RoutedEventArgs e)
        {
            FundDate.DisplayDate = FundDate.DisplayDate.AddYears(1);
            if (FundDate.SelectedDate != null)
                FundDate.SelectedDate = ((DateTime) FundDate.SelectedDate).AddYears(1);
        }

        private void buttonSubmit_Click(object sender, RoutedEventArgs e)
        {
            if (FundClassifys.SelectedItem == null)
            {
                MessageBox.Show("请选择经费类型！", "错误");
                return;
            }
            var fundClassify = (FundClassify) FundClassifys.SelectedItem;
            if ((fundClassify.FundClassify1 == "课题组报支") || (fundClassify.FundClassify1 == "管理费"))
            {
                MessageBox.Show("经费类型属于项目经费，请到项目经费管理模块处理！", "错误");
                return;
            }
            if (FundDate.SelectedDate == null)
            {
                MessageBox.Show("请选择日期！", "错误");
                return;
            }
            if (contractID <= 0)
            {
                MessageBox.Show("请选择合同！", "错误");
                return;
            }
            var funds = new ContractFunds();
            funds.ContractID = contractID;
            if (!string.IsNullOrWhiteSpace(textBoxContractNo.Text))
                funds.ContractNo = textBoxContractNo.Text.Trim();
            funds.Source = FundSource.Text;
            funds.FundClassifyID = ((FundClassify) FundClassifys.SelectedItem).FandClassifyId;
            try
            {
                funds.Money = (decimal) double.Parse(Money.Text);
            }
            catch (FormatException)
            {
                MessageBox.Show("金额格式不对！", "错误");
                return;
            }
            funds.Date = FundDate.SelectedDate;
            funds.Handled = Handled.Text.Trim();
            funds.SubPrincipal = SubPrincipal.Text.Trim();
            dataContent = new DataClassesProjectClassifyDataContext();
            dataContent.ContractFunds.InsertOnSubmit(funds);
            dataContent.SubmitChanges();
            dataGridFund.DataContext = dataContent.Funds.Where(f => f.ProjectID.Equals(ProjectID)).OrderBy(f => f.Date);
            ((MainWindow) Owner).DialogR = true;
            SetTotalMoney();
        }

        private void SetTotalMoney()
        {
            if (dataContent == null)
                dataContent = new DataClassesProjectClassifyDataContext();
            totalInComing = 0;
            var fundsInComing =
                dataContent.Funds.Where(f => f.ProjectID.Equals(ProjectID) && (f.FundClassify.FundClassify1 == "到账"));
            foreach (var funds in fundsInComing)
                if (funds.Money != null) totalInComing += (decimal) funds.Money;
            var contractComing =
                dataContent.ContractFunds.Where(
                    f => f.ProjectContracts.ProjectID.Equals(ProjectID) && (f.FundClassify.FundClassify1 == "到账"));
            foreach (var contractFunds in contractComing)
                if (contractFunds.Money != null) totalInComing += (decimal) contractFunds.Money;
            totalPayfor = 0;
            var fundsPayfor =
                dataContent.Funds.Where(f => f.ProjectID.Equals(ProjectID) && (f.FundClassify.FundClassify1 != "到账"));
            foreach (var funds in fundsPayfor)
                if (funds.Money != null) totalPayfor += (decimal) funds.Money;
            var contractPayfor =
                dataContent.ContractFunds.Where(
                    f => f.ProjectContracts.ProjectID.Equals(ProjectID) && (f.FundClassify.FundClassify1 != "到账"));
            foreach (var funds in contractPayfor)
                if (funds.Money != null) totalPayfor += (decimal) funds.Money;
            textBlockTotal.Text = "收入总计：" + $"{totalInComing:N2}" + "万元，支出总计：" +
                                  $"{totalPayfor:N2}" + "万元";
        }

        private void buttonUpdate_Click(object sender, RoutedEventArgs e)
        {
            if (fundID == 0)
                return;
            if (FundClassifys.SelectedItem == null)
            {
                MessageBox.Show("请选择经费类型！", "错误");
                return;
            }
            if (FundDate.SelectedDate == null)
            {
                MessageBox.Show("请选择日期！", "错误");
                return;
            }
            if (dataContent == null)
                dataContent = new DataClassesProjectClassifyDataContext();
            var fund = dataContent.ContractFunds.Single(f => f.Id.Equals(fundID));
            fund.FundClassifyID = ((FundClassify)FundClassifys.SelectedItem).FandClassifyId;
            fund.ContractNo = textBoxContractNo.Text.Trim();
            fund.Source = FundSource.Text.Trim();
            try
            {
                fund.Money = (decimal)double.Parse(Money.Text);
            }
            catch (FormatException)
            {
                MessageBox.Show("金额格式不对！", "错误");
                return;
            }
            fund.Date = FundDate.SelectedDate;
            fund.Handled = Handled.Text.Trim();
            fund.SubPrincipal = SubPrincipal.Text.Trim();
            dataContent.SubmitChanges();
            dataContent = new DataClassesProjectClassifyDataContext();
            dataGridFund.DataContext = dataContent.ContractFunds.Where(f => f.ProjectContracts.ProjectID.Equals(ProjectID)).OrderBy(f => f.Date);
            SetTotalMoney();
        }

        private void ButtonShift_Click(object sender, RoutedEventArgs e)
        {
            if (FundClassifys.SelectedItem == null)
                return;
            var fundClassify = (FundClassify)FundClassifys.SelectedItem;
            if ((fundClassify.FundClassify1 != "课题组报支") && (fundClassify.FundClassify1 != "管理费"))
                return;
            if (dataGridFund.SelectedItem == null)
                return;
            
            if (dataContent == null) dataContent = new DataClassesProjectClassifyDataContext();
            try
            {
                var contractFunds = (ContractFunds)dataGridFund.SelectedItem;
                var contract = dataContent.ProjectContracts.FirstOrDefault(c => c.ContractNo.Equals(contractFunds.ContractNo));
                if (contract == null) return;
                var funds = new Funds
                {
                    ContractNo = contractFunds.ContractNo,
                    FundClassifyID = contractFunds.FundClassifyID,
                    Date = contractFunds.Date,
                    Handled = contractFunds.Handled,
                    Money = contractFunds.Money,
                    Source = contractFunds.Source,
                    SubPrincipal = contractFunds.SubPrincipal
                };
                dataContent.Funds.InsertOnSubmit(funds);
                dataContent.ContractFunds.DeleteOnSubmit(contractFunds);
                dataContent.SubmitChanges();
                dataContent = new DataClassesProjectClassifyDataContext();                                                                           
                dataGridFund.DataContext =
                    dataContent.ContractFunds.Where(f => f.ProjectContracts.ProjectID.Equals(ProjectID)).OrderBy(f => f.Date);
                SetTotalMoney();
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void buttonDelete_Click(object sender, RoutedEventArgs e)
        {
            if (fundID == 0)
            {
                MessageBox.Show("请选择经费项！", "错误");
                return;
            }
            if (MessageBox.Show("该项经费将被删除！确认要删除该项经费信息？", "警告", MessageBoxButton.YesNo) == MessageBoxResult.No)
                return;
            dataContent = new DataClassesProjectClassifyDataContext();
            var fd = dataContent.ContractFunds.Single(f => f.Id.Equals(fundID));
            dataContent.ContractFunds.DeleteOnSubmit(fd);
            dataContent.SubmitChanges();
            dataContent = new DataClassesProjectClassifyDataContext();
            dataGridFund.DataContext = dataContent.ContractFunds.Where(f => f.ProjectContracts.ProjectID.Equals(ProjectID));
            ((MainWindow)Owner).DialogR = true;
            Clear();
            SetTotalMoney();
        }

        private void buttonCancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Clear()
        {
            fundID = 0;
            textBoxContractNo.Text = "";
            FundSource.Text = "";
            FundClassifys.SelectedIndex = 0;
            Money.Text = "";
            FundDate.SelectedDate = DateTime.Now;
            FundDate.DisplayDate = DateTime.Now;
            Handled.Text = "";
            SubPrincipal.Text = "";
        }
    }
}