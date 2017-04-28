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
using DocumentFormat.OpenXml.Drawing;

namespace HNCDIExternalProjectManage
{
    /// <summary>
    /// SelectPrizeEmployee.xaml 的交互逻辑
    /// </summary>
    public partial class SelectPrizeEmployee : Window
    {
        private List<Employee> _listEmployees;
        public string EmployeeName { get; set; }
        public string SelectedAccountName { get; set; }
        public string SelectedDepartment { get; set; }


        public SelectPrizeEmployee()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(EmployeeName))
            {
                this.Close();
            }
            DomainOperate doo = new DomainOperate("hncdi");
            List<string> accounts = new List<string>();

            accounts = doo.GetLoginIDByUserName(EmployeeName);
            if (!accounts.Any()) this.Close();
            _listEmployees = new List<Employee>();
            
            foreach (string account in accounts)
            {
                Employee employee = new Employee
                {
                    Name = EmployeeName,
                    AccountName = account,
                    Department = doo.GetOuByLoginID(account)
                };
                _listEmployees.Add(employee);
            }
            LabelMessage.Content = "域内存在多个姓名为 " + EmployeeName + " 的账号，请选择：";
            DataGridEmployeeList.ItemsSource = _listEmployees;
        }

        private void ButtonSelect_GotFocus(object sender, RoutedEventArgs e)
        {
            ButtonSelect.MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
        }

        private void DataGridEmployeeList_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = e.Row.GetIndex() + 1;
        }

        private void DataGridEmployeeList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (DataGridEmployeeList.SelectedItem == null)
            {
                ButtonSelect.IsEnabled = false;
            }
            else
            {
                ButtonSelect.IsEnabled = true;
            }
        }

        private void ButtonSelect_Click(object sender, RoutedEventArgs e)
        {
            if (DataGridEmployeeList.SelectedItem == null)
            {
                ((ImportPrizes) this.Owner).CurrentEmployee = null;
                return;
            }
            Employee employee = (Employee) DataGridEmployeeList.SelectedItem;

            SelectedAccountName = employee.AccountName;
            SelectedDepartment = employee.Department;
            this.Close();
        }

        private void ButtonCancel_GotFocus(object sender, RoutedEventArgs e)
        {
            ButtonCancel.MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
        }

        private void ButtonCancel_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("尚未选择员工，确定忽略该条记录？", "温馨提示", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                ((ImportPrizes) this.Owner).CurrentEmployeeName = "";
                this.Close();
            }
        }
    }
}
