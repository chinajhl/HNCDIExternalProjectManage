using System;
using System.Collections.Generic;
using System.Data;
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

namespace HNCDIExternalProjectManage
{
    /// <summary>
    /// PrizePayManage.xaml 的交互逻辑
    /// </summary>
    public partial class PrizePayManage : Window
    {
        public PrizePayManage()
        {
            InitializeComponent();
        }

        private List<Department> _listDepartment;
        private List<Employee> _listEmployee;
        private List<PrizeClassify> _listPrizeClassify;
        private DataClassesProjectClassifyDataContext dataContext;
        private DomainOperate _domain;
        public int ProjectId { get; set; }
        private int PrizeId { get; set; }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            dataContext = new DataClassesProjectClassifyDataContext();
            _listDepartment = new List<Department>();
            _domain = new DomainOperate("HNCDI");
            _domain.GetOU();
            foreach (DataRow o in _domain.ArrayOU.Rows)
            {
                _listDepartment.Add(new Department { Name = o["Text"].ToString(), Value = o["Value"].ToString() });
            }
            ListBoxDepartment.ItemsSource = _listDepartment;
            ProjectBase projectBase = dataContext.ProjectBase.FirstOrDefault(p => p.ProjectId.Equals(ProjectId));
            if (projectBase != null) ProjectName.Text = projectBase.ProjectName.Trim();
            //DataGridPrizes.ItemsSource = dataContext.Prizes.Where(p => p.ProjectID.Equals(ProjectId));
            _listPrizeClassify = new List<PrizeClassify>
            {
                new PrizeClassify {PrizeClassifyName = "国家、部、省科技进步奖/管理成果奖"},
                new PrizeClassify {PrizeClassifyName = "国家、部、省优秀项目奖"},
                new PrizeClassify {PrizeClassifyName = "院优项目奖"},
                new PrizeClassify {PrizeClassifyName = "专利奖励"},
                new PrizeClassify {PrizeClassifyName = "论著奖励"}
            };
            ListBoxPrizeClassify.ItemsSource = _listPrizeClassify;
            TextBoxYear.Text = DateTime.Now.Year.ToString();
        }

        private void ListBoxDepartment_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ListBoxDepartment.SelectedItem == null)
            {
                ListBoxEmployee.Items.Clear();
                return;
            }
            _listEmployee = new List<Employee>();
            Department department = (Department)ListBoxDepartment.SelectedItem;
            _domain = new DomainOperate("HNCDI");
            _domain.GetUsersByOU(department.Name);
            foreach (DataRow emRow in _domain.ArrayUser.Rows)
            {
                _listEmployee.Add(new Employee { AccountName = emRow["Value"].ToString(), Name = emRow["Text"].ToString() });
            }
            ListBoxEmployee.ItemsSource = _listEmployee;
            CheckEmpty();
        }

        private void ButtonCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void ButtonCancel_GotFocus(object sender, RoutedEventArgs e)
        {
            ButtonCancel.MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
        }

        private void ButtonPay_GotFocus(object sender, RoutedEventArgs e)
        {
            ButtonPay.MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
        }

        private void ButtonRemove_GotFocus(object sender, RoutedEventArgs e)
        {
            ButtonRemove.MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
        }

        private void TextBoxPrize_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Enter) return;
            TextBoxPrize.MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
            CheckEmpty();
        }

        private void CheckEmpty()
        {
            ButtonPay.IsEnabled = false;
            ButtonRemove.IsEnabled = false;
            if (ListBoxDepartment.SelectedItem == null) return;
            if (ListBoxEmployee.SelectedItem == null) return;
            if (ListBoxPrizeClassify.SelectedItem == null) return;
            if (string.IsNullOrWhiteSpace(TextBoxAwardName.Text.Trim())) return;
            if (string.IsNullOrWhiteSpace(TextBoxPrize.Text.Trim())) return;
            if (string.IsNullOrWhiteSpace(TextBoxYear.Text.Trim())) return;
            ButtonPay.IsEnabled = true;
            if (PrizeId > 0) ButtonRemove.IsEnabled = true;
        }

        private void DataGridPrizes_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = e.Row.GetIndex() + 1;
        }

        private void ButtonPay_Click(object sender, RoutedEventArgs e)
        {
            if (dataContext == null) dataContext = new DataClassesProjectClassifyDataContext();
            try
            {
                Prizes prizes = new Prizes
                {
                    //ProjectID = ProjectId,
                    Department = ((Department)ListBoxDepartment.SelectedItem).Name,
                    Name = ((Employee)ListBoxEmployee.SelectedItem).Name,
                    PrizeClassify = ((PrizeClassify)ListBoxPrizeClassify.SelectedItem).PrizeClassifyName,
                    AccountName = ((Employee)ListBoxEmployee.SelectedItem).AccountName,
                    AwardName = TextBoxAwardName.Text.Trim(),
                    Prize = Convert.ToDecimal(TextBoxPrize.Text.Trim()),
                    PayYear = TextBoxYear.Text.Trim()
                };
                dataContext.Prizes.InsertOnSubmit(prizes);
                dataContext.SubmitChanges();
                //DataGridPrizes.ItemsSource = dataContext.Prizes.Where(p => p.ProjectID.Equals(ProjectId));
            }
            catch (Exception)
            { }
        }

        private void ButtonRemove_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("确定删除该项奖金？", "温馨提示", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                if (dataContext == null) dataContext = new DataClassesProjectClassifyDataContext();
                dataContext.Prizes.DeleteOnSubmit((Prizes)DataGridPrizes.SelectedItem);
                dataContext.SubmitChanges();
                //DataGridPrizes.ItemsSource = dataContext.Prizes.Where(p => p.ProjectID.Equals(ProjectId));
            }
        }

        private void DataGridPrizes_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Prizes prizes = (Prizes)DataGridPrizes.SelectedItem;
            if (prizes == null)
            {
                GridDetail.DataContext = null;
                PrizeId = 0;
                ButtonRemove.IsEnabled = false;
                return;
            }
            GridDetail.DataContext = prizes;
            
            Department department = _listDepartment.FirstOrDefault(d => d.Name.Equals(prizes.Department));
            //ListBoxItem listItemDepartment = new ListBoxItem();
            //listItemDepartment.Content = new TextBlock() { Text = prizes.Department };
            //foreach (var item in ListBoxDepartment.Items)
            //{
            //    if (item == department)
            //    {
            //        ((ListBoxItem)item).IsSelected = true;
            //    }
            ListBoxDepartment.SelectedItem = department;
            if (department != null) ListBoxDepartment.ScrollIntoView(department);
            Employee employee = _listEmployee.FirstOrDefault(em => em.AccountName.Equals(prizes.AccountName));
            if (employee != null)
            {
                ListBoxEmployee.SelectedItem = employee;
                ListBoxEmployee.ScrollIntoView(employee);
            }
            PrizeClassify prizeClassify =
                _listPrizeClassify.FirstOrDefault(pc => pc.PrizeClassifyName.Equals(prizes.PrizeClassify));
            if (prizeClassify != null)
            {
                ListBoxPrizeClassify.SelectedItem = prizeClassify;
                ListBoxPrizeClassify.ScrollIntoView(prizeClassify);
            }
            ButtonPay.IsEnabled = false;
            ButtonRemove.IsEnabled = true;
        }

        private void ListBoxEmployee_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            CheckEmpty();
        }

        private void ListBoxPrizeClassify_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            CheckEmpty();
        }

        private void TextBoxAwardName_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Enter) return;
            TextBoxAwardName.MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
            CheckEmpty();
        }

        private void TextBoxAwardName_LostFocus(object sender, RoutedEventArgs e)
        {
            CheckEmpty();
        }

        private void TextBoxPrize_LostFocus(object sender, RoutedEventArgs e)
        {
            CheckEmpty();
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
    }
}
