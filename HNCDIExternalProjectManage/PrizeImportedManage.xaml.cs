using System;
using System.Collections;
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

namespace HNCDIExternalProjectManage
{
    /// <summary>
    /// PrizeImportedManage.xaml 的交互逻辑
    /// </summary>
    public partial class PrizeImportedManage : Window
    {
        private DataClassesProjectClassifyDataContext dataContext;
        private long _currentPrizeID = 0;
        private List<string> _listDepartments;
        private List<string> _listDeclareDeoartments;
        private List<string> _listProjects;
        private List<string> _listYear;
        IEnumerable<Prizes> _listPrizes;
        private Prize _currentPrize;


        public PrizeImportedManage()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            dataContext = new DataClassesProjectClassifyDataContext();
            _listDepartments = new List<string>();
            _listDeclareDeoartments = new List<string>();
            _listYear = new List<string>();
            _listProjects = new List<string>();
            _listDepartments = dataContext.Prizes.Select(p => p.Department).Distinct().ToList();
            _listDepartments.Add("全部");
            _listDeclareDeoartments = dataContext.Prizes.Select(p => p.DeclareDepartment).Distinct().ToList();
            _listDeclareDeoartments.Add("全部");
            _listYear = dataContext.Prizes.Select(p => p.PayYear).Distinct().ToList();
            _listYear.Add("全部");
            _listProjects = dataContext.Prizes.Select(p => p.Project).Distinct().ToList();
            _listProjects.Add("全部");

            ListBoxDepartment.ItemsSource = _listDepartments;
            ListBoxDeclareDepartment.ItemsSource = _listDeclareDeoartments;
            ListBoxYear.ItemsSource = _listYear;
            ListBoxProject.ItemsSource = _listProjects;
            //_listPrizes =new List<Prizes>();

            IEnumerable<Prize> _listPrizes =
                dataContext.Prizes.Select(p => new Prize
                {
                    PrizeClassify = p.PrizeClassify,
                    Project = p.Project,
                    AwardName = p.AwardName,
                    Name = p.Name,
                    AccountName = p.AccountName,
                    Department = p.Department,
                    DeclareDepartment = p.DeclareDepartment,
                    PayYear = p.PayYear,
                    PrizeValue = Convert.ToDecimal(p.Prize)
                }).AsEnumerable().Distinct(new PrizeComparer());
            DataGridPrizes.ItemsSource = _listPrizes;
        }

        private void DataGridPrizes_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = e.Row.GetIndex() + 1;
        }

        private void ButtonUpdate_GotFocus(object sender, RoutedEventArgs e)
        {
            ButtonUpdate.MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
        }

        private void ButtonRemove_GotFocus(object sender, RoutedEventArgs e)
        {
            ButtonRemove.MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
        }

        private void ButtonExit_GotFocus(object sender, RoutedEventArgs e)
        {
            ButtonExit.MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
        }

        private void ButtonExit_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void CheckEmpty()
        {
            ButtonUpdate.IsEnabled = false;
            ButtonRemove.IsEnabled = false;
            if (_currentPrizeID == 0) return;
            ButtonUpdate.IsEnabled = true;
            ButtonRemove.IsEnabled = true;
        }

        private void DataGridPrizes_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (DataGridPrizes.SelectedItem == null)
            {
                _currentPrizeID = 0;
                GridPrizeDetail.DataContext = null;
            }
            else
            {
                Prize prize = (Prize)DataGridPrizes.SelectedItem;
                _currentPrize = new Prize();
                _currentPrize = prize;
                Prizes prizes =
                    dataContext.Prizes.FirstOrDefault(
                        p => p.Name.Equals(prize.Name) && p.AccountName.Equals(prize.AccountName)
                             && p.Department.Equals(prize.Department) && p.PrizeClassify.Equals(prize.PrizeClassify) &&
                             p.Project.Equals(prize.Project)
                             & p.AwardName.Equals(prize.AwardName) && p.PayYear.Equals(prize.PayYear) &&
                             p.Prize.Equals(prize.PrizeValue));

                _currentPrizeID = prizes?.ID ?? 0;
                GridPrizeDetail.DataContext = prize;
            }
            CheckEmpty();
        }

        private void ListBoxDepartment_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            FreshData();
        }

        private void FreshData()
        {
            string filter = "";
            //_listPrizes = new List<Prizes>();

            var searchPredicate = PredicateExtensions.True<Prizes>();

            if (ListBoxDepartment.SelectedItem != null)
            {
                filter = (string)ListBoxDepartment.SelectedItem;
                if (filter != "全部")
                {
                    var filter1 = filter;
                    searchPredicate = searchPredicate.And(p => p.Department.Equals(filter1));
                }
            }

            if (ListBoxDeclareDepartment.SelectedItem != null)
            {
                filter = (string)ListBoxDeclareDepartment.SelectedItem;
                if (filter != "全部")
                {
                    var filter1 = filter;
                    searchPredicate = searchPredicate.And(p => p.DeclareDepartment.Equals(filter1));
                }
            }

            if (ListBoxYear.SelectedItem != null)
            {
                filter = (string)ListBoxYear.SelectedItem;
                if (filter != "全部")
                {
                    var filter1 = filter;
                    searchPredicate = searchPredicate.And(p => p.PayYear.Equals(filter1));
                }
            }

            if (ListBoxProject.SelectedItem != null)
            {
                filter = (string)ListBoxProject.SelectedItem;
                if (filter != "全部")
                {
                    var filter1 = filter;
                    searchPredicate = searchPredicate.And(p => p.Project.Equals(filter1));
                }
            }

            var prizess = dataContext.Prizes.Where(searchPredicate);
            List<Prize> _listPrizes =
                prizess.Select(p => new Prize
                {
                    PrizeClassify = p.PrizeClassify,
                    Project = p.Project,
                    AwardName = p.AwardName,
                    Name = p.Name,
                    AccountName = p.AccountName,
                    Department = p.Department,
                    DeclareDepartment = p.DeclareDepartment,
                    PayYear = p.PayYear,
                    PrizeValue = Convert.ToDecimal(p.Prize)
                }).AsEnumerable().Distinct(new PrizeComparer()).ToList();
            List<Prize> result = new List<Prize>();
            foreach (Prize prize in _listPrizes)
            {
                bool boolMatch = result.Any(p => prize.AccountName == p.AccountName && prize.AwardName == p.AwardName && prize.Department == p.Department && prize.PayYear == p.PayYear && prize.PrizeClassify == p.PrizeClassify && prize.Project == p.Project && prize.PrizeValue == p.PrizeValue && prize.Name == p.Name);

                if (!boolMatch)
                {
                    result.Add(prize);
                }
            }

            //_listPrizes = _listPrizes.Where(searchPredicate);

            //if (ListBoxDepartment.SelectedItem != null)
            //{
            //    filter = (string) ListBoxDepartment.SelectedItem;
            //    if (filter != "全部")
            //    {
            //        _listPrizes = _listPrizes.Where(p => p.Department.Equals(filter)).AsEnumerable().Distinct(new PrizeComparer()).ToList();
            //    }
            //}
            //if (ListBoxDeclareDepartment.SelectedItem != null)
            //{
            //    filter = (string) ListBoxDeclareDepartment.SelectedItem;
            //    if (filter != "全部")
            //    {
            //        _listPrizes = _listPrizes.Where(p => p.DeclareDepartment.Equals(filter)).AsEnumerable().Distinct(new PrizeDepartmentComparer()).ToList();
            //    }
            //}
            //if (ListBoxYear.SelectedItem != null)
            //{
            //    filter = (string) ListBoxYear.SelectedItem;
            //    if (filter != "全部")
            //    {
            //        _listPrizes = _listPrizes.Where(p => p.PayYear.Equals(filter)).AsEnumerable().Distinct(new PrizeComparer()).ToList();
            //    }
            //}
            //if (ListBoxProject.SelectedItem != null)
            //{
            //    filter = (string) ListBoxProject.SelectedItem;
            //    if (filter != "全部")
            //    {
            //        _listPrizes = _listPrizes.Where(p => p.Project.Equals(filter)).AsEnumerable().Distinct(new PrizeComparer()).ToList();
            //    }
            //}
            //_listPrizes = _listPrizes.Distinct(new PrizeComparer()).ToList();
            DataGridPrizes.ItemsSource = result;
        }

        private void ListBoxDeclareDepartment_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            FreshData();
        }

        private void ListBoxYear_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            FreshData();
        }

        private void ListBoxProject_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            FreshData();
        }

        private void TextBoxPrize_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Enter) return;
            TextBoxPrize.MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
            if (string.IsNullOrWhiteSpace(TextBoxPrize.Text))
            {
                ButtonUpdate.IsEnabled = false;
                return;
            }
            try
            {
                decimal tempMoney = Convert.ToDecimal(TextBoxPrize.Text);
                CheckEmpty();
            }
            catch (Exception)
            {
                MessageBox.Show("奖金格式不对，应为有效数字");
                ButtonUpdate.IsEnabled = false;
            }
        }

        private void ButtonUpdate_Click(object sender, RoutedEventArgs e)
        {
            if (dataContext == null) dataContext = new DataClassesProjectClassifyDataContext();
            Prizes prizes = dataContext.Prizes.FirstOrDefault(p => p.ID.Equals(_currentPrizeID));
            if (prizes != null)
            {
                var prizeses =
                    dataContext.Prizes.Where(
                        p =>
                            p.Name.Equals(prizes.Name) && p.AccountName.Equals(prizes.AccountName) && p.PayYear.Equals(prizes.PayYear) &&
                            p.Department.Equals(prizes.Department) && p.PrizeClassify.Equals(prizes.PrizeClassify) &&
                            p.Project.Equals(prizes.Project) && p.AwardName.Equals(prizes.AwardName));
                try
                {
                    foreach (var pr in prizeses)
                    {
                        pr.Prize = Convert.ToDecimal(TextBoxPrize.Text);
                    }
                    dataContext.SubmitChanges();
                    FreshData();
                }
                catch (Exception error)
                {
                    MessageBox.Show(error.Message);
                }
            }
        }

        private void TextBoxPrize_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TextBoxPrize.Text))
            {
                ButtonUpdate.IsEnabled = false;
                return;
            }
            try
            {
                decimal tempMoney = Convert.ToDecimal(TextBoxPrize.Text);
                CheckEmpty();
            }
            catch (Exception)
            {
                MessageBox.Show("奖金格式不对，应为有效数字");
                ButtonUpdate.IsEnabled = false;
            }
        }

        private void ButtonRemove_Click(object sender, RoutedEventArgs e)
        {
            if (dataContext == null) dataContext = new DataClassesProjectClassifyDataContext();
            Prizes prizes = dataContext.Prizes.FirstOrDefault(p => p.ID.Equals(_currentPrizeID));
            if (prizes != null)
            {
                var prizeses =
                    dataContext.Prizes.Where(
                        p =>
                            p.Name.Equals(prizes.Name) && p.AccountName.Equals(prizes.AccountName) &&
                            p.PayYear.Equals(prizes.PayYear) &&
                            p.Department.Equals(prizes.Department) && p.PrizeClassify.Equals(prizes.PrizeClassify) &&
                            p.Project.Equals(prizes.Project) && p.AwardName.Equals(prizes.AwardName));
                try
                {
                    foreach (var pr in prizeses)
                    {
                        dataContext.Prizes.DeleteOnSubmit(pr);
                        dataContext.SubmitChanges();
                    }
                    FreshData();
                }
                catch (Exception error)
                {
                    MessageBox.Show(error.Message);
                }
            }
        }
    }
}
