using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace HNCDIExternalProjectManage
{
    /// <summary>
    /// RateManage.xaml 的交互逻辑
    /// </summary>
    public partial class RateManage : Window
    {
        private DataClassesProjectClassifyDataContext dataContent;
        private int projectID;

        public int ProjectID
        {
            get { return projectID; }
            set { projectID = value; }
        }

        private int rateID;

        public RateManage()
        {
            this.InitializeComponent();

            // 在此点之下插入创建对象所需的代码。
        }

        private void dataGrigRate_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = e.Row.GetIndex() + 1;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            dataContent = new DataClassesProjectClassifyDataContext();
            RateClassifys.DisplayMemberPath = "RateClassify1";
            RateClassifys.SelectedValuePath = "RateClassifyID";
            RateClassifys.DataContext = dataContent.RateClassify;
            dataGridRate.DataContext = dataContent.RateResults.Where(r => r.ProjectID.Equals(projectID));
            ProjectBase projectBase = dataContent.ProjectBase.Single(pb => pb.ProjectId.Equals(projectID));
            this.Title = projectBase.ProjectName + "——鉴定管理";
        }

        private void dataGridRate_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            RateResults rateResults = (RateResults)dataGridRate.SelectedItem;
            if (rateResults != null)
            {
                rateID = rateResults.Id;
                RateDepartment.Text = rateResults.RateDepartment;
                RateClassifys.SelectedItem = (RateClassify)dataContent.RateClassify.Single(r => r.RateClassifyId.Equals(rateResults.RateClassifyID));
                RateDate.SelectedDate = rateResults.RateDate;
                RateDate.DisplayDate = (DateTime)rateResults.RateDate;
                Note.Text = rateResults.Note;
            }
        }

        private void SetBlackOutDate()
        {
            try
            {
                DateTime dt = RateDate.DisplayDate;
                RateDate.BlackoutDates.Clear();
                if (dt.Month == 1)
                {
                    RateDate.BlackoutDates.Add(new CalendarDateRange(new DateTime(dt.Year - 1, 12, 15), new DateTime(dt.Year - 1, 12, 31)));
                }
                else
                {
                    DateTime s = new DateTime(dt.Year, dt.Month - 1, 15);
                    DateTime e = new DateTime(dt.Year, dt.Month - 1, DateTime.DaysInMonth(dt.Year, dt.Month - 1));
                    RateDate.BlackoutDates.Add(new CalendarDateRange(s, e));
                }
                if (dt.Month == 12)
                {
                    RateDate.BlackoutDates.Add(new CalendarDateRange(new DateTime(dt.Year + 1, 1, 1), new DateTime(dt.Year + 1, 1, 15)));
                }
                else
                {
                    DateTime s = new DateTime(dt.Year, dt.Month + 1, 1);
                    DateTime e = new DateTime(dt.Year, dt.Month + 1, 15);
                    RateDate.BlackoutDates.Add(new CalendarDateRange(s, e));
                }
            }
            catch (Exception)
            {
            }
        }

        private void RateDate_Loaded(object sender, RoutedEventArgs e)
        {
            SetBlackOutDate();
        }

        private void RateDate_DisplayDateChanged(object sender, CalendarDateChangedEventArgs e)
        {
            SetBlackOutDate();
        }

        private void buttonSubmit_Click(object sender, RoutedEventArgs e)
        {
            if (RateClassifys.SelectedItem == null)
            {
                MessageBox.Show("请选择鉴定结论！", "错误");
                return;
            }
            if (RateDate.SelectedDate == null)
            {
                MessageBox.Show("请选择日期！", "错误");
                return;
            }
            RateResults rateResults = new RateResults();
            rateResults.ProjectID = projectID;
            rateResults.RateDepartment = RateDepartment.Text.Trim();
            rateResults.RateClassifyID = ((RateClassify)(RateClassifys.SelectedItem)).RateClassifyId;
            rateResults.RateDate = RateDate.SelectedDate;
            rateResults.Note = Note.Text.Trim();
            dataContent = new DataClassesProjectClassifyDataContext();
            dataContent.RateResults.InsertOnSubmit(rateResults);
            dataContent.SubmitChanges();
            dataGridRate.DataContext = dataContent.RateResults.Where(r => r.ProjectID.Equals(projectID));
            ((MainWindow)(this.Owner)).DialogR = true;
        }

        private void buttonDelete_Click(object sender, RoutedEventArgs e)
        {
            if (rateID == 0)
            {
                MessageBox.Show("请选择鉴定项！", "错误");
                return;
            }
            if (MessageBox.Show("该项鉴定将被删除！确认要删除该项鉴定信息？", "警告", MessageBoxButton.YesNo) == MessageBoxResult.No)
            {
                return;
            }

            dataContent = new DataClassesProjectClassifyDataContext();
            var rs = dataContent.RateResults.Single(r => r.Id.Equals(rateID));
            dataContent.RateResults.DeleteOnSubmit(rs);
            dataContent.SubmitChanges();
            dataContent = new DataClassesProjectClassifyDataContext();
            dataGridRate.DataContext = dataContent.RateResults.Where(r => r.ProjectID.Equals(projectID));
            ((MainWindow)(this.Owner)).DialogR = true;
            Clear();
        }

        private void Clear()
        {
            rateID = 0;
            RateDepartment.Text = "";
            RateClassifys.SelectedIndex = 0;
            RateDate.SelectedDate = DateTime.Now;
            RateDate.DisplayDate = DateTime.Now;
            Note.Text = "";
        }

        private void buttonPreYear_Click(object sender, RoutedEventArgs e)
        {
            RateDate.DisplayDate = RateDate.DisplayDate.AddYears(-1);
            if (RateDate.SelectedDate != null)
            {
                RateDate.SelectedDate = ((DateTime)(RateDate.SelectedDate)).AddYears(-1);
            }
        }

        private void buttonNextYear_Click(object sender, RoutedEventArgs e)
        {
            RateDate.DisplayDate = RateDate.DisplayDate.AddYears(1);
            if (RateDate.SelectedDate != null)
            {
                RateDate.SelectedDate = ((DateTime)(RateDate.SelectedDate)).AddYears(1);
            }
        }
    }
}