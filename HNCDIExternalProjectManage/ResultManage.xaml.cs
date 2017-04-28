using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace HNCDIExternalProjectManage
{
    /// <summary>
    /// ResultManage.xaml 的交互逻辑
    /// </summary>
    public partial class ResultManage : Window
    {
        private DataClassesProjectClassifyDataContext dataContent;

        private int projectID;

        public int ProjectID
        {
            get { return projectID; }
            set { projectID = value; }
        }

        private int ID;

        public ResultManage()
        {
            this.InitializeComponent();

            // 在此点之下插入创建对象所需的代码。
        }

        private void dataGridResults_LoadingRow(object sender, System.Windows.Controls.DataGridRowEventArgs e)
        {
            // 在此处添加事件处理程序实现。
            e.Row.Header = e.Row.GetIndex() + 1;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            dataContent = new DataClassesProjectClassifyDataContext();
            dataGridResults.DataContext = dataContent.Results.Where(r => r.ProjectID.Equals(projectID));
            ProjectBase projectBase = dataContent.ProjectBase.Single(p => p.ProjectId.Equals(projectID));
            this.Title = projectBase.ProjectName + "——成果登记管理";
        }

        private void RegistDate_Loaded(object sender, RoutedEventArgs e)
        {
            SetBlackOutDate();
        }

        private void RegistDate_DisplayDateChanged(object sender, CalendarDateChangedEventArgs e)
        {
            SetBlackOutDate();
        }

        private void SetBlackOutDate()
        {
            try
            {
                DateTime dt = RegistDate.DisplayDate;
                RegistDate.BlackoutDates.Clear();
                if (dt.Month == 1)
                {
                    RegistDate.BlackoutDates.Add(new CalendarDateRange(new DateTime(dt.Year - 1, 12, 15), new DateTime(dt.Year - 1, 12, 31)));
                }
                else
                {
                    DateTime s = new DateTime(dt.Year, dt.Month - 1, 15);
                    DateTime e = new DateTime(dt.Year, dt.Month - 1, DateTime.DaysInMonth(dt.Year, dt.Month - 1));
                    RegistDate.BlackoutDates.Add(new CalendarDateRange(s, e));
                }
                if (dt.Month == 12)
                {
                    RegistDate.BlackoutDates.Add(new CalendarDateRange(new DateTime(dt.Year + 1, 1, 1), new DateTime(dt.Year + 1, 1, 15)));
                }
                else
                {
                    DateTime s = new DateTime(dt.Year, dt.Month + 1, 1);
                    DateTime e = new DateTime(dt.Year, dt.Month + 1, 15);
                    RegistDate.BlackoutDates.Add(new CalendarDateRange(s, e));
                }
            }
            catch (Exception)
            {
            }
        }

        private void buttonDelete_Click(object sender, RoutedEventArgs e)
        {
            if (ID == 0)
            {
                MessageBox.Show("请选择成果登记项！", "错误");
                return;
            }
            dataContent = new DataClassesProjectClassifyDataContext();
            var re = dataContent.Results.Single(r => r.Id.Equals(ID));
            dataContent.Results.DeleteOnSubmit(re);
            dataContent.SubmitChanges();
            dataContent = new DataClassesProjectClassifyDataContext();
            dataGridResults.DataContext = dataContent.Results.Where(r => r.ProjectID.Equals(projectID));
            ((MainWindow)(this.Owner)).DialogR = true;
            Clear();
        }

        private void dataGridResults_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Results r = (Results)dataGridResults.SelectedItem;
            if (r != null)
            {
                ID = r.Id;
                RegistDate.SelectedDate = r.RegistDate;
                RegistDate.DisplayDate = (DateTime)r.RegistDate;
                RegistNo.Text = r.RegistNo;
            }
        }

        private void Clear()
        {
            ID = 0;
            RegistDate.SelectedDate = DateTime.Now;
            RegistDate.DisplayDate = DateTime.Now;
            RegistNo.Text = "";
        }

        private void buttonSubmit_Click(object sender, RoutedEventArgs e)
        {
            if (RegistDate.SelectedDate == null)
            {
                MessageBox.Show("请选择日期！", "错误");
                return;
            }
            if (String.IsNullOrEmpty(RegistNo.Text))
            {
                MessageBox.Show("请输入成果登记号！", "错误");
                return;
            }
            Results rs = new Results();
            rs.ProjectID = projectID;
            rs.RegistDate = RegistDate.SelectedDate;
            rs.RegistNo = RegistNo.Text.Trim();
            dataContent = new DataClassesProjectClassifyDataContext();
            dataContent.Results.InsertOnSubmit(rs);
            dataContent.SubmitChanges();
            dataGridResults.DataContext = dataContent.Results.Where(r => r.ProjectID.Equals(projectID));
            ((MainWindow)(this.Owner)).DialogR = true;
        }

        private void buttonPreYear_Click(object sender, RoutedEventArgs e)
        {
            RegistDate.DisplayDate = RegistDate.DisplayDate.AddYears(-1);
            if (RegistDate.SelectedDate != null)
            {
                RegistDate.SelectedDate = ((DateTime)(RegistDate.SelectedDate)).AddYears(-1);
            }
        }

        private void buttonNextYear_Click(object sender, RoutedEventArgs e)
        {
            RegistDate.DisplayDate = RegistDate.DisplayDate.AddYears(1);
            if (RegistDate.SelectedDate != null)
            {
                RegistDate.SelectedDate = ((DateTime)(RegistDate.SelectedDate)).AddYears(1);
            }
        }
    }
}