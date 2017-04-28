using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace HNCDIExternalProjectManage
{
    /// <summary>
    /// RewardManage.xaml 的交互逻辑
    /// </summary>
    public partial class RewardManage : Window
    {
        private DataClassesProjectClassifyDataContext dataContext;

        private int projectID;

        public int ProjectID
        {
            get { return projectID; }
            set { projectID = value; }
        }

        private int ID;

        public RewardManage()
        {
            this.InitializeComponent();

            // 在此点之下插入创建对象所需的代码。
        }

        private void dataGridRewards_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = e.Row.GetIndex() + 1;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            dataContext = new DataClassesProjectClassifyDataContext();
            RewardClasses.DisplayMemberPath = "RewardClass1";
            RewardClasses.SelectedValuePath = "RewradClassID";
            RewardClasses.DataContext = dataContext.RewardClass;
            RewardClassifies.DisplayMemberPath = "RewardClassify1";
            RewardClassifies.SelectedValuePath = "RewardClassifyID";
            RewardClassifies.DataContext = dataContext.RewardClassify;
            dataGridRewards.DataContext = dataContext.Reward.Where(r => r.ProjectID.Equals(projectID));
            ProjectBase projectBase = dataContext.ProjectBase.Single(pb => pb.ProjectId.Equals(projectID));
            this.Title = projectBase.ProjectName + "——获奖记录管理";
        }

        private void SetBlackOutDate()
        {
            try
            {
                DateTime dt = RewardDate.DisplayDate;
                RewardDate.BlackoutDates.Clear();
                if (dt.Month == 1)
                {
                    RewardDate.BlackoutDates.Add(new CalendarDateRange(new DateTime(dt.Year - 1, 12, 15), new DateTime(dt.Year - 1, 12, 31)));
                }
                else
                {
                    DateTime s = new DateTime(dt.Year, dt.Month - 1, 15);
                    DateTime e = new DateTime(dt.Year, dt.Month - 1, DateTime.DaysInMonth(dt.Year, dt.Month - 1));
                    RewardDate.BlackoutDates.Add(new CalendarDateRange(s, e));
                }
                if (dt.Month == 12)
                {
                    RewardDate.BlackoutDates.Add(new CalendarDateRange(new DateTime(dt.Year + 1, 1, 1), new DateTime(dt.Year + 1, 1, 15)));
                }
                else
                {
                    DateTime s = new DateTime(dt.Year, dt.Month + 1, 1);
                    DateTime e = new DateTime(dt.Year, dt.Month + 1, 15);
                    RewardDate.BlackoutDates.Add(new CalendarDateRange(s, e));
                }
            }
            catch (Exception)
            {
            }
        }

        private void RewardDate_DisplayDateChanged(object sender, CalendarDateChangedEventArgs e)
        {
            SetBlackOutDate();
        }

        private void RewardDate_Loaded(object sender, RoutedEventArgs e)
        {
            SetBlackOutDate();
        }

        private void dataGridRewards_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Reward rw = (Reward)dataGridRewards.SelectedItem;
            if (rw != null)
            {
                ID = rw.Id;
                RewardName.Text = rw.RewardName;
                RewardClasses.SelectedItem = (RewardClass)dataContext.RewardClass.Single(rc => rc.RewardClassId.Equals(rw.RewardClassID));
                RewardClassifies.SelectedItem = (RewardClassify)dataContext.RewardClassify.Single(rcf => rcf.RewardClassifyID.Equals(rw.RewardClassifyID));
                RewardDepartment.Text = rw.RewardDepartment;
                RewardYear.Text = rw.ReawardYear;
                if (rw.RewardDate != null)
                {
                    RewardDate.SelectedDate = rw.RewardDate;
                    RewardDate.DisplayDate = (DateTime)rw.RewardDate;
                }
                Department.Text = rw.Department;
                Workers.Text = rw.Workers;
                Note.Text = rw.Note;
            }
        }

        private void buttonDelete_Click(object sender, RoutedEventArgs e)
        {
            if (ID == 0)
            {
                MessageBox.Show("请选择奖项！", "错误");
                return;
            }
            if (MessageBox.Show("该项奖项将被删除！确认要删除该项奖项信息？", "警告", MessageBoxButton.YesNo) == MessageBoxResult.No)
            {
                return;
            }
            dataContext = new DataClassesProjectClassifyDataContext();
            var rw = dataContext.Reward.Single(r => r.Id.Equals(ID));
            dataContext.Reward.DeleteOnSubmit(rw);
            dataContext.SubmitChanges();
            dataContext = new DataClassesProjectClassifyDataContext();
            dataGridRewards.DataContext = dataContext.Reward.Where(r => r.ProjectID.Equals(projectID));
            ((MainWindow)(this.Owner)).DialogR = true;
            Clear();
        }

        private void Clear()
        {
            ID = 0;
            RewardName.Text = "";
            RewardClasses.SelectedIndex = 0;
            RewardClassifies.SelectedIndex = 0;
            RewardDepartment.Text = "";
            RewardYear.Text = "";
            RewardDate.SelectedDate = DateTime.Now;
            RewardDate.DisplayDate = DateTime.Now;
            Department.Text = "";
            Workers.Text = "";
            Note.Text = "";
        }

        private void buttonSubmit_Click(object sender, RoutedEventArgs e)
        {
            if (RewardClasses.SelectedItem == null)
            {
                MessageBox.Show("请选择奖项级别！", "错误");
                return;
            }
            if (RewardClassifies.SelectedItem == null)
            {
                MessageBox.Show("请选择奖项等别！", "错误");
                return;
            }
            //if(RewardDate.SelectedDate == null)
            //{
            //    MessageBox.Show("请选择授奖日期！", "错误");
            //    return;
            //}
            if (String.IsNullOrEmpty(RewardDepartment.Text))
            {
                MessageBox.Show("请填写授奖机构！", "错误");
                return;
            }
            if (String.IsNullOrEmpty(RewardName.Text))
            {
                MessageBox.Show("请填写奖项名称！", "错误");
                return;
            }
            dataContext = new DataClassesProjectClassifyDataContext();
            Reward rw = new Reward();
            rw.ProjectID = projectID;
            rw.RewardClassID = ((RewardClass)RewardClasses.SelectedItem).RewardClassId;
            rw.RewardClassifyID = ((RewardClassify)RewardClassifies.SelectedItem).RewardClassifyID;
            rw.RewardName = RewardName.Text.Trim();
            rw.ReawardYear = RewardYear.Text.Trim();
            if (RewardDate.SelectedDate != null)
            {
                rw.RewardDate = RewardDate.SelectedDate;
            }
            rw.RewardDepartment = RewardDepartment.Text.Trim();
            rw.Department = Department.Text.Trim();
            rw.Workers = Workers.Text.Trim();
            rw.Note = Note.Text.Trim();
            dataContext.Reward.InsertOnSubmit(rw);
            dataContext.SubmitChanges();
            dataGridRewards.DataContext = dataContext.Reward.Where(r => r.ProjectID.Equals(projectID));
            ((MainWindow)(this.Owner)).DialogR = true;
        }

        private void buttonPreYear_Click(object sender, RoutedEventArgs e)
        {
            RewardDate.DisplayDate = RewardDate.DisplayDate.AddYears(-1);
            if (RewardDate.SelectedDate != null)
            {
                RewardDate.SelectedDate = ((DateTime)(RewardDate.SelectedDate)).AddYears(-1);
            }
        }

        private void buttonNextYear_Click(object sender, RoutedEventArgs e)
        {
            RewardDate.DisplayDate = RewardDate.DisplayDate.AddYears(1);
            if (RewardDate.SelectedDate != null)
            {
                RewardDate.SelectedDate = ((DateTime)(RewardDate.SelectedDate)).AddYears(1);
            }
        }
    }
}