using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Linq;
using System.Linq;
using System.Text;
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
	/// PatentManage.xaml 的交互逻辑
	/// </summary>
	public partial class PatentManage : Window
	{
        DataClassesProjectClassifyDataContext dataContext;

        private int projectID;

        public int ProjectID
        {
            get { return projectID; }
            set { projectID = value; }
        }

        private int patentsID;
		public PatentManage()
		{
			this.InitializeComponent();
			
			// 在此点之下插入创建对象所需的代码。
		}

        private void dataGridPatents_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = e.Row.GetIndex() + 1;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            dataContext = new DataClassesProjectClassifyDataContext();
            PatentClassifies.DisplayMemberPath = "PatentClassify1";
            PatentClassifies.SelectedValuePath = "PatentClassifyID";
            PatentClassifies.DataContext = dataContext.PatentClassify;
            dataGridPatents.DataContext = dataContext.Patents.Where(p => p.ProjectID.Equals(projectID));
            ProjectBase projectBase = dataContext.ProjectBase.Single(p => p.ProjectId.Equals(projectID));
            this.Title = projectBase.ProjectName + "——知识产权管理";
        }

        private void SetBlackOutDate()
        {
            try
            {
                DateTime dt = PatentDate.DisplayDate;
                PatentDate.BlackoutDates.Clear();
                if (dt.Month == 1)
                {
                    PatentDate.BlackoutDates.Add(new CalendarDateRange(new DateTime(dt.Year - 1, 12, 15), new DateTime(dt.Year - 1, 12, 31)));
                }
                else
                {
                    DateTime s = new DateTime(dt.Year, dt.Month - 1, 15);
                    DateTime e = new DateTime(dt.Year, dt.Month - 1, DateTime.DaysInMonth(dt.Year, dt.Month - 1));
                    PatentDate.BlackoutDates.Add(new CalendarDateRange(s, e));
                }
                if (dt.Month == 12)
                {
                    PatentDate.BlackoutDates.Add(new CalendarDateRange(new DateTime(dt.Year + 1, 1, 1), new DateTime(dt.Year + 1, 1, 15)));
                }
                else
                {
                    DateTime s = new DateTime(dt.Year, dt.Month + 1, 1);
                    DateTime e = new DateTime(dt.Year, dt.Month + 1, 15);
                    PatentDate.BlackoutDates.Add(new CalendarDateRange(s, e));
                }
            }
            catch (Exception)
            {

            }
        }

        private void PatentDate_Loaded(object sender, RoutedEventArgs e)
        {
            SetBlackOutDate();
        }

        private void PatentDate_DisplayDateChanged(object sender, CalendarDateChangedEventArgs e)
        {
            SetBlackOutDate();
        }

        private void dataGridPatents_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if(dataGridPatents.SelectedItem != null)
            {
                Patents patents = (Patents)dataGridPatents.SelectedItem;
                patentsID = patents.PatentsId;
                PatentClassifies.SelectedItem = dataContext.PatentClassify.Single(p => p.PatentClassifyID.Equals(patents.PatentClassifyID));
                PatentNo.Text = patents.PatentNo;
                PatentName.Text = patents.PatentName;
                PatentDepartment.Text = patents.PatendDepartment;
                PatentDate.SelectedDate = patents.PatentDate;
                PatentDate.DisplayDate = (DateTime)patents.PatentDate;
                Note.Text = patents.Note;
            }
        }

        private void buttonDelete_Click(object sender, RoutedEventArgs e)
        {
            if (patentsID == 0)
            {
                MessageBox.Show("请选择知识产权项！", "错误");
                return;
            }
            if (MessageBox.Show("该项知识产权项将被删除！确认要删除该项知识产权信息？", "警告", MessageBoxButton.YesNo) == MessageBoxResult.No)
            {
                return;
            }
            dataContext = new DataClassesProjectClassifyDataContext();
            Patents patents = dataContext.Patents.Single(p => p.PatentsId.Equals(patentsID));
            dataContext.Patents.DeleteOnSubmit(patents);
            dataContext.SubmitChanges();
            dataContext = new DataClassesProjectClassifyDataContext();
            dataGridPatents.DataContext = dataContext.Patents.Where(p => p.ProjectID.Equals(projectID));
            ((MainWindow)(this.Owner)).DialogR = true;
            Clear();
        }

        private void Clear()
        {
            patentsID = 0;
            PatentClassifies.SelectedIndex = 0;
            PatentNo.Text = "";
            PatentName.Text = "";
            PatentDepartment.Text = "";
            PatentDate.SelectedDate = DateTime.Now;
            PatentDate.DisplayDate = DateTime.Now;
            Note.Text = "";
        }

        private void buttonSubmit_Click(object sender, RoutedEventArgs e)
        {
            if(PatentClassifies.SelectedItem == null)
            {
                MessageBox.Show("请选择知识产权项类别！", "错误");
                return;
            }
            if(PatentDate.SelectedDate == null)
            {
                MessageBox.Show("请选择颁布时间！", "错误");
                return;
            }
            Patents patents = new Patents();
            patents.ProjectID = projectID;
            patents.PatentClassifyID = ((PatentClassify)PatentClassifies.SelectedItem).PatentClassifyID;
            patents.PatentNo = PatentNo.Text.Trim();
            patents.PatentName = PatentName.Text.Trim();
            patents.PatendDepartment = PatentDepartment.Text.Trim();
            patents.PatentDate = PatentDate.SelectedDate;
            patents.Note = Note.Text.Trim();
            dataContext = new DataClassesProjectClassifyDataContext();
            dataContext.Patents.InsertOnSubmit(patents);
            dataContext.SubmitChanges();
            dataGridPatents.DataContext = dataContext.Patents.Where(p => p.ProjectID.Equals(projectID));
            ((MainWindow)(this.Owner)).DialogR = true;
        }

        private void buttonPreYear_Click(object sender, RoutedEventArgs e)
        {
            PatentDate.DisplayDate = PatentDate.DisplayDate.AddYears(-1);
            if (PatentDate.SelectedDate != null)
            {
                PatentDate.SelectedDate = ((DateTime)(PatentDate.SelectedDate)).AddYears(-1);
            }
        }

        private void buttonNextYear_Click(object sender, RoutedEventArgs e)
        {
            PatentDate.DisplayDate = PatentDate.DisplayDate.AddYears(1);
            if (PatentDate.SelectedDate != null)
            {
                PatentDate.SelectedDate = ((DateTime)(PatentDate.SelectedDate)).AddYears(1);
            }
        }
	}
}