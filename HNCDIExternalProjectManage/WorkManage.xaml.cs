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
	/// WorkManage.xaml 的交互逻辑
	/// </summary>
	public partial class WorkManage : Window
	{
        DataClassesProjectClassifyDataContext dataContext;

        private int projectID;

        public int ProjectID
        {
            get { return projectID; }
            set { projectID = value; }
        }

        private int ID;
		public WorkManage()
		{
			this.InitializeComponent();
			
			// 在此点之下插入创建对象所需的代码。
		}

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            dataContext = new DataClassesProjectClassifyDataContext();
            Workers.DisplayMemberPath = "WorkerName";
            Workers.SelectedValuePath = "Id";
            Workers.DataContext = dataContext.TeamWorkers.Where(t => t.ProjectID.Equals(projectID));
        }

        private void Workers_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if(Workers.SelectedItem != null)
            {
                TeamWorkers teamWorkers = (TeamWorkers)Workers.SelectedItem;
                ID = teamWorkers.Id;
                Worker.Text = teamWorkers.WorkerName;
            }
            else
            {
                Clear();
            }
        }

        private void Clear()
        {
            ID = 0;
            Worker.Text = "";
        }

        private void buttonRemove_Click(object sender, RoutedEventArgs e)
        {
            if (ID == 0)
            {
                MessageBox.Show("请选择团队成员！", "错误");
                return;
            }
            if (MessageBox.Show("该团队成员将被删除！确认要删除该团队成员信息？", "警告", MessageBoxButton.YesNo) == MessageBoxResult.No)
            {
                return;
            }
            dataContext = new DataClassesProjectClassifyDataContext();
            TeamWorkers teamWorkers = (TeamWorkers)Workers.SelectedItem;
            dataContext.TeamWorkers.DeleteOnSubmit(teamWorkers);
            dataContext.SubmitChanges();
            dataContext = new DataClassesProjectClassifyDataContext();
            Workers.DataContext = dataContext.TeamWorkers.Where(t => t.ProjectID.Equals(projectID));
            ((MainWindow)(this.Owner)).DialogR = true;
            Clear();
        }

        private void buttonAdd_Click(object sender, RoutedEventArgs e)
        {
            if (String.IsNullOrEmpty(Worker.Text))
            {
                MessageBox.Show("请输入团队成员姓名！", "错误");
                return;
            }
            dataContext = new DataClassesProjectClassifyDataContext();
            TeamWorkers teamWorkers = new TeamWorkers();
            teamWorkers.ProjectID = projectID;
            teamWorkers.WorkerName = Worker.Text;
            dataContext.TeamWorkers.InsertOnSubmit(teamWorkers);
            dataContext.SubmitChanges();
            Workers.DataContext = dataContext.TeamWorkers.Where(t => t.ProjectID.Equals(projectID));
            ((MainWindow)(this.Owner)).DialogR = true;
        }

	}
}