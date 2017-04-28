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
	/// TeamDepartmentManage.xaml 的交互逻辑
	/// </summary>
	public partial class TeamDepartmentManage : Window
	{
        DataClassesProjectClassifyDataContext dataContext;

        private int projectID;

        public int ProjectID
        {
            get { return projectID; }
            set { projectID = value; }
        }

        private int ID;
		public TeamDepartmentManage()
		{
			this.InitializeComponent();
			
			// 在此点之下插入创建对象所需的代码。
		}

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            dataContext = new DataClassesProjectClassifyDataContext();
            TeamDepartment.DisplayMemberPath = "Department";
            TeamDepartment.SelectedValuePath = "Id";
            TeamDepartment.DataContext = dataContext.TeamDepartments.Where(t => t.ProjectID.Equals(projectID));
        }

        private void TeamDepartments_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if(TeamDepartment.SelectedItem != null)
            {
                TeamDepartments teamDepartments = (TeamDepartments)TeamDepartment.SelectedItem;
                ID = teamDepartments.Id;
                Department.Text = teamDepartments.Department;
            }
            else
            {
                Clear();
            }
        }

        private void Clear()
        {
            ID = 0;
            Department.Text = "";
        }

        private void buttonRemove_Click(object sender, RoutedEventArgs e)
        {
            if (ID == 0)
            {
                MessageBox.Show("请选择协作单位！", "错误");
                return;
            }
            if (MessageBox.Show("该协作单位将被删除！确认要删除该协作单位信息？", "警告", MessageBoxButton.YesNo) == MessageBoxResult.No)
            {
                return;
            }
            dataContext = new DataClassesProjectClassifyDataContext();
            TeamDepartments teamDepartments = (TeamDepartments)TeamDepartment.SelectedItem;
            dataContext.TeamDepartments.DeleteOnSubmit(teamDepartments);
            dataContext.SubmitChanges();
            dataContext = new DataClassesProjectClassifyDataContext();
            TeamDepartment.DataContext = dataContext.TeamDepartments.Where(t => t.ProjectID.Equals(projectID));
            ((MainWindow)(this.Owner)).DialogR = true;
            Clear();
        }

        private void buttonAdd_Click(object sender, RoutedEventArgs e)
        {
            if(String.IsNullOrEmpty(Department.Text))
            {
                MessageBox.Show("请输入协作单位名称！", "错误");
                return;
            }
            dataContext = new DataClassesProjectClassifyDataContext();
            TeamDepartments teamDepartments = new TeamDepartments();
            teamDepartments.ProjectID = projectID;
            teamDepartments.Department = Department.Text;
            dataContext.TeamDepartments.InsertOnSubmit(teamDepartments);
            dataContext.SubmitChanges();
            TeamDepartment.DataContext = dataContext.TeamDepartments.Where(t => t.ProjectID.Equals(projectID));
            ((MainWindow)(this.Owner)).DialogR = true;
        }
	}
}