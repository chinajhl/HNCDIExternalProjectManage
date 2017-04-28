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
	/// ProjectClassifyManage.xaml 的交互逻辑
	/// </summary>
	public partial class ProjectClassifyManage : Window
	{
        DataClassesProjectClassifyDataContext dataContext;
		public ProjectClassifyManage()
		{
			this.InitializeComponent();
			
			// 在此点之下插入创建对象所需的代码。
		}

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            dataContext = new DataClassesProjectClassifyDataContext();
            Classifies.DisplayMemberPath = "ProjectClassify1";
            Classifies.SelectedValuePath = "ClassifyID";
            Classifies.DataContext = dataContext.ProjectClassify;
        }

        private void buttonAdd_Click(object sender, RoutedEventArgs e)
        {
            if(String.IsNullOrEmpty(ClassifyName.Text))
            {
                MessageBox.Show("请输入项目类别名称", "错误");
                return;
            }
            else
            {
                var pc = dataContext.ProjectClassify.Count(p => p.ProjectClassify1.Equals(ClassifyName.Text.Trim()));
                if(pc > 0)
                {
                    MessageBox.Show("已存在相同项目类别名称", "错误");
                    return;
                }
                ProjectClassify projectClassify = new ProjectClassify();
                projectClassify.ProjectClassify1 = ClassifyName.Text.Trim();
                dataContext.ProjectClassify.InsertOnSubmit(projectClassify);
                dataContext.SubmitChanges();
                dataContext = new DataClassesProjectClassifyDataContext();
                Classifies.DataContext = dataContext.ProjectClassify;
                ((MainWindow)(this.Owner)).DialogR = true;
            }
        }
	}
}