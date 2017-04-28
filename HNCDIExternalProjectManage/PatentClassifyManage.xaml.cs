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
	/// PatentClassifyManage.xaml 的交互逻辑
	/// </summary>
	public partial class PatentClassifyManage : Window
	{
        DataClassesProjectClassifyDataContext dataContext;
		public PatentClassifyManage()
		{
			this.InitializeComponent();
			
			// 在此点之下插入创建对象所需的代码。
		}

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            dataContext = new DataClassesProjectClassifyDataContext();
            Classifies.DisplayMemberPath = "PatentClassify1";
            Classifies.SelectedValuePath = "PatentClassifyID";
            Classifies.DataContext = dataContext.PatentClassify;
        }

        private void buttonAdd_Click(object sender, RoutedEventArgs e)
        {
            if (String.IsNullOrEmpty(ClassifyName.Text))
            {
                MessageBox.Show("请输入知识产权类别名称", "错误");
                return;
            }
            else
            {
                var pc = dataContext.PatentClassify.Count(p => p.PatentClassify1.Equals(ClassifyName.Text.Trim()));
                if (pc > 0)
                {
                    MessageBox.Show("已存在相同知识产权类别", "错误");
                    return;
                }
                PatentClassify patentClassify = new PatentClassify();
                patentClassify.PatentClassify1 = ClassifyName.Text.Trim();
                dataContext.PatentClassify.InsertOnSubmit(patentClassify);
                dataContext.SubmitChanges();
                dataContext = new DataClassesProjectClassifyDataContext();
                Classifies.DataContext = dataContext.PatentClassify;
            }
        }
	}
}