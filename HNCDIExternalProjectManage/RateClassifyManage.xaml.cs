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
	/// RateClassifyManage.xaml 的交互逻辑
	/// </summary>
	public partial class RateClassifyManage : Window
	{
        DataClassesProjectClassifyDataContext dataContext;
		public RateClassifyManage()
		{
			this.InitializeComponent();
			
			// 在此点之下插入创建对象所需的代码。
		}

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            dataContext = new DataClassesProjectClassifyDataContext();
            Classifies.DisplayMemberPath = "RateClassify1";
            Classifies.SelectedValuePath = "RateClassifyID";
            Classifies.DataContext = dataContext.RateClassify;
        }

        private void buttonAdd_Click(object sender, RoutedEventArgs e)
        {
            if (String.IsNullOrEmpty(ClassifyName.Text))
            {
                MessageBox.Show("请输入鉴定等级", "错误");
                return;
            }
            else
            {
                var pc = dataContext.RateClassify.Count(p => p.RateClassify1.Equals(ClassifyName.Text.Trim()));
                if (pc > 0)
                {
                    MessageBox.Show("已存在相同鉴定等级", "错误");
                    return;
                }
                RateClassify rateClassify = new RateClassify();
                rateClassify.RateClassify1 = ClassifyName.Text.Trim();
                dataContext.RateClassify.InsertOnSubmit(rateClassify);
                dataContext.SubmitChanges();
                dataContext = new DataClassesProjectClassifyDataContext();
                Classifies.DataContext = dataContext.RateClassify;
            }
        }
	}
}