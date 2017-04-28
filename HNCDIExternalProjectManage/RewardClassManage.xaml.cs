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
	/// RewardClassManage.xaml 的交互逻辑
	/// </summary>
	public partial class RewardClassManage : Window
	{
        DataClassesProjectClassifyDataContext dataContext;
		public RewardClassManage()
		{
			this.InitializeComponent();
			
			// 在此点之下插入创建对象所需的代码。
		}

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            dataContext = new DataClassesProjectClassifyDataContext();
            Classifies.DisplayMemberPath = "RewardClass1";
            Classifies.SelectedValuePath = "RewardClassID";
            Classifies.DataContext = dataContext.RewardClass;
        }

        private void buttonAdd_Click(object sender, RoutedEventArgs e)
        {
            if (String.IsNullOrEmpty(ClassifyName.Text))
            {
                MessageBox.Show("请输入奖项级别", "错误");
                return;
            }
            else
            {
                var pc = dataContext.RewardClass.Count(p => p.RewardClass1.Equals(ClassifyName.Text.Trim()));
                if (pc > 0)
                {
                    MessageBox.Show("已存在相同奖项级别", "错误");
                    return;
                }
                RewardClass rewardClass = new RewardClass();
                rewardClass.RewardClass1 = ClassifyName.Text.Trim();
                dataContext.RewardClass.InsertOnSubmit(rewardClass);
                dataContext.SubmitChanges();
                dataContext = new DataClassesProjectClassifyDataContext();
                Classifies.DataContext = dataContext.RewardClass;
            }
        }
	}
}