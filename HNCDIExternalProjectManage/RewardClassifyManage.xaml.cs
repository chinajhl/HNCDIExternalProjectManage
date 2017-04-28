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
	/// RewardClassifyManage.xaml 的交互逻辑
	/// </summary>
	public partial class RewardClassifyManage : Window
	{
        DataClassesProjectClassifyDataContext dataContext;
		public RewardClassifyManage()
		{
			this.InitializeComponent();
			
			// 在此点之下插入创建对象所需的代码。
		}

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            dataContext = new DataClassesProjectClassifyDataContext();
            Classifies.DisplayMemberPath = "RewardClassify1";
            Classifies.SelectedValuePath = "RewardClassifyID";
            Classifies.DataContext = dataContext.RewardClassify;
        }

        private void buttonAdd_Click(object sender, RoutedEventArgs e)
        {
            if (String.IsNullOrEmpty(ClassifyName.Text))
            {
                MessageBox.Show("请输入奖项等别", "错误");
                return;
            }
            else
            {
                var pc = dataContext.RewardClassify.Count(p => p.RewardClassify1.Equals(ClassifyName.Text.Trim()));
                if (pc > 0)
                {
                    MessageBox.Show("已存在相同奖项等别", "错误");
                    return;
                }
                RewardClassify rewardClassify = new RewardClassify();
                rewardClassify.RewardClassify1 = ClassifyName.Text.Trim();
                dataContext.RewardClassify.InsertOnSubmit(rewardClassify);
                dataContext.SubmitChanges();
                dataContext = new DataClassesProjectClassifyDataContext();
                Classifies.DataContext = dataContext.RewardClassify;
            }
        }
	}
}