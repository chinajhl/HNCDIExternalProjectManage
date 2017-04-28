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
	/// FundClassifyManage.xaml 的交互逻辑
	/// </summary>
	public partial class FundClassifyManage : Window
	{
        DataClassesProjectClassifyDataContext dataContext;
		public FundClassifyManage()
		{
			this.InitializeComponent();
			
			// 在此点之下插入创建对象所需的代码。
		}

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            dataContext = new DataClassesProjectClassifyDataContext();
            Classifies.DisplayMemberPath = "FundClassify1";
            Classifies.SelectedValuePath = "FundClassifyID";
            Classifies.DataContext = dataContext.FundClassify;
        }

        private void buttonAdd_Click(object sender, RoutedEventArgs e)
        {
            if (String.IsNullOrEmpty(ClassifyName.Text))
            {
                MessageBox.Show("请输入经费类别", "错误");
                return;
            }
            else
            {
                var pc = dataContext.FundClassify.Count(p => p.FundClassify1.Equals(ClassifyName.Text.Trim()));
                if (pc > 0)
                {
                    MessageBox.Show("已存在相同经费类别", "错误");
                    return;
                }
                FundClassify fundClassify = new FundClassify();
                fundClassify.FundClassify1 = ClassifyName.Text.Trim();
                fundClassify.IncomeOrPay = IsInCome.IsChecked;
                dataContext.FundClassify.InsertOnSubmit(fundClassify);
                dataContext.SubmitChanges();
                dataContext = new DataClassesProjectClassifyDataContext();
                Classifies.DataContext = dataContext.FundClassify;
            }
        }

        private void Classifies_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if(Classifies.SelectedItem != null)
            {
                FundClassify fundClassify = (FundClassify)Classifies.SelectedItem;
                ClassifyName.Text = fundClassify.FundClassify1;
                IsInCome.IsChecked = fundClassify.IncomeOrPay;
            }
        }
	}
}