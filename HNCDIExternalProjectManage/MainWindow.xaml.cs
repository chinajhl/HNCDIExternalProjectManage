using System;
using System.Collections.Generic;
using System.Data.Linq;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;

namespace HNCDIExternalProjectManage
{
    /// <summary>
    ///     MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        private DataClassesProjectClassifyDataContext dataContext;
        public bool DialogR;
        private bool isCreateOrUpdateProject;
        private List<int> LinkProject;
        private List<int> LinkProjectClassify;
        private List<int> LinkProjectForSearch;
        private int projectClassifyID;

        private List<string> searchText;

        private string selectedType = "";

        //树控件选择的节点类型
        //项目类别ID
        private bool subProjectCanCreate;

        public MainWindow()
        {
            InitializeComponent();
        }

        public int ProjectID { get; set; }

        /// <summary>
        ///     创建当前项目路径列表
        /// </summary>
        /// <param name="pid">项目ID</param>
        private void BuildLinkProjectList(int pid)
        {
            LinkProject.Clear();
            LinkProject.Add(pid);
            var pb = dataContext.ProjectBase.Single(p => p.ProjectId.Equals(pid));
            while (pb.ParentID != null)
            {
                pb = dataContext.ProjectBase.Single(p => p.ProjectId.Equals(pb.ParentID));
                LinkProject.Add(pb.ProjectId);
            }
        }

        private void buildSearchTreeProjectBase(TreeViewItem tvi)
        {
            if (tvi.ItemContainerGenerator.Status != GeneratorStatus.ContainersGenerated)
            {
                tvi.UpdateLayout();
            }
            var itemsControl = tvi as ItemsControl;
            foreach (ProjectBase oit in tvi.Items)
            {
                var container = itemsControl.ItemContainerGenerator.ContainerFromItem(oit) as TreeViewItem;
                if (container != null)
                {
                    if (!container.HasItems)
                    {
                        if (LinkProjectForSearch.Contains(oit.ProjectId))
                        {
                            if (container != null)
                            {
                                container.IsExpanded = true;
                                var b = new SolidColorBrush(Colors.Blue);
                                var f = new SolidColorBrush(Colors.White);
                                container.Background = b;
                                container.Foreground = f;
                            }
                        }
                    }
                    else
                    {
                        if (LinkProjectForSearch.Contains(oit.ProjectId))
                        {
                            container.IsExpanded = true;
                            if (container.ItemContainerGenerator.Status != GeneratorStatus.ContainersGenerated)
                            {
                                container.UpdateLayout();
                            }
                            foreach (var str in searchText)
                            {
                                if (oit.ProjectName.Contains(str))
                                {
                                    var b = new SolidColorBrush(Colors.Blue);
                                    var f = new SolidColorBrush(Colors.White);
                                    container.Background = b;
                                    container.Foreground = f;
                                    break;
                                }
                            }
                            buildSearchTreeProjectBase(container);
                        }
                    }
                }
            }
        }

        /// <summary>
        ///     建立搜索后的TreeView
        /// </summary>
        private void buildSearchTreeProjectClassify()
        {
            var itemsControl = MainTreeView as ItemsControl;
            foreach (ProjectClassify oit in MainTreeView.Items)
            {
                if (LinkProjectClassify.Contains(oit.ClassifyId))
                {
                    var container = itemsControl.ItemContainerGenerator.ContainerFromItem(oit) as TreeViewItem;
                    if (container != null)
                    {
                        container.IsExpanded = true;
                        if (container.ItemContainerGenerator.Status != GeneratorStatus.ContainersGenerated)
                        {
                            container.UpdateLayout();
                        }
                        buildSearchTreeProjectBase(container);
                    }
                }
            }
        }

        private void buttonContractManage_Click(object sender, RoutedEventArgs e)
        {
            ProjectContractsManage();
        }

        private void buttonFundManage_Click(object sender, RoutedEventArgs e)
        {
            FundManage();
        }

        private void buttonNewProjectBase_Click(object sender, RoutedEventArgs e)
        {
            NewProjectBase();
        }

        private void buttonPatentManage_Click(object sender, RoutedEventArgs e)
        {
            PatentsManage();
        }

        private void buttonRateManage_Click(object sender, RoutedEventArgs e)
        {
            RatesManage();
        }

        private void buttonResultManage_Click(object sender, RoutedEventArgs e)
        {
            ResultsManage();
        }

        private void buttonReward_Click(object sender, RoutedEventArgs e)
        {
            RewardsManage();
        }

        /// <summary>
        ///     搜索
        /// </summary>
        private void buttonSearch_Click()
        {
            if (string.IsNullOrEmpty(searchTextBox.Text))
            {
                dataContext = new DataClassesProjectClassifyDataContext();
                MainTreeView.DataContext = dataContext.ProjectClassify;
                WindowDataBind();
                return;
            }
            dataContext = new DataClassesProjectClassifyDataContext();
            var searchPredicate = PredicateExtensions.True<ProjectBase>();
            searchText = new List<string>();
            searchText = Regex.Split(searchTextBox.Text, @"[' ']+").ToList();
            //foreach (string str in searchText)
            //{
            //    searchPredicate = searchPredicate.Or(p => p.ProjectName.Contains(str));
            //}
            //var pb = dataContext.ProjectBase.Where(searchPredicate);
            var pb = dataContext.ProjectBase.Where(p => p.ProjectName.Contains(searchText[0]));
            var list = pb.ToList();
            if (searchText.Count > 1)
            {
                foreach (var str in searchText)
                {
                    if (str == searchText[0]) continue;
                    pb = pb.Where(c => c.ProjectName.Contains(str));
                    list = pb.ToList();
                    //合并结果示例
                    //pb = pb.Union<ProjectBase>(dataContext.ProjectBase.Where(p => p.ProjectName.Contains(str)));
                }
            }

            //创建符合条件的项目的所有关联项目的项目树ProjectID列表
            LinkProjectForSearch = new List<int>();
            LinkProjectClassify = new List<int>();
            foreach (var p in pb)
            {
                var pp = p;
                if (!LinkProjectForSearch.Contains(pp.ProjectId))
                {
                    LinkProjectForSearch.Add(pp.ProjectId);
                    while (pp.ParentID != null)
                    {
                        pp = dataContext.ProjectBase.Single(ppb => ppb.ProjectId.Equals(pp.ParentID));
                        if (!LinkProjectForSearch.Contains(pp.ProjectId))
                        {
                            LinkProjectForSearch.Add(pp.ProjectId);
                        }
                    }
                }
                if (pp.ProjectClassifyID != null && !LinkProjectClassify.Contains((int) pp.ProjectClassifyID))
                {
                    LinkProjectClassify.Add((int) pp.ProjectClassifyID);
                }
            }
            //dataContext = new DataClassesProjectClassifyDataContext();
            MainTreeView.DataContext = dataContext.ProjectClassify;
            buildSearchTreeProjectClassify();
            if (LinkProjectForSearch.Count > 0)
            {
                ProjectID = LinkProjectForSearch[0];
                WindowDataBind();
            }
        }

        private void buttonSearch_Click_1(object sender, RoutedEventArgs e)
        {
            buttonSearch_Click();
        }

        private void buttonTeamManage_Click(object sender, RoutedEventArgs e)
        {
            TeamDepartsManage();
        }

        private void buttonUpdateProjectBase_Click(object sender, RoutedEventArgs e)
        {
            UpdateProjectBase();
        }

        private void buttonWorkerManage_Click(object sender, RoutedEventArgs e)
        {
            WorkersManage();
        }

        private void datagridContractIn_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = e.Row.GetIndex() + 1;
        }

        private void datagridContractPay_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = e.Row.GetIndex() + 1;
        }

        private void dataGridFund_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = e.Row.GetIndex() + 1;
        }

        private void dataGridPatents_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = e.Row.GetIndex() + 1;
        }

        private void dataGridResults_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = e.Row.GetIndex() + 1;
        }

        private void dataGridRewards_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = e.Row.GetIndex() + 1;
        }

        private void dataGrigRate_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = e.Row.GetIndex() + 1;
        }

        private void dataGrigWorkers_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = e.Row.GetIndex() + 1;
        }

        private void dataTeamDepartment_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = e.Row.GetIndex() + 1;
        }

        private void FundManage()
        {
            if (ProjectID == 0)
            {
                MessageBox.Show("请选择项目！", "错误");
                return;
            }
            var fundsManage = new FundsManage();
            fundsManage.ProjectID = ProjectID;
            fundsManage.Owner = this;
            fundsManage.ShowDialog();
            RefreshData();
        }

        private void MainForm_Loaded(object sender, RoutedEventArgs e)
        {
            dataContext = new DataClassesProjectClassifyDataContext();
            MainTreeView.DataContext = dataContext.ProjectClassify;
            LinkProject = new List<int>();
            //var query = dataContext.ProjectBase.Where(p => p.ProjectName.)
        }

        private void MainTreeView_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            switch (MainTreeView.SelectedItem.ToString())
            {
                case "HNCDIExternalProjectManage.ProjectClassify":
                    var projectClassify = (ProjectClassify) MainTreeView.SelectedItem;
                    selectedType = "ProjectClassify";
                    if (!isCreateOrUpdateProject)
                    {
                        ProjectID = 0;
                        projectClassifyID = projectClassify.ClassifyId;
                        LinkProject.Clear();
                    }
                    else
                    {
                        isCreateOrUpdateProject = false;
                    }
                    WindowDataBind();

                    subProjectCanCreate = false;
                    break;

                case "HNCDIExternalProjectManage.ProjectBase":
                    var projectBase = (ProjectBase) MainTreeView.SelectedItem;
                    selectedType = "ProjectBase";
                    if (!isCreateOrUpdateProject)
                    {
                        ProjectID = projectBase.ProjectId;
                        LinkProject.Add(projectBase.ProjectId);
                        WindowDataBind();
                        subProjectCanCreate = true;
                        int parentID;
                        while (projectBase.ParentID != null)
                        {
                            //子项目
                            parentID = Convert.ToInt32(projectBase.ParentID);
                            LinkProject.Add(parentID);
                            projectBase = dataContext.ProjectBase.Single(p => p.ProjectId.Equals(parentID));
                        }
                        projectClassifyID = Convert.ToInt32(projectBase.ProjectClassifyID);
                    }
                    else
                    {
                        isCreateOrUpdateProject = false;
                    }
                    break;

                default:
                    break;
            }
        }

        private void menuitemContractManage_Click(object sender, RoutedEventArgs e)
        {
            ProjectContractsManage();
        }

        private void menuitemExit_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        private void menuitemExportData_Click(object sender, RoutedEventArgs e)
        {
            var exportToExcel = new ExportToExcel();
            exportToExcel.ShowDialog();
        }

        private void menuitemMoneyDetail_Click(object sender, RoutedEventArgs e)
        {
            if (ProjectID == 0)
            {
                MessageBox.Show("请选择要导出经费明细的项目！", "错误");
                return;
            }
            var mdToExcel = new MoneyDetailToExcel();
            mdToExcel.Owner = this;
            mdToExcel.ProjectID = ProjectID;
            mdToExcel.ShowDialog();
        }

        private void menuitemNewFund_Click(object sender, RoutedEventArgs e)
        {
            FundManage();
        }

        private void menuitemNewFundClassify_Click(object sender, RoutedEventArgs e)
        {
            var fundClassifyManage = new FundClassifyManage();
            fundClassifyManage.ShowDialog();
        }

        private void menuitemNewPatentClassify_Click(object sender, RoutedEventArgs e)
        {
            var patentClassifyManage = new PatentClassifyManage();
            patentClassifyManage.ShowDialog();
        }

        private void menuitemNewProjectBase_Click(object sender, RoutedEventArgs e)
        {
            NewProjectBase();
        }

        //可否创建子项目
        //private DataSetMain dataSetMain;
        private void menuitemNewProjectClassify_Click(object sender, RoutedEventArgs e)
        {
            var projectClassifyManage = new ProjectClassifyManage();
            projectClassifyManage.Owner = this;
            projectClassifyManage.ShowDialog();
            if (DialogR)
            {
                dataContext = new DataClassesProjectClassifyDataContext();
                MainTreeView.DataContext = dataContext.ProjectClassify;
                WindowDataBind();
                var pr = dataContext.ProjectBase.OrderByDescending(p => p.ProjectId).FirstOrDefault();
                ProjectID = pr.ProjectId;
                BuildLinkProjectList(ProjectID);
                projectClassifyID = (int) pr.ProjectClassifyID;
                SetTreeViewFocus();
                DialogR = false;
            }
        }

        private void menuitemNewRateClassify_Click(object sender, RoutedEventArgs e)
        {
            var rateClassifyManage = new RateClassifyManage();
            rateClassifyManage.ShowDialog();
        }

        private void menuitemNewRewardClass_Click(object sender, RoutedEventArgs e)
        {
            var rewardClassManage = new RewardClassManage();
            rewardClassManage.ShowDialog();
        }

        private void menuitemNewRewardClassify_Click(object sender, RoutedEventArgs e)
        {
            var rewardClassifyManage = new RewardClassifyManage();
            rewardClassifyManage.ShowDialog();
        }

        private void menuitemPatentManage_Click(object sender, RoutedEventArgs e)
        {
            PatentsManage();
        }

        private void menuitemRateManage_Click(object sender, RoutedEventArgs e)
        {
            RatesManage();
        }

        private void menuitemResultManage_Click(object sender, RoutedEventArgs e)
        {
            ResultsManage();
        }

        private void menuitemRewardManage_Click(object sender, RoutedEventArgs e)
        {
            RewardsManage();
        }

        private void menuitemTeamManage_Click(object sender, RoutedEventArgs e)
        {
            TeamDepartsManage();
        }

        private void menuitemUpdateProjectBase_Click(object sender, RoutedEventArgs e)
        {
            UpdateProjectBase();
        }

        private void menuitemWorkerManage_Click(object sender, RoutedEventArgs e)
        {
            WorkersManage();
        }

        private void NewProjectBase()
        {
            isCreateOrUpdateProject = true;
            if (ProjectID != 0)
            {
                var projectBase = dataContext.ProjectBase.Single(p => p.ProjectId.Equals(ProjectID));
                if (MessageBox.Show("新建 " + projectBase.ProjectName + " 下的子项目？", "温馨提示", MessageBoxButton.OKCancel) ==
                    MessageBoxResult.OK)
                {
                    var newProjectBase = new NewProjectBase();
                    newProjectBase.Owner = this;
                    newProjectBase.IsCreate = true;
                    newProjectBase.IsSubProject = true;
                    newProjectBase.ParentID = ProjectID;
                    newProjectBase.ShowDialog();
                    if (DialogR)
                    {
                        dataContext = new DataClassesProjectClassifyDataContext();
                        dataContext.Refresh(RefreshMode.OverwriteCurrentValues, dataContext);
                        MainTreeView.DataContext = dataContext.ProjectClassify;
                        WindowDataBind();
                        var pr = dataContext.ProjectBase.OrderByDescending(p => p.ProjectId).FirstOrDefault();
                        ProjectID = pr.ProjectId;
                        BuildLinkProjectList(ProjectID);
                        while (pr.ParentID != null)
                        {
                            pr = dataContext.ProjectBase.Single(p => p.ProjectId.Equals(pr.ParentID));
                        }
                        projectClassifyID = (int) pr.ProjectClassifyID;
                        SetTreeViewFocus();
                        DialogR = false;
                    }
                }
            }
            else
            {
                var newProjectBase = new NewProjectBase();
                newProjectBase.Owner = this;
                newProjectBase.IsCreate = true;
                newProjectBase.IsSubProject = false;
                newProjectBase.ProjectClassifyID = projectClassifyID;
                newProjectBase.ShowDialog();
                if (DialogR)
                {
                    dataContext = new DataClassesProjectClassifyDataContext();
                    // dataContext.Refresh(System.Data.Linq.RefreshMode.OverwriteCurrentValues, dataContext.ProjectBase);
                    //List<ProjectBase> lProjectBase = dataContext.ProjectBase.ToList<ProjectBase>();
                    MainTreeView.DataContext = dataContext.ProjectClassify;
                    WindowDataBind();
                    var pr = dataContext.ProjectBase.OrderByDescending(p => p.ProjectId).FirstOrDefault();
                    ProjectID = pr.ProjectId;
                    BuildLinkProjectList(ProjectID);
                    projectClassifyID = (int) pr.ProjectClassifyID;
                    SetTreeViewFocus();
                    DialogR = false;
                }
            }
        }

        private void PatentsManage()
        {
            if (ProjectID == 0)
            {
                MessageBox.Show("请选择项目！", "错误");
                return;
            }
            var patentManage = new PatentManage();
            patentManage.ProjectID = ProjectID;
            patentManage.Owner = this;
            patentManage.ShowDialog();
            RefreshData();
        }

        private void ProjectContractsManage()
        {
            if (ProjectID == 0)
            {
                MessageBox.Show("请选择项目！", "错误");
                return;
            }
            var contractManage = new ContractManage();
            contractManage.ProjectID = ProjectID;
            contractManage.Owner = this;
            contractManage.ShowDialog();
            RefreshData();
        }

        private void RatesManage()
        {
            if (ProjectID == 0)
            {
                MessageBox.Show("请选择项目！", "错误");
                return;
            }
            var rateManage = new RateManage();
            rateManage.ProjectID = ProjectID;
            rateManage.Owner = this;
            rateManage.ShowDialog();
            RefreshData();
        }

        private void RefreshData()
        {
            if (DialogR)
            {
                dataContext = new DataClassesProjectClassifyDataContext();
                WindowDataBind();
                DialogR = false;
            }
        }

        private void ResultsManage()
        {
            if (ProjectID == 0)
            {
                MessageBox.Show("请选择项目！", "错误");
                return;
            }
            var resultManage = new ResultManage();
            resultManage.ProjectID = ProjectID;
            resultManage.Owner = this;
            resultManage.ShowDialog();
            RefreshData();
        }

        private void RewardsManage()
        {
            if (ProjectID == 0)
            {
                MessageBox.Show("请选择项目！", "错误");
                return;
            }
            var rewardManage = new RewardManage();
            rewardManage.ProjectID = ProjectID;
            rewardManage.Owner = this;
            rewardManage.ShowDialog();
            RefreshData();
        }

        private void searchTextBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                searchTextBox.MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
            }
        }

        private void SetProjectSelect(TreeViewItem tvi)
        {
            if (tvi.ItemContainerGenerator.Status != GeneratorStatus.ContainersGenerated)
            {
                tvi.UpdateLayout();
            }
            var itemsControl = tvi as ItemsControl;
            foreach (ProjectBase oit in tvi.Items)
            {
                var container = itemsControl.ItemContainerGenerator.ContainerFromItem(oit) as TreeViewItem;
                if (container != null)
                {
                    if (!container.HasItems)
                    {
                        if (oit.ProjectId == ProjectID)
                        {
                            if (container != null)
                            {
                                container.IsExpanded = true;
                                container.IsSelected = true;
                                return;
                            }
                        }
                    }
                    else
                    {
                        if (LinkProject.Contains(oit.ProjectId))
                        {
                            container.IsExpanded = true;
                            if (container.ItemContainerGenerator.Status != GeneratorStatus.ContainersGenerated)
                            {
                                container.UpdateLayout();
                            }
                            SetProjectSelect(container);
                        }
                    }
                }
            }
        }

        private void SetTreeViewFocus()
        {
            var itemsControl = MainTreeView as ItemsControl;
            foreach (ProjectClassify oit in MainTreeView.Items)
            {
                if (oit.ClassifyId == projectClassifyID)
                {
                    var container = itemsControl.ItemContainerGenerator.ContainerFromItem(oit) as TreeViewItem;
                    if (container != null)
                    {
                        container.IsExpanded = true;
                        if (container.ItemContainerGenerator.Status != GeneratorStatus.ContainersGenerated)
                        {
                            container.UpdateLayout();
                        }
                        SetProjectSelect(container);
                    }
                }
            }
        }

        private void TeamDepartsManage()
        {
            if (ProjectID == 0)
            {
                MessageBox.Show("请选择项目！", "错误");
                return;
            }
            var teamDepartmentManage = new TeamDepartmentManage();
            teamDepartmentManage.ProjectID = ProjectID;
            teamDepartmentManage.Owner = this;
            teamDepartmentManage.ShowDialog();
            RefreshData();
        }

        private void UpdateProjectBase()
        {
            isCreateOrUpdateProject = true;
            if (ProjectID == 0)
            {
                MessageBox.Show("请选择项目！", "错误");
                return;
            }
            dataContext = new DataClassesProjectClassifyDataContext();
            var pb = dataContext.ProjectBase.Single(p => p.ProjectId.Equals(ProjectID));

            var newProjectBase = new NewProjectBase();
            newProjectBase.Owner = this;
            newProjectBase.IsCreate = false;
            if (pb.ParentID != null)
            {
                newProjectBase.ParentID = (int) pb.ParentID;
            }
            newProjectBase.ProjectID = ProjectID;
            newProjectBase.ShowDialog();
            if (DialogR)
            {
                dataContext = new DataClassesProjectClassifyDataContext();
                //dataContext.Refresh(System.Data.Linq.RefreshMode.KeepChanges, dataContext.ProjectBase);

                MainTreeView.DataContext = dataContext.ProjectClassify;

                WindowDataBind();
                BuildLinkProjectList(ProjectID);
                SetTreeViewFocus();
                DialogR = false;
            }
        }

        private void WindowDataBind()
        {
            tabProjectBase.DataContext = dataContext.ProjectBase.Where(p => p.ProjectId.Equals(ProjectID));
            tabFunds.DataContext = dataContext.View_Funds.Where(f => f.ProjectID.Equals(ProjectID)).OrderBy(f => f.Date);
            tabRates.DataContext = dataContext.View_Rates.Where(r => r.ProjectID.Equals(ProjectID));
            tabResults.DataContext = dataContext.Results.Where(r => r.ProjectID.Equals(ProjectID));
            tabRewards.DataContext = dataContext.View_Rewards.Where(r => r.ProjectID.Equals(ProjectID));
            tabPatents.DataContext = dataContext.View_Patents.Where(p => p.ProjectID.Equals(ProjectID));
            tabTeam.DataContext = dataContext.TeamDepartments.Where(t => t.ProjectID.Equals(ProjectID));
            tabWorkers.DataContext = dataContext.TeamWorkers.Where(t => t.ProjectID.Equals(ProjectID));
            datagridContractIn.DataContext =
                dataContext.ProjectContracts.Where(pc => pc.ProjectID.Equals(ProjectID) && pc.TypeID.Equals(1));
            datagridContractPay.DataContext =
                dataContext.ProjectContracts.Where(pc => pc.ProjectID.Equals(ProjectID) && pc.TypeID.Equals(2));
        }

        private void WorkersManage()
        {
            if (ProjectID == 0)
            {
                MessageBox.Show("请选择项目！", "错误");
                return;
            }
            var workManage = new WorkManage();
            workManage.ProjectID = ProjectID;
            workManage.Owner = this;
            workManage.ShowDialog();
            RefreshData();
        }

        private void buttonDeleteProject_Click(object sender, RoutedEventArgs e)
        {
            DeleteProject();
        }

        private void DeleteProject()
        {
            if (ProjectID == 0)
            {
                MessageBox.Show("请选择项目！", "错误");
                return;
            }
            var pb = dataContext.ProjectBase.Single(p => p.ProjectId.Equals(ProjectID));
            if (MessageBox.Show("确认删除项目 " + pb.ProjectName + " ？！ 删除后将不能恢复！", "警告", MessageBoxButton.YesNoCancel) !=
                MessageBoxResult.Yes)
            {
                return;
            }
            dataContext.ProjectBase.DeleteOnSubmit(pb);
            dataContext.SubmitChanges();
            ProjectID = 0;
            DialogR = true;
            RefreshData();
        }

        private void menuitemDeleteProject_Click(object sender, RoutedEventArgs e)
        {
            DeleteProject();
        }

        private void menuitemMoneyDetailYear_Click(object sender, RoutedEventArgs e)
        {
            // 在此处添加事件处理程序实现。
            var moneyDetailYear = new MoneyDetailYear();
            moneyDetailYear.ShowDialog();
        }

        private void buttonPrizePay_Click(object sender, RoutedEventArgs e)
        {
            PrizePay();
        }

        private void PrizePay()
        {
            if (ProjectID == 0)
            {
                MessageBox.Show("请选择项目！", "错误");
                return;
            }
            var prizePayManage = new PrizePayManage();
            prizePayManage.ProjectId = ProjectID;
            prizePayManage.ShowDialog();
        }

        private void menuItemPrizesDetailYear_Click(object sender, RoutedEventArgs e)
        {
            var prizesDetailYear = new PrizesDetailYear();
            prizesDetailYear.ShowDialog();
        }

        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            var importPrizes = new ImportPrizes();
            importPrizes.ShowDialog();
        }

        private void MenuItemPrizesDetailYear_Click_1(object sender, RoutedEventArgs e)
        {
            var prizesDetailYear = new PrizesDetailYear();
            prizesDetailYear.ShowDialog();
        }

        private void MenuItemPrizeImportedManage_Click(object sender, RoutedEventArgs e)
        {
            var prizeImportedManage = new PrizeImportedManage();
            prizeImportedManage.ShowDialog();
        }

        private void ButtonContractFundManage_Click(object sender, RoutedEventArgs e)
        {
        }
    }
}