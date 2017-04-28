using System;
using System.Linq;
using System.Windows;

namespace HNCDIExternalProjectManage
{
    /// <summary>
    /// NewProjectBase.xaml 的交互逻辑
    /// </summary>
    public partial class NewProjectBase : Window
    {
        private DataClassesProjectClassifyDataContext dataContent;
        private int projectID;

        public int ProjectID
        {
            get { return projectID; }
            set { projectID = value; }
        }

        private int parentID;

        public int ParentID
        {
            get { return parentID; }
            set { parentID = value; }
        }

        private bool isCreate;

        public bool IsCreate
        {
            get { return isCreate; }
            set { isCreate = value; }
        }

        private bool isSubProject;

        public bool IsSubProject
        {
            get { return isSubProject; }
            set { isSubProject = value; }
        }

        private int projectClassifyID;

        public int ProjectClassifyID
        {
            get { return projectClassifyID; }
            set { projectClassifyID = value; }
        }

        public NewProjectBase()
        {
            this.InitializeComponent();

            // 在此点之下插入创建对象所需的代码。
        }

        private void buttonSubmit_Click(object sender, RoutedEventArgs e)
        {
            if (!validDate())
            {
                MessageBox.Show("日期或数字格式不对！", "警告");
                return;
            }

            if (isCreate)
            {
                if (MessageBox.Show("确认新建项目？", "温馨提示", MessageBoxButton.OKCancel) != MessageBoxResult.OK)
                {
                    return;
                }
                //新建项目
                if (!isSubProject)
                {
                    //非子项目
                    dataContent = new DataClassesProjectClassifyDataContext();
                    var pb = dataContent.ProjectBase.Where(p => p.ProjectNo.Trim().Equals(ProjectNo.Text.Trim()) && p.ProjectNo.Trim() != "");
                    if (pb.Count() > 0)
                    {
                        MessageBox.Show("院编号重复，已经录入该项目？", "错误");
                        return;
                    }
                    ProjectBase projectBase = new ProjectBase();
                    projectBase.ProjectClassifyID = projectClassifyID;
                    projectBase.ProjectNo = ProjectNo.Text.Trim();
                    projectBase.ContractNo = ContractNo.Text.Trim();
                    projectBase.FirstParty = FirstParty.Text.Trim();
                    projectBase.SecondParty = SecondParty.Text.Trim();
                    projectBase.SetupYear = SetupYear.Text.Trim();
                    projectBase.ProjectName = ProjectName.Text.Trim();
                    if (!String.IsNullOrEmpty(StartDate.Text.Trim()))
                    {
                        projectBase.StartDate = DateTime.Parse(StartDate.Text.Trim());
                    }
                    if (!String.IsNullOrEmpty(PlanFinishDate.Text.Trim()))
                    {
                        projectBase.PlanFinishDate = DateTime.Parse(PlanFinishDate.Text.Trim());
                    }
                    if (!String.IsNullOrEmpty(FinishDate.Text.Trim()))
                    {
                        projectBase.FinishDate = DateTime.Parse(FinishDate.Text.Trim());
                    }
                    projectBase.Principal = Principal.Text.Trim();
                    if (!String.IsNullOrEmpty(SumMoney.Text.Trim()))
                    {
                        projectBase.SumMoney = Convert.ToDecimal(SumMoney.Text.Trim());
                    }
                    if (!String.IsNullOrEmpty(Ministry.Text.Trim()))
                    {
                        projectBase.Ministry = Convert.ToDecimal(Ministry.Text.Trim());
                    }
                    if (!String.IsNullOrEmpty(Transportation.Text))
                    {
                        projectBase.Transportation = Convert.ToDecimal(Transportation.Text.Trim());
                    }
                    if (!String.IsNullOrEmpty(Science.Text.Trim()))
                    {
                        projectBase.Science = Convert.ToDecimal(Science.Text.Trim());
                    }
                    if (!String.IsNullOrEmpty(SupportEngineering.Text.Trim()))
                    {
                        projectBase.SupportEngineering = Convert.ToDecimal(SupportEngineering.Text);
                    }
                    if (!String.IsNullOrEmpty(Other.Text.Trim()))
                    {
                        projectBase.Other = Convert.ToDecimal(Other.Text.Trim());
                    }
                    if (!String.IsNullOrEmpty(AuthrizeMoney.Text.Trim()))
                    {
                        projectBase.AuthorizeMoney = Convert.ToDecimal(AuthrizeMoney.Text.Trim());
                    }
                    projectBase.AnchoredDepartment = AnchoredDepartment.Text.Trim();
                    if ((bool)(IsKnot1.IsChecked))
                    {
                        projectBase.IsKnot = "验收";
                    }
                    if ((bool)(IsKnot2.IsChecked))
                    {
                        projectBase.IsKnot = "鉴定";
                    }
                    if ((bool)(IsKnot3.IsChecked))
                    {
                        projectBase.IsKnot = "尚未结题";
                    }
                    if ((bool)(IsKnot5.IsChecked))
                    {
                        projectBase.IsKnot = "结清";
                    }
                    projectBase.IsMainResearch = IsMainSearch.IsChecked;
                    projectBase.IsFiled = IsFiled.IsChecked;
                    projectBase.Note = Note.Text.Trim();
                    dataContent.ProjectBase.InsertOnSubmit(projectBase);
                    dataContent.SubmitChanges();
                    ProjectID = projectBase.ProjectId;
                    ((MainWindow)(this.Owner)).ProjectID = projectID;
                }
                else
                {
                    //子项目
                    dataContent = new DataClassesProjectClassifyDataContext();
                    var pb = dataContent.ProjectBase.Where(p => p.ProjectNo.Trim().Equals(ProjectNo.Text.Trim()) && p.ProjectNo.Trim() != "");
                    if (pb.Count() > 0)
                    {
                        MessageBox.Show("院编号重复，已经录入该项目？", "错误");
                        return;
                    }
                    ProjectBase projectBase = new ProjectBase();
                    projectBase.ParentID = ParentID;
                    projectBase.ProjectNo = ProjectNo.Text.Trim();
                    projectBase.ContractNo = ContractNo.Text.Trim();
                    projectBase.FirstParty = FirstParty.Text.Trim();
                    projectBase.SecondParty = SecondParty.Text.Trim();
                    projectBase.SetupYear = SetupYear.Text.Trim();
                    projectBase.ProjectName = ProjectName.Text.Trim();
                    if (!String.IsNullOrEmpty(StartDate.Text.Trim()))
                    {
                        projectBase.StartDate = DateTime.Parse(StartDate.Text.Trim());
                    }
                    if (!String.IsNullOrEmpty(PlanFinishDate.Text.Trim()))
                    {
                        projectBase.PlanFinishDate = DateTime.Parse(PlanFinishDate.Text.Trim());
                    }
                    if (!String.IsNullOrEmpty(FinishDate.Text.Trim()))
                    {
                        projectBase.FinishDate = DateTime.Parse(FinishDate.Text.Trim());
                    }
                    projectBase.Principal = Principal.Text.Trim();
                    if (!String.IsNullOrEmpty(SumMoney.Text.Trim()))
                    {
                        projectBase.SumMoney = Convert.ToDecimal(SumMoney.Text.Trim());
                    }
                    if (!String.IsNullOrEmpty(Ministry.Text.Trim()))
                    {
                        projectBase.Ministry = Convert.ToDecimal(Ministry.Text.Trim());
                    }
                    if (!String.IsNullOrEmpty(Transportation.Text))
                    {
                        projectBase.Transportation = Convert.ToDecimal(Transportation.Text.Trim());
                    }
                    if (!String.IsNullOrEmpty(Science.Text.Trim()))
                    {
                        projectBase.Science = Convert.ToDecimal(Science.Text.Trim());
                    }
                    if (!String.IsNullOrEmpty(SupportEngineering.Text.Trim()))
                    {
                        projectBase.SupportEngineering = Convert.ToDecimal(SupportEngineering.Text);
                    }
                    if (!String.IsNullOrEmpty(Other.Text.Trim()))
                    {
                        projectBase.Other = Convert.ToDecimal(Other.Text.Trim());
                    }
                    if (!String.IsNullOrEmpty(AuthrizeMoney.Text.Trim()))
                    {
                        projectBase.AuthorizeMoney = Convert.ToDecimal(AuthrizeMoney.Text.Trim());
                    }
                    projectBase.AnchoredDepartment = AnchoredDepartment.Text.Trim();
                    if ((bool)(IsKnot1.IsChecked))
                    {
                        projectBase.IsKnot = "验收";
                    }
                    if ((bool)(IsKnot2.IsChecked))
                    {
                        projectBase.IsKnot = "鉴定";
                    }
                    if ((bool)(IsKnot3.IsChecked))
                    {
                        projectBase.IsKnot = "尚未结题";
                    }
                    if ((bool)(IsKnot5.IsChecked))
                    {
                        projectBase.IsKnot = "结清";
                    }
                    projectBase.IsMainResearch = IsMainSearch.IsChecked;
                    projectBase.IsFiled = IsFiled.IsChecked;
                    projectBase.Note = Note.Text.Trim();
                    dataContent.ProjectBase.InsertOnSubmit(projectBase);
                    dataContent.SubmitChanges();
                    ProjectID = projectBase.ProjectId;
                    ((MainWindow)(this.Owner)).ProjectID = projectID;
                }
            }
            else
            {
                //修改项目
                if (dataContent == null)
                {
                    dataContent = new DataClassesProjectClassifyDataContext();
                }
                var pb = dataContent.ProjectBase.Where(p => p.ProjectNo.Trim().Equals(ProjectNo.Text.Trim()) && p.ProjectNo.Trim() != "" && p.ProjectId != projectID);
                if (pb.Count() > 0)
                {
                    MessageBox.Show("项目编号重复！", "错误");
                    return;
                }
                ProjectBase projectBase = dataContent.ProjectBase.Single(p => p.ProjectId.Equals(ProjectID));
                projectBase.ProjectNo = ProjectNo.Text.Trim();
                projectBase.ContractNo = ContractNo.Text.Trim();
                projectBase.FirstParty = FirstParty.Text.Trim();
                projectBase.SecondParty = SecondParty.Text.Trim();
                projectBase.SetupYear = SetupYear.Text.Trim();
                projectBase.ProjectName = ProjectName.Text.Trim();
                if (!String.IsNullOrEmpty(StartDate.Text.Trim()))
                {
                    projectBase.StartDate = DateTime.Parse(StartDate.Text.Trim());
                }
                if (!String.IsNullOrEmpty(PlanFinishDate.Text.Trim()))
                {
                    projectBase.PlanFinishDate = DateTime.Parse(PlanFinishDate.Text.Trim());
                }
                if (!String.IsNullOrEmpty(FinishDate.Text.Trim()))
                {
                    projectBase.FinishDate = DateTime.Parse(FinishDate.Text.Trim());
                }
                projectBase.Principal = Principal.Text.Trim();
                if (!String.IsNullOrEmpty(SumMoney.Text.Trim()))
                {
                    projectBase.SumMoney = Convert.ToDecimal(SumMoney.Text.Trim());
                }
                if (!String.IsNullOrEmpty(Ministry.Text.Trim()))
                {
                    projectBase.Ministry = Convert.ToDecimal(Ministry.Text.Trim());
                }
                if (!String.IsNullOrEmpty(Transportation.Text))
                {
                    projectBase.Transportation = Convert.ToDecimal(Transportation.Text.Trim());
                }
                if (!String.IsNullOrEmpty(Science.Text.Trim()))
                {
                    projectBase.Science = Convert.ToDecimal(Science.Text.Trim());
                }
                if (!String.IsNullOrEmpty(SupportEngineering.Text.Trim()))
                {
                    projectBase.SupportEngineering = Convert.ToDecimal(SupportEngineering.Text);
                }
                if (!String.IsNullOrEmpty(Other.Text.Trim()))
                {
                    projectBase.Other = Convert.ToDecimal(Other.Text.Trim());
                }
                if (!String.IsNullOrEmpty(AuthrizeMoney.Text.Trim()))
                {
                    projectBase.AuthorizeMoney = Convert.ToDecimal(AuthrizeMoney.Text.Trim());
                }
                projectBase.AnchoredDepartment = AnchoredDepartment.Text.Trim();
                if ((bool)(IsKnot1.IsChecked))
                {
                    projectBase.IsKnot = "验收";
                }
                if ((bool)(IsKnot2.IsChecked))
                {
                    projectBase.IsKnot = "鉴定";
                }
                if ((bool)(IsKnot3.IsChecked))
                {
                    projectBase.IsKnot = "尚未结题";
                }
                if ((bool)(IsKnot5.IsChecked))
                {
                    projectBase.IsKnot = "结清";
                }
                projectBase.IsMainResearch = IsMainSearch.IsChecked;
                projectBase.IsFiled = IsFiled.IsChecked;
                projectBase.Note = Note.Text.Trim();
                dataContent.SubmitChanges();
            }
            ((MainWindow)(this.Owner)).DialogR = true;
            this.Close();
        }

        private void buttonCancel_Click(object sender, RoutedEventArgs e)
        {
            ((MainWindow)(this.Owner)).DialogR = false;
            this.Close();
        }

        /// <summary>
        /// 数据验证
        /// </summary>
        /// <returns></returns>
        private bool validDate()
        {
            try
            {
                if (!String.IsNullOrEmpty(StartDate.Text))
                {
                    DateTime dt = Convert.ToDateTime(StartDate.Text);
                }
                if (!String.IsNullOrEmpty(PlanFinishDate.Text))
                {
                    DateTime dt = Convert.ToDateTime(PlanFinishDate.Text);
                }
                if (!String.IsNullOrEmpty(FinishDate.Text))
                {
                    DateTime dt = Convert.ToDateTime(FinishDate.Text);
                }
                if (!String.IsNullOrEmpty(SumMoney.Text))
                {
                    double d = Convert.ToDouble(SumMoney.Text);
                }
                if (!String.IsNullOrEmpty(Ministry.Text))
                {
                    double d = Convert.ToDouble(Ministry.Text);
                }
                if (!String.IsNullOrEmpty(Transportation.Text))
                {
                    double d = Convert.ToDouble(Transportation.Text);
                }
                if (!String.IsNullOrEmpty(Science.Text))
                {
                    double d = Convert.ToDouble(Science.Text);
                }
                if (!String.IsNullOrEmpty(SupportEngineering.Text))
                {
                    double d = Convert.ToDouble(SupportEngineering.Text);
                }
                if (!String.IsNullOrEmpty(Other.Text))
                {
                    double d = Convert.ToDouble(Other.Text);
                }
                if (!String.IsNullOrEmpty(AuthrizeMoney.Text))
                {
                    double d = Convert.ToDouble(AuthrizeMoney.Text);
                }
                if (!String.IsNullOrEmpty(SetupYear.Text))
                {
                    int d = Convert.ToInt32(SetupYear.Text);
                    if (d < 1900 || d > 9999)
                    {
                        return false;
                    }
                }
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }

        private void buttonCancel_Loaded(object sender, RoutedEventArgs e)
        {
            if (!IsCreate)
            {
                dataContent = new DataClassesProjectClassifyDataContext();
                ProjectBase projectBase = dataContent.ProjectBase.Single(p => p.ProjectId.Equals(ProjectID));
                ProjectNo.Text = projectBase.ProjectNo;
                ContractNo.Text = projectBase.ContractNo;
                FirstParty.Text = projectBase.FirstParty;
                SecondParty.Text = projectBase.SecondParty;
                SetupYear.Text = projectBase.SetupYear;
                IsMainSearch.IsChecked = projectBase.IsMainResearch;
                ProjectName.Text = projectBase.ProjectName;
                if (projectBase.StartDate != null)
                {
                    StartDate.Text = ((DateTime)(projectBase.StartDate)).ToShortDateString();
                }
                if (projectBase.PlanFinishDate != null)
                {
                    PlanFinishDate.Text = ((DateTime)(projectBase.PlanFinishDate)).ToShortDateString();
                }
                if (projectBase.FinishDate != null)
                {
                    FinishDate.Text = ((DateTime)(projectBase.FinishDate)).ToShortDateString();
                }
                Principal.Text = projectBase.Principal;
                SumMoney.Text = projectBase.SumMoney.ToString();
                Ministry.Text = projectBase.Ministry.ToString();
                Transportation.Text = projectBase.Transportation.ToString();
                Science.Text = projectBase.Science.ToString();
                SupportEngineering.Text = projectBase.SupportEngineering.ToString();
                Other.Text = projectBase.Other.ToString();
                AuthrizeMoney.Text = projectBase.AuthorizeMoney.ToString();
                AnchoredDepartment.Text = projectBase.AnchoredDepartment;
                if (projectBase.IsKnot == "验收")
                {
                    IsKnot1.IsChecked = true;
                }
                if (projectBase.IsKnot == "鉴定")
                {
                    IsKnot2.IsChecked = true;
                }
                if (projectBase.IsKnot == "尚未结题")
                {
                    IsKnot3.IsChecked = true;
                }
                if (projectBase.IsKnot == "结清")
                {
                    IsKnot5.IsChecked = true;
                }
                IsFiled.IsChecked = projectBase.IsFiled;
                Note.Text = projectBase.Note;
            }
        }

        private void NewProjectBaseWindow_Loaded(object sender, RoutedEventArgs e)
        {
            dataContent = new DataClassesProjectClassifyDataContext();
            string title = "";
            if (isCreate)
            {
                if (!isSubProject)
                {
                    ProjectClassify projectClassify = dataContent.ProjectClassify.Single(pc => pc.ClassifyId.Equals(projectClassifyID));
                    title = "新建“" + projectClassify.ProjectClassify1 + "”类别项目";
                }
                else
                {
                    ProjectBase projectBase = dataContent.ProjectBase.Single(pb => pb.ProjectId.Equals(parentID));
                    title = "新建项目“" + projectBase.ProjectName + "”的子项目";
                }
            }
            else
            {
                ProjectBase projectBase = dataContent.ProjectBase.Single(pb => pb.ProjectId.Equals(projectID));
                title = projectBase.ProjectName + "——修改项目基本信息";
            }
            this.Title = title;
        }
    }
}