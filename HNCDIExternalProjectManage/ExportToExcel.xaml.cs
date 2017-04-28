using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Linq;
using System.IO;
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
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using A = DocumentFormat.OpenXml.Drawing;
using Op = DocumentFormat.OpenXml.CustomProperties;

namespace HNCDIExternalProjectManage
{
    /// <summary>
    /// ExportToExcel.xaml 的交互逻辑
    /// </summary>
    public partial class ExportToExcel : Window
    {
        DataClassesProjectClassifyDataContext dataContext;

        private List<int> projectIDList; //要导出的项目ID列表
        private List<int> projectToExcelIDListSorted; //排好序的要导出的项目ID列表
        private List<int> LinkProjectClassify; //要导出的项目类型列表
        private List<int> classifyProjects; //各类项目数量列表
        private string projectClassifyName = "";
        private int projects = 0;
        FileInfo fileToCreate; //要创建的文件
        private int rows = 0; //电子表格行数
        private int cols = 0; //要导出的字段数
        private string lastCellName = ""; //表格最后单元格名
        private string lastColName = ""; //最后一列名
        private string ExcelTitle = ""; //表格标题
        private int mergeCellsCount = 0; //合并单元格个数
        private List<Fields> listSelectedFields; //选择的字段
        public ExportToExcel()
        {
            this.InitializeComponent();
            
            // 在此点之下插入创建对象所需的代码。
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            dataContext = new DataClassesProjectClassifyDataContext();
            ListFileType.DisplayMemberPath = "ExcelFileType1";
            ListFileType.SelectedValuePath = "Id";
            ListFileType.DataContext = dataContext.ExcelFileType;
            System.Text.ASCIIEncoding asciiEncoding = new System.Text.ASCIIEncoding();
            setListSelectFields();
        }

        /// <summary>
        /// 初始化选择字段listSelectedFields列表
        /// </summary>
        private void setListSelectFields()
        {
            listSelectedFields = new List<Fields>();
            foreach(UIElement checkBox in fields.Children)
            {
                if(checkBox is CheckBox)
                {
                    CheckBox c = (CheckBox)checkBox;
                    if (c.IsChecked == true)
                    {
                        listSelectedFields.Add(new Fields { fieldName = (string)c.Content, fieldValue = c.Name });
                    }
                }
            }
            ListSourceFields.DisplayMemberPath = "fieldName";
            ListSourceFields.SelectedValuePath = "fieldValue";
            ListSourceFields.ItemsSource = listSelectedFields;
            if(listSelectedFields.ElementAt(0) != null)
            {
                ListSourceFields.SelectedIndex = 0;
            }
        }

        private void ListFileType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ExcelFileType excelFileType = (ExcelFileType)ListFileType.SelectedItem;
            if (excelFileType != null)
            {
                switch (excelFileType.ExcelFileType1)
                {
                    case "HNCDI院外立项课题明细":
                        ProjectNo.IsChecked = true;
                        FirstParty.IsChecked = true;
                        SetupYear.IsChecked = true;
                        ProjectName.IsChecked = true;
                        Principal.IsChecked = true;
                        ContractPeriod.IsChecked = true;
                        SumMoney.IsChecked = true;
                        MoneySourceDetail.IsChecked = true;
                        MoneyDetail.IsChecked = false;
                        AnchoredDepartment.IsChecked = true;
                        Workers.IsChecked = true;
                        TeamDepartment.IsChecked = true;
                        Note.IsChecked = true;
                        SecondParty.IsChecked = false;
                        ContractNo.IsChecked = false;
                        IsTimeCheck.IsChecked = false;
                        isKnoteReq.IsChecked = false;
                        IsFiledReq.IsChecked = false;
                        //CompleteDepartment.IsChecked = false;
                        //CompleteWorks.IsChecked = false;
                        //FinishState.IsChecked = true;
                        RateState.IsChecked = false;
                        FactFinishDate.IsChecked = false;
                        RewardState.IsChecked = false;
                        PatentState.IsChecked = false;
                        MainResearchState.IsChecked = false;
                        KnoteState.IsChecked = false;
                        FiledState.IsChecked = false;
                        setListSelectFields();
                        break;
                    case "HNCDI已结题科研项目情况一览表":
                        ProjectNo.IsChecked = false;
                        FirstParty.IsChecked = false;
                        SecondParty.IsChecked = false;
                        SetupYear.IsChecked = false;
                        ProjectName.IsChecked = true;
                        Principal.IsChecked = true;
                        ContractPeriod.IsChecked = false;
                        SumMoney.IsChecked = true;
                        MoneySourceDetail.IsChecked = false;
                        MoneyDetail.IsChecked = true;
                        AnchoredDepartment.IsChecked = false;
                        Workers.IsChecked = true;
                        TeamDepartment.IsChecked = false;
                        Note.IsChecked = true;
                        ContractNo.IsChecked = true;
                        IsTimeCheck.IsChecked = false;
                        isKnoteReq.IsChecked = false;
                        IsFiledReq.IsChecked = false;
                        //CompleteDepartment.IsChecked = true;
                        //CompleteWorks.IsChecked = true;
                        //FinishState.IsChecked = false;
                        RateState.IsChecked = true;
                        FactFinishDate.IsChecked = true;
                        RewardState.IsChecked = true;
                        PatentState.IsChecked = true;
                        MainResearchState.IsChecked = false;
                        KnoteState.IsChecked = true;
                        FiledState.IsChecked = true;
                        setListSelectFields();
                        break;
                    case "自定义":
                    default:
                        ProjectNo.IsChecked = true;
                        FirstParty.IsChecked = true;
                        SetupYear.IsChecked = true;
                        ProjectName.IsChecked = true;
                        Principal.IsChecked = true;
                        ContractPeriod.IsChecked = true;
                        SumMoney.IsChecked = true;
                        SecondParty.IsChecked = false;
                        MoneySourceDetail.IsChecked = false;
                        MoneyDetail.IsChecked = false;
                        AnchoredDepartment.IsChecked = false;
                        Workers.IsChecked = false;
                        TeamDepartment.IsChecked = false;
                        Note.IsChecked = false;
                        ContractNo.IsChecked = false;
                        IsTimeCheck.IsChecked = false;
                        isKnoteReq.IsChecked = false;
                        IsFiledReq.IsChecked = false;
                        //CompleteDepartment.IsChecked = false;
                        //CompleteWorks.IsChecked = false;
                        //FinishState.IsChecked = false;
                        RateState.IsChecked = false;
                        FactFinishDate.IsChecked = false;
                        RewardState.IsChecked = false;
                        PatentState.IsChecked = false;
                        MainResearchState.IsChecked = false;
                        KnoteState.IsChecked = false;
                        FiledState.IsChecked = false;
                        setListSelectFields();
                        break;
                }
            }
        }

        private void buttonOK_Click(object sender, RoutedEventArgs e)
        {
            ExcelFileType excelFileType = (ExcelFileType)ListFileType.SelectedItem;
            if (excelFileType != null)
            {
                ExcelTitle = excelFileType.ExcelFileType1;
            }
            else
            {
                MessageBox.Show("请选择导出的Excle表格类型");
                return;
            }
            dataContext = new DataClassesProjectClassifyDataContext();
            var searchPredicate = PredicateExtensions.True<ProjectBase>();

            SetTimeReq(ref searchPredicate); //处理时间要求
            searchPredicate = SetKnoteReq(searchPredicate); //处理结题要求
            searchPredicate = SetFiledReq(searchPredicate); //处理归档要求

            var projectBases = dataContext.ProjectBase.Where(searchPredicate);

            //创建项目ID列表
            projectIDList = new List<int>();
            LinkProjectClassify = new List<int>();
            foreach(ProjectBase p in projectBases)
            {
                var pp = p;
                if (!projectIDList.Contains(pp.ProjectId))
                {
                    projectIDList.Add(pp.ProjectId);
                    while (pp.ParentID != null)
                    {
                        pp = dataContext.ProjectBase.Single(ppb => ppb.ProjectId.Equals(pp.ParentID));
                        if (!projectIDList.Contains(pp.ProjectId))
                        {
                            projectIDList.Add(pp.ProjectId);
                        }
                    }
                }
                if (!LinkProjectClassify.Contains((int)pp.ProjectClassifyID))
                {
                    LinkProjectClassify.Add((int)pp.ProjectClassifyID);
                }
            }

            rows = LinkProjectClassify.Count + projectIDList.Count + 2;
            GetColsAndLastCellName();
            getMergeCellCount();
            

            //排列项目输出顺序
            LinkProjectClassify.Sort();
            projectToExcelIDListSorted = new List<int>();
            var prc = dataContext.ProjectClassify.OrderBy(pc => pc.ClassifyId);
            classifyProjects = new List<int>();
            foreach(var pc in prc)
            {
                projects = 0;
                if(LinkProjectClassify.Contains(pc.ClassifyId))
                {
                    var pbs = dataContext.ProjectBase.Where(p => p.ProjectClassifyID.Equals(pc.ClassifyId)).OrderBy(p => p.ProjectId);
                    foreach(ProjectBase pb in pbs)
                    {
                        if(!projectIDList.Contains(pb.ProjectId))
                        {
                            continue;
                        }
                        if (pb.ProjectClassifyID != null)
                        {
                            //顶级项目
                            if(!projectToExcelIDListSorted.Contains(pb.ProjectId))
                            {
                                projectToExcelIDListSorted.Add(pb.ProjectId);
                                projects += 1;
                            }
                            AddSubProjectToprojectToExcelListSorted(pb);
                        }
                    }
                }
                classifyProjects.Add(projects);
            }

            try
            {
                Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog();
                saveFileDialog.Filter = "Execl files (*.xlsx)|*.xlsx";
                saveFileDialog.FilterIndex = 0;
                saveFileDialog.RestoreDirectory = true;
                saveFileDialog.CreatePrompt = true;
                saveFileDialog.Title = "要创建Excel文件";
                saveFileDialog.ShowDialog();
                if (saveFileDialog.FileName == "")
                {
                    MessageBox.Show("错误", "请选择文件或输入文件名",MessageBoxButton.OK);
                    return;
                }
                fileToCreate = new FileInfo(saveFileDialog.FileName);
                if (fileToCreate.Exists)
                {
                    try
                    {
                        fileToCreate.Delete();
                    }
                    catch (Exception error)
                    {
                        MessageBox.Show(error.Message, "删除失败 ", MessageBoxButton.OK);
                        return;
                    }
                }
                if(ExcelTitle == "自定义")
                {
                    ExcelTitle = saveFileDialog.FileName.Substring(saveFileDialog.FileName.LastIndexOf(@"\") + 1);
                    ExcelTitle = ExcelTitle.Remove(ExcelTitle.LastIndexOf("."));
                }
                CreatePackage(saveFileDialog.FileName);
                MessageBox.Show("导出成功！");
            }
            catch(Exception error)
            {
                MessageBox.Show(error.Message, "导出失败 ", MessageBoxButton.OK);
                return;
            }
        }

        /// <summary>
        /// 获取合并单元格个数
        /// </summary>
        private void getMergeCellCount()
        {
            mergeCellsCount = new Int32();
            mergeCellsCount = LinkProjectClassify.Count + 1;
        }

        /// <summary>
        /// 将子项目插入projectToExcelListSorted
        /// </summary>
        /// <param name="project"></param>
        private void AddSubProjectToprojectToExcelListSorted(ProjectBase project)
        {
            var ps = dataContext.ProjectBase.Where(pp => pp.ParentID.Equals(project.ProjectId)).OrderBy(pp => pp.ProjectId);
            foreach(var pb in ps)
            {
                if(projectIDList.Contains(pb.ProjectId))
                {
                    if(!projectToExcelIDListSorted.Contains(pb.ProjectId))
                    {
                        projectToExcelIDListSorted.Add(pb.ProjectId);
                        projects += 1;
                        AddSubProjectToprojectToExcelListSorted(pb);
                    }
                }
            }
        }
        /// <summary>
        /// 获取字段数，设定最后单元格名称
        /// </summary>
        private void GetColsAndLastCellName()
        {
            cols = 1;
            if (ProjectNo.IsChecked == true) cols += 1;
            if (FirstParty.IsChecked == true) cols += 1;
            if (SetupYear.IsChecked == true) cols += 1;
            if (ProjectName.IsChecked == true) cols += 1;
            if (SecondParty.IsChecked == true) cols += 1;
            if (ContractNo.IsChecked == true) cols += 1;
            if (Principal.IsChecked == true) cols += 1;
            if (ContractPeriod.IsChecked == true) cols += 1;
            if (SumMoney.IsChecked == true) cols += 1;
            if (MoneySourceDetail.IsChecked == true) cols += 5;
            if (MoneyDetail.IsChecked == true) cols += 5;
            if (AnchoredDepartment.IsChecked == true) cols += 1;
            if (Workers.IsChecked == true) cols += 1;
            if (TeamDepartment.IsChecked == true) cols += 1;
            //if (CompleteDepartment.IsChecked == true) cols += 1;
            //if (CompleteWorks.IsChecked == true) cols += 1;
            //if (FinishState.IsChecked == true) cols += 1;
            if (RateState.IsChecked == true) cols += 1;
            if (FactFinishDate.IsChecked == true) cols += 1;
            if (RewardState.IsChecked == true) cols += 6;
            if (PatentState.IsChecked == true) cols += 1;
            if (MainResearchState.IsChecked == true) cols += 1;
            if (KnoteState.IsChecked == true) cols += 1;
            if (FiledState.IsChecked == true) cols += 1;
            if (Note.IsChecked == true) cols += 1;
            System.Text.ASCIIEncoding asciiEncoding = new System.Text.ASCIIEncoding();
            lastCellName = "";
            byte[] btNumber;
            if(cols <= 26)
            {
                btNumber = new byte[] { (byte)(cols + 64) };
                lastCellName = asciiEncoding.GetString(btNumber);
            }
            else
            {
                btNumber = new byte[] { (byte)(cols / 26 + 64) };
                lastCellName = asciiEncoding.GetString(btNumber);
                btNumber = new byte[] { (byte)(cols % 26 + 64) };
                lastCellName += asciiEncoding.GetString(btNumber);
            }
            lastColName = lastCellName;
            lastCellName += rows.ToString();
        }

        /// <summary>
        /// 处理归档要求
        /// </summary>
        /// <param name="searchPredicate"></param>
        /// <returns></returns>
        private System.Linq.Expressions.Expression<Func<ProjectBase, bool>> SetFiledReq(System.Linq.Expressions.Expression<Func<ProjectBase, bool>> searchPredicate)
        {
            if (IsFiledReq.IsChecked == true)
            {
                //有归档要求
                if (IsFiled.IsChecked == true)
                {
                    //已归档
                    searchPredicate = searchPredicate.And(p => p.IsFiled.Equals(true));
                }
                else
                {
                    //未归档
                    searchPredicate = searchPredicate.And(p => p.IsFiled.Equals(true));
                }
            }
            return searchPredicate;
        }

        /// <summary>
        /// 处理结题要求
        /// </summary>
        /// <param name="searchPredicate"></param>
        /// <returns></returns>
        private System.Linq.Expressions.Expression<Func<ProjectBase, bool>> SetKnoteReq(System.Linq.Expressions.Expression<Func<ProjectBase, bool>> searchPredicate)
        {
            if (isKnoteReq.IsChecked == true)
            {
                //有结题要求
                bool isSeted = false;
                if (CheckAndAccept.IsChecked == true)
                {
                    //验收
                    searchPredicate = searchPredicate.And(p => p.IsKnot.Equals("验收"));
                    isSeted = true;
                }
                if (Authenticate.IsChecked == true)
                {
                    //鉴定
                    if (isSeted)
                    {
                        searchPredicate = searchPredicate.Or(p => p.IsKnot.Equals("鉴定"));
                    }
                    else
                    {
                        searchPredicate = searchPredicate.And(p => p.IsKnot.Equals("鉴定"));
                        isSeted = true;
                    }
                }
                if(Finished.IsChecked == true)
                {
                    //已结题（鉴定或验收）
                    if (isSeted)
                    {
                        searchPredicate = searchPredicate.Or(p => p.IsKnot.Equals("鉴定") || p.IsKnot.Equals("验收"));
                    }
                    else
                    {
                        searchPredicate = searchPredicate.And(p => p.IsKnot.Equals("鉴定") || p.IsKnot.Equals("验收"));
                        isSeted = true;
                    }
                }
                if (Unfinished.IsChecked == true)
                {
                    //尚未结题
                    if (isSeted)
                    {
                        searchPredicate = searchPredicate.Or(p => p.IsKnot.Equals("尚未结题"));
                    }
                    else
                    {
                        searchPredicate = searchPredicate.And(p => p.IsKnot.Equals("尚未结题"));
                        isSeted = true;
                    }
                }
                if(FundFinished.IsChecked == true)
                {
                    //结清
                    if (isSeted)
                    {
                        searchPredicate = searchPredicate.Or(p => p.IsKnot.Equals("结清"));
                    }
                    else
                    {
                        searchPredicate = searchPredicate.And(p => p.IsKnot.Equals("结清"));
                    }
                }
            }
            return searchPredicate;
        }

        /// <summary>
        /// 处理时间要求
        /// </summary>
        /// <param name="searchPredicate"></param>
        private void SetTimeReq(ref System.Linq.Expressions.Expression<Func<ProjectBase, bool>> searchPredicate)
        {
            if (IsTimeCheck.IsChecked == true)
            {
                //有时间要求
                if (DateSetupYear.IsChecked == true)
                {

                    //立项年度
                    if (String.IsNullOrEmpty(StartYear.Text) && String.IsNullOrEmpty(EndYear.Text))
                    {
                        MessageBox.Show("年度格式错误！应为四位有效年度数字", "错误");
                        return;
                    }
                    try
                    {
                        int startYear = 0;
                        int endYear = 0;
                        if(!String.IsNullOrEmpty(StartYear.Text))
                        {
                            startYear = Int32.Parse(StartYear.Text);
                        }
                        if(!String.IsNullOrEmpty(EndYear.Text))
                        {
                            endYear = Int32.Parse(EndYear.Text);
                        }
                        if(startYear == 0 && endYear == 0)
                        {
                            return;
                        }
                        if (endYear < startYear)
                        {
                            int i = startYear;
                            startYear = endYear;
                            endYear = i;
                        }
                        if(startYear > 0)
                        {
                            if (startYear < 1900 || startYear > 9999)
                            {
                                MessageBox.Show("年度格式错误！应为四位有效年度数字", "错误");
                                return;
                            }
                        }
                        if(endYear > 0)
                        {
                            if (endYear < 1900 || endYear > 9999)
                            {
                                MessageBox.Show("年度格式错误！应为四位有效年度数字", "错误");
                                return;
                            }
                        }
                        if (startYear == endYear)
                        {
                            searchPredicate = searchPredicate.And(p => (Convert.ToInt32(p.SetupYear)).Equals(startYear));
                        }
                        else
                        {
                            if (startYear == 0)
                            {
                                searchPredicate = searchPredicate.And(p => (Convert.ToInt32(p.SetupYear)).Equals(endYear));
                            }
                            else
                            {
                                searchPredicate = searchPredicate.And(p => (Convert.ToInt32(p.SetupYear)) >= startYear);
                                searchPredicate = searchPredicate.And(p => (Convert.ToInt32(p.SetupYear)) <= endYear);
                            }
                        }
                    }
                    catch (FormatException)
                    {
                        MessageBox.Show("年度格式错误！应为四位有效年度数字", "错误");
                        return;
                    }
                }
                else
                {
                    //合同期限
                    if (String.IsNullOrEmpty(FirstDate.Text) || String.IsNullOrEmpty(FinalDate.Text))
                    {
                        MessageBox.Show("时间格式错误！应为有效日期格式，如：2010.1.1", "错误");
                        return;
                    }
                    try
                    {
                        DateTime startDate = DateTime.Parse(FirstDate.Text);
                        DateTime endDate = DateTime.Parse(FinalDate.Text);
                        if (endDate < startDate)
                        {
                            DateTime t = startDate;
                            startDate = endDate;
                            endDate = t;
                        }
                        searchPredicate = searchPredicate.And(p => p.StartDate >= startDate);
                        searchPredicate = searchPredicate.And(p => p.StartDate <= endDate);
                    }
                    catch (FormatException)
                    {
                        if (String.IsNullOrEmpty(FirstDate.Text) || String.IsNullOrEmpty(FinalDate.Text))
                        {
                            MessageBox.Show("时间格式错误！应为有效日期格式，如：2010.1.1", "错误");
                            return;
                        }
                    }
                }
            }
        }

        public void CreatePackage(string filePath)
        {
            using (SpreadsheetDocument package = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                try
                {
                    CreateParts(package);
                }
                catch(Exception)
                {
                    package.Close();
                }
            }
        }

        // Adds child parts and generates content of the specified part.
        private void CreateParts(SpreadsheetDocument document)
        {
            ExtendedFilePropertiesPart extendedFilePropertiesPart1 = document.AddNewPart<ExtendedFilePropertiesPart>("rId3");
            GenerateExtendedFilePropertiesPart1Content(extendedFilePropertiesPart1);

            WorkbookPart workbookPart1 = document.AddWorkbookPart();
            GenerateWorkbookPart1Content(workbookPart1);

            WorkbookStylesPart workbookStylesPart1 = workbookPart1.AddNewPart<WorkbookStylesPart>("rId3");
            GenerateWorkbookStylesPart1Content(workbookStylesPart1);

            ThemePart themePart1 = workbookPart1.AddNewPart<ThemePart>("rId2");
            GenerateThemePart1Content(themePart1);

            WorksheetPart worksheetPart1 = workbookPart1.AddNewPart<WorksheetPart>("rId1");
            GenerateWorksheetPart1Content(worksheetPart1);

            SpreadsheetPrinterSettingsPart spreadsheetPrinterSettingsPart1 = worksheetPart1.AddNewPart<SpreadsheetPrinterSettingsPart>("rId1");
            GenerateSpreadsheetPrinterSettingsPart1Content(spreadsheetPrinterSettingsPart1);

            SharedStringTablePart sharedStringTablePart1 = workbookPart1.AddNewPart<SharedStringTablePart>("rId4");
            GenerateSharedStringTablePart1Content(sharedStringTablePart1);

            SetPackageProperties(document);
        }

        // Generates content of extendedFilePropertiesPart1.
        private void GenerateExtendedFilePropertiesPart1Content(ExtendedFilePropertiesPart extendedFilePropertiesPart1)
        {
            Ap.Properties properties1 = new Ap.Properties();
            properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            Ap.Application application1 = new Ap.Application();
            application1.Text = "Microsoft Excel";
            Ap.DocumentSecurity documentSecurity1 = new Ap.DocumentSecurity();
            documentSecurity1.Text = "0";
            Ap.ScaleCrop scaleCrop1 = new Ap.ScaleCrop();
            scaleCrop1.Text = "false";

            Ap.HeadingPairs headingPairs1 = new Ap.HeadingPairs();

            Vt.VTVector vTVector1 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Variant, Size = (UInt32Value)2U };

            Vt.Variant variant1 = new Vt.Variant();
            Vt.VTLPSTR vTLPSTR1 = new Vt.VTLPSTR();
            vTLPSTR1.Text = "工作表";

            variant1.Append(vTLPSTR1);

            Vt.Variant variant2 = new Vt.Variant();
            Vt.VTInt32 vTInt321 = new Vt.VTInt32();
            vTInt321.Text = "1";

            variant2.Append(vTInt321);

            vTVector1.Append(variant1);
            vTVector1.Append(variant2);

            headingPairs1.Append(vTVector1);

            Ap.TitlesOfParts titlesOfParts1 = new Ap.TitlesOfParts();

            Vt.VTVector vTVector2 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Lpstr, Size = (UInt32Value)1U };
            Vt.VTLPSTR vTLPSTR2 = new Vt.VTLPSTR();
            vTLPSTR2.Text = "Sheet1";

            vTVector2.Append(vTLPSTR2);

            titlesOfParts1.Append(vTVector2);
            Ap.LinksUpToDate linksUpToDate1 = new Ap.LinksUpToDate();
            linksUpToDate1.Text = "false";
            Ap.SharedDocument sharedDocument1 = new Ap.SharedDocument();
            sharedDocument1.Text = "false";
            Ap.HyperlinksChanged hyperlinksChanged1 = new Ap.HyperlinksChanged();
            hyperlinksChanged1.Text = "false";
            Ap.ApplicationVersion applicationVersion1 = new Ap.ApplicationVersion();
            applicationVersion1.Text = "14.0300";

            properties1.Append(application1);
            properties1.Append(documentSecurity1);
            properties1.Append(scaleCrop1);
            properties1.Append(headingPairs1);
            properties1.Append(titlesOfParts1);
            properties1.Append(linksUpToDate1);
            properties1.Append(sharedDocument1);
            properties1.Append(hyperlinksChanged1);
            properties1.Append(applicationVersion1);

            extendedFilePropertiesPart1.Properties = properties1;
        }

        // Generates content of workbookPart1.
        private void GenerateWorkbookPart1Content(WorkbookPart workbookPart1)
        {
            Workbook workbook1 = new Workbook();
            workbook1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            FileVersion fileVersion1 = new FileVersion() { ApplicationName = "xl", LastEdited = "5", LowestEdited = "5", BuildVersion = "9303" };
            WorkbookProperties workbookProperties1 = new WorkbookProperties() { DefaultThemeVersion = (UInt32Value)124226U };

            BookViews bookViews1 = new BookViews();
            WorkbookView workbookView1 = new WorkbookView() { XWindow = 480, YWindow = 375, WindowWidth = (UInt32Value)24555U, WindowHeight = (UInt32Value)11610U };

            bookViews1.Append(workbookView1);

            Sheets sheets1 = new Sheets();
            Sheet sheet1 = new Sheet() { Name = "Sheet1", SheetId = (UInt32Value)1U, Id = "rId1" };

            sheets1.Append(sheet1);
            CalculationProperties calculationProperties1 = new CalculationProperties() { CalculationId = (UInt32Value)145621U };

            workbook1.Append(fileVersion1);
            workbook1.Append(workbookProperties1);
            workbook1.Append(bookViews1);
            workbook1.Append(sheets1);
            workbook1.Append(calculationProperties1);

            workbookPart1.Workbook = workbook1;
        }

        // Generates content of workbookStylesPart1.
        private void GenerateWorkbookStylesPart1Content(WorkbookStylesPart workbookStylesPart1)
        {
            Stylesheet stylesheet1 = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            DocumentFormat.OpenXml.Spreadsheet.Fonts fonts1 = new DocumentFormat.OpenXml.Spreadsheet.Fonts() { Count = (UInt32Value)14U, KnownFonts = true };

            Font font1 = new Font();
            FontSize fontSize1 = new FontSize() { Val = 11D };
            DocumentFormat.OpenXml.Spreadsheet.Color color1 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Theme = (UInt32Value)1U };
            FontName fontName1 = new FontName() { Val = "宋体" };
            FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet1 = new FontCharSet() { Val = 134 };
            FontScheme fontScheme1 = new FontScheme() { Val = FontSchemeValues.Minor };

            font1.Append(fontSize1);
            font1.Append(color1);
            font1.Append(fontName1);
            font1.Append(fontFamilyNumbering1);
            font1.Append(fontCharSet1);
            font1.Append(fontScheme1);

            Font font2 = new Font();
            DocumentFormat.OpenXml.Spreadsheet.Bold bold1 = new DocumentFormat.OpenXml.Spreadsheet.Bold();
            FontSize fontSize2 = new FontSize() { Val = 16D };
            FontName fontName2 = new FontName() { Val = "宋体" };
            FontFamilyNumbering fontFamilyNumbering2 = new FontFamilyNumbering() { Val = 3 };
            FontCharSet fontCharSet2 = new FontCharSet() { Val = 134 };

            font2.Append(bold1);
            font2.Append(fontSize2);
            font2.Append(fontName2);
            font2.Append(fontFamilyNumbering2);
            font2.Append(fontCharSet2);

            Font font3 = new Font();
            FontSize fontSize3 = new FontSize() { Val = 9D };
            FontName fontName3 = new FontName() { Val = "宋体" };
            FontFamilyNumbering fontFamilyNumbering3 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet3 = new FontCharSet() { Val = 134 };
            FontScheme fontScheme2 = new FontScheme() { Val = FontSchemeValues.Minor };

            font3.Append(fontSize3);
            font3.Append(fontName3);
            font3.Append(fontFamilyNumbering3);
            font3.Append(fontCharSet3);
            font3.Append(fontScheme2);

            Font font4 = new Font();
            FontSize fontSize4 = new FontSize() { Val = 9D };
            FontName fontName4 = new FontName() { Val = "宋体" };
            FontFamilyNumbering fontFamilyNumbering4 = new FontFamilyNumbering() { Val = 3 };
            FontCharSet fontCharSet4 = new FontCharSet() { Val = 134 };

            font4.Append(fontSize4);
            font4.Append(fontName4);
            font4.Append(fontFamilyNumbering4);
            font4.Append(fontCharSet4);

            Font font5 = new Font();
            DocumentFormat.OpenXml.Spreadsheet.Bold bold2 = new DocumentFormat.OpenXml.Spreadsheet.Bold();
            FontSize fontSize5 = new FontSize() { Val = 16D };
            FontName fontName5 = new FontName() { Val = "Times New Roman" };
            FontFamilyNumbering fontFamilyNumbering5 = new FontFamilyNumbering() { Val = 1 };

            font5.Append(bold2);
            font5.Append(fontSize5);
            font5.Append(fontName5);
            font5.Append(fontFamilyNumbering5);

            Font font6 = new Font();
            DocumentFormat.OpenXml.Spreadsheet.Bold bold3 = new DocumentFormat.OpenXml.Spreadsheet.Bold();
            FontSize fontSize6 = new FontSize() { Val = 10D };
            FontName fontName6 = new FontName() { Val = "宋体" };
            FontFamilyNumbering fontFamilyNumbering6 = new FontFamilyNumbering() { Val = 3 };
            FontCharSet fontCharSet5 = new FontCharSet() { Val = 134 };

            font6.Append(bold3);
            font6.Append(fontSize6);
            font6.Append(fontName6);
            font6.Append(fontFamilyNumbering6);
            font6.Append(fontCharSet5);

            Font font7 = new Font();
            DocumentFormat.OpenXml.Spreadsheet.Bold bold4 = new DocumentFormat.OpenXml.Spreadsheet.Bold();
            FontSize fontSize7 = new FontSize() { Val = 8D };
            FontName fontName7 = new FontName() { Val = "宋体" };
            FontFamilyNumbering fontFamilyNumbering7 = new FontFamilyNumbering() { Val = 3 };
            FontCharSet fontCharSet6 = new FontCharSet() { Val = 134 };

            font7.Append(bold4);
            font7.Append(fontSize7);
            font7.Append(fontName7);
            font7.Append(fontFamilyNumbering7);
            font7.Append(fontCharSet6);

            Font font8 = new Font();
            FontSize fontSize8 = new FontSize() { Val = 10D };
            FontName fontName8 = new FontName() { Val = "宋体" };
            FontFamilyNumbering fontFamilyNumbering8 = new FontFamilyNumbering() { Val = 3 };
            FontCharSet fontCharSet7 = new FontCharSet() { Val = 134 };

            font8.Append(fontSize8);
            font8.Append(fontName8);
            font8.Append(fontFamilyNumbering8);
            font8.Append(fontCharSet7);

            Font font9 = new Font();
            FontSize fontSize9 = new FontSize() { Val = 8D };
            FontName fontName9 = new FontName() { Val = "宋体" };
            FontFamilyNumbering fontFamilyNumbering9 = new FontFamilyNumbering() { Val = 3 };
            FontCharSet fontCharSet8 = new FontCharSet() { Val = 134 };

            font9.Append(fontSize9);
            font9.Append(fontName9);
            font9.Append(fontFamilyNumbering9);
            font9.Append(fontCharSet8);

            Font font10 = new Font();
            FontSize fontSize10 = new FontSize() { Val = 8D };
            FontName fontName10 = new FontName() { Val = "Times New Roman" };
            FontFamilyNumbering fontFamilyNumbering10 = new FontFamilyNumbering() { Val = 1 };

            font10.Append(fontSize10);
            font10.Append(fontName10);
            font10.Append(fontFamilyNumbering10);

            Font font11 = new Font();
            FontSize fontSize11 = new FontSize() { Val = 7D };
            FontName fontName11 = new FontName() { Val = "Times New Roman" };
            FontFamilyNumbering fontFamilyNumbering11 = new FontFamilyNumbering() { Val = 1 };

            font11.Append(fontSize11);
            font11.Append(fontName11);
            font11.Append(fontFamilyNumbering11);

            Font font12 = new Font();
            FontSize fontSize12 = new FontSize() { Val = 7D };
            FontName fontName12 = new FontName() { Val = "宋体" };
            FontFamilyNumbering fontFamilyNumbering12 = new FontFamilyNumbering() { Val = 3 };
            FontCharSet fontCharSet9 = new FontCharSet() { Val = 134 };

            font12.Append(fontSize12);
            font12.Append(fontName12);
            font12.Append(fontFamilyNumbering12);
            font12.Append(fontCharSet9);

            Font font13 = new Font();
            FontSize fontSize13 = new FontSize() { Val = 6D };
            FontName fontName13 = new FontName() { Val = "宋体" };
            FontFamilyNumbering fontFamilyNumbering13 = new FontFamilyNumbering() { Val = 3 };
            FontCharSet fontCharSet10 = new FontCharSet() { Val = 134 };

            font13.Append(fontSize13);
            font13.Append(fontName13);
            font13.Append(fontFamilyNumbering13);
            font13.Append(fontCharSet10);

            Font font14 = new Font();
            FontSize fontSize14 = new FontSize() { Val = 9D };
            FontName fontName14 = new FontName() { Val = "Times New Roman" };
            FontFamilyNumbering fontFamilyNumbering14 = new FontFamilyNumbering() { Val = 1 };

            font14.Append(fontSize14);
            font14.Append(fontName14);
            font14.Append(fontFamilyNumbering14);

            fonts1.Append(font1);
            fonts1.Append(font2);
            fonts1.Append(font3);
            fonts1.Append(font4);
            fonts1.Append(font5);
            fonts1.Append(font6);
            fonts1.Append(font7);
            fonts1.Append(font8);
            fonts1.Append(font9);
            fonts1.Append(font10);
            fonts1.Append(font11);
            fonts1.Append(font12);
            fonts1.Append(font13);
            fonts1.Append(font14);

            Fills fills1 = new Fills() { Count = (UInt32Value)2U };

            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };

            fill1.Append(patternFill1);

            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };

            fill2.Append(patternFill2);

            fills1.Append(fill1);
            fills1.Append(fill2);

            Borders borders1 = new Borders() { Count = (UInt32Value)3U };

            DocumentFormat.OpenXml.Spreadsheet.Border border1 = new DocumentFormat.OpenXml.Spreadsheet.Border();
            LeftBorder leftBorder1 = new LeftBorder();
            RightBorder rightBorder1 = new RightBorder();
            TopBorder topBorder1 = new TopBorder();
            BottomBorder bottomBorder1 = new BottomBorder();
            DiagonalBorder diagonalBorder1 = new DiagonalBorder();

            border1.Append(leftBorder1);
            border1.Append(rightBorder1);
            border1.Append(topBorder1);
            border1.Append(bottomBorder1);
            border1.Append(diagonalBorder1);

            DocumentFormat.OpenXml.Spreadsheet.Border border2 = new DocumentFormat.OpenXml.Spreadsheet.Border();
            LeftBorder leftBorder2 = new LeftBorder();
            RightBorder rightBorder2 = new RightBorder();
            TopBorder topBorder2 = new TopBorder();

            BottomBorder bottomBorder2 = new BottomBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color2 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Indexed = (UInt32Value)64U };

            bottomBorder2.Append(color2);
            DiagonalBorder diagonalBorder2 = new DiagonalBorder();

            border2.Append(leftBorder2);
            border2.Append(rightBorder2);
            border2.Append(topBorder2);
            border2.Append(bottomBorder2);
            border2.Append(diagonalBorder2);

            DocumentFormat.OpenXml.Spreadsheet.Border border3 = new DocumentFormat.OpenXml.Spreadsheet.Border();

            LeftBorder leftBorder3 = new LeftBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color3 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Indexed = (UInt32Value)64U };

            leftBorder3.Append(color3);

            RightBorder rightBorder3 = new RightBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color4 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Indexed = (UInt32Value)64U };

            rightBorder3.Append(color4);

            TopBorder topBorder3 = new TopBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color5 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Indexed = (UInt32Value)64U };

            topBorder3.Append(color5);

            BottomBorder bottomBorder3 = new BottomBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color6 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Indexed = (UInt32Value)64U };

            bottomBorder3.Append(color6);
            DiagonalBorder diagonalBorder3 = new DiagonalBorder();

            border3.Append(leftBorder3);
            border3.Append(rightBorder3);
            border3.Append(topBorder3);
            border3.Append(bottomBorder3);
            border3.Append(diagonalBorder3);

            borders1.Append(border1);
            borders1.Append(border2);
            borders1.Append(border3);

            CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)1U };

            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };
            Alignment alignment1 = new Alignment() { Vertical = VerticalAlignmentValues.Center };

            cellFormat1.Append(alignment1);

            cellStyleFormats1.Append(cellFormat1);

            CellFormats cellFormats1 = new CellFormats() { Count = (UInt32Value)21U };

            CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            Alignment alignment2 = new Alignment() { Vertical = VerticalAlignmentValues.Center };

            cellFormat2.Append(alignment2);

            CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)5U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)2U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment3 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

            cellFormat3.Append(alignment3);

            CellFormat cellFormat4 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)5U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)2U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment4 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true };

            cellFormat4.Append(alignment4);

            CellFormat cellFormat5 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)6U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)2U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment5 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true };

            cellFormat5.Append(alignment5);

            CellFormat cellFormat6 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)5U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)2U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment6 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true };

            cellFormat6.Append(alignment6);

            CellFormat cellFormat7 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)5U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)2U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment7 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

            cellFormat7.Append(alignment7);

            CellFormat cellFormat8 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)2U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true, ApplyProtection = true };
            Alignment alignment8 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true };
            Protection protection1 = new Protection() { Locked = false };

            cellFormat8.Append(alignment8);
            cellFormat8.Append(protection1);

            CellFormat cellFormat9 = new CellFormat() { NumberFormatId = (UInt32Value)49U, FontId = (UInt32Value)9U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)2U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment9 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center, WrapText = true };

            cellFormat9.Append(alignment9);

            CellFormat cellFormat10 = new CellFormat() { NumberFormatId = (UInt32Value)49U, FontId = (UInt32Value)10U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)2U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment10 = new Alignment() { Vertical = VerticalAlignmentValues.Center, WrapText = true };

            cellFormat10.Append(alignment10);

            CellFormat cellFormat11 = new CellFormat() { NumberFormatId = (UInt32Value)49U, FontId = (UInt32Value)11U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)2U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment11 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center, WrapText = true };

            cellFormat11.Append(alignment11);

            CellFormat cellFormat12 = new CellFormat() { NumberFormatId = (UInt32Value)49U, FontId = (UInt32Value)12U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)2U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment12 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true };

            cellFormat12.Append(alignment12);

            CellFormat cellFormat13 = new CellFormat() { NumberFormatId = (UInt32Value)49U, FontId = (UInt32Value)13U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)2U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment13 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true };

            cellFormat13.Append(alignment13);

            CellFormat cellFormat14 = new CellFormat() { NumberFormatId = (UInt32Value)49U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)2U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment14 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true };

            cellFormat14.Append(alignment14);

            CellFormat cellFormat15 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment15 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat15.Append(alignment15);

            CellFormat cellFormat16 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)4U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment16 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat16.Append(alignment16);

            CellFormat cellFormat17 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment17 = new Alignment() { Vertical = VerticalAlignmentValues.Center };

            cellFormat17.Append(alignment17);

            CellFormat cellFormat18 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)5U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)2U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment18 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

            cellFormat18.Append(alignment18);

            CellFormat cellFormat19 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)2U, FormatId = (UInt32Value)0U, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment19 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

            cellFormat19.Append(alignment19);

            CellFormat cellFormat20 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)7U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)2U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment20 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true };

            cellFormat20.Append(alignment20);

            CellFormat cellFormat21 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)8U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)2U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment21 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center, WrapText = true };

            cellFormat21.Append(alignment21);

            CellFormat cellFormat22 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)8U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)2U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment22 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true };

            cellFormat22.Append(alignment22);

            cellFormats1.Append(cellFormat2);
            cellFormats1.Append(cellFormat3);
            cellFormats1.Append(cellFormat4);
            cellFormats1.Append(cellFormat5);
            cellFormats1.Append(cellFormat6);
            cellFormats1.Append(cellFormat7);
            cellFormats1.Append(cellFormat8);
            cellFormats1.Append(cellFormat9);
            cellFormats1.Append(cellFormat10);
            cellFormats1.Append(cellFormat11);
            cellFormats1.Append(cellFormat12);
            cellFormats1.Append(cellFormat13);
            cellFormats1.Append(cellFormat14);
            cellFormats1.Append(cellFormat15);
            cellFormats1.Append(cellFormat16);
            cellFormats1.Append(cellFormat17);
            cellFormats1.Append(cellFormat18);
            cellFormats1.Append(cellFormat19);
            cellFormats1.Append(cellFormat20);
            cellFormats1.Append(cellFormat21);
            cellFormats1.Append(cellFormat22);

            CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)1U };
            CellStyle cellStyle1 = new CellStyle() { Name = "常规", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };

            cellStyles1.Append(cellStyle1);
            DifferentialFormats differentialFormats1 = new DifferentialFormats() { Count = (UInt32Value)0U };
            TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleLight16" };

            StylesheetExtensionList stylesheetExtensionList1 = new StylesheetExtensionList();

            StylesheetExtension stylesheetExtension1 = new StylesheetExtension() { Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" };
            stylesheetExtension1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            X14.SlicerStyles slicerStyles1 = new X14.SlicerStyles() { DefaultSlicerStyle = "SlicerStyleLight1" };

            stylesheetExtension1.Append(slicerStyles1);

            stylesheetExtensionList1.Append(stylesheetExtension1);

            stylesheet1.Append(fonts1);
            stylesheet1.Append(fills1);
            stylesheet1.Append(borders1);
            stylesheet1.Append(cellStyleFormats1);
            stylesheet1.Append(cellFormats1);
            stylesheet1.Append(cellStyles1);
            stylesheet1.Append(differentialFormats1);
            stylesheet1.Append(tableStyles1);
            stylesheet1.Append(stylesheetExtensionList1);

            workbookStylesPart1.Stylesheet = stylesheet1;
        }

        // Generates content of themePart1.
        private void GenerateThemePart1Content(ThemePart themePart1)
        {
            A.Theme theme1 = new A.Theme() { Name = "Office 主题​​" };
            theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.ThemeElements themeElements1 = new A.ThemeElements();

            A.ColorScheme colorScheme1 = new A.ColorScheme() { Name = "Office" };

            A.Dark1Color dark1Color1 = new A.Dark1Color();
            A.SystemColor systemColor1 = new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = "000000" };

            dark1Color1.Append(systemColor1);

            A.Light1Color light1Color1 = new A.Light1Color();
            A.SystemColor systemColor2 = new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" };

            light1Color1.Append(systemColor2);

            A.Dark2Color dark2Color1 = new A.Dark2Color();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "1F497D" };

            dark2Color1.Append(rgbColorModelHex1);

            A.Light2Color light2Color1 = new A.Light2Color();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "EEECE1" };

            light2Color1.Append(rgbColorModelHex2);

            A.Accent1Color accent1Color1 = new A.Accent1Color();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "4F81BD" };

            accent1Color1.Append(rgbColorModelHex3);

            A.Accent2Color accent2Color1 = new A.Accent2Color();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "C0504D" };

            accent2Color1.Append(rgbColorModelHex4);

            A.Accent3Color accent3Color1 = new A.Accent3Color();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "9BBB59" };

            accent3Color1.Append(rgbColorModelHex5);

            A.Accent4Color accent4Color1 = new A.Accent4Color();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "8064A2" };

            accent4Color1.Append(rgbColorModelHex6);

            A.Accent5Color accent5Color1 = new A.Accent5Color();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "4BACC6" };

            accent5Color1.Append(rgbColorModelHex7);

            A.Accent6Color accent6Color1 = new A.Accent6Color();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "F79646" };

            accent6Color1.Append(rgbColorModelHex8);

            A.Hyperlink hyperlink1 = new A.Hyperlink();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex() { Val = "0000FF" };

            hyperlink1.Append(rgbColorModelHex9);

            A.FollowedHyperlinkColor followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex() { Val = "800080" };

            followedHyperlinkColor1.Append(rgbColorModelHex10);

            colorScheme1.Append(dark1Color1);
            colorScheme1.Append(light1Color1);
            colorScheme1.Append(dark2Color1);
            colorScheme1.Append(light2Color1);
            colorScheme1.Append(accent1Color1);
            colorScheme1.Append(accent2Color1);
            colorScheme1.Append(accent3Color1);
            colorScheme1.Append(accent4Color1);
            colorScheme1.Append(accent5Color1);
            colorScheme1.Append(accent6Color1);
            colorScheme1.Append(hyperlink1);
            colorScheme1.Append(followedHyperlinkColor1);

            A.FontScheme fontScheme3 = new A.FontScheme() { Name = "Office" };

            A.MajorFont majorFont1 = new A.MajorFont();
            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "Cambria" };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont1 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ Ｐゴシック" };
            A.SupplementalFont supplementalFont2 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont3 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont4 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont5 = new A.SupplementalFont() { Script = "Arab", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont6 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont7 = new A.SupplementalFont() { Script = "Thai", Typeface = "Tahoma" };
            A.SupplementalFont supplementalFont8 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont9 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont10 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont11 = new A.SupplementalFont() { Script = "Khmr", Typeface = "MoolBoran" };
            A.SupplementalFont supplementalFont12 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont13 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont14 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont15 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont16 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont17 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont18 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont19 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont20 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont21 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont22 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont23 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont24 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont25 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont26 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont27 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont28 = new A.SupplementalFont() { Script = "Viet", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont29 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont30 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };

            majorFont1.Append(latinFont1);
            majorFont1.Append(eastAsianFont1);
            majorFont1.Append(complexScriptFont1);
            majorFont1.Append(supplementalFont1);
            majorFont1.Append(supplementalFont2);
            majorFont1.Append(supplementalFont3);
            majorFont1.Append(supplementalFont4);
            majorFont1.Append(supplementalFont5);
            majorFont1.Append(supplementalFont6);
            majorFont1.Append(supplementalFont7);
            majorFont1.Append(supplementalFont8);
            majorFont1.Append(supplementalFont9);
            majorFont1.Append(supplementalFont10);
            majorFont1.Append(supplementalFont11);
            majorFont1.Append(supplementalFont12);
            majorFont1.Append(supplementalFont13);
            majorFont1.Append(supplementalFont14);
            majorFont1.Append(supplementalFont15);
            majorFont1.Append(supplementalFont16);
            majorFont1.Append(supplementalFont17);
            majorFont1.Append(supplementalFont18);
            majorFont1.Append(supplementalFont19);
            majorFont1.Append(supplementalFont20);
            majorFont1.Append(supplementalFont21);
            majorFont1.Append(supplementalFont22);
            majorFont1.Append(supplementalFont23);
            majorFont1.Append(supplementalFont24);
            majorFont1.Append(supplementalFont25);
            majorFont1.Append(supplementalFont26);
            majorFont1.Append(supplementalFont27);
            majorFont1.Append(supplementalFont28);
            majorFont1.Append(supplementalFont29);
            majorFont1.Append(supplementalFont30);

            A.MinorFont minorFont1 = new A.MinorFont();
            A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "Calibri" };
            A.EastAsianFont eastAsianFont2 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont31 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ Ｐゴシック" };
            A.SupplementalFont supplementalFont32 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont33 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont34 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont35 = new A.SupplementalFont() { Script = "Arab", Typeface = "Arial" };
            A.SupplementalFont supplementalFont36 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Arial" };
            A.SupplementalFont supplementalFont37 = new A.SupplementalFont() { Script = "Thai", Typeface = "Tahoma" };
            A.SupplementalFont supplementalFont38 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont39 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont40 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont41 = new A.SupplementalFont() { Script = "Khmr", Typeface = "DaunPenh" };
            A.SupplementalFont supplementalFont42 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont43 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont44 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont45 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont46 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont47 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont48 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont49 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont50 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont51 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont52 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont53 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont54 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont55 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont56 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont57 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont58 = new A.SupplementalFont() { Script = "Viet", Typeface = "Arial" };
            A.SupplementalFont supplementalFont59 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont60 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };

            minorFont1.Append(latinFont2);
            minorFont1.Append(eastAsianFont2);
            minorFont1.Append(complexScriptFont2);
            minorFont1.Append(supplementalFont31);
            minorFont1.Append(supplementalFont32);
            minorFont1.Append(supplementalFont33);
            minorFont1.Append(supplementalFont34);
            minorFont1.Append(supplementalFont35);
            minorFont1.Append(supplementalFont36);
            minorFont1.Append(supplementalFont37);
            minorFont1.Append(supplementalFont38);
            minorFont1.Append(supplementalFont39);
            minorFont1.Append(supplementalFont40);
            minorFont1.Append(supplementalFont41);
            minorFont1.Append(supplementalFont42);
            minorFont1.Append(supplementalFont43);
            minorFont1.Append(supplementalFont44);
            minorFont1.Append(supplementalFont45);
            minorFont1.Append(supplementalFont46);
            minorFont1.Append(supplementalFont47);
            minorFont1.Append(supplementalFont48);
            minorFont1.Append(supplementalFont49);
            minorFont1.Append(supplementalFont50);
            minorFont1.Append(supplementalFont51);
            minorFont1.Append(supplementalFont52);
            minorFont1.Append(supplementalFont53);
            minorFont1.Append(supplementalFont54);
            minorFont1.Append(supplementalFont55);
            minorFont1.Append(supplementalFont56);
            minorFont1.Append(supplementalFont57);
            minorFont1.Append(supplementalFont58);
            minorFont1.Append(supplementalFont59);
            minorFont1.Append(supplementalFont60);

            fontScheme3.Append(majorFont1);
            fontScheme3.Append(minorFont1);

            A.FormatScheme formatScheme1 = new A.FormatScheme() { Name = "Office" };

            A.FillStyleList fillStyleList1 = new A.FillStyleList();

            A.SolidFill solidFill1 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill1.Append(schemeColor1);

            A.GradientFill gradientFill1 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList1 = new A.GradientStopList();

            A.GradientStop gradientStop1 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint1 = new A.Tint() { Val = 50000 };
            A.SaturationModulation saturationModulation1 = new A.SaturationModulation() { Val = 300000 };

            schemeColor2.Append(tint1);
            schemeColor2.Append(saturationModulation1);

            gradientStop1.Append(schemeColor2);

            A.GradientStop gradientStop2 = new A.GradientStop() { Position = 35000 };

            A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint2 = new A.Tint() { Val = 37000 };
            A.SaturationModulation saturationModulation2 = new A.SaturationModulation() { Val = 300000 };

            schemeColor3.Append(tint2);
            schemeColor3.Append(saturationModulation2);

            gradientStop2.Append(schemeColor3);

            A.GradientStop gradientStop3 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint3 = new A.Tint() { Val = 15000 };
            A.SaturationModulation saturationModulation3 = new A.SaturationModulation() { Val = 350000 };

            schemeColor4.Append(tint3);
            schemeColor4.Append(saturationModulation3);

            gradientStop3.Append(schemeColor4);

            gradientStopList1.Append(gradientStop1);
            gradientStopList1.Append(gradientStop2);
            gradientStopList1.Append(gradientStop3);
            A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill() { Angle = 16200000, Scaled = true };

            gradientFill1.Append(gradientStopList1);
            gradientFill1.Append(linearGradientFill1);

            A.GradientFill gradientFill2 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList2 = new A.GradientStopList();

            A.GradientStop gradientStop4 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor5 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade1 = new A.Shade() { Val = 51000 };
            A.SaturationModulation saturationModulation4 = new A.SaturationModulation() { Val = 130000 };

            schemeColor5.Append(shade1);
            schemeColor5.Append(saturationModulation4);

            gradientStop4.Append(schemeColor5);

            A.GradientStop gradientStop5 = new A.GradientStop() { Position = 80000 };

            A.SchemeColor schemeColor6 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade2 = new A.Shade() { Val = 93000 };
            A.SaturationModulation saturationModulation5 = new A.SaturationModulation() { Val = 130000 };

            schemeColor6.Append(shade2);
            schemeColor6.Append(saturationModulation5);

            gradientStop5.Append(schemeColor6);

            A.GradientStop gradientStop6 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor7 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade3 = new A.Shade() { Val = 94000 };
            A.SaturationModulation saturationModulation6 = new A.SaturationModulation() { Val = 135000 };

            schemeColor7.Append(shade3);
            schemeColor7.Append(saturationModulation6);

            gradientStop6.Append(schemeColor7);

            gradientStopList2.Append(gradientStop4);
            gradientStopList2.Append(gradientStop5);
            gradientStopList2.Append(gradientStop6);
            A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill() { Angle = 16200000, Scaled = false };

            gradientFill2.Append(gradientStopList2);
            gradientFill2.Append(linearGradientFill2);

            fillStyleList1.Append(solidFill1);
            fillStyleList1.Append(gradientFill1);
            fillStyleList1.Append(gradientFill2);

            A.LineStyleList lineStyleList1 = new A.LineStyleList();

            A.Outline outline1 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill2 = new A.SolidFill();

            A.SchemeColor schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade4 = new A.Shade() { Val = 95000 };
            A.SaturationModulation saturationModulation7 = new A.SaturationModulation() { Val = 105000 };

            schemeColor8.Append(shade4);
            schemeColor8.Append(saturationModulation7);

            solidFill2.Append(schemeColor8);
            A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline1.Append(solidFill2);
            outline1.Append(presetDash1);

            A.Outline outline2 = new A.Outline() { Width = 25400, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.SchemeColor schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill3.Append(schemeColor9);
            A.PresetDash presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline2.Append(solidFill3);
            outline2.Append(presetDash2);

            A.Outline outline3 = new A.Outline() { Width = 38100, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill4.Append(schemeColor10);
            A.PresetDash presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline3.Append(solidFill4);
            outline3.Append(presetDash3);

            lineStyleList1.Append(outline1);
            lineStyleList1.Append(outline2);
            lineStyleList1.Append(outline3);

            A.EffectStyleList effectStyleList1 = new A.EffectStyleList();

            A.EffectStyle effectStyle1 = new A.EffectStyle();

            A.EffectList effectList1 = new A.EffectList();

            A.OuterShadow outerShadow1 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha1 = new A.Alpha() { Val = 38000 };

            rgbColorModelHex11.Append(alpha1);

            outerShadow1.Append(rgbColorModelHex11);

            effectList1.Append(outerShadow1);

            effectStyle1.Append(effectList1);

            A.EffectStyle effectStyle2 = new A.EffectStyle();

            A.EffectList effectList2 = new A.EffectList();

            A.OuterShadow outerShadow2 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 23000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex12 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha2 = new A.Alpha() { Val = 35000 };

            rgbColorModelHex12.Append(alpha2);

            outerShadow2.Append(rgbColorModelHex12);

            effectList2.Append(outerShadow2);

            effectStyle2.Append(effectList2);

            A.EffectStyle effectStyle3 = new A.EffectStyle();

            A.EffectList effectList3 = new A.EffectList();

            A.OuterShadow outerShadow3 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 23000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex13 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha3 = new A.Alpha() { Val = 35000 };

            rgbColorModelHex13.Append(alpha3);

            outerShadow3.Append(rgbColorModelHex13);

            effectList3.Append(outerShadow3);

            A.Scene3DType scene3DType1 = new A.Scene3DType();

            A.Camera camera1 = new A.Camera() { Preset = A.PresetCameraValues.OrthographicFront };
            A.Rotation rotation1 = new A.Rotation() { Latitude = 0, Longitude = 0, Revolution = 0 };

            camera1.Append(rotation1);

            A.LightRig lightRig1 = new A.LightRig() { Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };
            A.Rotation rotation2 = new A.Rotation() { Latitude = 0, Longitude = 0, Revolution = 1200000 };

            lightRig1.Append(rotation2);

            scene3DType1.Append(camera1);
            scene3DType1.Append(lightRig1);

            A.Shape3DType shape3DType1 = new A.Shape3DType();
            A.BevelTop bevelTop1 = new A.BevelTop() { Width = 63500L, Height = 25400L };

            shape3DType1.Append(bevelTop1);

            effectStyle3.Append(effectList3);
            effectStyle3.Append(scene3DType1);
            effectStyle3.Append(shape3DType1);

            effectStyleList1.Append(effectStyle1);
            effectStyleList1.Append(effectStyle2);
            effectStyleList1.Append(effectStyle3);

            A.BackgroundFillStyleList backgroundFillStyleList1 = new A.BackgroundFillStyleList();

            A.SolidFill solidFill5 = new A.SolidFill();
            A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill5.Append(schemeColor11);

            A.GradientFill gradientFill3 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList3 = new A.GradientStopList();

            A.GradientStop gradientStop7 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint4 = new A.Tint() { Val = 40000 };
            A.SaturationModulation saturationModulation8 = new A.SaturationModulation() { Val = 350000 };

            schemeColor12.Append(tint4);
            schemeColor12.Append(saturationModulation8);

            gradientStop7.Append(schemeColor12);

            A.GradientStop gradientStop8 = new A.GradientStop() { Position = 40000 };

            A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint5 = new A.Tint() { Val = 45000 };
            A.Shade shade5 = new A.Shade() { Val = 99000 };
            A.SaturationModulation saturationModulation9 = new A.SaturationModulation() { Val = 350000 };

            schemeColor13.Append(tint5);
            schemeColor13.Append(shade5);
            schemeColor13.Append(saturationModulation9);

            gradientStop8.Append(schemeColor13);

            A.GradientStop gradientStop9 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade6 = new A.Shade() { Val = 20000 };
            A.SaturationModulation saturationModulation10 = new A.SaturationModulation() { Val = 255000 };

            schemeColor14.Append(shade6);
            schemeColor14.Append(saturationModulation10);

            gradientStop9.Append(schemeColor14);

            gradientStopList3.Append(gradientStop7);
            gradientStopList3.Append(gradientStop8);
            gradientStopList3.Append(gradientStop9);

            A.PathGradientFill pathGradientFill1 = new A.PathGradientFill() { Path = A.PathShadeValues.Circle };
            A.FillToRectangle fillToRectangle1 = new A.FillToRectangle() { Left = 50000, Top = -80000, Right = 50000, Bottom = 180000 };

            pathGradientFill1.Append(fillToRectangle1);

            gradientFill3.Append(gradientStopList3);
            gradientFill3.Append(pathGradientFill1);

            A.GradientFill gradientFill4 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList4 = new A.GradientStopList();

            A.GradientStop gradientStop10 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint6 = new A.Tint() { Val = 80000 };
            A.SaturationModulation saturationModulation11 = new A.SaturationModulation() { Val = 300000 };

            schemeColor15.Append(tint6);
            schemeColor15.Append(saturationModulation11);

            gradientStop10.Append(schemeColor15);

            A.GradientStop gradientStop11 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor16 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade7 = new A.Shade() { Val = 30000 };
            A.SaturationModulation saturationModulation12 = new A.SaturationModulation() { Val = 200000 };

            schemeColor16.Append(shade7);
            schemeColor16.Append(saturationModulation12);

            gradientStop11.Append(schemeColor16);

            gradientStopList4.Append(gradientStop10);
            gradientStopList4.Append(gradientStop11);

            A.PathGradientFill pathGradientFill2 = new A.PathGradientFill() { Path = A.PathShadeValues.Circle };
            A.FillToRectangle fillToRectangle2 = new A.FillToRectangle() { Left = 50000, Top = 50000, Right = 50000, Bottom = 50000 };

            pathGradientFill2.Append(fillToRectangle2);

            gradientFill4.Append(gradientStopList4);
            gradientFill4.Append(pathGradientFill2);

            backgroundFillStyleList1.Append(solidFill5);
            backgroundFillStyleList1.Append(gradientFill3);
            backgroundFillStyleList1.Append(gradientFill4);

            formatScheme1.Append(fillStyleList1);
            formatScheme1.Append(lineStyleList1);
            formatScheme1.Append(effectStyleList1);
            formatScheme1.Append(backgroundFillStyleList1);

            themeElements1.Append(colorScheme1);
            themeElements1.Append(fontScheme3);
            themeElements1.Append(formatScheme1);
            A.ObjectDefaults objectDefaults1 = new A.ObjectDefaults();
            A.ExtraColorSchemeList extraColorSchemeList1 = new A.ExtraColorSchemeList();

            theme1.Append(themeElements1);
            theme1.Append(objectDefaults1);
            theme1.Append(extraColorSchemeList1);

            themePart1.Theme = theme1;
        }

        // Generates content of spreadsheetPrinterSettingsPart1.
        private void GenerateSpreadsheetPrinterSettingsPart1Content(SpreadsheetPrinterSettingsPart spreadsheetPrinterSettingsPart1)
        {
            System.IO.Stream data = GetBinaryDataStream(spreadsheetPrinterSettingsPart1Data);
            spreadsheetPrinterSettingsPart1.FeedData(data);
            data.Close();
        }

        // Generates content of sharedStringTablePart1.
        private void GenerateSharedStringTablePart1Content(SharedStringTablePart sharedStringTablePart1)
        {
            SharedStringTable sharedStringTable1 = new SharedStringTable() { Count = (UInt32Value)61U, UniqueCount = (UInt32Value)56U };

            SharedStringItem sharedStringItem1 = new SharedStringItem();
            Text text1 = new Text();
            text1.Text = "HNCDI院外立项课题明细";
            PhoneticProperties phoneticProperties2 = new PhoneticProperties() { FontId = (UInt32Value)3U, Type = PhoneticValues.NoConversion };

            sharedStringItem1.Append(text1);
            sharedStringItem1.Append(phoneticProperties2);

            SharedStringItem sharedStringItem2 = new SharedStringItem();
            Text text2 = new Text();
            text2.Text = "序号";
            PhoneticProperties phoneticProperties3 = new PhoneticProperties() { FontId = (UInt32Value)3U, Type = PhoneticValues.NoConversion };

            sharedStringItem2.Append(text2);
            sharedStringItem2.Append(phoneticProperties3);

            SharedStringItem sharedStringItem3 = new SharedStringItem();
            Text text3 = new Text();
            text3.Text = "项目编号";
            PhoneticProperties phoneticProperties4 = new PhoneticProperties() { FontId = (UInt32Value)3U, Type = PhoneticValues.NoConversion };

            sharedStringItem3.Append(text3);
            sharedStringItem3.Append(phoneticProperties4);

            SharedStringItem sharedStringItem4 = new SharedStringItem();
            Text text4 = new Text();
            text4.Text = "立项部门";
            PhoneticProperties phoneticProperties5 = new PhoneticProperties() { FontId = (UInt32Value)3U, Type = PhoneticValues.NoConversion };

            sharedStringItem4.Append(text4);
            sharedStringItem4.Append(phoneticProperties5);

            SharedStringItem sharedStringItem5 = new SharedStringItem();
            Text text5 = new Text();
            text5.Text = "立项时间";
            PhoneticProperties phoneticProperties6 = new PhoneticProperties() { FontId = (UInt32Value)3U, Type = PhoneticValues.NoConversion };

            sharedStringItem5.Append(text5);
            sharedStringItem5.Append(phoneticProperties6);

            SharedStringItem sharedStringItem6 = new SharedStringItem();
            Text text6 = new Text();
            text6.Text = "项目名称";
            PhoneticProperties phoneticProperties7 = new PhoneticProperties() { FontId = (UInt32Value)3U, Type = PhoneticValues.NoConversion };

            sharedStringItem6.Append(text6);
            sharedStringItem6.Append(phoneticProperties7);

            SharedStringItem sharedStringItem7 = new SharedStringItem();
            Text text7 = new Text();
            text7.Text = "乙方";
            PhoneticProperties phoneticProperties8 = new PhoneticProperties() { FontId = (UInt32Value)2U, Type = PhoneticValues.NoConversion };

            sharedStringItem7.Append(text7);
            sharedStringItem7.Append(phoneticProperties8);

            SharedStringItem sharedStringItem8 = new SharedStringItem();
            Text text8 = new Text();
            text8.Text = "负责人";
            PhoneticProperties phoneticProperties9 = new PhoneticProperties() { FontId = (UInt32Value)3U, Type = PhoneticValues.NoConversion };

            sharedStringItem8.Append(text8);
            sharedStringItem8.Append(phoneticProperties9);

            SharedStringItem sharedStringItem9 = new SharedStringItem();
            Text text9 = new Text();
            text9.Text = "合同时限";
            PhoneticProperties phoneticProperties10 = new PhoneticProperties() { FontId = (UInt32Value)3U, Type = PhoneticValues.NoConversion };

            sharedStringItem9.Append(text9);
            sharedStringItem9.Append(phoneticProperties10);

            SharedStringItem sharedStringItem10 = new SharedStringItem();
            Text text10 = new Text();
            text10.Text = "合同额(万)";
            PhoneticProperties phoneticProperties11 = new PhoneticProperties() { FontId = (UInt32Value)3U, Type = PhoneticValues.NoConversion };

            sharedStringItem10.Append(text10);
            sharedStringItem10.Append(phoneticProperties11);

            SharedStringItem sharedStringItem11 = new SharedStringItem();
            Text text11 = new Text();
            text11.Text = "其中：\n交通部";
            PhoneticProperties phoneticProperties12 = new PhoneticProperties() { FontId = (UInt32Value)3U, Type = PhoneticValues.NoConversion };

            sharedStringItem11.Append(text11);
            sharedStringItem11.Append(phoneticProperties12);

            SharedStringItem sharedStringItem12 = new SharedStringItem();
            Text text12 = new Text();
            text12.Text = "其中：\n交通厅";
            PhoneticProperties phoneticProperties13 = new PhoneticProperties() { FontId = (UInt32Value)3U, Type = PhoneticValues.NoConversion };

            sharedStringItem12.Append(text12);
            sharedStringItem12.Append(phoneticProperties13);

            SharedStringItem sharedStringItem13 = new SharedStringItem();
            Text text13 = new Text();
            text13.Text = "其中：\n科技厅";
            PhoneticProperties phoneticProperties14 = new PhoneticProperties() { FontId = (UInt32Value)2U, Type = PhoneticValues.NoConversion };

            sharedStringItem13.Append(text13);
            sharedStringItem13.Append(phoneticProperties14);

            SharedStringItem sharedStringItem14 = new SharedStringItem();
            Text text14 = new Text();
            text14.Text = "其中：工程";
            PhoneticProperties phoneticProperties15 = new PhoneticProperties() { FontId = (UInt32Value)3U, Type = PhoneticValues.NoConversion };

            sharedStringItem14.Append(text14);
            sharedStringItem14.Append(phoneticProperties15);

            SharedStringItem sharedStringItem15 = new SharedStringItem();
            Text text15 = new Text();
            text15.Text = "其中：其他";
            PhoneticProperties phoneticProperties16 = new PhoneticProperties() { FontId = (UInt32Value)3U, Type = PhoneticValues.NoConversion };

            sharedStringItem15.Append(text15);
            sharedStringItem15.Append(phoneticProperties16);

            SharedStringItem sharedStringItem16 = new SharedStringItem();
            Text text16 = new Text();
            text16.Text = "挂靠处室";
            PhoneticProperties phoneticProperties17 = new PhoneticProperties() { FontId = (UInt32Value)3U, Type = PhoneticValues.NoConversion };

            sharedStringItem16.Append(text16);
            sharedStringItem16.Append(phoneticProperties17);

            SharedStringItem sharedStringItem17 = new SharedStringItem();
            Text text17 = new Text();
            text17.Text = "主要人员";
            PhoneticProperties phoneticProperties18 = new PhoneticProperties() { FontId = (UInt32Value)3U, Type = PhoneticValues.NoConversion };

            sharedStringItem17.Append(text17);
            sharedStringItem17.Append(phoneticProperties18);

            SharedStringItem sharedStringItem18 = new SharedStringItem();
            Text text18 = new Text();
            text18.Text = "协作单位";
            PhoneticProperties phoneticProperties19 = new PhoneticProperties() { FontId = (UInt32Value)3U, Type = PhoneticValues.NoConversion };

            sharedStringItem18.Append(text18);
            sharedStringItem18.Append(phoneticProperties19);

            SharedStringItem sharedStringItem19 = new SharedStringItem();
            Text text19 = new Text();
            text19.Text = "合同号";
            PhoneticProperties phoneticProperties20 = new PhoneticProperties() { FontId = (UInt32Value)3U, Type = PhoneticValues.NoConversion };

            sharedStringItem19.Append(text19);
            sharedStringItem19.Append(phoneticProperties20);

            SharedStringItem sharedStringItem20 = new SharedStringItem();
            Text text20 = new Text();
            text20.Text = "鉴定情况";
            PhoneticProperties phoneticProperties21 = new PhoneticProperties() { FontId = (UInt32Value)3U, Type = PhoneticValues.NoConversion };

            sharedStringItem20.Append(text20);
            sharedStringItem20.Append(phoneticProperties21);

            SharedStringItem sharedStringItem21 = new SharedStringItem();
            Text text21 = new Text();
            text21.Text = "成果登记";
            PhoneticProperties phoneticProperties22 = new PhoneticProperties() { FontId = (UInt32Value)2U, Type = PhoneticValues.NoConversion };

            sharedStringItem21.Append(text21);
            sharedStringItem21.Append(phoneticProperties22);

            SharedStringItem sharedStringItem22 = new SharedStringItem();
            Text text22 = new Text();
            text22.Text = "获奖情况";
            PhoneticProperties phoneticProperties23 = new PhoneticProperties() { FontId = (UInt32Value)3U, Type = PhoneticValues.NoConversion };

            sharedStringItem22.Append(text22);
            sharedStringItem22.Append(phoneticProperties23);

            SharedStringItem sharedStringItem23 = new SharedStringItem();
            Text text23 = new Text();
            text23.Text = "知识产权";
            PhoneticProperties phoneticProperties24 = new PhoneticProperties() { FontId = (UInt32Value)2U, Type = PhoneticValues.NoConversion };

            sharedStringItem23.Append(text23);
            sharedStringItem23.Append(phoneticProperties24);

            SharedStringItem sharedStringItem24 = new SharedStringItem();
            Text text24 = new Text();
            text24.Text = "完成单位";
            PhoneticProperties phoneticProperties25 = new PhoneticProperties() { FontId = (UInt32Value)3U, Type = PhoneticValues.NoConversion };

            sharedStringItem24.Append(text24);
            sharedStringItem24.Append(phoneticProperties25);

            SharedStringItem sharedStringItem25 = new SharedStringItem();
            Text text25 = new Text();
            text25.Text = "完成人员";
            PhoneticProperties phoneticProperties26 = new PhoneticProperties() { FontId = (UInt32Value)3U, Type = PhoneticValues.NoConversion };

            sharedStringItem25.Append(text25);
            sharedStringItem25.Append(phoneticProperties26);

            SharedStringItem sharedStringItem26 = new SharedStringItem();
            Text text26 = new Text();
            text26.Text = "到账\n(万元)";
            PhoneticProperties phoneticProperties27 = new PhoneticProperties() { FontId = (UInt32Value)3U, Type = PhoneticValues.NoConversion };

            sharedStringItem26.Append(text26);
            sharedStringItem26.Append(phoneticProperties27);

            SharedStringItem sharedStringItem27 = new SharedStringItem();
            Text text27 = new Text();
            text27.Text = "支付外协(万元)";
            PhoneticProperties phoneticProperties28 = new PhoneticProperties() { FontId = (UInt32Value)3U, Type = PhoneticValues.NoConversion };

            sharedStringItem27.Append(text27);
            sharedStringItem27.Append(phoneticProperties28);

            SharedStringItem sharedStringItem28 = new SharedStringItem();
            Text text28 = new Text();
            text28.Text = "课题组报支(万元)";
            PhoneticProperties phoneticProperties29 = new PhoneticProperties() { FontId = (UInt32Value)3U, Type = PhoneticValues.NoConversion };

            sharedStringItem28.Append(text28);
            sharedStringItem28.Append(phoneticProperties29);

            SharedStringItem sharedStringItem29 = new SharedStringItem();
            Text text29 = new Text();
            text29.Text = "提取管理费(万元)";
            PhoneticProperties phoneticProperties30 = new PhoneticProperties() { FontId = (UInt32Value)3U, Type = PhoneticValues.NoConversion };

            sharedStringItem29.Append(text29);
            sharedStringItem29.Append(phoneticProperties30);

            SharedStringItem sharedStringItem30 = new SharedStringItem();
            Text text30 = new Text();
            text30.Text = "经费节余(万元)";
            PhoneticProperties phoneticProperties31 = new PhoneticProperties() { FontId = (UInt32Value)3U, Type = PhoneticValues.NoConversion };

            sharedStringItem30.Append(text30);
            sharedStringItem30.Append(phoneticProperties31);

            SharedStringItem sharedStringItem31 = new SharedStringItem();
            Text text31 = new Text();
            text31.Text = "是否主研";
            PhoneticProperties phoneticProperties32 = new PhoneticProperties() { FontId = (UInt32Value)2U, Type = PhoneticValues.NoConversion };

            sharedStringItem31.Append(text31);
            sharedStringItem31.Append(phoneticProperties32);

            SharedStringItem sharedStringItem32 = new SharedStringItem();
            Text text32 = new Text();
            text32.Text = "完成情况";
            PhoneticProperties phoneticProperties33 = new PhoneticProperties() { FontId = (UInt32Value)2U, Type = PhoneticValues.NoConversion };

            sharedStringItem32.Append(text32);
            sharedStringItem32.Append(phoneticProperties33);

            SharedStringItem sharedStringItem33 = new SharedStringItem();
            Text text33 = new Text();
            text33.Text = "归档情况";

            sharedStringItem33.Append(text33);

            SharedStringItem sharedStringItem34 = new SharedStringItem();
            Text text34 = new Text();
            text34.Text = "结题情况";
            PhoneticProperties phoneticProperties34 = new PhoneticProperties() { FontId = (UInt32Value)2U, Type = PhoneticValues.NoConversion };

            sharedStringItem34.Append(text34);
            sharedStringItem34.Append(phoneticProperties34);

            SharedStringItem sharedStringItem35 = new SharedStringItem();
            Text text35 = new Text();
            text35.Text = "备注";
            PhoneticProperties phoneticProperties35 = new PhoneticProperties() { FontId = (UInt32Value)3U, Type = PhoneticValues.NoConversion };

            sharedStringItem35.Append(text35);
            sharedStringItem35.Append(phoneticProperties35);

            SharedStringItem sharedStringItem36 = new SharedStringItem();
            Text text36 = new Text();
            text36.Text = "一";
            PhoneticProperties phoneticProperties36 = new PhoneticProperties() { FontId = (UInt32Value)3U, Type = PhoneticValues.NoConversion };

            sharedStringItem36.Append(text36);
            sharedStringItem36.Append(phoneticProperties36);

            SharedStringItem sharedStringItem37 = new SharedStringItem();
            Text text37 = new Text();
            text37.Text = "西部项目";
            PhoneticProperties phoneticProperties37 = new PhoneticProperties() { FontId = (UInt32Value)2U, Type = PhoneticValues.NoConversion };

            sharedStringItem37.Append(text37);
            sharedStringItem37.Append(phoneticProperties37);

            SharedStringItem sharedStringItem38 = new SharedStringItem();
            Text text38 = new Text();
            text38.Text = "西部中心";
            PhoneticProperties phoneticProperties38 = new PhoneticProperties() { FontId = (UInt32Value)3U, Type = PhoneticValues.NoConversion };

            sharedStringItem38.Append(text38);
            sharedStringItem38.Append(phoneticProperties38);

            SharedStringItem sharedStringItem39 = new SharedStringItem();
            Text text39 = new Text();
            text39.Text = "京珠高速公路湘潭至耒阳段红砂岩地带路基修筑技术研究";

            sharedStringItem39.Append(text39);

            SharedStringItem sharedStringItem40 = new SharedStringItem();
            Text text40 = new Text();
            text40.Text = "湖南省交通规划勘察设计院";
            PhoneticProperties phoneticProperties39 = new PhoneticProperties() { FontId = (UInt32Value)2U, Type = PhoneticValues.NoConversion };

            sharedStringItem40.Append(text40);
            sharedStringItem40.Append(phoneticProperties39);

            SharedStringItem sharedStringItem41 = new SharedStringItem();
            Text text41 = new Text();
            text41.Text = "赵明华";

            sharedStringItem41.Append(text41);

            SharedStringItem sharedStringItem42 = new SharedStringItem();
            Text text42 = new Text();
            text42.Text = "2010.10-2012.10";
            PhoneticProperties phoneticProperties40 = new PhoneticProperties() { FontId = (UInt32Value)2U, Type = PhoneticValues.NoConversion };

            sharedStringItem42.Append(text42);
            sharedStringItem42.Append(phoneticProperties40);

            SharedStringItem sharedStringItem43 = new SharedStringItem();
            Text text43 = new Text();
            text43.Text = "a";
            PhoneticProperties phoneticProperties41 = new PhoneticProperties() { FontId = (UInt32Value)2U, Type = PhoneticValues.NoConversion };

            sharedStringItem43.Append(text43);
            sharedStringItem43.Append(phoneticProperties41);

            SharedStringItem sharedStringItem44 = new SharedStringItem();
            Text text44 = new Text();
            text44.Text = "b";
            PhoneticProperties phoneticProperties42 = new PhoneticProperties() { FontId = (UInt32Value)2U, Type = PhoneticValues.NoConversion };

            sharedStringItem44.Append(text44);
            sharedStringItem44.Append(phoneticProperties42);

            SharedStringItem sharedStringItem45 = new SharedStringItem();
            Text text45 = new Text();
            text45.Text = "c";
            PhoneticProperties phoneticProperties43 = new PhoneticProperties() { FontId = (UInt32Value)2U, Type = PhoneticValues.NoConversion };

            sharedStringItem45.Append(text45);
            sharedStringItem45.Append(phoneticProperties43);

            SharedStringItem sharedStringItem46 = new SharedStringItem();
            Text text46 = new Text();
            text46.Text = "ds";
            PhoneticProperties phoneticProperties44 = new PhoneticProperties() { FontId = (UInt32Value)2U, Type = PhoneticValues.NoConversion };

            sharedStringItem46.Append(text46);
            sharedStringItem46.Append(phoneticProperties44);

            SharedStringItem sharedStringItem47 = new SharedStringItem();
            Text text47 = new Text();
            text47.Text = "dh";
            PhoneticProperties phoneticProperties45 = new PhoneticProperties() { FontId = (UInt32Value)2U, Type = PhoneticValues.NoConversion };

            sharedStringItem47.Append(text47);
            sharedStringItem47.Append(phoneticProperties45);

            SharedStringItem sharedStringItem48 = new SharedStringItem();
            Text text48 = new Text();
            text48.Text = "a";
            PhoneticProperties phoneticProperties46 = new PhoneticProperties() { FontId = (UInt32Value)3U, Type = PhoneticValues.NoConversion };

            sharedStringItem48.Append(text48);
            sharedStringItem48.Append(phoneticProperties46);

            SharedStringItem sharedStringItem49 = new SharedStringItem();
            Text text49 = new Text();
            text49.Text = "1000";
            PhoneticProperties phoneticProperties47 = new PhoneticProperties() { FontId = (UInt32Value)2U, Type = PhoneticValues.NoConversion };

            sharedStringItem49.Append(text49);
            sharedStringItem49.Append(phoneticProperties47);

            SharedStringItem sharedStringItem50 = new SharedStringItem();
            Text text50 = new Text();
            text50.Text = "100";
            PhoneticProperties phoneticProperties48 = new PhoneticProperties() { FontId = (UInt32Value)2U, Type = PhoneticValues.NoConversion };

            sharedStringItem50.Append(text50);
            sharedStringItem50.Append(phoneticProperties48);

            SharedStringItem sharedStringItem51 = new SharedStringItem();
            Text text51 = new Text();
            text51.Text = "800";
            PhoneticProperties phoneticProperties49 = new PhoneticProperties() { FontId = (UInt32Value)2U, Type = PhoneticValues.NoConversion };

            sharedStringItem51.Append(text51);
            sharedStringItem51.Append(phoneticProperties49);

            SharedStringItem sharedStringItem52 = new SharedStringItem();
            Text text52 = new Text();
            text52.Text = "0";
            PhoneticProperties phoneticProperties50 = new PhoneticProperties() { FontId = (UInt32Value)2U, Type = PhoneticValues.NoConversion };

            sharedStringItem52.Append(text52);
            sharedStringItem52.Append(phoneticProperties50);

            SharedStringItem sharedStringItem53 = new SharedStringItem();
            Text text53 = new Text();
            text53.Text = "是";
            PhoneticProperties phoneticProperties51 = new PhoneticProperties() { FontId = (UInt32Value)2U, Type = PhoneticValues.NoConversion };

            sharedStringItem53.Append(text53);
            sharedStringItem53.Append(phoneticProperties51);

            SharedStringItem sharedStringItem54 = new SharedStringItem();
            Text text54 = new Text();
            text54.Text = "已完成";
            PhoneticProperties phoneticProperties52 = new PhoneticProperties() { FontId = (UInt32Value)2U, Type = PhoneticValues.NoConversion };

            sharedStringItem54.Append(text54);
            sharedStringItem54.Append(phoneticProperties52);

            SharedStringItem sharedStringItem55 = new SharedStringItem();
            Text text55 = new Text();
            text55.Text = "已归档";
            PhoneticProperties phoneticProperties53 = new PhoneticProperties() { FontId = (UInt32Value)3U, Type = PhoneticValues.NoConversion };

            sharedStringItem55.Append(text55);
            sharedStringItem55.Append(phoneticProperties53);

            SharedStringItem sharedStringItem56 = new SharedStringItem();
            Text text56 = new Text();
            text56.Text = "已结题";
            PhoneticProperties phoneticProperties54 = new PhoneticProperties() { FontId = (UInt32Value)2U, Type = PhoneticValues.NoConversion };

            sharedStringItem56.Append(text56);
            sharedStringItem56.Append(phoneticProperties54);

            sharedStringTable1.Append(sharedStringItem1);
            sharedStringTable1.Append(sharedStringItem2);
            sharedStringTable1.Append(sharedStringItem3);
            sharedStringTable1.Append(sharedStringItem4);
            sharedStringTable1.Append(sharedStringItem5);
            sharedStringTable1.Append(sharedStringItem6);
            sharedStringTable1.Append(sharedStringItem7);
            sharedStringTable1.Append(sharedStringItem8);
            sharedStringTable1.Append(sharedStringItem9);
            sharedStringTable1.Append(sharedStringItem10);
            sharedStringTable1.Append(sharedStringItem11);
            sharedStringTable1.Append(sharedStringItem12);
            sharedStringTable1.Append(sharedStringItem13);
            sharedStringTable1.Append(sharedStringItem14);
            sharedStringTable1.Append(sharedStringItem15);
            sharedStringTable1.Append(sharedStringItem16);
            sharedStringTable1.Append(sharedStringItem17);
            sharedStringTable1.Append(sharedStringItem18);
            sharedStringTable1.Append(sharedStringItem19);
            sharedStringTable1.Append(sharedStringItem20);
            sharedStringTable1.Append(sharedStringItem21);
            sharedStringTable1.Append(sharedStringItem22);
            sharedStringTable1.Append(sharedStringItem23);
            sharedStringTable1.Append(sharedStringItem24);
            sharedStringTable1.Append(sharedStringItem25);
            sharedStringTable1.Append(sharedStringItem26);
            sharedStringTable1.Append(sharedStringItem27);
            sharedStringTable1.Append(sharedStringItem28);
            sharedStringTable1.Append(sharedStringItem29);
            sharedStringTable1.Append(sharedStringItem30);
            sharedStringTable1.Append(sharedStringItem31);
            sharedStringTable1.Append(sharedStringItem32);
            sharedStringTable1.Append(sharedStringItem33);
            sharedStringTable1.Append(sharedStringItem34);
            sharedStringTable1.Append(sharedStringItem35);
            sharedStringTable1.Append(sharedStringItem36);
            sharedStringTable1.Append(sharedStringItem37);
            sharedStringTable1.Append(sharedStringItem38);
            sharedStringTable1.Append(sharedStringItem39);
            sharedStringTable1.Append(sharedStringItem40);
            sharedStringTable1.Append(sharedStringItem41);
            sharedStringTable1.Append(sharedStringItem42);
            sharedStringTable1.Append(sharedStringItem43);
            sharedStringTable1.Append(sharedStringItem44);
            sharedStringTable1.Append(sharedStringItem45);
            sharedStringTable1.Append(sharedStringItem46);
            sharedStringTable1.Append(sharedStringItem47);
            sharedStringTable1.Append(sharedStringItem48);
            sharedStringTable1.Append(sharedStringItem49);
            sharedStringTable1.Append(sharedStringItem50);
            sharedStringTable1.Append(sharedStringItem51);
            sharedStringTable1.Append(sharedStringItem52);
            sharedStringTable1.Append(sharedStringItem53);
            sharedStringTable1.Append(sharedStringItem54);
            sharedStringTable1.Append(sharedStringItem55);
            sharedStringTable1.Append(sharedStringItem56);

            sharedStringTablePart1.SharedStringTable = sharedStringTable1;
        }

        private void SetPackageProperties(OpenXmlPackage document)
        {
            document.PackageProperties.Creator = "蒋惠林";
            document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2014-07-22T09:26:27Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2014-07-22T23:19:23Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.LastModifiedBy = "蒋惠林";
        }

        #region Binary Data
        private string spreadsheetPrinterSettingsPart1Data = "SABQACAATABhAHMAZQByAEoAZQB0ACAANQAyADAAMAAgAFAAQwBMADYAIABDAGwAYQBzAHMAIABEAHIAAAAAAAEEAwbcAAgEQ78AAgEACQCaCzQIZAABAA8A//8BAAEA//8DAAEAQQA0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAAAAAABAAAAAgAAAEQBAAD/////R0lTNAAAAAAAAAAAAAAAAERJTlUiAJAB7AMcAN55QG4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADQAAAAEAAAAAAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACQAQAAU01USgAAAAAQAIABewAzADgARQA3AEIANwA0ADYALQA0ADYARgBFAC0ANABhADEAZAAtADkANgA3AEQALQA1AEIAQgA2ADEAQwBEAEMARQA3ADQANQB9AAAASW5wdXRCaW4AQXV0b1NlbGVjdABSRVNETEwAVW5pcmVzRExMAFBhcGVyU2l6ZQBMRVRURVIAT3JpZW50YXRpb24AUE9SVFJBSVQATWVkaWFUeXBlAEF1dG8AUmVzb2x1dGlvbgA2MDBEUEkAUGFnZU91dHB1dFF1YWxpdHkATm9ybWFsAENvbG9yTW9kZQBNb25vAERvY3VtZW50TlVwADEAQ29sbGF0ZQBPTgBEdXBsZXgATk9ORQBPdXRwdXRCaW4AQXV0bwBTdGFwbGluZwBOb25lAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAHAAAAFY0RE0BAAAAAAAAAAAAAAAAAAAAAAAAAA==";

        private System.IO.Stream GetBinaryDataStream(string base64String)
        {
            return new System.IO.MemoryStream(System.Convert.FromBase64String(base64String));
        }

        #endregion

        private void GenerateWorksheetPart1Content(WorksheetPart worksheetPart3)
        {
            Worksheet worksheet3 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet3.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet3.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet3.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            SheetDimension sheetDimension1 = new SheetDimension() { Reference = "A1:" + lastCellName };
            

            SheetViews sheetViews1 = new SheetViews();

            SheetView sheetView1 = new SheetView() { TabSelected = true, WorkbookViewId = (UInt32Value)0U };
            Selection selection1 = new Selection() { ActiveCell = "A1", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "A1" } };

            sheetView1.Append(selection1);

            sheetViews1.Append(sheetView1);
            SheetFormatProperties sheetFormatProperties1 = new SheetFormatProperties() { DefaultRowHeight = 13.5D, DyDescent = 0.15D };

            Columns columns1 = new Columns();
            Column column1 = new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)1U, Width = 4.625D, CustomWidth = true }; //序号
            columns1.Append(column1);
            //添加列宽度定义
            int startcol = 2;
            foreach (Fields field in listSelectedFields)
            {
                AddColumnsOther(columns1, field, ref startcol);
            }

            SheetData sheetData1 = new SheetData();

            MergeCells mergeCells1 = new MergeCells() { Count = (UInt32Value)(Convert.ToUInt32(mergeCellsCount)) };
            //MergeCells mergeCells1 = new MergeCells() { Count = (UInt32Value)1U };

            System.Text.ASCIIEncoding asciiEncoding = new System.Text.ASCIIEncoding();
            

            Row row1 = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:" + cols.ToString() }, Height = 21D, DyDescent = 0.3D };

            Cell cell1 = new Cell() { CellReference = "A1", StyleIndex = (UInt32Value)13U, DataType = CellValues.String };
            CellValue cellValue1 = new CellValue();
            cellValue1.Text = ExcelTitle;
            
            cell1.Append(cellValue1);
            row1.Append(cell1);

            //添加第一行空单元格
            startcol = 2;
            foreach (Fields field in listSelectedFields)
            {
                AddTitleCells(asciiEncoding, row1,ref startcol, field);
            }

            sheetData1.Append(row1);
            MergeCell mergeCell = new MergeCell { Reference = "A1:" + lastColName + "1" };
            mergeCells1.Append(mergeCell);

            //添加列标题
            Row row2 = new Row() { RowIndex = (UInt32Value)2U, Spans = new ListValue<StringValue>() { InnerText = "1:" + cols.ToString() }, Height = 24D, DyDescent = 0.15D };

            //序号
            Cell cell35 = new Cell() { CellReference = "A2", StyleIndex = (UInt32Value)1U, DataType = CellValues.String };
            CellValue cellValue2 = new CellValue();
            cellValue2.Text = "序号";

            cell35.Append(cellValue2);
            row2.Append(cell35);

            //添加其他字段列标题
            startcol = 2;
            foreach (Fields field in listSelectedFields)
            {
                AddColHeaders(asciiEncoding, row2, ref startcol, field);
            }

            sheetData1.Append(row2);

            LinkProjectClassify.Sort();
            int projectClassifyNo = 0; //项目类型序号
            int sheetRows = 3; //行号，从3开始
            int classNo = 0;
            
            foreach (int pcid in LinkProjectClassify)
            {
                classNo += 1;
                int classProjects = 0;
                var prc = dataContext.ProjectClassify.Single(p => p.ClassifyId.Equals(pcid));
                projectClassifyName = prc.ProjectClassify1;
                //添加项目类别标题
                projectClassifyNo += 1;
                Row row3 = new Row() { RowIndex = (UInt32Value)Convert.ToUInt32(sheetRows), Spans = new ListValue<StringValue>() { InnerText = "1:" + cols.ToString() }, CustomFormat = true, DyDescent = 0.15D };

                //序号
                DigitToChnText dtt = new DigitToChnText();

                Cell cell69 = new Cell() { CellReference = "A" + sheetRows.ToString(), StyleIndex = (UInt32Value)5U, DataType = CellValues.String };
                CellValue cellValue36 = new CellValue();
                cellValue36.Text = dtt.Convert(projectClassifyNo.ToString(), false);

                cell69.Append(cellValue36);

                row3.Append(cell69);

                //添加项目类别行空白单元格
                startcol = 2;
                foreach (Fields field in listSelectedFields)
                {
                    AddProjectClassifyBlankCells(asciiEncoding, sheetRows, row3, ref startcol, field);
                }
                sheetData1.Append(row3);

                mergeCell = new MergeCell { Reference = "B" + sheetRows.ToString() + ":" + lastColName + sheetRows.ToString() };
                mergeCells1.Append(mergeCell);

                sheetRows += 1;
                //添加项目
                var pbs = dataContext.ProjectBase.Where(p => p.ProjectClassifyID.Equals(pcid)).OrderBy(p => p.ProjectId);
                foreach (var pb in pbs)
                {
                    if (projectToExcelIDListSorted.Contains(pb.ProjectId))
                    {
                        AddProject(sheetData1, asciiEncoding, ref sheetRows, ref classProjects, pb);
                    }
                }
            }
            
            //sheetData3.Append(row3);
            //sheetData3.Append(row4);


            PhoneticProperties phoneticProperties3 = new PhoneticProperties() { FontId = (UInt32Value)2U, Type = PhoneticValues.NoConversion };
            PageMargins pageMargins3 = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            PageSetup pageSetup1 = new PageSetup() { PaperSize = (UInt32Value)8U, Orientation = OrientationValues.Landscape, HorizontalDpi = (UInt32Value)4294967295U, VerticalDpi = (UInt32Value)4294967295U, Id = "rId1" };

            worksheet3.Append(sheetDimension1);
            worksheet3.Append(sheetViews1);
            worksheet3.Append(sheetFormatProperties1);
            worksheet3.Append(columns1);
            worksheet3.Append(sheetData1);
            worksheet3.Append(mergeCells1);
            worksheet3.Append(phoneticProperties3);
            worksheet3.Append(pageMargins3);
            worksheet3.Append(pageSetup1);

            worksheetPart3.Worksheet = worksheet3;
        }

        private void AddColumnsOther(Columns columns1, Fields field, ref int startcol)
        {
            switch (field.fieldValue)
            {
                case "ProjectNo":
                    //项目编号列
                    Column column2 = new Column() { Min = (UInt32Value)Convert.ToUInt32(startcol), Max = (UInt32Value)Convert.ToUInt32(startcol), Width = 9.625D, CustomWidth = true }; //项目编号列
                    columns1.Append(column2);
                    startcol += 1;
                    break;
                case "FirstParty":
                    //甲方列
                    Column column3 = new Column() { Min = (UInt32Value)Convert.ToUInt32(startcol), Max = (UInt32Value)Convert.ToUInt32(startcol), Width = 12.625D, CustomWidth = true };
                    columns1.Append(column3);
                    startcol += 1;
                    break;
                case "SetupYear":
                    //立项时间列
                    Column column4 = new Column() { Min = (UInt32Value)Convert.ToUInt32(startcol), Max = (UInt32Value)Convert.ToUInt32(startcol), Width = 8.125D, CustomWidth = true };
                    columns1.Append(column4);
                    startcol += 1;
                    break;
                case "ProjectName":
                    //项目名称列
                    Column column5 = new Column() { Min = (UInt32Value)Convert.ToUInt32(startcol), Max = (UInt32Value)Convert.ToUInt32(startcol), Width = 50.625D, CustomWidth = true };
                    columns1.Append(column5);
                    startcol += 1;
                    break;
                case "SecondParty":
                    //乙方列
                    Column column6 = new Column() { Min = (UInt32Value)Convert.ToUInt32(startcol), Max = (UInt32Value)Convert.ToUInt32(startcol), Width = 30.625D, CustomWidth = true };
                    columns1.Append(column6);
                    startcol += 1;
                    break;
                case "Principal":
                    //项目负责人列
                    Column column7 = new Column() { Min = (UInt32Value)Convert.ToUInt32(startcol), Max = (UInt32Value)Convert.ToUInt32(startcol), Width = 6.625D, CustomWidth = true };
                    columns1.Append(column7);
                    startcol += 1;
                    break;
                case "ContractPeriod":
                    //合同时限列
                    Column column8 = new Column() { Min = (UInt32Value)Convert.ToUInt32(startcol), Max = (UInt32Value)Convert.ToUInt32(startcol), Width = 15.625D, CustomWidth = true };
                    columns1.Append(column8);
                    startcol += 1;
                    break;
                case "SumMoney":
                    //合同额列
                    Column column9 = new Column() { Min = (UInt32Value)Convert.ToUInt32(startcol), Max = (UInt32Value)Convert.ToUInt32(startcol), Width = 10.625D, CustomWidth = true };
                    columns1.Append(column9);
                    startcol += 1;
                    break;
                case "MoneySourceDetail":
                    //经费来源明细列
                    Column column10 = new Column() { Min = (UInt32Value)Convert.ToUInt32(startcol), Max = (UInt32Value)Convert.ToUInt32(startcol + 4), Width = 10.625D, CustomWidth = true };
                    columns1.Append(column10);
                    startcol += 5;
                    break;
                case "AnchoredDepartment":
                    //挂靠处室列
                    Column column11 = new Column() { Min = (UInt32Value)Convert.ToUInt32(startcol), Max = (UInt32Value)Convert.ToUInt32(startcol), Width = 10.625D, CustomWidth = true };
                    columns1.Append(column11);
                    startcol += 1;
                    break;
                case "Workers":
                    //主要人员列
                    Column column12 = new Column() { Min = (UInt32Value)Convert.ToUInt32(startcol), Max = (UInt32Value)Convert.ToUInt32(startcol), Width = 10.625D, CustomWidth = true };
                    columns1.Append(column12);
                    startcol += 1;
                    break;
                case "TeamDepartment":
                    //协作单位列
                    Column column13 = new Column() { Min = (UInt32Value)Convert.ToUInt32(startcol), Max = (UInt32Value)Convert.ToUInt32(startcol), Width = 12.625D, CustomWidth = true };
                    columns1.Append(column13);
                    startcol += 1;
                    break;
                case "ContractNo":
                    //合同号列
                    Column column14 = new Column() { Min = (UInt32Value)Convert.ToUInt32(startcol), Max = (UInt32Value)Convert.ToUInt32(startcol), Width = 10.625D, CustomWidth = true };
                    columns1.Append(column14);
                    startcol += 1;
                    break;
                case "RateState":
                    //鉴定情况列
                    Column column15 = new Column() { Min = (UInt32Value)Convert.ToUInt32(startcol), Max = (UInt32Value)Convert.ToUInt32(startcol), Width = 12.625D, CustomWidth = true };
                    columns1.Append(column15);
                    startcol += 1;
                    break;
                case "FactFinishDate":
                    //实际完成时间
                    Column column16 = new Column() { Min = (UInt32Value)Convert.ToUInt32(startcol), Max = (UInt32Value)Convert.ToUInt32(startcol), Width = 12.625D, CustomWidth = true };
                    columns1.Append(column16);
                    startcol += 1;
                    break;
                case "RewardState":
                    //获奖情况列
                    Column column17 = new Column() { Min = (UInt32Value)Convert.ToUInt32(startcol), Max = (UInt32Value)Convert.ToUInt32(startcol), Width = 12.625D, CustomWidth = true };
                    columns1.Append(column17);
                    startcol += 1;
                    Column column28 = new Column() { Min = (UInt32Value)Convert.ToUInt32(startcol), Max = (UInt32Value)Convert.ToUInt32(startcol), Width = 20.625D, CustomWidth = true };
                    columns1.Append(column28);
                    startcol += 1;
                    Column column29 = new Column() { Min = (UInt32Value)Convert.ToUInt32(startcol), Max = (UInt32Value)Convert.ToUInt32(startcol + 2), Width = 12.625D, CustomWidth = true };
                    columns1.Append(column29);
                    startcol += 3;
                    Column column27 = new Column() { Min = (UInt32Value)Convert.ToUInt32(startcol), Max = (UInt32Value)Convert.ToUInt32(startcol), Width = 20.625D, CustomWidth = true };
                    columns1.Append(column27);
                    startcol += 1;
                    break;
                case "PatentState":
                    //知识产权列
                    Column column18 = new Column() { Min = (UInt32Value)Convert.ToUInt32(startcol), Max = (UInt32Value)Convert.ToUInt32(startcol), Width = 12.625D, CustomWidth = true };
                    columns1.Append(column18);
                    startcol += 1;
                    break;
                case "CompleteDepartment":
                    //完成单位列
                    Column column19 = new Column() { Min = (UInt32Value)Convert.ToUInt32(startcol), Max = (UInt32Value)Convert.ToUInt32(startcol), Width = 12.625D, CustomWidth = true };
                    columns1.Append(column19);
                    startcol += 1;
                    break;
                case "CompleteWorks":
                    //完成人员列
                    Column column20 = new Column() { Min = (UInt32Value)Convert.ToUInt32(startcol), Max = (UInt32Value)Convert.ToUInt32(startcol), Width = 12.625D, CustomWidth = true };
                    columns1.Append(column20);
                    startcol += 1;
                    break;
                case "MoneyDetail":
                    //经费使用统计列
                    Column column21 = new Column() { Min = (UInt32Value)Convert.ToUInt32(startcol), Max = (UInt32Value)Convert.ToUInt32(startcol + 4), Width = 10.625D, CustomWidth = true };
                    columns1.Append(column21);
                    startcol += 5;
                    break;
                case "MainResearchState":
                    //是否主研列
                    Column column22 = new Column() { Min = (UInt32Value)Convert.ToUInt32(startcol), Max = (UInt32Value)Convert.ToUInt32(startcol), Width = 8.625D, CustomWidth = true };
                    columns1.Append(column22);
                    startcol += 1;
                    break;
                case "FinishState":
                    //完成情况列
                    Column column23 = new Column() { Min = (UInt32Value)Convert.ToUInt32(startcol), Max = (UInt32Value)Convert.ToUInt32(startcol), Width = 8.625D, CustomWidth = true };
                    columns1.Append(column23);
                    startcol += 1;
                    break;
                case "FiledState":
                    //归档情况列
                    Column column24 = new Column() { Min = (UInt32Value)Convert.ToUInt32(startcol), Max = (UInt32Value)Convert.ToUInt32(startcol), Width = 8.625D, CustomWidth = true };
                    columns1.Append(column24);
                    startcol += 1;
                    break;
                case "KnoteState":
                    //结题情况列
                    Column column25 = new Column() { Min = (UInt32Value)Convert.ToUInt32(startcol), Max = (UInt32Value)Convert.ToUInt32(startcol), Width = 8.625D, CustomWidth = true };
                    columns1.Append(column25);
                    startcol += 1;
                    break;
                case "Note":
                    //备注列
                    Column column26 = new Column() { Min = (UInt32Value)Convert.ToUInt32(startcol), Max = (UInt32Value)Convert.ToUInt32(startcol), Width = 10.625D, CustomWidth = true };
                    columns1.Append(column26);
                    startcol += 1;
                    break;
            }

        }

        private void AddProject(SheetData sheetData3, System.Text.ASCIIEncoding asciiEncoding, ref int sheetRows, ref int classProjects, ProjectBase pb)
        {
            int startCol = 2;
            byte[] btNumber = new byte[] { (byte)startCol };
            Row row4 = new Row() { RowIndex = (UInt32Value)Convert.ToUInt32(sheetRows), Spans = new ListValue<StringValue>() { InnerText = "1:" + cols.ToString() }, DyDescent = 0.15D };

            classProjects += 1;
            //序号
            Cell cell103 = new Cell() { CellReference = "A" + sheetRows.ToString(), StyleIndex = (UInt32Value)18U, DataType = CellValues.String };
            CellValue cellValue38 = new CellValue();
            cellValue38.Text = classProjects.ToString();

            cell103.Append(cellValue38);
            row4.Append(cell103);

            foreach (Fields field in listSelectedFields)
            {
                AddProjectOtherCell(asciiEncoding, sheetRows, pb, ref startCol, ref btNumber, row4, field);
            }

            sheetData3.Append(row4);
            sheetRows += 1;
            var pbs = dataContext.ProjectBase.Where(p => p.ParentID.Equals(pb.ProjectId));
            foreach(var p in pbs)
            {
                if(projectToExcelIDListSorted.Contains(p.ProjectId))
                {
                    AddProject(sheetData3, asciiEncoding, ref sheetRows, ref classProjects, p);
                }
            }
        }

        private void AddProjectOtherCell(System.Text.ASCIIEncoding asciiEncoding, int sheetRows, ProjectBase pb, ref int startCol, ref byte[] btNumber, Row row4, Fields field)
        {
            string columnName = "";
            if (startCol <= 26)
            {
                btNumber = new byte[] { (byte)(startCol + 64) };
                columnName = asciiEncoding.GetString(btNumber);
            }
            else
            {
                btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                columnName += asciiEncoding.GetString(btNumber);
                btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                columnName += asciiEncoding.GetString(btNumber);
            }
            switch(field.fieldValue)
            {
                case "ProjectNo":
                    //项目编号
                    Cell cell104 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)18U, DataType = CellValues.String };
                    CellValue cellValue39 = new CellValue();
                    cellValue39.Text = pb.ProjectNo;

                    startCol += 1;
                    cell104.Append(cellValue39);
                    row4.Append(cell104);
                    break;
                case "FirstParty":
                    //甲方
                    
                    Cell cell105 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)18U, DataType = CellValues.String };
                    CellValue cellValue40 = new CellValue();
                    cellValue40.Text = pb.FirstParty;

                    startCol += 1;
                    cell105.Append(cellValue40);
                    row4.Append(cell105);
                    break;
                case "SetupYear":
                    //立项时间
                    
                    Cell cell106 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)18U, DataType = CellValues.String };
                    CellValue cellValue41 = new CellValue();
                    cellValue41.Text = pb.SetupYear;

                    startCol += 1;
                    cell106.Append(cellValue41);
                    row4.Append(cell106);
                    break;
                case "ProjectName":
                    //项目名称
                    
                    Cell cell107 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)19U, DataType = CellValues.String };
                    CellValue cellValue42 = new CellValue();
                    cellValue42.Text = pb.ProjectName;

                    startCol += 1;
                    cell107.Append(cellValue42);
                    row4.Append(cell107);
                    break;
                case "SecondParty":
                    //乙方
                    
                    Cell cell108 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)19U, DataType = CellValues.String };
                    CellValue cellValue43 = new CellValue();
                    cellValue43.Text = pb.SecondParty;

                    startCol += 1;
                    cell108.Append(cellValue43);
                    row4.Append(cell108);
                    break;
                case "Principal":
                    //负责人
                    
                    Cell cell109 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)20U, DataType = CellValues.String };
                    CellValue cellValue44 = new CellValue();
                    cellValue44.Text = pb.Principal;

                    startCol += 1;
                    cell109.Append(cellValue44);
                    row4.Append(cell109);
                    break;
                case "ContractPeriod":
                    //合同时限
                    
                    Cell cell110 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)20U, DataType = CellValues.String };
                    CellValue cellValue45 = new CellValue();
                    if (pb.StartDate != null && pb.PlanFinishDate != null)
                    {
                        string text = ((DateTime)(pb.StartDate)).ToShortDateString().Replace(@"/", ".") + "-" + ((DateTime)(pb.PlanFinishDate)).ToShortDateString().Replace(@"/", ".");;
                        cellValue45.Text = text;
                    }
                    else
                    {
                        cellValue45.Text = "";
                    }

                    startCol += 1;
                    cell110.Append(cellValue45);
                    row4.Append(cell110);
                    break;
                case "SumMoney":
                    //合同额
                    
                    Cell cell111 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)18U, DataType = CellValues.Number };
                    CellValue cellValue46 = new CellValue();
                    cellValue46.Text = pb.SumMoney.ToString();

                    startCol += 1;
                    cell111.Append(cellValue46);
                    row4.Append(cell111);
                    break;
                case "MoneySourceDetail":
                    //经费来源统计
                    
                    Cell cell112 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)20U, DataType = CellValues.Number };
                    CellValue cellValue47 = new CellValue();
                    cellValue47.Text = pb.Ministry.ToString();

                    startCol += 1;
                    cell112.Append(cellValue47);
                    row4.Append(cell112);

                    if (startCol <= 26)
                    {
                        btNumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btNumber);
                    }
                    else
                    {
                        btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                        btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                    }
                    Cell cell113 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)20U, DataType = CellValues.Number };
                    CellValue cellValue48 = new CellValue();
                    cellValue48.Text = pb.Transportation.ToString();

                    startCol += 1;
                    cell113.Append(cellValue48);
                    row4.Append(cell113);

                    if (startCol <= 26)
                    {
                        btNumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btNumber);
                    }
                    else
                    {
                        btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                        btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                    }
                    Cell cell114 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)20U, DataType = CellValues.Number };
                    CellValue cellValue49 = new CellValue();
                    cellValue49.Text = pb.Science.ToString();

                    startCol += 1;
                    cell114.Append(cellValue49);
                    row4.Append(cell114);

                    if (startCol <= 26)
                    {
                        btNumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btNumber);
                    }
                    else
                    {
                        btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                        btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                    }
                    Cell cell115 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)20U, DataType = CellValues.Number };
                    CellValue cellValue50 = new CellValue();
                    cellValue50.Text = pb.SupportEngineering.ToString();

                    startCol += 1;
                    cell115.Append(cellValue50);
                    row4.Append(cell115);

                    if (startCol <= 26)
                    {
                        btNumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btNumber);
                    }
                    else
                    {
                        btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                        btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                    }
                    Cell cell116 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)20U, DataType = CellValues.Number };
                    CellValue cellValue51 = new CellValue();
                    cellValue51.Text = pb.Other.ToString();

                    startCol += 1;
                    cell116.Append(cellValue51);
                    row4.Append(cell116);
                    break;
                case "AnchoredDepartment":
                    //挂靠处室
                    
                    Cell cell117 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)20U, DataType = CellValues.String };
                    CellValue cellValue52 = new CellValue();
                    cellValue52.Text = pb.AnchoredDepartment;

                    startCol += 1;
                    cell117.Append(cellValue52);
                    row4.Append(cell117);
                    break;
                case "Workers":
                    //主要人员
                    var workers = dataContext.TeamWorkers.Where(w => w.ProjectID.Equals(pb.ProjectId));
                    string workerlist = "";
                    foreach (TeamWorkers tw in workers)
                    {
                        workerlist += tw.WorkerName + "、";
                    }
                    if (workerlist.Length > 0)
                    {
                        workerlist = workerlist.Substring(0, workerlist.Length - 1);
                    }
                    
                    Cell cell118 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)20U, DataType = CellValues.String };
                    CellValue cellValue53 = new CellValue();
                    cellValue53.Text = workerlist;

                    startCol += 1;
                    cell118.Append(cellValue53);
                    row4.Append(cell118);
                    break;
                case "TeamDepartment":
                    //协作单位
                    var teams = dataContext.TeamDepartments.Where(t => t.ProjectID.Equals(pb.ProjectId));
                    string teamlist = "";
                    foreach (TeamDepartments t in teams)
                    {
                        teamlist += t.Department + "、";
                    }
                    if (teamlist.Length > 0)
                    {
                        teamlist = teamlist.Substring(0, teamlist.Length - 1);
                    }
                    
                    Cell cell119 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)18U, DataType = CellValues.String };
                    CellValue cellValue54 = new CellValue();
                    cellValue54.Text = teamlist;

                    startCol += 1;
                    cell119.Append(cellValue54);
                    row4.Append(cell119);
                    break;
                case "ContractNo":
                    //合同号
                    
                    Cell cell120 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)6U, DataType = CellValues.String };
                    CellValue cellValue55 = new CellValue();
                    cellValue55.Text = pb.ContractNo;

                    startCol += 1;
                    cell120.Append(cellValue55);
                    row4.Append(cell120);
                    break;
                case "RateState":
                    //鉴定情况
                    string ratelist = "";
                    var rates = dataContext.RateResults.Where(r => r.ProjectID.Equals(pb.ProjectId));
                    foreach (var rate in rates)
                    {
                        if (rate.RateDate != null)
                        {
                            ratelist += ((DateTime)(rate.RateDate)).ToLongDateString();
                        }
                        if (rate.RateDepartment != null)
                        {
                            ratelist += "经" + rate.RateDepartment + "鉴定，鉴定结果：" + rate.RateClassify.RateClassify1 + "，";
                        }
                        else
                        {
                            ratelist += "鉴定结果：" + rate.RateClassify.RateClassify1 + ",";
                        }
                        if(!String.IsNullOrEmpty(rate.Note))
                        {
                            ratelist += "（备注）" + rate.Note + "；";
                        }
                        else
                        {
                            ratelist = ratelist.Substring(0, ratelist.Length - 1) + "；";
                        }
                    }
                    string resultlist = "";
                    var results = dataContext.Results.Where(r => r.ProjectID.Equals(pb.ProjectId));
                    foreach (var r in results)
                    {
                        if (r.RegistDate != null)
                        {
                            resultlist += " 成果登记日期：" + ((DateTime)r.RegistDate).ToLongDateString();
                        }
                        if (r.RegistNo != null)
                        {
                            resultlist += " 成果登记号：" + r.RegistNo + "；";
                        }
                    }
                    if (resultlist.Length > 0)
                    {
                        resultlist = resultlist.Substring(0, resultlist.Length - 1);
                        ratelist += resultlist;
                    }
                    else
                    {
                        if (ratelist.Length > 0)
                        {
                            ratelist = ratelist.Substring(0, ratelist.Length - 1);
                        }
                    }
                    
                    Cell cell121 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)7U, DataType = CellValues.String };
                    CellValue cellValue56 = new CellValue();
                    cellValue56.Text = ratelist;

                    startCol += 1;
                    cell121.Append(cellValue56);
                    row4.Append(cell121);
                    break;
                case "FactFinishDate":
                    //实际完成时间
                    
                    Cell cell122 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)12U, DataType = CellValues.String };
                    CellValue cellValue57 = new CellValue();
                    if(pb.FinishDate != null)
                    {
                        cellValue57.Text = ((DateTime)pb.FinishDate).ToString("yyyy.MM.dd");
                    }
                    else
                    {
                        cellValue57.Text = "";
                    }

                    startCol += 1;
                    cell122.Append(cellValue57);
                    row4.Append(cell122);
                    break;
                case "RewardState":
                    //获奖情况
                    string rewardYear = ""; //获奖年度
                    string rewardName = ""; //奖项名称
                    string rewardDepartment = ""; //授奖单位
                    string rewardClassify = ""; //授奖等别
                    string department = ""; //本单位排名
                    string rewardWorkers = ""; //获奖人员
                    var rewards = dataContext.Reward.Where(r => r.ProjectID.Equals(pb.ProjectId));
                    foreach (var r in rewards)
                    {
                        if (!string.IsNullOrEmpty(r.ReawardYear))
                        {
                            //有获奖年度
                            rewardYear += r.ReawardYear.Trim() + "\r\n";
                        }
                        else
                        {
                            rewardYear += " \r\n";
                        }

                        if (!string.IsNullOrEmpty(r.RewardName))
                        {
                            //有奖项名称
                            rewardName += r.RewardName.Trim() + "\r\n";
                        }
                        else
                        {
                            rewardName += " \r\n";
                        }
                        if (!String.IsNullOrEmpty(r.RewardDepartment))
                        {
                            //颁奖单位
                            rewardDepartment += r.RewardDepartment.Trim() + "\r\n";
                        }
                        else
                        {
                            rewardDepartment += " \r\n";
                        }
                        if (!string.IsNullOrEmpty(r.RewardClassify.RewardClassify1))
                        {
                            //奖励等别
                            rewardClassify += r.RewardClassify.RewardClassify1.Trim() + "\r\n";
                        }
                        else
                        {
                            rewardClassify += " \r\n";
                        }
                        if (!string.IsNullOrEmpty(r.Department))
                        {
                            //本单位排名
                            department += r.Department.Trim() + "\r\n";
                        }
                        else
                        {
                            department += " \r\n";
                        }
                        if (!string.IsNullOrEmpty(r.Workers))
                        {
                            //获奖人员
                            rewardWorkers += r.Workers.Trim() + "\r\n";
                        }
                        else
                        {
                            rewardWorkers += " \r\n";
                        }
                    }
                    if (!string.IsNullOrEmpty(rewardYear))
                    {
                        rewardYear = rewardYear.Substring(0, rewardYear.Length - 1);
                    }
                    if (!string.IsNullOrEmpty(rewardName))
                    {
                        rewardName = rewardName.Substring(0, rewardName.Length - 1);
                    }
                    if (!string.IsNullOrEmpty(rewardDepartment))
                    {
                        rewardDepartment = rewardDepartment.Substring(0, rewardDepartment.Length - 1);
                    }
                    if (!string.IsNullOrEmpty(rewardClassify))
                    {
                        rewardClassify = rewardClassify.Substring(0, rewardClassify.Length - 1);
                    }
                    if (!string.IsNullOrEmpty(department))
                    {
                        department = department.Substring(0, department.Length - 1);
                    }
                    if (!string.IsNullOrEmpty(rewardWorkers))
                    {
                        rewardWorkers = rewardWorkers.Substring(0, rewardWorkers.Length - 1);
                    }
                    
                    //获奖年度
                    Cell cell123 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)12U, DataType = CellValues.String };
                    CellValue cellValue58 = new CellValue();
                    cellValue58.Text = rewardYear;

                    startCol += 1;
                    cell123.Append(cellValue58);
                    row4.Append(cell123);

                    //奖项名称
                    if (startCol <= 26)
                    {
                        btNumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btNumber);
                    }
                    else
                    {
                        btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                        btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                    }
                    Cell cell144 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)12U, DataType = CellValues.String };
                    CellValue cellValue84 = new CellValue();
                    cellValue84.Text = rewardName;

                    startCol += 1;
                    cell144.Append(cellValue84);
                    row4.Append(cell144);

                    //颁奖单位
                    if (startCol <= 26)
                    {
                        btNumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btNumber);
                    }
                    else
                    {
                        btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                        btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                    }
                    Cell cell140 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)12U, DataType = CellValues.String };
                    CellValue cellValue80 = new CellValue();
                    cellValue80.Text = rewardDepartment;

                    startCol += 1;
                    cell140.Append(cellValue80);
                    row4.Append(cell140);

                    //奖励等别
                    if (startCol <= 26)
                    {
                        btNumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btNumber);
                    }
                    else
                    {
                        btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                        btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                    }
                    Cell cell141 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)12U, DataType = CellValues.String };
                    CellValue cellValue81 = new CellValue();
                    cellValue81.Text = rewardClassify;

                    startCol += 1;
                    cell141.Append(cellValue81);
                    row4.Append(cell141);

                    //本单位排名
                    if (startCol <= 26)
                    {
                        btNumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btNumber);
                    }
                    else
                    {
                        btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                        btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                    }
                    Cell cell142 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)12U, DataType = CellValues.String };
                    CellValue cellValue82 = new CellValue();
                    cellValue82.Text = department;

                    startCol += 1;
                    cell142.Append(cellValue82);
                    row4.Append(cell142);

                    //获奖人员
                    if (startCol <= 26)
                    {
                        btNumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btNumber);
                    }
                    else
                    {
                        btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                        btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                    }
                    Cell cell143 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)8U, DataType = CellValues.String };
                    CellValue cellValue83 = new CellValue();
                    cellValue83.Text = rewardWorkers;

                    startCol += 1;
                    cell143.Append(cellValue83);
                    row4.Append(cell143);
                    break;
                case "PatentState":
                    //知识产权
                    string patentlist = "";
                    var patents = dataContext.Patents.Where(p => p.ProjectID.Equals(pb.ProjectId));
                    foreach (var pt in patents)
                    {
                        if (pt.PatentDate != null)
                        {
                            patentlist += ((DateTime)pt.PatentDate).ToLongDateString();
                        }
                        if (pt.PatendDepartment != null)
                        {
                            if (pt.PatentDate != null)
                            {
                                patentlist += "，";
                            }
                            patentlist += pt.PatendDepartment + "颁布";
                        }
                        if (pt.PatentClassifyID != null)
                        {
                            patentlist += pt.PatentClassify.PatentClassify1 + "：";
                        }
                        if (pt.PatentName != null)
                        {
                            if (pt.PatentNo == null)
                            {
                                patentlist += pt.PatentName + "；";
                            }
                            else
                            {
                                patentlist += pt.PatentName + "，编号：" + pt.PatentNo;
                            }
                        }
                        patentlist += "；";
                    }
                    if (patentlist.Length > 0)
                    {
                        patentlist = patentlist.Substring(0, patentlist.Length - 1);
                    }
                    
                    Cell cell124 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)8U, DataType = CellValues.String };
                    CellValue cellValue59 = new CellValue();
                    cellValue59.Text = patentlist;

                    startCol += 1;
                    cell124.Append(cellValue59);
                    row4.Append(cell124);
                    break;
                case "CompleteDepartment":
                    //完成单位
                    string departmentlist = "";
                    if (pb.SecondParty != null)
                    {
                        departmentlist += pb.SecondParty;
                    }
                    var departments = dataContext.TeamDepartments.Where(t => t.ProjectID.Equals(pb.ProjectId));
                    foreach (var d in departments)
                    {
                        if (departmentlist.Length > 0)
                        {
                            departmentlist += "、" + d.Department;
                        }
                        else
                        {
                            departmentlist += d.Department;
                        }
                    }
                    
                    Cell cell125 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)9U, DataType = CellValues.String };
                    CellValue cellValue60 = new CellValue();
                    cellValue60.Text = departmentlist;

                    startCol += 1;
                    cell125.Append(cellValue60);
                    row4.Append(cell125);
                    break;
                case "CompleteWorks":
                    //完成人员
                    string workerlist1 = "";
                    var workers1 = dataContext.TeamWorkers.Where(w => w.ProjectID.Equals(pb.ProjectId));
                    foreach (var w in workers1)
                    {
                        if (workerlist1.Length > 0)
                        {
                            workerlist1 += "、" + w.WorkerName;
                        }
                        else
                        {
                            workerlist1 += w.WorkerName;
                        }
                    }
                    
                    Cell cell126 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)10U, DataType = CellValues.String };
                    CellValue cellValue61 = new CellValue();
                    cellValue61.Text = workerlist1;

                    startCol += 1;
                    cell126.Append(cellValue61);
                    row4.Append(cell126);
                    break;
                case "MoneyDetail":
                    //经费使用统计
                    //到账
                    View_SubTotalFund subtotal = new View_SubTotalFund();
                    string substring = "";
                    var v = dataContext.View_SubTotalFund.Where(s => s.ProjectID.Equals(pb.ProjectId) && s.FundClassify.Equals("到账"));
                    if (v.Count() > 0)
                    {
                        subtotal = dataContext.View_SubTotalFund.Single(s => s.ProjectID.Equals(pb.ProjectId) && s.FundClassify.Equals("到账"));
                        if (subtotal != null)
                        {
                            substring = subtotal.SubTotalMoney.ToString();
                        }
                    }
                    
                    Cell cell127 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)12U, DataType = CellValues.Number };
                    CellValue cellValue62 = new CellValue();
                    cellValue62.Text = substring;

                    startCol += 1;
                    cell127.Append(cellValue62);
                    row4.Append(cell127);

                    //支付外协
                    v = dataContext.View_SubTotalFund.Where(s => s.ProjectID.Equals(pb.ProjectId) && s.FundClassify.Equals("支付外协"));
                    if (v.Count() > 0)
                    {
                        subtotal = dataContext.View_SubTotalFund.Single(s => s.ProjectID.Equals(pb.ProjectId) && s.FundClassify.Equals("支付外协"));
                        if (subtotal != null)
                        {
                            substring = subtotal.SubTotalMoney.ToString();
                        }
                    }
                    else
                    {
                        substring = "";
                    }
                    if (startCol <= 26)
                    {
                        btNumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btNumber);
                    }
                    else
                    {
                        btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                        btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                    }
                    Cell cell128 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)12U, DataType = CellValues.Number };
                    CellValue cellValue63 = new CellValue();
                    cellValue63.Text = substring;

                    startCol += 1;
                    cell128.Append(cellValue63);
                    row4.Append(cell128);

                    //课题组报支
                    v = dataContext.View_SubTotalFund.Where(s => s.ProjectID.Equals(pb.ProjectId) && s.FundClassify.Equals("课题组报支"));
                    if (v.Count() > 0)
                    {
                        subtotal = dataContext.View_SubTotalFund.Single(s => s.ProjectID.Equals(pb.ProjectId) && s.FundClassify.Equals("课题组报支"));
                        if (subtotal != null)
                        {
                            substring = subtotal.SubTotalMoney.ToString();
                        }
                    }
                    else
                    {
                        substring = "";
                    }
                    if (startCol <= 26)
                    {
                        btNumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btNumber);
                    }
                    else
                    {
                        btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                        btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                    }
                    Cell cell129 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)12U, DataType = CellValues.Number };
                    CellValue cellValue64 = new CellValue();
                    cellValue64.Text = substring;

                    startCol += 1;
                    cell129.Append(cellValue64);
                    row4.Append(cell129);

                    //提取管理费
                    v = dataContext.View_SubTotalFund.Where(s => s.ProjectID.Equals(pb.ProjectId) && s.FundClassify.Equals("管理费"));
                    if (v.Count() > 0)
                    {
                        subtotal = dataContext.View_SubTotalFund.Single(s => s.ProjectID.Equals(pb.ProjectId) && s.FundClassify.Equals("管理费"));
                        if (subtotal != null)
                        {
                            substring = subtotal.SubTotalMoney.ToString();
                        }
                    }
                    else
                    {
                        substring = "";
                    }
                    if (startCol <= 26)
                    {
                        btNumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btNumber);
                    }
                    else
                    {
                        btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                        btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                    }
                    Cell cell130 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)12U, DataType = CellValues.Number };
                    CellValue cellValue65 = new CellValue();
                    cellValue65.Text = substring;

                    startCol += 1;
                    cell130.Append(cellValue65);
                    row4.Append(cell130);

                    //经费结余
                    var sub = dataContext.View_SubTotalFund.Where(s => s.ProjectID.Equals(pb.ProjectId));
                    decimal submoney = 0.0M;
                    foreach (var s in sub)
                    {
                        if (s.IncomeOrPay == true)
                        {
                            submoney += (decimal)s.SubTotalMoney;
                        }
                        else
                        {
                            submoney -= (decimal)s.SubTotalMoney;
                        }
                    }
                    substring = submoney.ToString();
                    if (startCol <= 26)
                    {
                        btNumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btNumber);
                    }
                    else
                    {
                        btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                        btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                    }
                    Cell cell131 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)11U, DataType = CellValues.Number };
                    CellValue cellValue66 = new CellValue();
                    cellValue66.Text = substring;

                    startCol += 1;
                    cell131.Append(cellValue66);
                    row4.Append(cell131);
                    break;
                case "MainResearchState":
                    //是否主研
                    
                    Cell cell132 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)12U, DataType = CellValues.String };
                    CellValue cellValue67 = new CellValue();
                    if (pb.IsMainResearch == true)
                    {
                        cellValue67.Text = "主研";
                    }
                    else
                    {
                        cellValue67.Text = "非主研";
                    }

                    startCol += 1;
                    cell132.Append(cellValue67);
                    row4.Append(cell132);
                    break;
                case "FinishState":
                    //完成情况
                    
                    Cell cell133 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)12U, DataType = CellValues.String };
                    CellValue cellValue68 = new CellValue();
                    if (pb.IsKnot == "鉴定" || pb.IsKnot == "验收" || pb.IsKnot == "结清")
                    {
                        cellValue68.Text = "已完成";
                    }
                    else
                    {
                        cellValue68.Text = "未完成";
                    }

                    startCol += 1;
                    cell133.Append(cellValue68);
                    row4.Append(cell133);
                    break;
                case "FiledState":
                    //归档情况
                    
                    Cell cell134 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)12U, DataType = CellValues.String };
                    CellValue cellValue69 = new CellValue();
                    if (pb.IsFiled == true)
                    {
                        cellValue69.Text = "已归档";
                    }
                    else
                    {
                        cellValue69.Text = "未归档";
                    }

                    startCol += 1;
                    cell134.Append(cellValue69);
                    row4.Append(cell134);
                    break;
                case "KnoteState":
                    //结题情况
                    
                    Cell cell135 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)12U, DataType = CellValues.String };
                    CellValue cellValue70 = new CellValue();
                    cellValue70.Text = pb.IsKnot;

                    startCol += 1;
                    cell135.Append(cellValue70);
                    row4.Append(cell135);
                    break;
                case "Note":
                    //备注
                    
                    Cell cell136 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)12U, DataType = CellValues.String };
                    CellValue cellValue71 = new CellValue();
                    cellValue71.Text = pb.Note;

                    cell136.Append(cellValue71);
                    row4.Append(cell136);
                    break;
            }
        }

        /// <summary>
        /// //添加项目类别行空白单元格
        /// </summary>
        /// <param name="asciiEncoding"></param>
        /// <param name="sheetRows"></param>
        /// <param name="row3"></param>
        private void AddProjectClassifyBlankCells(System.Text.ASCIIEncoding asciiEncoding, int sheetRows, Row row3, ref int startCol, Fields field)
        {
            byte[] btNumber;
            string columnName = "";
            if (startCol <= 26)
            {
                btNumber = new byte[] { (byte)(startCol + 64) };
                columnName = asciiEncoding.GetString(btNumber);
            }
            else
            {
                btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                columnName += asciiEncoding.GetString(btNumber);
                btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                columnName += asciiEncoding.GetString(btNumber);
            }
            switch (field.fieldValue)
            {
                case "ProjectNo":
                    //项目编号
                    //如果有同时填入项目分类标题
                    
                    Cell cell1 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)16U, DataType = CellValues.String };
                    CellValue cellValue1 = new CellValue();
                    cellValue1.Text = projectClassifyName;

                    cell1.Append(cellValue1);
                    row3.Append(cell1);
                    startCol += 1;
                    break;
                case "FirstParty":
                    //甲方
                    
                    if (columnName == "B")
                    {
                        Cell cell70 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U, DataType = CellValues.String };
                        CellValue cellValue37 = new CellValue();
                        cellValue37.Text = projectClassifyName;

                        cell70.Append(cellValue37);
                        row3.Append(cell70);
                    }
                    else
                    {
                        Cell cell71 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
                        row3.Append(cell71);
                    }
                    startCol += 1;
                    break;
                case "SetupYear":
                    //立项时间
                    
                    if (columnName == "B")
                    {
                        Cell cell70 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U, DataType = CellValues.String };
                        CellValue cellValue37 = new CellValue();
                        cellValue37.Text = projectClassifyName;

                        cell70.Append(cellValue37);
                        row3.Append(cell70);
                    }
                    else
                    {
                        Cell cell71 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
                        row3.Append(cell71);
                    }
                    startCol += 1;
                    break;
                case "ProjectName":
                    //项目名称
                    
                    if (columnName == "B")
                    {
                        Cell cell70 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U, DataType = CellValues.String };
                        CellValue cellValue37 = new CellValue();
                        cellValue37.Text = projectClassifyName;

                        cell70.Append(cellValue37);
                        row3.Append(cell70);
                    }
                    else
                    {
                        Cell cell71 = new Cell() { CellReference = asciiEncoding.GetString(btNumber) + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
                        row3.Append(cell71);
                    }
                    startCol += 1;
                    break;
                case "SecondParty":
                    //乙方
                    
                    if (columnName == "B")
                    {
                        Cell cell70 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U, DataType = CellValues.String };
                        CellValue cellValue37 = new CellValue();
                        cellValue37.Text = projectClassifyName;

                        cell70.Append(cellValue37);
                        row3.Append(cell70);
                    }
                    else
                    {
                        Cell cell71 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
                        row3.Append(cell71);
                    }
                    startCol += 1;
                    break;
                case "Principal":
                    //负责人
                    
                    if (columnName == "B")
                    {
                        Cell cell70 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U, DataType = CellValues.String };
                        CellValue cellValue37 = new CellValue();
                        cellValue37.Text = projectClassifyName;

                        cell70.Append(cellValue37);
                        row3.Append(cell70);
                    }
                    else
                    {
                        Cell cell71 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
                        row3.Append(cell71);
                    }
                    startCol += 1;
                    break;
                case "ContractPeriod":
                    //合同时限
                    
                    if (columnName == "B")
                    {
                        Cell cell70 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U, DataType = CellValues.String };
                        CellValue cellValue37 = new CellValue();
                        cellValue37.Text = projectClassifyName;

                        cell70.Append(cellValue37);
                        row3.Append(cell70);
                    }
                    else
                    {
                        Cell cell71 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
                        row3.Append(cell71);
                    }
                    startCol += 1;
                    break;
                case "SumMoney":
                    //合同额
                    
                    if (columnName == "B")
                    {
                        Cell cell70 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U, DataType = CellValues.String };
                        CellValue cellValue37 = new CellValue();
                        cellValue37.Text = projectClassifyName;

                        cell70.Append(cellValue37);
                        row3.Append(cell70);
                    }
                    else
                    {
                        Cell cell71 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
                        row3.Append(cell71);
                    }
                    startCol += 1;
                    break;
                case "MoneySourceDetail":
                    //经费来源统计
                    
                    if (columnName == "B")
                    {
                        Cell cell70 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U, DataType = CellValues.String };
                        CellValue cellValue37 = new CellValue();
                        cellValue37.Text = projectClassifyName;

                        cell70.Append(cellValue37);
                        row3.Append(cell70);
                    }
                    else
                    {
                        Cell cell71 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
                        row3.Append(cell71);
                    }
                    startCol += 1;
                    if (startCol <= 26)
                    {
                        btNumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btNumber);
                    }
                    else
                    {
                        btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                        btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                    }
                    if (columnName == "B")
                    {
                        Cell cell70 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U, DataType = CellValues.String };
                        CellValue cellValue37 = new CellValue();
                        cellValue37.Text = projectClassifyName;

                        cell70.Append(cellValue37);
                        row3.Append(cell70);
                    }
                    else
                    {
                        Cell cell71 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
                        row3.Append(cell71);
                    }
                    startCol += 1;
                    if (startCol <= 26)
                    {
                        btNumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btNumber);
                    }
                    else
                    {
                        btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                        btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                    }
                    if (columnName == "B")
                    {
                        Cell cell70 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U, DataType = CellValues.String };
                        CellValue cellValue37 = new CellValue();
                        cellValue37.Text = projectClassifyName;

                        cell70.Append(cellValue37);
                        row3.Append(cell70);
                    }
                    else
                    {
                        Cell cell71 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
                        row3.Append(cell71);
                    }
                    startCol += 1;
                    if (startCol <= 26)
                    {
                        btNumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btNumber);
                    }
                    else
                    {
                        btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                        btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                    }
                    if (columnName == "B")
                    {
                        Cell cell70 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U, DataType = CellValues.String };
                        CellValue cellValue37 = new CellValue();
                        cellValue37.Text = projectClassifyName;

                        cell70.Append(cellValue37);
                        row3.Append(cell70);
                    }
                    else
                    {
                        Cell cell71 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
                        row3.Append(cell71);
                    }
                    startCol += 1;
                    if (startCol <= 26)
                    {
                        btNumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btNumber);
                    }
                    else
                    {
                        btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                        btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                    }
                    if (columnName == "B")
                    {
                        Cell cell70 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U, DataType = CellValues.String };
                        CellValue cellValue37 = new CellValue();
                        cellValue37.Text = projectClassifyName;

                        cell70.Append(cellValue37);
                        row3.Append(cell70);
                    }
                    else
                    {
                        Cell cell71 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
                        row3.Append(cell71);
                    }
                    startCol += 1;
                    break;
                case "AnchoredDepartment":
                    //挂靠处室
                    
                    if (columnName == "B")
                    {
                        Cell cell70 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U, DataType = CellValues.String };
                        CellValue cellValue37 = new CellValue();
                        cellValue37.Text = projectClassifyName;

                        cell70.Append(cellValue37);
                        row3.Append(cell70);
                    }
                    else
                    {
                        Cell cell71 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
                        row3.Append(cell71);
                    }
                    startCol += 1;
                    break;
                case "Workers":
                    //主要人员
                    
                    if (columnName == "B")
                    {
                        Cell cell70 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U, DataType = CellValues.String };
                        CellValue cellValue37 = new CellValue();
                        cellValue37.Text = projectClassifyName;

                        cell70.Append(cellValue37);
                        row3.Append(cell70);
                    }
                    else
                    {
                        Cell cell71 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
                        row3.Append(cell71);
                    }
                    startCol += 1;
                    break;
                case "TeamDepartment":
                    //协作单位
                    
                    if (columnName == "B")
                    {
                        Cell cell70 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U, DataType = CellValues.String };
                        CellValue cellValue37 = new CellValue();
                        cellValue37.Text = projectClassifyName;

                        cell70.Append(cellValue37);
                        row3.Append(cell70);
                    }
                    else
                    {
                        Cell cell71 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
                        row3.Append(cell71);
                    }
                    startCol += 1;
                    break;
                case "ContractNo":
                    //合同号
                    
                    if (columnName == "B")
                    {
                        Cell cell70 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U, DataType = CellValues.String };
                        CellValue cellValue37 = new CellValue();
                        cellValue37.Text = projectClassifyName;

                        cell70.Append(cellValue37);
                        row3.Append(cell70);
                    }
                    else
                    {
                        Cell cell71 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
                        row3.Append(cell71);
                    }
                    startCol += 1;
                    break;
                case "RateState":
                    //鉴定情况
                    
                    if (columnName == "B")
                    {
                        Cell cell70 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U, DataType = CellValues.String };
                        CellValue cellValue37 = new CellValue();
                        cellValue37.Text = projectClassifyName;

                        cell70.Append(cellValue37);
                        row3.Append(cell70);
                    }
                    else
                    {
                        Cell cell71 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
                        row3.Append(cell71);
                    }
                    startCol += 1;
                    break;
                case "FactFinishDate":
                    //成果登记
                    
                    if (columnName == "B")
                    {
                        Cell cell70 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U, DataType = CellValues.String };
                        CellValue cellValue37 = new CellValue();
                        cellValue37.Text = projectClassifyName;

                        cell70.Append(cellValue37);
                        row3.Append(cell70);
                    }
                    else
                    {
                        Cell cell71 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
                        row3.Append(cell71);
                    }
                    startCol += 1;
                    break;
                case "RewardState":
                    //获奖情况
                    
                    if (columnName == "B")
                    {
                        Cell cell70 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U, DataType = CellValues.String };
                        CellValue cellValue37 = new CellValue();
                        cellValue37.Text = projectClassifyName;

                        cell70.Append(cellValue37);
                        row3.Append(cell70);
                    }
                    else
                    {
                        Cell cell71 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
                        row3.Append(cell71);
                    }
                    startCol += 1;
                    if (startCol <= 26)
                    {
                        btNumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btNumber);
                    }
                    else
                    {
                        btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                        btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                    }
                    if (columnName == "B")
                    {
                        Cell cell70 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U, DataType = CellValues.String };
                        CellValue cellValue37 = new CellValue();
                        cellValue37.Text = projectClassifyName;

                        cell70.Append(cellValue37);
                        row3.Append(cell70);
                    }
                    else
                    {
                        Cell cell71 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
                        row3.Append(cell71);
                    }
                    startCol += 1;
                    if (startCol <= 26)
                    {
                        btNumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btNumber);
                    }
                    else
                    {
                        btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                        btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                    }
                    Cell cell72 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
                    row3.Append(cell72);
                    startCol += 1;
                    if (startCol <= 26)
                    {
                        btNumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btNumber);
                    }
                    else
                    {
                        btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                        btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                    }
                    Cell cell73 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
                    row3.Append(cell73);
                    startCol += 1;
                    if (startCol <= 26)
                    {
                        btNumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btNumber);
                    }
                    else
                    {
                        btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                        btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                    }
                    Cell cell74 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
                    row3.Append(cell74);
                    startCol += 1;
                    if (startCol <= 26)
                    {
                        btNumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btNumber);
                    }
                    else
                    {
                        btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                        btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                    }
                    Cell cell75 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
                    row3.Append(cell75);
                    startCol += 1;
                    break;
                case "PatentState":
                    //知识产权
                    
                    if (columnName == "B")
                    {
                        Cell cell70 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U, DataType = CellValues.String };
                        CellValue cellValue37 = new CellValue();
                        cellValue37.Text = projectClassifyName;

                        cell70.Append(cellValue37);
                        row3.Append(cell70);
                    }
                    else
                    {
                        Cell cell71 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
                        row3.Append(cell71);
                    }
                    startCol += 1;
                    break;
                case "CompleteDepartment":
                    //完成单位
                    
                    if (columnName == "B")
                    {
                        Cell cell70 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U, DataType = CellValues.String };
                        CellValue cellValue37 = new CellValue();
                        cellValue37.Text = projectClassifyName;

                        cell70.Append(cellValue37);
                        row3.Append(cell70);
                    }
                    else
                    {
                        Cell cell71 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
                        row3.Append(cell71);
                    }
                    startCol += 1;
                    break;
                case "CompleteWorks":
                    //完成人员
                    
                    if (columnName == "B")
                    {
                        Cell cell70 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U, DataType = CellValues.String };
                        CellValue cellValue37 = new CellValue();
                        cellValue37.Text = projectClassifyName;

                        cell70.Append(cellValue37);
                        row3.Append(cell70);
                    }
                    else
                    {
                        Cell cell71 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
                        row3.Append(cell71);
                    }
                    startCol += 1;
                    break;
                case "MoneyDetail":
                    //经费使用统计
                    
                    if (columnName == "B")
                    {
                        Cell cell70 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U, DataType = CellValues.String };
                        CellValue cellValue37 = new CellValue();
                        cellValue37.Text = projectClassifyName;

                        cell70.Append(cellValue37);
                        row3.Append(cell70);
                    }
                    else
                    {
                        Cell cell71 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
                        row3.Append(cell71);
                    }
                    startCol += 1;
                    if (startCol <= 26)
                    {
                        btNumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btNumber);
                    }
                    else
                    {
                        btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                        btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                    }
                    if (columnName == "B")
                    {
                        Cell cell70 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U, DataType = CellValues.String };
                        CellValue cellValue37 = new CellValue();
                        cellValue37.Text = projectClassifyName;

                        cell70.Append(cellValue37);
                        row3.Append(cell70);
                    }
                    else
                    {
                        Cell cell71 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
                        row3.Append(cell71);
                    }
                    startCol += 1;
                    if (startCol <= 26)
                    {
                        btNumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btNumber);
                    }
                    else
                    {
                        btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                        btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                    }
                    if (columnName == "B")
                    {
                        Cell cell70 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U, DataType = CellValues.String };
                        CellValue cellValue37 = new CellValue();
                        cellValue37.Text = projectClassifyName;

                        cell70.Append(cellValue37);
                        row3.Append(cell70);
                    }
                    else
                    {
                        Cell cell71 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
                        row3.Append(cell71);
                    }
                    startCol += 1;
                    if (startCol <= 26)
                    {
                        btNumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btNumber);
                    }
                    else
                    {
                        btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                        btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                    }
                    if (columnName == "B")
                    {
                        Cell cell70 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U, DataType = CellValues.String };
                        CellValue cellValue37 = new CellValue();
                        cellValue37.Text = projectClassifyName;

                        cell70.Append(cellValue37);
                        row3.Append(cell70);
                    }
                    else
                    {
                        Cell cell71 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
                        row3.Append(cell71);
                    }
                    startCol += 1;
                    if (startCol <= 26)
                    {
                        btNumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btNumber);
                    }
                    else
                    {
                        btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                        btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                    }
                    if (columnName == "B")
                    {
                        Cell cell70 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U, DataType = CellValues.String };
                        CellValue cellValue37 = new CellValue();
                        cellValue37.Text = projectClassifyName;

                        cell70.Append(cellValue37);
                        row3.Append(cell70);
                    }
                    else
                    {
                        Cell cell71 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
                        row3.Append(cell71);
                    }
                    startCol += 1;
                    break;
                case "MainResearchState":
                    //主研情况
                    
                    if (columnName == "B")
                    {
                        Cell cell70 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U, DataType = CellValues.String };
                        CellValue cellValue37 = new CellValue();
                        cellValue37.Text = projectClassifyName;

                        cell70.Append(cellValue37);
                        row3.Append(cell70);
                    }
                    else
                    {
                        Cell cell71 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
                        row3.Append(cell71);
                    }
                    startCol += 1;
                    break;
                case "FinishState":
                    //完成情况
                    
                    if (columnName == "B")
                    {
                        Cell cell70 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U, DataType = CellValues.String };
                        CellValue cellValue37 = new CellValue();
                        cellValue37.Text = projectClassifyName;

                        cell70.Append(cellValue37);
                        row3.Append(cell70);
                    }
                    else
                    {
                        Cell cell71 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
                        row3.Append(cell71);
                    }
                    startCol += 1;
                    break;
                case "FiledState":
                    //归档情况
                    
                    if (columnName == "B")
                    {
                        Cell cell70 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U, DataType = CellValues.String };
                        CellValue cellValue37 = new CellValue();
                        cellValue37.Text = projectClassifyName;

                        cell70.Append(cellValue37);
                        row3.Append(cell70);
                    }
                    else
                    {
                        Cell cell71 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
                        row3.Append(cell71);
                    }
                    startCol += 1;
                    break;
                case "KnoteState":
                    //结题情况

                    if (columnName == "B")
                    {
                        Cell cell70 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U, DataType = CellValues.String };
                        CellValue cellValue37 = new CellValue();
                        cellValue37.Text = projectClassifyName;

                        cell70.Append(cellValue37);
                        row3.Append(cell70);
                    }
                    else
                    {
                        Cell cell71 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
                        row3.Append(cell71);
                    }
                    startCol += 1;
                    break;
                case "Note":
                    //备注
                    
                    if (columnName == "B")
                    {
                        Cell cell70 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U, DataType = CellValues.String };
                        CellValue cellValue37 = new CellValue();
                        cellValue37.Text = projectClassifyName;

                        cell70.Append(cellValue37);
                        row3.Append(cell70);
                    }
                    else
                    {
                        Cell cell71 = new Cell() { CellReference = columnName + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
                        row3.Append(cell71);
                    }
                    startCol += 1;
                    break;
            }
        }

        /// <summary>
        /// 添加表头列
        /// </summary>
        /// <param name="asciiEncoding"></param>
        /// <param name="row2"></param>
        private void AddColHeaders(System.Text.ASCIIEncoding asciiEncoding, Row row2, ref int startCol, Fields field)
        {
            byte[] btnumber;
            string columnName = "";
            if (startCol <= 26)
            {
                btnumber = new byte[] { (byte)(startCol + 64) };
                columnName = asciiEncoding.GetString(btnumber);
            }
            else
            {
                btnumber = new byte[] { (byte)(startCol / 26 + 64) };
                columnName += asciiEncoding.GetString(btnumber);
                btnumber = new byte[] { (byte)(startCol % 26 + 64) };
                columnName += asciiEncoding.GetString(btnumber);
            }
            switch (field.fieldValue)
            {
                case "ProjectNo":
                    //项目编号
                    
                    Cell cell1 = new Cell() { CellReference = columnName + "2", StyleIndex = (UInt32Value)1U, DataType = CellValues.String };
                    CellValue cellValue1 = new CellValue();
                    cellValue1.Text = "项目编号";
                    cell1.Append(cellValue1);

                    startCol += 1;
                    row2.Append(cell1);
                    break;
                case "FirstParty":
                    //甲方
                    
                    Cell cell2 = new Cell() { CellReference = columnName + "2", StyleIndex = (UInt32Value)1U, DataType = CellValues.String };
                    CellValue cellValue2 = new CellValue();
                    cellValue2.Text = "立项部门";
                    cell2.Append(cellValue2);

                    startCol += 1;
                    row2.Append(cell2);
                    break;
                case "SetupYear":
                    //立项时间
                    
                    Cell cell3 = new Cell() { CellReference = columnName + "2", StyleIndex = (UInt32Value)4U, DataType = CellValues.String };
                    CellValue cellValue3 = new CellValue();
                    cellValue3.Text = "立项时间";
                    cell3.Append(cellValue3);

                    startCol += 1;
                    row2.Append(cell3);
                    break;
                case "ProjectName":
                    //项目名称
                    
                    Cell cell4 = new Cell() { CellReference = columnName + "2", StyleIndex = (UInt32Value)1U, DataType = CellValues.String };
                    CellValue cellValue4 = new CellValue();
                    cellValue4.Text = "项目名称";
                    cell4.Append(cellValue4);

                    startCol += 1;
                    row2.Append(cell4);
                    break;
                case "SecondParty":
                    //乙方
                    
                    Cell cell5 = new Cell() { CellReference = columnName + "2", StyleIndex = (UInt32Value)1U, DataType = CellValues.String };
                    CellValue cellValue5 = new CellValue();
                    cellValue5.Text = "乙方";
                    cell5.Append(cellValue5);

                    startCol += 1;
                    row2.Append(cell5);
                    break;
                case "Principal":
                    //负责人
                    
                    Cell cell6 = new Cell() { CellReference = columnName + "2", StyleIndex = (UInt32Value)1U, DataType = CellValues.String };
                    CellValue cellValue6 = new CellValue();
                    cellValue6.Text = "负责人";
                    cell6.Append(cellValue6);

                    startCol += 1;
                    row2.Append(cell6);
                    break;
                case "ContractPeriod":
                    //合同时限
                    
                    Cell cell7 = new Cell() { CellReference = columnName + "2", StyleIndex = (UInt32Value)1U, DataType = CellValues.String };
                    CellValue cellValue7 = new CellValue();
                    cellValue7.Text = "合同时限";
                    cell7.Append(cellValue7);

                    startCol += 1;
                    row2.Append(cell7);
                    break;
                case "SumMoney":
                    //合同额
                    
                    Cell cell8 = new Cell() { CellReference = columnName + "2", StyleIndex = (UInt32Value)2U, DataType = CellValues.String };
                    CellValue cellValue8 = new CellValue();
                    cellValue8.Text = "合同额\n(万元)";
                    cell8.Append(cellValue8);

                    startCol += 1;
                    row2.Append(cell8);
                    break;
                case "MoneySourceDetail":
                    //经费来源统计
                    
                    Cell cell9 = new Cell() { CellReference = columnName + "2", StyleIndex = (UInt32Value)3U, DataType = CellValues.String };
                    CellValue cellValue9 = new CellValue();
                    cellValue9.Text = "其中：\n交通部";
                    cell9.Append(cellValue9);

                    startCol += 1;
                    row2.Append(cell9);

                    if (startCol <= 26)
                    {
                        btnumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btnumber);
                    }
                    else
                    {
                        btnumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btnumber);
                        btnumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btnumber);
                    }
                    Cell cell10 = new Cell() { CellReference = columnName + "2", StyleIndex = (UInt32Value)3U, DataType = CellValues.String };
                    CellValue cellValue10 = new CellValue();
                    cellValue10.Text = "其中：\n交通厅";
                    cell10.Append(cellValue10);

                    startCol += 1;
                    row2.Append(cell10);

                    if (startCol <= 26)
                    {
                        btnumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btnumber);
                    }
                    else
                    {
                        btnumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btnumber);
                        btnumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btnumber);
                    }
                    Cell cell11 = new Cell() { CellReference = columnName + "2", StyleIndex = (UInt32Value)3U, DataType = CellValues.String };
                    CellValue cellValue11 = new CellValue();
                    cellValue11.Text = "其中：\n科技厅";
                    cell11.Append(cellValue11);

                    startCol += 1;
                    row2.Append(cell11);

                    if (startCol <= 26)
                    {
                        btnumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btnumber);
                    }
                    else
                    {
                        btnumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btnumber);
                        btnumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btnumber);
                    }
                    Cell cell12 = new Cell() { CellReference = columnName + "2", StyleIndex = (UInt32Value)3U, DataType = CellValues.String };
                    CellValue cellValue12 = new CellValue();
                    cellValue12.Text = "其中：\n依托工程";
                    cell12.Append(cellValue12);

                    startCol += 1;
                    row2.Append(cell12);

                    if (startCol <= 26)
                    {
                        btnumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btnumber);
                    }
                    else
                    {
                        btnumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btnumber);
                        btnumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btnumber);
                    }
                    Cell cell13 = new Cell() { CellReference = columnName + "2", StyleIndex = (UInt32Value)3U, DataType = CellValues.String };
                    CellValue cellValue13 = new CellValue();
                    cellValue13.Text = "其中：\n其他";
                    cell13.Append(cellValue13);

                    startCol += 1;
                    row2.Append(cell13);
                    break;
                case "AnchoredDepartment":
                    //挂靠处室
                    
                    Cell cell14 = new Cell() { CellReference = columnName + "2", StyleIndex = (UInt32Value)1U, DataType = CellValues.String };
                    CellValue cellValue14 = new CellValue();
                    cellValue14.Text = "挂靠处室";
                    cell14.Append(cellValue14);

                    startCol += 1;
                    row2.Append(cell14);
                    break;
                case "Workers":
                    //主要人员
                    
                    Cell cell15 = new Cell() { CellReference = columnName + "2", StyleIndex = (UInt32Value)1U, DataType = CellValues.String };
                    CellValue cellValue15 = new CellValue();
                    cellValue15.Text = "主要人员";
                    cell15.Append(cellValue15);

                    startCol += 1;
                    row2.Append(cell15);
                    break;
                case "TeamDepartment":
                    //协作单位
                    
                    Cell cell16 = new Cell() { CellReference = columnName + "2", StyleIndex = (UInt32Value)1U, DataType = CellValues.String };
                    CellValue cellValue16 = new CellValue();
                    cellValue16.Text = "协作单位";
                    cell16.Append(cellValue16);

                    startCol += 1;
                    row2.Append(cell16);
                    break;
                case "ContractNo":
                    //合同号
                    
                    Cell cell17 = new Cell() { CellReference = columnName + "2", StyleIndex = (UInt32Value)1U, DataType = CellValues.String };
                    CellValue cellValue17 = new CellValue();
                    cellValue17.Text = "合同号";
                    cell17.Append(cellValue17);

                    startCol += 1;
                    row2.Append(cell17);
                    break;
                case "RateState":
                    //鉴定情况
                    
                    Cell cell18 = new Cell() { CellReference = columnName + "2", StyleIndex = (UInt32Value)1U, DataType = CellValues.String };
                    CellValue cellValue18 = new CellValue();
                    cellValue18.Text = "鉴定情况";
                    cell18.Append(cellValue18);

                    startCol += 1;
                    row2.Append(cell18);
                    break;
                case "FactFinishDate":
                    //实际完成时间
                    
                    Cell cell19 = new Cell() { CellReference = columnName + "2", StyleIndex = (UInt32Value)1U, DataType = CellValues.String };
                    CellValue cellValue19 = new CellValue();
                    cellValue19.Text = "实际完成时间";
                    cell19.Append(cellValue19);

                    startCol += 1;
                    row2.Append(cell19);
                    break;
                case "RewardState":
                    //获奖情况
                    
                    Cell cell20 = new Cell() { CellReference = columnName + "2", StyleIndex = (UInt32Value)1U, DataType = CellValues.String };
                    CellValue cellValue20 = new CellValue();
                    cellValue20.Text = "获奖年度";
                    cell20.Append(cellValue20);

                    startCol += 1;
                    row2.Append(cell20);

                    if (startCol <= 26)
                    {
                        btnumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btnumber);
                    }
                    else
                    {
                        btnumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btnumber);
                        btnumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btnumber);
                    }

                    Cell cell38 = new Cell() { CellReference = columnName + "2", StyleIndex = (UInt32Value)1U, DataType = CellValues.String };
                    CellValue cellValue38 = new CellValue();
                    cellValue38.Text = "奖项名称";
                    cell38.Append(cellValue38);

                    startCol += 1;
                    row2.Append(cell38);

                    if (startCol <= 26)
                    {
                        btnumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btnumber);
                    }
                    else
                    {
                        btnumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btnumber);
                        btnumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btnumber);
                    }

                    Cell cell34 = new Cell() { CellReference = columnName + "2", StyleIndex = (UInt32Value)1U, DataType = CellValues.String };
                    CellValue cellValue34 = new CellValue();
                    cellValue34.Text = "授奖单位";
                    cell34.Append(cellValue34);

                    startCol += 1;
                    row2.Append(cell34);

                    if (startCol <= 26)
                    {
                        btnumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btnumber);
                    }
                    else
                    {
                        btnumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btnumber);
                        btnumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btnumber);
                    }

                    Cell cell35 = new Cell() { CellReference = columnName + "2", StyleIndex = (UInt32Value)1U, DataType = CellValues.String };
                    CellValue cellValue35 = new CellValue();
                    cellValue35.Text = "奖励等别";
                    cell35.Append(cellValue35);

                    startCol += 1;
                    row2.Append(cell35);

                    if (startCol <= 26)
                    {
                        btnumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btnumber);
                    }
                    else
                    {
                        btnumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btnumber);
                        btnumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btnumber);
                    }

                    Cell cell36 = new Cell() { CellReference = columnName + "2", StyleIndex = (UInt32Value)1U, DataType = CellValues.String };
                    CellValue cellValue36 = new CellValue();
                    cellValue36.Text = "本单位排名";
                    cell36.Append(cellValue36);

                    startCol += 1;
                    row2.Append(cell36);

                    if (startCol <= 26)
                    {
                        btnumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btnumber);
                    }
                    else
                    {
                        btnumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btnumber);
                        btnumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btnumber);
                    }

                    Cell cell37 = new Cell() { CellReference = columnName + "2", StyleIndex = (UInt32Value)1U, DataType = CellValues.String };
                    CellValue cellValue37 = new CellValue();
                    cellValue37.Text = "获奖人员";
                    cell37.Append(cellValue37);

                    startCol += 1;
                    row2.Append(cell37);

                    break;
                case "PatentState":
                    //知识产权
                    
                    Cell cell21 = new Cell() { CellReference = columnName + "2", StyleIndex = (UInt32Value)1U, DataType = CellValues.String };
                    CellValue cellValue21 = new CellValue();
                    cellValue21.Text = "知识产权";
                    cell21.Append(cellValue21);

                    startCol += 1;
                    row2.Append(cell21);
                    break;
                case "CompleteDepartment":
                    //完成单位
                    
                    Cell cell22 = new Cell() { CellReference = columnName + "2", StyleIndex = (UInt32Value)1U, DataType = CellValues.String };
                    CellValue cellValue22 = new CellValue();
                    cellValue22.Text = "完成单位";
                    cell22.Append(cellValue22);

                    startCol += 1;
                    row2.Append(cell22);
                    break;
                case "CompleteWorks":
                    //完成人员
                    
                    Cell cell23 = new Cell() { CellReference = columnName + "2", StyleIndex = (UInt32Value)1U, DataType = CellValues.String };
                    CellValue cellValue23 = new CellValue();
                    cellValue23.Text = "完成人员";
                    cell23.Append(cellValue23);

                    startCol += 1;
                    row2.Append(cell23);
                    break;
                case "MoneyDetail":
                    //费用使用统计
                    
                    Cell cell24 = new Cell() { CellReference = columnName + "2", StyleIndex = (UInt32Value)2U, DataType = CellValues.String };
                    CellValue cellValue24 = new CellValue();
                    cellValue24.Text = "到账\n(万元)";
                    cell24.Append(cellValue24);

                    startCol += 1;
                    row2.Append(cell24);

                    if (startCol <= 26)
                    {
                        btnumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btnumber);
                    }
                    else
                    {
                        btnumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btnumber);
                        btnumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btnumber);
                    }
                    Cell cell25 = new Cell() { CellReference = columnName + "2", StyleIndex = (UInt32Value)2U, DataType = CellValues.String };
                    CellValue cellValue25 = new CellValue();
                    cellValue25.Text = "支付外协\n(万元)";
                    cell25.Append(cellValue25);

                    startCol += 1;
                    row2.Append(cell25);

                    if (startCol <= 26)
                    {
                        btnumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btnumber);
                    }
                    else
                    {
                        btnumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btnumber);
                        btnumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btnumber);
                    }
                    Cell cell26 = new Cell() { CellReference = columnName + "2", StyleIndex = (UInt32Value)2U, DataType = CellValues.String };
                    CellValue cellValue26 = new CellValue();
                    cellValue26.Text = "课题组报支\n(万元)";
                    cell26.Append(cellValue26);

                    startCol += 1;
                    row2.Append(cell26);
                    

                    if (startCol <= 26)
                    {
                        btnumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btnumber);
                    }
                    else
                    {
                        btnumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btnumber);
                        btnumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btnumber);
                    }
                    
                    Cell cell27 = new Cell() { CellReference = columnName + "2", StyleIndex = (UInt32Value)2U, DataType = CellValues.String };
                    CellValue cellValue27 = new CellValue();
                    cellValue27.Text = "提取管理费\n(万元)";
                    cell27.Append(cellValue27);

                    startCol += 1;
                    row2.Append(cell27);
                    

                    if (startCol <= 26)
                    {
                        btnumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btnumber);
                    }
                    else
                    {
                        btnumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btnumber);
                        btnumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btnumber);
                    }
                    Cell cell28 = new Cell() { CellReference = columnName + "2", StyleIndex = (UInt32Value)2U, DataType = CellValues.String };
                    CellValue cellValue28 = new CellValue();
                    cellValue28.Text = "经费结余\n(万元)";
                    cell28.Append(cellValue28);

                    startCol += 1;
                    row2.Append(cell28);
                    
                    break;
                case "MainResearchState":
                    //是否主研
                    
                    Cell cell29 = new Cell() { CellReference = columnName + "2", StyleIndex = (UInt32Value)4U, DataType = CellValues.String };
                    CellValue cellValue29 = new CellValue();
                    cellValue29.Text = "是否主研";

                    cell29.Append(cellValue29);
                    row2.Append(cell29);
                    
                    startCol += 1;
                    break;
                case "FinishState":
                    //完成情况
                    
                    Cell cell30 = new Cell() { CellReference = columnName + "2", StyleIndex = (UInt32Value)4U, DataType = CellValues.String };
                    CellValue cellValue30 = new CellValue();
                    cellValue30.Text = "完成情况";

                    cell30.Append(cellValue30);
                    row2.Append(cell30);
                    
                    startCol += 1;
                    break;
                case "FiledState":
                    //归档情况
                    
                    Cell cell31 = new Cell() { CellReference = columnName + "2", StyleIndex = (UInt32Value)4U, DataType = CellValues.String };
                    CellValue cellValue31 = new CellValue();
                    cellValue31.Text = "归档情况";

                    cell31.Append(cellValue31);
                    row2.Append(cell31);
                    
                    startCol += 1;
                    break;
                case "KnoteState":
                    //结题情况
                    
                    Cell cell32 = new Cell() { CellReference = columnName + "2", StyleIndex = (UInt32Value)4U, DataType = CellValues.String };
                    CellValue cellValue32 = new CellValue();
                    cellValue32.Text = "结题情况";

                    cell32.Append(cellValue32);
                    row2.Append(cell32);
                    
                    startCol += 1;
                    break;
                case "Note":
                    //备注
                    
                    Cell cell33 = new Cell() { CellReference = columnName + "2", StyleIndex = (UInt32Value)1U, DataType = CellValues.String };
                    CellValue cellValue33 = new CellValue();
                    cellValue33.Text = "备注";

                    cell33.Append(cellValue33);
                    row2.Append(cell33);

                    startCol += 1;
                    break;
            }
        }

        /// <summary>
        /// 添加标题空白列
        /// </summary>
        /// <param name="asciiEncoding"></param>
        /// <param name="row1"></param>
        /// <param name="startCol"></param>
        /// <param name="field"></param>
        private void AddTitleCells(System.Text.ASCIIEncoding asciiEncoding, Row row1, ref int startCol, Fields field)
        {
            byte[] btNumber;
            string columnName = "";
            if (startCol <= 26)
            {
                btNumber = new byte[] { (byte)(startCol + 64) };
                columnName = asciiEncoding.GetString(btNumber);
            }
            else
            {
                btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                columnName += asciiEncoding.GetString(btNumber);
                btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                columnName += asciiEncoding.GetString(btNumber);
            }
            switch(field.fieldValue)
            {
                case "ProjectNo":
                    //项目编号
                    
                    Cell cell2 = new Cell() { CellReference = columnName + "1", StyleIndex = (UInt32Value)14U };
                    startCol += 1;
                    row1.Append(cell2);
                    break;
                
                case "FirstParty":
                    //甲方

                    Cell cell3 = new Cell() { CellReference = columnName + "1", StyleIndex = (UInt32Value)14U };
                    startCol += 1;
                    row1.Append(cell3);
                    break;
                case "SetupYear":
                    //立项时间
                    btNumber = new byte[] { (byte)(startCol + 64) };
                    Cell cell4 = new Cell() { CellReference = asciiEncoding.GetString(btNumber) + "1", StyleIndex = (UInt32Value)14U };
                    startCol += 1;
                    row1.Append(cell4);
                    break;
                case "ProjectName":
                    //项目名称
                    
                    Cell cell5 = new Cell() { CellReference = columnName + "1", StyleIndex = (UInt32Value)14U };
                    startCol += 1;
                    row1.Append(cell5);
                    break;
                case "SecondParty":
                    //乙方
                    
                    Cell cell6 = new Cell() { CellReference = columnName + "1", StyleIndex = (UInt32Value)14U };
                    startCol += 1;
                    row1.Append(cell6);
                    break;
                case "Principal":
                    //负责人
                    
                    Cell cell7 = new Cell() { CellReference = columnName + "1", StyleIndex = (UInt32Value)14U };
                    startCol += 1;
                    row1.Append(cell7);
                    break;
                case "ContractPeriod":
                    //合同时限
                    
                    Cell cell8 = new Cell() { CellReference = columnName + "1", StyleIndex = (UInt32Value)14U };
                    startCol += 1;
                    row1.Append(cell8);
                    break;
                case "SumMoney":
                    //合同额
                    
                    Cell cell9 = new Cell() { CellReference = columnName + "1", StyleIndex = (UInt32Value)14U };
                    startCol += 1;
                    row1.Append(cell9);
                    break;
                case "MoneySourceDetail":
                    //经费来源统计
                    
                    Cell cell10 = new Cell() { CellReference = columnName + "1", StyleIndex = (UInt32Value)14U };
                    startCol += 1;
                    row1.Append(cell10);

                    if (startCol <= 26)
                    {
                        btNumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btNumber);
                    }
                    else
                    {
                        btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                        btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                    }

                    Cell cell11 = new Cell() { CellReference = columnName + "1", StyleIndex = (UInt32Value)14U };
                    startCol += 1;
                    row1.Append(cell11);

                    if (startCol <= 26)
                    {
                        btNumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btNumber);
                    }
                    else
                    {
                        btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                        btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                    }
                    Cell cell12 = new Cell() { CellReference = columnName + "1", StyleIndex = (UInt32Value)14U };
                    startCol += 1;
                    row1.Append(cell12);

                    if (startCol <= 26)
                    {
                        btNumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btNumber);
                    }
                    else
                    {
                        btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                        btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                    }
                    Cell cell13 = new Cell() { CellReference = columnName + "1", StyleIndex = (UInt32Value)14U };
                    startCol += 1;
                    row1.Append(cell13);

                    if (startCol <= 26)
                    {
                        btNumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btNumber);
                    }
                    else
                    {
                        btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                        btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                    }
                    Cell cell14 = new Cell() { CellReference = columnName + "1", StyleIndex = (UInt32Value)14U };
                    startCol += 1;
                    row1.Append(cell14);
                    break;
                case "AnchoredDepartment":
                    //挂靠处室
                    
                    Cell cell15 = new Cell() { CellReference = columnName + "1", StyleIndex = (UInt32Value)14U };
                    startCol += 1;
                    row1.Append(cell15);
                    break;
                case "Workers":
                    //主要人员
                    
                    Cell cell16 = new Cell() { CellReference = columnName + "1", StyleIndex = (UInt32Value)14U };
                    startCol += 1;
                    row1.Append(cell16);
                    break;
                case "TeamDepartment":
                    //协作单位
                    
                    Cell cell17 = new Cell() { CellReference = columnName + "1", StyleIndex = (UInt32Value)14U };
                    startCol += 1;
                    row1.Append(cell17);
                    break;
                case "ContractNo":
                    //合同号
                    
                    Cell cell18 = new Cell() { CellReference = columnName + "1", StyleIndex = (UInt32Value)15U };
                    startCol += 1;
                    row1.Append(cell18);
                    break;
                case "RateState":
                    //鉴定情况
                    
                    Cell cell19 = new Cell() { CellReference = columnName + "1", StyleIndex = (UInt32Value)15U };
                    startCol += 1;
                    row1.Append(cell19);
                    break;
                case "FactFinishDate":
                    //成果登记
                    
                    Cell cell20 = new Cell() { CellReference = columnName + "1", StyleIndex = (UInt32Value)15U };
                    startCol += 1;
                    row1.Append(cell20);
                    break;
                case "RewardState":
                    //获奖情况
                    
                    Cell cell21 = new Cell() { CellReference = columnName + "1", StyleIndex = (UInt32Value)15U };
                    startCol += 1;
                    row1.Append(cell21);
                    if (startCol <= 26)
                    {
                        btNumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btNumber);
                    }
                    else
                    {
                        btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                        btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                    }
                    Cell cell31 = new Cell() { CellReference = columnName + "1", StyleIndex = (UInt32Value)15U };
                    startCol += 1;
                    row1.Append(cell31);
                    if (startCol <= 26)
                    {
                        btNumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btNumber);
                    }
                    else
                    {
                        btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                        btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                    }
                    Cell cell32 = new Cell() { CellReference = columnName + "1", StyleIndex = (UInt32Value)15U };
                    startCol += 1;
                    row1.Append(cell32);
                    if (startCol <= 26)
                    {
                        btNumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btNumber);
                    }
                    else
                    {
                        btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                        btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                    }
                    Cell cell33 = new Cell() { CellReference = columnName + "1", StyleIndex = (UInt32Value)15U };
                    startCol += 1;
                    row1.Append(cell33);
                    if (startCol <= 26)
                    {
                        btNumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btNumber);
                    }
                    else
                    {
                        btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                        btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                    }
                    Cell cell34 = new Cell() { CellReference = columnName + "1", StyleIndex = (UInt32Value)15U };
                    startCol += 1;
                    row1.Append(cell34);
                    if (startCol <= 26)
                    {
                        btNumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btNumber);
                    }
                    else
                    {
                        btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                        btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                    }
                    Cell cell40 = new Cell() { CellReference = columnName + "1", StyleIndex = (UInt32Value)15U };
                    startCol += 1;
                    row1.Append(cell40);
                    break;
                case "PatentState":
                    //知识产权
                    
                    Cell cell22 = new Cell() { CellReference = columnName + "1", StyleIndex = (UInt32Value)15U };
                    startCol += 1;
                    row1.Append(cell22);
                    break;
                case "CompleteDepartment":
                    //完成单位
                    
                    Cell cell23 = new Cell() { CellReference = columnName + "1", StyleIndex = (UInt32Value)15U };
                    startCol += 1;
                    row1.Append(cell23);
                    break;
                case "CompleteWorks":
                    //完成人员
                    
                    Cell cell24 = new Cell() { CellReference = columnName + "1", StyleIndex = (UInt32Value)15U };
                    startCol += 1;
                    row1.Append(cell24);
                    break;
                case "MoneyDetail":
                    //经费使用统计
                    
                    Cell cell25 = new Cell() { CellReference = columnName + "1", StyleIndex = (UInt32Value)15U };
                    startCol += 1;
                    row1.Append(cell25);
                    if (startCol <= 26)
                    {
                        btNumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btNumber);
                    }
                    else
                    {
                        btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                        btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                    }
                    Cell cell26 = new Cell() { CellReference = columnName + "1", StyleIndex = (UInt32Value)15U };
                    startCol += 1;
                    row1.Append(cell26);

                    if (startCol <= 26)
                    {
                        btNumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btNumber);
                    }
                    else
                    {
                        btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                        btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                    }
                    Cell cell27 = new Cell() { CellReference = columnName + "1", StyleIndex = (UInt32Value)15U };
                    startCol += 1;
                    row1.Append(cell27);
                    
                    if (startCol <= 26)
                    {
                        btNumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btNumber);
                    }
                    else
                    {
                        btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                        btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                    }
                    startCol += 1;
                    Cell cell28 = new Cell() { CellReference = columnName + "1", StyleIndex = (UInt32Value)15U };
                    row1.Append(cell28);
                    
                    if (startCol <= 26)
                    {
                        btNumber = new byte[] { (byte)(startCol + 64) };
                        columnName = asciiEncoding.GetString(btNumber);
                    }
                    else
                    {
                        btNumber = new byte[] { (byte)(startCol / 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                        btNumber = new byte[] { (byte)(startCol % 26 + 64) };
                        columnName += asciiEncoding.GetString(btNumber);
                    }
                    startCol += 1;
                    Cell cell29 = new Cell() { CellReference = columnName + "1", StyleIndex = (UInt32Value)15U };
                    row1.Append(cell29);
                    
                    break;
                case "MainResearchState":
                    //是否主研
                    
                    startCol += 1;
                    Cell cell35 = new Cell() { CellReference = columnName + "1", StyleIndex = (UInt32Value)15U };
                    row1.Append(cell35);
                    
                    break;
                case "FinishState":
                    //完成情况
                    
                    startCol += 1;
                    Cell cell36 = new Cell() { CellReference = columnName + "1", StyleIndex = (UInt32Value)15U };
                    row1.Append(cell36);
                    
                    break;
                case "FiledState":
                    //归档情况
                    
                    startCol += 1;
                    Cell cell37 = new Cell() { CellReference = columnName + "1", StyleIndex = (UInt32Value)15U };
                    row1.Append(cell37);
                    
                    break;
                case "KnoteState":
                    //结题情况
                    
                    startCol += 1;
                    Cell cell38 = new Cell() { CellReference = columnName + "1", StyleIndex = (UInt32Value)15U };
                    row1.Append(cell38);
                    
                    break;
                case "Note":
                    //备注
                    
                    startCol += 1;
                    Cell cell39 = new Cell() { CellReference = columnName + "1", StyleIndex = (UInt32Value)15U };
                    row1.Append(cell39);
                    
                    break;
            }
        }

        // Generates content of spreadsheetPrinterSettingsPart1.
        class Fields
        {
            public string fieldName {get; set;}
            public string fieldValue {get; set;}
        }

        private void ProjectNo_Click(object sender, RoutedEventArgs e)
        {
            setListSelectFields();
        }

        private void FirstParty_Click(object sender, RoutedEventArgs e)
        {
            setListSelectFields();
        }

        private void SetupYear_Click(object sender, RoutedEventArgs e)
        {
            setListSelectFields();
        }

        private void ProjectName_Click(object sender, RoutedEventArgs e)
        {
            setListSelectFields();
        }

        private void SecondParty_Click(object sender, RoutedEventArgs e)
        {
            setListSelectFields();
        }

        private void ContractNo_Click(object sender, RoutedEventArgs e)
        {
            setListSelectFields();
        }

        private void Principal_Click(object sender, RoutedEventArgs e)
        {
            setListSelectFields();
        }

        private void ContractPeriod_Click(object sender, RoutedEventArgs e)
        {
            setListSelectFields();
        }

        private void SumMoney_Click(object sender, RoutedEventArgs e)
        {
            setListSelectFields();
        }

        private void MoneySourceDetail_Click(object sender, RoutedEventArgs e)
        {
            setListSelectFields();
        }

        private void MoneyDetail_Click(object sender, RoutedEventArgs e)
        {
            setListSelectFields();
        }

        private void AnchoredDepartment_Click(object sender, RoutedEventArgs e)
        {
            setListSelectFields();
        }

        private void Workers_Click(object sender, RoutedEventArgs e)
        {
            setListSelectFields();
        }

        private void TeamDepartment_Click(object sender, RoutedEventArgs e)
        {
            setListSelectFields();
        }

        private void CompleteDepartment_Click(object sender, RoutedEventArgs e)
        {
            setListSelectFields();
        }

        private void CompleteWorks_Click(object sender, RoutedEventArgs e)
        {
            setListSelectFields();
        }

        private void FinishState_Click(object sender, RoutedEventArgs e)
        {
            setListSelectFields();
        }

        private void RateState_Click(object sender, RoutedEventArgs e)
        {
            setListSelectFields();
        }

        private void FactFinishDate_Click(object sender, RoutedEventArgs e)
        {
            setListSelectFields();
        }

        private void RewardState_Click(object sender, RoutedEventArgs e)
        {
            setListSelectFields();
        }

        private void PatentState_Click(object sender, RoutedEventArgs e)
        {
            setListSelectFields();
        }

        private void MainResearchState_Click(object sender, RoutedEventArgs e)
        {
            setListSelectFields();
        }

        private void KnoteState_Click(object sender, RoutedEventArgs e)
        {
            setListSelectFields();
        }

        private void FiledState_Click(object sender, RoutedEventArgs e)
        {
            setListSelectFields();
        }

        private void Note_Click(object sender, RoutedEventArgs e)
        {
            setListSelectFields();
        }

        private void buttonToLeft_Click(object sender, RoutedEventArgs e)
        {
            if(ListSourceFields.SelectedItem == null)
            {
                MessageBox.Show("请在左侧列表选择要移动的源字段！", "错误");
                return;
            }
            int selectedIndex = ListSourceFields.SelectedIndex;
            if(selectedIndex == 0)
            {
                return;
            }
            
            Fields field = (Fields)ListSourceFields.SelectedItem;
            listSelectedFields.Remove(field);
            listSelectedFields.Insert(selectedIndex - 1, field);
            ListSourceFields.ItemsSource = null;
            ListSourceFields.Items.Clear();
            ListSourceFields.DisplayMemberPath = "fieldName";
            ListSourceFields.SelectedValuePath = "fieldValue";
            ListSourceFields.ItemsSource = listSelectedFields;
            ListSourceFields.SelectedIndex = selectedIndex - 1;
        }

        private void buttonToRight_Click(object sender, RoutedEventArgs e)
        {
            if (ListSourceFields.SelectedItem == null)
            {
                MessageBox.Show("请在左侧列表选择要移动的源字段！", "错误");
                return;
            }
            int selectedIndex = ListSourceFields.SelectedIndex;
            if (selectedIndex == ListSourceFields.Items.Count - 1)
            {
                return;
            }
            Fields field = (Fields)ListSourceFields.SelectedItem;
            listSelectedFields.Remove(field);
            listSelectedFields.Insert(selectedIndex + 1, field);
            ListSourceFields.ItemsSource = null;
            ListSourceFields.Items.Clear();
            ListSourceFields.DisplayMemberPath = "fieldName";
            ListSourceFields.SelectedValuePath = "fieldValue";
            ListSourceFields.ItemsSource = listSelectedFields;
            ListSourceFields.SelectedIndex = selectedIndex + 1;
        }
    }
}