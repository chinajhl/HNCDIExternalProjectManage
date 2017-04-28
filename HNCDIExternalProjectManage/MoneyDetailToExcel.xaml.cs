using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.IO;
using System.Linq;
using System.Windows;
using A = DocumentFormat.OpenXml.Drawing;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace HNCDIExternalProjectManage
{
    /// <summary>
    /// MoneyDetailToExcel.xaml 的交互逻辑
    /// </summary>
    public partial class MoneyDetailToExcel : Window
    {
        private DataClassesProjectClassifyDataContext dataContext;
        private int projectID;
        private FileInfo fileToCreate; //要创建的文件
        private int rows; //表格行数
        private string lastCellName; //最后单元格名字
        private int mergeCellCount; //合并单元格数目

        public int ProjectID
        {
            get { return projectID; }
            set { projectID = value; }
        }

        public MoneyDetailToExcel()
        {
            this.InitializeComponent();

            // 在此点之下插入创建对象所需的代码。
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            dataContext = new DataClassesProjectClassifyDataContext();
            var pb = dataContext.ProjectBase.Single(p => p.ProjectId.Equals(projectID));
            getRowsAndLastCellNameAndMergeCellCount();
            try
            {
                Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog();
                saveFileDialog.Filter = "Execl files (*.xlsx)|*.xlsx";
                saveFileDialog.FilterIndex = 0;
                saveFileDialog.RestoreDirectory = true;
                saveFileDialog.FileName = pb.ProjectName + "-经费明细表";
                saveFileDialog.CreatePrompt = true;
                saveFileDialog.Title = "要创建Excel文件";
                saveFileDialog.ShowDialog();
                if (saveFileDialog.FileName == "")
                {
                    MessageBox.Show("错误", "请选择文件或输入文件名", MessageBoxButton.OK);
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
                CreatePackage(saveFileDialog.FileName);
                MessageBox.Show("导出成功！");
                this.Close();
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message, "导出失败 ", MessageBoxButton.OK);
                this.Close();
            }
        }

        /// <summary>
        /// 获取表格行数、最后单元格名字、合并单元格个数
        /// </summary>
        private void getRowsAndLastCellNameAndMergeCellCount()
        {
            rows = 3; //表头行
            mergeCellCount = 3;
            dataContext = new DataClassesProjectClassifyDataContext();
            var fundClassifyCount = dataContext.Funds.Select(fc => new { fc.ProjectID, fc.FundClassifyID }).Where(fc => fc.ProjectID.Equals(projectID)).GroupBy(fc => fc.ProjectID, fc => fc.FundClassifyID);
            rows += fundClassifyCount.Count() * 3; //经费类别、题头及小计行
            mergeCellCount += fundClassifyCount.Count();
            var funds = dataContext.Funds.Where(f => f.ProjectID.Equals(projectID));
            rows += funds.Count(); //经费行
            //收入总览
            var projectFundSumIncoming = dataContext.View_ProjectFundSumIncoming.Where(pf => pf.ProjectID.Equals(projectID) && pf.IncomeOrPay.Equals(true));
            if (projectFundSumIncoming.Count() > 0)
            {
                rows += projectFundSumIncoming.Count() + 3;
                mergeCellCount += projectFundSumIncoming.Count() * 2 + 5;
            }
            //支出总览
            var projectFundPay = dataContext.View_ProjectFundSumIncoming.Where(pf => pf.ProjectID.Equals(projectID) && pf.IncomeOrPay.Equals(false));
            if (projectFundPay.Count() > 0)
            {
                rows += projectFundPay.Count() + 3;
                mergeCellCount += projectFundPay.Count() * 2 + 5;
            }
            rows += 1; //合计行
            mergeCellCount += 1;
            lastCellName = "F" + rows.ToString();
        }

        // Creates a SpreadsheetDocument.
        public void CreatePackage(string filePath)
        {
            using (SpreadsheetDocument package = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                CreateParts(package);
            }
        }

        // Generates content of worksheetPart3.
        private void GenerateWorksheetPart1Content(WorksheetPart worksheetPart3)
        {
            Worksheet worksheet3 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet3.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet3.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet3.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            SheetDimension sheetDimension3 = new SheetDimension() { Reference = "A1:" + lastCellName };

            SheetViews sheetViews3 = new SheetViews();

            SheetView sheetView3 = new SheetView() { TabSelected = true, WorkbookViewId = (UInt32Value)0U };
            Selection selection1 = new Selection() { SequenceOfReferences = new ListValue<StringValue>() { InnerText = "A1:F1" } };

            sheetView3.Append(selection1);

            sheetViews3.Append(sheetView3);
            SheetFormatProperties sheetFormatProperties3 = new SheetFormatProperties() { DefaultRowHeight = 13.5D, DyDescent = 0.15D };

            Columns columns1 = new Columns();
            Column column1 = new Column() { Min = (UInt32Value)2U, Max = (UInt32Value)2U, Width = 29.75D, CustomWidth = true };
            Column column2 = new Column() { Min = (UInt32Value)3U, Max = (UInt32Value)3U, Width = 12.375D, CustomWidth = true };
            Column column3 = new Column() { Min = (UInt32Value)4U, Max = (UInt32Value)4U, Width = 15D, CustomWidth = true };
            Column column4 = new Column() { Min = (UInt32Value)5U, Max = (UInt32Value)5U, Width = 13D, CustomWidth = true };
            Column column5 = new Column() { Min = (UInt32Value)6U, Max = (UInt32Value)6U, Width = 15D, CustomWidth = true };

            columns1.Append(column1);
            columns1.Append(column2);
            columns1.Append(column3);
            columns1.Append(column4);
            columns1.Append(column5);

            SheetData sheetData1 = new SheetData();
            MergeCells mergeCells1 = new MergeCells() { Count = (UInt32Value)Convert.ToUInt32(mergeCellCount) }; //合并单元格定义
            Row row1 = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:6" }, Height = 22.5D, DyDescent = 0.15D };

            Cell cell1 = new Cell() { CellReference = "A1", StyleIndex = (UInt32Value)27U, DataType = CellValues.SharedString };
            CellValue cellValue1 = new CellValue();
            cellValue1.Text = "0";

            cell1.Append(cellValue1);
            Cell cell2 = new Cell() { CellReference = "B1", StyleIndex = (UInt32Value)28U };
            Cell cell3 = new Cell() { CellReference = "C1", StyleIndex = (UInt32Value)28U };
            Cell cell4 = new Cell() { CellReference = "D1", StyleIndex = (UInt32Value)28U };
            Cell cell5 = new Cell() { CellReference = "E1", StyleIndex = (UInt32Value)28U };
            Cell cell6 = new Cell() { CellReference = "F1", StyleIndex = (UInt32Value)28U };

            row1.Append(cell1);
            row1.Append(cell2);
            row1.Append(cell3);
            row1.Append(cell4);
            row1.Append(cell5);
            row1.Append(cell6);
            sheetData1.Append(row1);

            MergeCell mergeCell1 = new MergeCell() { Reference = "A1:F1" };
            mergeCells1.Append(mergeCell1);

            //项目名称行
            Row row2 = new Row() { RowIndex = (UInt32Value)2U, Spans = new ListValue<StringValue>() { InnerText = "1:6" }, Height = 24.95D, ThickBot = true, CustomHeight = true, DyDescent = 0.2D };

            Cell cell7 = new Cell() { CellReference = "A2", StyleIndex = (UInt32Value)10U, DataType = CellValues.SharedString };
            CellValue cellValue2 = new CellValue();
            cellValue2.Text = "1";

            cell7.Append(cellValue2);

            //项目名称
            var projectBase = dataContext.ProjectBase.Single(p => p.ProjectId.Equals(projectID));
            Cell cell8 = new Cell() { CellReference = "B2", StyleIndex = (UInt32Value)29U, DataType = CellValues.String };
            CellValue cellValue3 = new CellValue();
            cellValue3.Text = projectBase.ProjectName;

            cell8.Append(cellValue3);
            Cell cell9 = new Cell() { CellReference = "C2", StyleIndex = (UInt32Value)29U };
            Cell cell10 = new Cell() { CellReference = "D2", StyleIndex = (UInt32Value)29U };
            Cell cell11 = new Cell() { CellReference = "E2", StyleIndex = (UInt32Value)29U };
            Cell cell12 = new Cell() { CellReference = "F2", StyleIndex = (UInt32Value)10U, DataType = CellValues.String };
            CellValue cellValueProjectNo = new CellValue();
            cellValueProjectNo.Text = "No." + projectBase.ProjectNo;
            cell12.Append(cellValueProjectNo);

            row2.Append(cell7);
            row2.Append(cell8);
            row2.Append(cell9);
            row2.Append(cell10);
            row2.Append(cell11);
            row2.Append(cell12);
            sheetData1.Append(row2);
            MergeCell mergeCell2 = new MergeCell() { Reference = "B2:E2" };
            mergeCells1.Append(mergeCell2);

            //合同总额
            Row row = new Row() { RowIndex = (UInt32Value)3U, Spans = new ListValue<StringValue>() { InnerText = "1:6" }, Height = 24.95D, CustomHeight = true, ThickBot = true, DyDescent = 0.2D };

            Cell cell100 = new Cell() { CellReference = "A3", StyleIndex = (UInt32Value)41U, DataType = CellValues.SharedString };
            CellValue cellValue100 = new CellValue();
            cellValue100.Text = "14";

            cell100.Append(cellValue100);
            row.Append(cell100);

            //合同总额
            Cell cell101 = new Cell() { CellReference = "B3", StyleIndex = (UInt32Value)42U, DataType = CellValues.String };
            CellValue cellValue101 = new CellValue();
            cellValue101.Text = projectBase.SumMoney.ToString() + "万元";
            cell101.Append(cellValue101);
            row.Append(cell101);

            Cell cell102 = new Cell() { CellReference = "C3", StyleIndex = (UInt32Value)43U };
            Cell cell103 = new Cell() { CellReference = "D3", StyleIndex = (UInt32Value)43U };
            Cell cell104 = new Cell() { CellReference = "E3", StyleIndex = (UInt32Value)43U };
            Cell cell105 = new Cell() { CellReference = "F3", StyleIndex = (UInt32Value)44U };

            row.Append(cell102);
            row.Append(cell103);
            row.Append(cell104);
            row.Append(cell105);

            sheetData1.Append(row);
            MergeCell mergeCell = new MergeCell() { Reference = "B3:F3" };
            mergeCells1.Append(mergeCell);

            var fundClassifies = dataContext.View_ProjectFundClassifies.Where(f => f.ProjectID.Equals(projectID));
            int sheetRowNo = 3;
            int fundClassifyNo = 0; //经费类型序号
            decimal incoming = 0; //收入合计
            decimal outcoming = 0; //支出合计
            foreach (var fc in fundClassifies)
            {
                sheetRowNo += 1;
                fundClassifyNo += 1;
                Row row4 = new Row() { RowIndex = (UInt32Value)Convert.ToUInt32(sheetRowNo), Spans = new ListValue<StringValue>() { InnerText = "1:6" }, Height = 20.0D, CustomHeight = true, DyDescent = 0.15D };

                //序号
                DigitToChnText dtt = new DigitToChnText();

                Cell cell19 = new Cell() { CellReference = "A" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)38U, DataType = CellValues.String };
                CellValue cellValue10 = new CellValue();
                cellValue10.Text = dtt.Convert(fundClassifyNo.ToString(), false);

                cell19.Append(cellValue10);

                //经费类型
                var ff = dataContext.FundClassify.Single(f => f.FandClassifyId.Equals(fc.FundClassifyID));
                Cell cell20 = new Cell() { CellReference = "B" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)39U, DataType = CellValues.String };
                CellValue cellValue11 = new CellValue();
                if (ff.IncomeOrPay == true)
                {
                    cellValue11.Text = ff.FundClassify1 + "(收入)";
                }
                else
                {
                    cellValue11.Text = ff.FundClassify1 + "(支出)";
                }

                cell20.Append(cellValue11);
                Cell cell21 = new Cell() { CellReference = "C" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)39U };
                Cell cell22 = new Cell() { CellReference = "D" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)39U };
                Cell cell23 = new Cell() { CellReference = "E" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)39U };
                Cell cell24 = new Cell() { CellReference = "F" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)40U };

                row4.Append(cell19);
                row4.Append(cell20);
                row4.Append(cell21);
                row4.Append(cell22);
                row4.Append(cell23);
                row4.Append(cell24);
                sheetData1.Append(row4);
                MergeCell mergeCell3 = new MergeCell() { Reference = "B" + sheetRowNo.ToString() + ":F" + sheetRowNo.ToString() };
                mergeCells1.Append(mergeCell3);

                //加入各类费用表头
                sheetRowNo += 1;
                Row row20 = new Row() { RowIndex = (UInt32Value)Convert.ToUInt32(sheetRowNo), Spans = new ListValue<StringValue>() { InnerText = "1:6" }, Height = 20.0D, CustomHeight = true, DyDescent = 0.15D };

                //序号
                Cell cell106 = new Cell() { CellReference = "A" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
                CellValue cellValue106 = new CellValue();
                cellValue106.Text = "2";
                cell106.Append(cellValue106);

                row20.Append(cell106);

                //甲方/乙方
                Cell cell107 = new Cell() { CellReference = "B" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)2U, DataType = CellValues.String };
                CellValue cellValue107 = new CellValue();
                if (ff.IncomeOrPay == true)
                {
                    cellValue107.Text = "甲方";
                }
                else
                {
                    cellValue107.Text = "乙方";
                }
                cell107.Append(cellValue107);
                row20.Append(cell107);

                //金额
                Cell cell108 = new Cell() { CellReference = "C" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
                CellValue cellValue108 = new CellValue();
                cellValue108.Text = "3";
                cell108.Append(cellValue108);
                row20.Append(cell108);

                //日期
                Cell cell109 = new Cell() { CellReference = "D" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
                CellValue cellValue109 = new CellValue();
                cellValue109.Text = "4";
                cell109.Append(cellValue109);
                row20.Append(cell109);

                //经手人
                Cell cell110 = new Cell() { CellReference = "E" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
                CellValue cellValue110 = new CellValue();
                cellValue110.Text = "5";
                cell110.Append(cellValue110);
                row20.Append(cell110);

                //经手项目负责人
                Cell cell111 = new Cell() { CellReference = "F" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)3U, DataType = CellValues.SharedString };
                CellValue cellValue111 = new CellValue();
                cellValue111.Text = "6";
                cell111.Append(cellValue111);
                row20.Append(cell111);
                sheetData1.Append(row20);

                //该类经费明细
                var pfcs = dataContext.Funds.Where(pf => pf.ProjectID.Equals(projectID) && pf.FundClassifyID.Equals(ff.FandClassifyId)).OrderBy(pf => pf.Date);
                int fundRows = 0;
                decimal subMoney = 0;
                foreach (Funds fund in pfcs)
                {
                    //经费行
                    fundRows += 1;
                    sheetRowNo += 1;
                    Row row5 = new Row() { RowIndex = (UInt32Value)Convert.ToUInt32(sheetRowNo), Spans = new ListValue<StringValue>() { InnerText = "1:6" }, Height = 20.0D, CustomHeight = true, DyDescent = 0.15D };

                    //序号
                    Cell cell25 = new Cell() { CellReference = "A" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)4U, DataType = CellValues.String };
                    CellValue cellValue12 = new CellValue();
                    cellValue12.Text = fundRows.ToString();

                    cell25.Append(cellValue12);

                    //来源
                    Cell cell26 = new Cell() { CellReference = "B" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)6U, DataType = CellValues.String };
                    CellValue cellValue13 = new CellValue();
                    cellValue13.Text = fund.Source.Trim();

                    cell26.Append(cellValue13);

                    //金额
                    Cell cell27 = new Cell() { CellReference = "C" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)7U, DataType = CellValues.Number };
                    CellValue cellValue14 = new CellValue();
                    cellValue14.Text = fund.Money.ToString();
                    subMoney += (decimal)fund.Money;
                    if (fund.FundClassify.IncomeOrPay == true)
                    {
                        incoming += (decimal)fund.Money;
                    }
                    else
                    {
                        outcoming += (decimal)fund.Money;
                    }

                    cell27.Append(cellValue14);

                    //日期
                    Cell cell28 = new Cell() { CellReference = "D" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)7U, DataType = CellValues.String };
                    CellValue cellValue15 = new CellValue();
                    cellValue15.Text = ((DateTime)(fund.Date)).ToLongDateString();

                    cell28.Append(cellValue15);

                    //经手人
                    Cell cell29 = new Cell() { CellReference = "E" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)7U, DataType = CellValues.String };
                    CellValue cellValue16 = new CellValue();
                    cellValue16.Text = fund.Handled.Trim();

                    cell29.Append(cellValue16);

                    //经手负责人
                    Cell cell30 = new Cell() { CellReference = "F" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)8U, DataType = CellValues.String };
                    CellValue cellValue17 = new CellValue();
                    cellValue17.Text = fund.SubPrincipal.Trim();

                    cell30.Append(cellValue17);

                    row5.Append(cell25);
                    row5.Append(cell26);
                    row5.Append(cell27);
                    row5.Append(cell28);
                    row5.Append(cell29);
                    row5.Append(cell30);
                    sheetData1.Append(row5);
                }

                //小计行
                sheetRowNo += 1;
                Row row6 = new Row() { RowIndex = (UInt32Value)Convert.ToUInt32(sheetRowNo), Spans = new ListValue<StringValue>() { InnerText = "1:6" }, Height = 20.100000000000001D, CustomHeight = true, DyDescent = 0.15D };
                Cell cell31 = new Cell() { CellReference = "A" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)4U };

                Cell cell32 = new Cell() { CellReference = "B" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)21U, DataType = CellValues.String };
                CellValue cellValue18 = new CellValue();
                cellValue18.Text = "小计";

                cell32.Append(cellValue18);

                //小计金额
                Cell cell33 = new Cell() { CellReference = "C" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)7U, DataType = CellValues.Number };
                CellValue cellValue19 = new CellValue();
                cellValue19.Text = subMoney.ToString();

                cell33.Append(cellValue19);
                Cell cell34 = new Cell() { CellReference = "D" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)7U };
                Cell cell35 = new Cell() { CellReference = "E" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)7U };
                Cell cell36 = new Cell() { CellReference = "F" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)8U };

                row6.Append(cell31);
                row6.Append(cell32);
                row6.Append(cell33);
                row6.Append(cell34);
                row6.Append(cell35);
                row6.Append(cell36);
                sheetData1.Append(row6);
            }

            var sumMoney = dataContext.View_ProjectFundSumIncoming.Where(s => s.IncomeOrPay.Equals(true) && s.ProjectID.Equals(projectID));
            if (sumMoney.Count() > 0)
            {
                //收入总览
                fundClassifyNo += 1; //序号
                sheetRowNo += 1;
                DigitToChnText dt = new DigitToChnText();
                Row row21 = new Row() { RowIndex = (UInt32Value)Convert.ToUInt32(sheetRowNo), Spans = new ListValue<StringValue>() { InnerText = "1:6" }, Height = 20.0D, CustomHeight = true, ThickTop = true, DyDescent = 0.15D };
                //中文序号
                Cell cell112 = new Cell() { CellReference = "A" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)18U, DataType = CellValues.String };
                CellValue cellValue112 = new CellValue();
                cellValue112.Text = dt.Convert(fundClassifyNo.ToString(), false);
                cell112.Append(cellValue112);
                row21.Append(cell112);

                Cell cell113 = new Cell() { CellReference = "B" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)31U, DataType = CellValues.String };
                CellValue cellValue113 = new CellValue();
                cellValue113.Text = "收入总览";
                cell113.Append(cellValue113);
                row21.Append(cell113);

                Cell cell114 = new Cell() { CellReference = "C" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)31U };
                Cell cell115 = new Cell() { CellReference = "D" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)31U };
                Cell cell116 = new Cell() { CellReference = "E" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)31U };
                Cell cell117 = new Cell() { CellReference = "F" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)32U };
                row21.Append(cell114);
                row21.Append(cell115);
                row21.Append(cell116);
                row21.Append(cell117);

                sheetData1.Append(row21);

                MergeCell mergeCell4 = new MergeCell() { Reference = "B" + sheetRowNo.ToString() + ":F" + sheetRowNo.ToString() };
                mergeCells1.Append(mergeCell4);

                sheetRowNo += 1;
                //题头行
                Row row9 = new Row() { RowIndex = (UInt32Value)Convert.ToUInt32(sheetRowNo), Spans = new ListValue<StringValue>() { InnerText = "1:6" }, Height = 20.0D, CustomHeight = true, DyDescent = 0.15D };
                Cell cell49 = new Cell() { CellReference = "A" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)4U };
                Cell cell50 = new Cell() { CellReference = "B" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)5U, DataType = CellValues.String };
                CellValue cellValue50 = new CellValue();
                cellValue50.Text = "经费类型";
                cell50.Append(cellValue50);

                Cell cell51 = new Cell() { CellReference = "C" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)30U };

                Cell cell52 = new Cell() { CellReference = "D" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)33U };
                Cell cell53 = new Cell() { CellReference = "E" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)30U, DataType = CellValues.String };
                CellValue cellValue53 = new CellValue();
                cellValue53.Text = "金额";
                cell53.Append(cellValue53);

                Cell cell54 = new Cell() { CellReference = "F" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)34U };
                row9.Append(cell49);
                row9.Append(cell50);
                row9.Append(cell51);
                row9.Append(cell52);
                row9.Append(cell53);
                row9.Append(cell54);

                sheetData1.Append(row9);

                MergeCell mergeCell5 = new MergeCell() { Reference = "B" + sheetRowNo.ToString() + ":D" + sheetRowNo.ToString() };
                MergeCell mergeCell6 = new MergeCell() { Reference = "E" + sheetRowNo.ToString() + ":F" + sheetRowNo.ToString() };
                mergeCells1.Append(mergeCell5);
                mergeCells1.Append(mergeCell6);

                int sums = 0;
                Decimal sumIn = 0;
                foreach (var v in sumMoney)
                {
                    sums += 1;
                    sheetRowNo += 1;
                    Row row22 = new Row() { RowIndex = (UInt32Value)Convert.ToUInt32(sheetRowNo), Spans = new ListValue<StringValue>() { InnerText = "1:6" }, Height = 20.0D, CustomHeight = true, DyDescent = 0.15D };
                    //序号
                    Cell cell55 = new Cell() { CellReference = "A" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)4U, DataType = CellValues.String };
                    CellValue cellValue55 = new CellValue();
                    cellValue55.Text = sums.ToString();
                    cell55.Append(cellValue55);

                    //经费类型
                    Cell cell56 = new Cell() { CellReference = "B" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)17U, DataType = CellValues.String };
                    CellValue cellValue56 = new CellValue();
                    cellValue56.Text = v.FundClassify;
                    cell56.Append(cellValue56);

                    Cell cell57 = new Cell() { CellReference = "C" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)35U };
                    Cell cell58 = new Cell() { CellReference = "D" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)35U };

                    //金额
                    Cell cell59 = new Cell() { CellReference = "E" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)35U, DataType = CellValues.Number };
                    CellValue cellValue59 = new CellValue();
                    cellValue59.Text = v.SumMoney.ToString();
                    sumIn += (Decimal)v.SumMoney;
                    cell59.Append(cellValue59);
                    Cell cell60 = new Cell() { CellReference = "F" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)36U };

                    row22.Append(cell55);
                    row22.Append(cell56);
                    row22.Append(cell57);
                    row22.Append(cell58);
                    row22.Append(cell59);
                    row22.Append(cell60);

                    sheetData1.Append(row22);

                    MergeCell mergeCell7 = new MergeCell() { Reference = "B" + sheetRowNo.ToString() + ":D" + sheetRowNo.ToString() };
                    MergeCell mergeCell8 = new MergeCell() { Reference = "E" + sheetRowNo.ToString() + ":F" + sheetRowNo.ToString() };
                    mergeCells1.Append(mergeCell7);
                    mergeCells1.Append(mergeCell8);
                }
                //合计
                sheetRowNo += 1;
                Row row23 = new Row() { RowIndex = (UInt32Value)Convert.ToUInt32(sheetRowNo), Spans = new ListValue<StringValue>() { InnerText = "1:6" }, Height = 20.0D, CustomHeight = true, DyDescent = 0.15D };
                Cell cell61 = new Cell() { CellReference = "A" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)9U };
                Cell cell62 = new Cell() { CellReference = "B" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)21U, DataType = CellValues.String };
                CellValue cellValue62 = new CellValue();
                cellValue62.Text = "合计";
                cell62.Append(cellValue62);
                Cell cell63 = new Cell() { CellReference = "C" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)22U };
                Cell cell64 = new Cell() { CellReference = "D" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)23U };
                Cell cell65 = new Cell() { CellReference = "E" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)37U, DataType = CellValues.Number };
                CellValue cellValue65 = new CellValue();
                cellValue65.Text = sumIn.ToString();
                cell65.Append(cellValue65);
                Cell cell66 = new Cell() { CellReference = "F" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)36U };

                row23.Append(cell61);
                row23.Append(cell62);
                row23.Append(cell63);
                row23.Append(cell64);
                row23.Append(cell65);
                row23.Append(cell66);

                sheetData1.Append(row23);
                MergeCell mergeCell9 = new MergeCell() { Reference = "B" + sheetRowNo.ToString() + ":D" + sheetRowNo.ToString() };
                MergeCell mergeCell10 = new MergeCell() { Reference = "E" + sheetRowNo.ToString() + ":F" + sheetRowNo.ToString() };
                mergeCells1.Append(mergeCell9);
                mergeCells1.Append(mergeCell10);
            }

            //支出总览
            var sumPay = dataContext.View_ProjectFundSumIncoming.Where(s => s.IncomeOrPay.Equals(false) && s.ProjectID.Equals(projectID));
            if (sumPay.Count() > 0)
            {
                fundClassifyNo += 1; //序号
                sheetRowNo += 1;
                DigitToChnText dt = new DigitToChnText();
                Row row21 = new Row() { RowIndex = (UInt32Value)Convert.ToUInt32(sheetRowNo), Spans = new ListValue<StringValue>() { InnerText = "1:6" }, Height = 20.0D, CustomHeight = true, ThickTop = true, DyDescent = 0.15D };
                //中文序号
                Cell cell112 = new Cell() { CellReference = "A" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)4U, DataType = CellValues.String };
                CellValue cellValue112 = new CellValue();
                cellValue112.Text = dt.Convert(fundClassifyNo.ToString(), false);
                cell112.Append(cellValue112);
                row21.Append(cell112);

                Cell cell113 = new Cell() { CellReference = "B" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)30U, DataType = CellValues.String };
                CellValue cellValue113 = new CellValue();
                cellValue113.Text = "支出总览";
                cell113.Append(cellValue113);
                row21.Append(cell113);

                Cell cell114 = new Cell() { CellReference = "C" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)33U };
                Cell cell115 = new Cell() { CellReference = "D" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)33U };
                Cell cell116 = new Cell() { CellReference = "E" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)33U };
                Cell cell117 = new Cell() { CellReference = "F" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)34U };
                row21.Append(cell114);
                row21.Append(cell115);
                row21.Append(cell116);
                row21.Append(cell117);

                sheetData1.Append(row21);

                MergeCell mergeCell4 = new MergeCell() { Reference = "B" + sheetRowNo.ToString() + ":F" + sheetRowNo.ToString() };
                mergeCells1.Append(mergeCell4);

                sheetRowNo += 1;
                //题头行
                Row row9 = new Row() { RowIndex = (UInt32Value)Convert.ToUInt32(sheetRowNo), Spans = new ListValue<StringValue>() { InnerText = "1:6" }, Height = 20.0D, CustomHeight = true, DyDescent = 0.15D };
                Cell cell49 = new Cell() { CellReference = "A" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)4U };
                Cell cell50 = new Cell() { CellReference = "B" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
                CellValue cellValue50 = new CellValue();
                cellValue50.Text = "23";
                cell50.Append(cellValue50);

                Cell cell51 = new Cell() { CellReference = "C" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)30U };

                Cell cell52 = new Cell() { CellReference = "D" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)33U };
                Cell cell53 = new Cell() { CellReference = "E" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)30U, DataType = CellValues.SharedString };
                CellValue cellValue53 = new CellValue();
                cellValue53.Text = "24";
                cell53.Append(cellValue53);

                Cell cell54 = new Cell() { CellReference = "F" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)34U };
                row9.Append(cell49);
                row9.Append(cell50);
                row9.Append(cell51);
                row9.Append(cell52);
                row9.Append(cell53);
                row9.Append(cell54);

                sheetData1.Append(row9);

                MergeCell mergeCell5 = new MergeCell() { Reference = "B" + sheetRowNo.ToString() + ":D" + sheetRowNo.ToString() };
                MergeCell mergeCell6 = new MergeCell() { Reference = "E" + sheetRowNo.ToString() + ":F" + sheetRowNo.ToString() };
                mergeCells1.Append(mergeCell5);
                mergeCells1.Append(mergeCell6);

                int sums = 0;
                Decimal sumPayfor = 0;
                foreach (var v in sumPay)
                {
                    sums += 1;
                    sheetRowNo += 1;
                    Row row22 = new Row() { RowIndex = (UInt32Value)Convert.ToUInt32(sheetRowNo), Spans = new ListValue<StringValue>() { InnerText = "1:6" }, Height = 20.0D, CustomHeight = true, DyDescent = 0.15D };
                    //序号
                    Cell cell55 = new Cell() { CellReference = "A" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)4U, DataType = CellValues.String };
                    CellValue cellValue55 = new CellValue();
                    cellValue55.Text = sums.ToString();
                    cell55.Append(cellValue55);

                    //经费类型
                    Cell cell56 = new Cell() { CellReference = "B" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)17U, DataType = CellValues.String };
                    CellValue cellValue56 = new CellValue();
                    cellValue56.Text = v.FundClassify;
                    cell56.Append(cellValue56);

                    Cell cell57 = new Cell() { CellReference = "C" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)35U };
                    Cell cell58 = new Cell() { CellReference = "D" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)35U };
                    Cell cell59 = new Cell() { CellReference = "E" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)35U, DataType = CellValues.Number };
                    CellValue cellValue59 = new CellValue();
                    cellValue59.Text = v.SumMoney.ToString();
                    sumPayfor += (Decimal)v.SumMoney;
                    cell59.Append(cellValue59);
                    Cell cell60 = new Cell() { CellReference = "F" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)36U };

                    row22.Append(cell55);
                    row22.Append(cell56);
                    row22.Append(cell57);
                    row22.Append(cell58);
                    row22.Append(cell59);
                    row22.Append(cell60);

                    sheetData1.Append(row22);

                    MergeCell mergeCell7 = new MergeCell() { Reference = "B" + sheetRowNo.ToString() + ":D" + sheetRowNo.ToString() };
                    MergeCell mergeCell8 = new MergeCell() { Reference = "E" + sheetRowNo.ToString() + ":F" + sheetRowNo.ToString() };
                    mergeCells1.Append(mergeCell7);
                    mergeCells1.Append(mergeCell8);
                }
                //合计
                sheetRowNo += 1;
                Row row23 = new Row() { RowIndex = (UInt32Value)Convert.ToUInt32(sheetRowNo), Spans = new ListValue<StringValue>() { InnerText = "1:6" }, Height = 20.0D, CustomHeight = true, DyDescent = 0.15D };
                Cell cell61 = new Cell() { CellReference = "A" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)16U };
                Cell cell62 = new Cell() { CellReference = "B" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)24U, DataType = CellValues.String };
                CellValue cellValue62 = new CellValue();
                cellValue62.Text = "合计";
                cell62.Append(cellValue62);
                Cell cell63 = new Cell() { CellReference = "C" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)25U };
                Cell cell64 = new Cell() { CellReference = "D" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)26U };
                Cell cell65 = new Cell() { CellReference = "E" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)19U, DataType = CellValues.Number };
                CellValue cellValue65 = new CellValue();
                cellValue65.Text = sumPayfor.ToString();
                cell65.Append(cellValue65);
                Cell cell66 = new Cell() { CellReference = "F" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)20U };

                row23.Append(cell61);
                row23.Append(cell62);
                row23.Append(cell63);
                row23.Append(cell64);
                row23.Append(cell65);
                row23.Append(cell66);

                sheetData1.Append(row23);
                MergeCell mergeCell9 = new MergeCell() { Reference = "B" + sheetRowNo.ToString() + ":D" + sheetRowNo.ToString() };
                MergeCell mergeCell10 = new MergeCell() { Reference = "E" + sheetRowNo.ToString() + ":F" + sheetRowNo.ToString() };
                mergeCells1.Append(mergeCell9);
                mergeCells1.Append(mergeCell10);
            }

            //合计行
            sheetRowNo += 1;
            Row row7 = new Row() { RowIndex = (UInt32Value)Convert.ToUInt32(sheetRowNo), Spans = new ListValue<StringValue>() { InnerText = "1:6" }, Height = 20.100000000000001D, CustomHeight = true, ThickBot = true, DyDescent = 0.2D };
            Cell cell37 = new Cell() { CellReference = "A" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)45U };

            //合计行
            string sumtext = "收入合计：" + incoming.ToString() + "万元，支出合计：" + outcoming.ToString() + "万元，结余：" + (incoming - outcoming).ToString() + "万元";

            Cell cell38 = new Cell() { CellReference = "B" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)46U, DataType = CellValues.String };
            CellValue cellValue20 = new CellValue();
            cellValue20.Text = sumtext;

            cell38.Append(cellValue20);
            Cell cell39 = new Cell() { CellReference = "C" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)46U };
            Cell cell40 = new Cell() { CellReference = "D" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)46U };
            Cell cell41 = new Cell() { CellReference = "E" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)46U };
            Cell cell42 = new Cell() { CellReference = "F" + sheetRowNo.ToString(), StyleIndex = (UInt32Value)47U };

            row7.Append(cell37);
            row7.Append(cell38);
            row7.Append(cell39);
            row7.Append(cell40);
            row7.Append(cell41);
            row7.Append(cell42);

            sheetData1.Append(row7);

            MergeCell mergeCell100 = new MergeCell() { Reference = "B" + sheetRowNo.ToString() + ":F" + sheetRowNo.ToString() };

            mergeCells1.Append(mergeCell100);
            PhoneticProperties phoneticProperties3 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };
            PageMargins pageMargins3 = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };

            worksheet3.Append(sheetDimension3);
            worksheet3.Append(sheetViews3);
            worksheet3.Append(sheetFormatProperties3);
            worksheet3.Append(columns1);
            worksheet3.Append(sheetData1);
            worksheet3.Append(mergeCells1);
            worksheet3.Append(phoneticProperties3);
            worksheet3.Append(pageMargins3);

            worksheetPart3.Worksheet = worksheet3;
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
            WorkbookView workbookView1 = new WorkbookView() { XWindow = 240, YWindow = 90, WindowWidth = (UInt32Value)24795U, WindowHeight = (UInt32Value)11895U };

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

            DocumentFormat.OpenXml.Spreadsheet.Fonts fonts1 = new DocumentFormat.OpenXml.Spreadsheet.Fonts() { Count = (UInt32Value)9U, KnownFonts = true };

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
            FontSize fontSize2 = new FontSize() { Val = 9D };
            FontName fontName2 = new FontName() { Val = "宋体" };
            FontFamilyNumbering fontFamilyNumbering2 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet2 = new FontCharSet() { Val = 134 };
            FontScheme fontScheme2 = new FontScheme() { Val = FontSchemeValues.Minor };

            font2.Append(fontSize2);
            font2.Append(fontName2);
            font2.Append(fontFamilyNumbering2);
            font2.Append(fontCharSet2);
            font2.Append(fontScheme2);

            Font font3 = new Font();
            DocumentFormat.OpenXml.Spreadsheet.Bold bold1 = new DocumentFormat.OpenXml.Spreadsheet.Bold();
            FontSize fontSize3 = new FontSize() { Val = 18D };
            DocumentFormat.OpenXml.Spreadsheet.Color color2 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Theme = (UInt32Value)1U };
            FontName fontName3 = new FontName() { Val = "宋体" };
            FontFamilyNumbering fontFamilyNumbering3 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet3 = new FontCharSet() { Val = 134 };
            FontScheme fontScheme3 = new FontScheme() { Val = FontSchemeValues.Minor };

            font3.Append(bold1);
            font3.Append(fontSize3);
            font3.Append(color2);
            font3.Append(fontName3);
            font3.Append(fontFamilyNumbering3);
            font3.Append(fontCharSet3);
            font3.Append(fontScheme3);

            Font font4 = new Font();
            DocumentFormat.OpenXml.Spreadsheet.Bold bold2 = new DocumentFormat.OpenXml.Spreadsheet.Bold();
            FontSize fontSize4 = new FontSize() { Val = 18D };
            DocumentFormat.OpenXml.Spreadsheet.Color color3 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Theme = (UInt32Value)1U };
            FontName fontName4 = new FontName() { Val = "宋体" };
            FontFamilyNumbering fontFamilyNumbering4 = new FontFamilyNumbering() { Val = 3 };
            FontCharSet fontCharSet4 = new FontCharSet() { Val = 134 };
            FontScheme fontScheme4 = new FontScheme() { Val = FontSchemeValues.Minor };

            font4.Append(bold2);
            font4.Append(fontSize4);
            font4.Append(color3);
            font4.Append(fontName4);
            font4.Append(fontFamilyNumbering4);
            font4.Append(fontCharSet4);
            font4.Append(fontScheme4);

            Font font5 = new Font();
            DocumentFormat.OpenXml.Spreadsheet.Bold bold3 = new DocumentFormat.OpenXml.Spreadsheet.Bold();
            FontSize fontSize5 = new FontSize() { Val = 14D };
            DocumentFormat.OpenXml.Spreadsheet.Color color4 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Theme = (UInt32Value)1U };
            FontName fontName5 = new FontName() { Val = "宋体" };
            FontFamilyNumbering fontFamilyNumbering5 = new FontFamilyNumbering() { Val = 3 };
            FontCharSet fontCharSet5 = new FontCharSet() { Val = 134 };
            FontScheme fontScheme5 = new FontScheme() { Val = FontSchemeValues.Minor };

            font5.Append(bold3);
            font5.Append(fontSize5);
            font5.Append(color4);
            font5.Append(fontName5);
            font5.Append(fontFamilyNumbering5);
            font5.Append(fontCharSet5);
            font5.Append(fontScheme5);

            Font font6 = new Font();
            DocumentFormat.OpenXml.Spreadsheet.Bold bold4 = new DocumentFormat.OpenXml.Spreadsheet.Bold();
            FontSize fontSize6 = new FontSize() { Val = 11D };
            DocumentFormat.OpenXml.Spreadsheet.Color color5 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Theme = (UInt32Value)1U };
            FontName fontName6 = new FontName() { Val = "宋体" };
            FontFamilyNumbering fontFamilyNumbering6 = new FontFamilyNumbering() { Val = 3 };
            FontCharSet fontCharSet6 = new FontCharSet() { Val = 134 };
            FontScheme fontScheme6 = new FontScheme() { Val = FontSchemeValues.Minor };

            font6.Append(bold4);
            font6.Append(fontSize6);
            font6.Append(color5);
            font6.Append(fontName6);
            font6.Append(fontFamilyNumbering6);
            font6.Append(fontCharSet6);
            font6.Append(fontScheme6);

            Font font7 = new Font();
            DocumentFormat.OpenXml.Spreadsheet.Bold bold5 = new DocumentFormat.OpenXml.Spreadsheet.Bold();
            FontSize fontSize7 = new FontSize() { Val = 12D };
            DocumentFormat.OpenXml.Spreadsheet.Color color6 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Theme = (UInt32Value)1U };
            FontName fontName7 = new FontName() { Val = "宋体" };
            FontFamilyNumbering fontFamilyNumbering7 = new FontFamilyNumbering() { Val = 3 };
            FontCharSet fontCharSet7 = new FontCharSet() { Val = 134 };
            FontScheme fontScheme7 = new FontScheme() { Val = FontSchemeValues.Minor };

            font7.Append(bold5);
            font7.Append(fontSize7);
            font7.Append(color6);
            font7.Append(fontName7);
            font7.Append(fontFamilyNumbering7);
            font7.Append(fontCharSet7);
            font7.Append(fontScheme7);

            Font font8 = new Font();
            FontSize fontSize8 = new FontSize() { Val = 11D };
            DocumentFormat.OpenXml.Spreadsheet.Color color7 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Theme = (UInt32Value)1U };
            FontName fontName8 = new FontName() { Val = "宋体" };
            FontFamilyNumbering fontFamilyNumbering8 = new FontFamilyNumbering() { Val = 3 };
            FontCharSet fontCharSet8 = new FontCharSet() { Val = 134 };
            FontScheme fontScheme8 = new FontScheme() { Val = FontSchemeValues.Minor };

            font8.Append(fontSize8);
            font8.Append(color7);
            font8.Append(fontName8);
            font8.Append(fontFamilyNumbering8);
            font8.Append(fontCharSet8);
            font8.Append(fontScheme8);

            Font font9 = new Font();
            FontSize fontSize9 = new FontSize() { Val = 12D };
            DocumentFormat.OpenXml.Spreadsheet.Color color8 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Theme = (UInt32Value)1U };
            FontName fontName9 = new FontName() { Val = "宋体" };
            FontFamilyNumbering fontFamilyNumbering9 = new FontFamilyNumbering() { Val = 3 };
            FontCharSet fontCharSet9 = new FontCharSet() { Val = 134 };
            FontScheme fontScheme9 = new FontScheme() { Val = FontSchemeValues.Minor };

            font9.Append(fontSize9);
            font9.Append(color8);
            font9.Append(fontName9);
            font9.Append(fontFamilyNumbering9);
            font9.Append(fontCharSet9);
            font9.Append(fontScheme9);

            fonts1.Append(font1);
            fonts1.Append(font2);
            fonts1.Append(font3);
            fonts1.Append(font4);
            fonts1.Append(font5);
            fonts1.Append(font6);
            fonts1.Append(font7);
            fonts1.Append(font8);
            fonts1.Append(font9);

            Fills fills1 = new Fills() { Count = (UInt32Value)2U };

            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };

            fill1.Append(patternFill1);

            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };

            fill2.Append(patternFill2);

            fills1.Append(fill1);
            fills1.Append(fill2);

            Borders borders1 = new Borders() { Count = (UInt32Value)25U };

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

            LeftBorder leftBorder2 = new LeftBorder() { Style = BorderStyleValues.Medium };
            DocumentFormat.OpenXml.Spreadsheet.Color color9 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            leftBorder2.Append(color9);

            RightBorder rightBorder2 = new RightBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color10 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            rightBorder2.Append(color10);

            TopBorder topBorder2 = new TopBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color11 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            topBorder2.Append(color11);

            BottomBorder bottomBorder2 = new BottomBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color12 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            bottomBorder2.Append(color12);
            DiagonalBorder diagonalBorder2 = new DiagonalBorder();

            border2.Append(leftBorder2);
            border2.Append(rightBorder2);
            border2.Append(topBorder2);
            border2.Append(bottomBorder2);
            border2.Append(diagonalBorder2);

            DocumentFormat.OpenXml.Spreadsheet.Border border3 = new DocumentFormat.OpenXml.Spreadsheet.Border();

            LeftBorder leftBorder3 = new LeftBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color13 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            leftBorder3.Append(color13);

            RightBorder rightBorder3 = new RightBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color14 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            rightBorder3.Append(color14);

            TopBorder topBorder3 = new TopBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color15 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            topBorder3.Append(color15);

            BottomBorder bottomBorder3 = new BottomBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color16 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            bottomBorder3.Append(color16);
            DiagonalBorder diagonalBorder3 = new DiagonalBorder();

            border3.Append(leftBorder3);
            border3.Append(rightBorder3);
            border3.Append(topBorder3);
            border3.Append(bottomBorder3);
            border3.Append(diagonalBorder3);

            DocumentFormat.OpenXml.Spreadsheet.Border border4 = new DocumentFormat.OpenXml.Spreadsheet.Border();

            LeftBorder leftBorder4 = new LeftBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color17 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            leftBorder4.Append(color17);

            RightBorder rightBorder4 = new RightBorder() { Style = BorderStyleValues.Medium };
            DocumentFormat.OpenXml.Spreadsheet.Color color18 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            rightBorder4.Append(color18);

            TopBorder topBorder4 = new TopBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color19 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            topBorder4.Append(color19);

            BottomBorder bottomBorder4 = new BottomBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color20 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            bottomBorder4.Append(color20);
            DiagonalBorder diagonalBorder4 = new DiagonalBorder();

            border4.Append(leftBorder4);
            border4.Append(rightBorder4);
            border4.Append(topBorder4);
            border4.Append(bottomBorder4);
            border4.Append(diagonalBorder4);

            DocumentFormat.OpenXml.Spreadsheet.Border border5 = new DocumentFormat.OpenXml.Spreadsheet.Border();

            LeftBorder leftBorder5 = new LeftBorder() { Style = BorderStyleValues.Medium };
            DocumentFormat.OpenXml.Spreadsheet.Color color21 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            leftBorder5.Append(color21);

            RightBorder rightBorder5 = new RightBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color22 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            rightBorder5.Append(color22);

            TopBorder topBorder5 = new TopBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color23 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            topBorder5.Append(color23);
            BottomBorder bottomBorder5 = new BottomBorder();
            DiagonalBorder diagonalBorder5 = new DiagonalBorder();

            border5.Append(leftBorder5);
            border5.Append(rightBorder5);
            border5.Append(topBorder5);
            border5.Append(bottomBorder5);
            border5.Append(diagonalBorder5);

            DocumentFormat.OpenXml.Spreadsheet.Border border6 = new DocumentFormat.OpenXml.Spreadsheet.Border();

            LeftBorder leftBorder6 = new LeftBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color24 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            leftBorder6.Append(color24);

            RightBorder rightBorder6 = new RightBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color25 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            rightBorder6.Append(color25);

            TopBorder topBorder6 = new TopBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color26 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            topBorder6.Append(color26);
            BottomBorder bottomBorder6 = new BottomBorder();
            DiagonalBorder diagonalBorder6 = new DiagonalBorder();

            border6.Append(leftBorder6);
            border6.Append(rightBorder6);
            border6.Append(topBorder6);
            border6.Append(bottomBorder6);
            border6.Append(diagonalBorder6);

            DocumentFormat.OpenXml.Spreadsheet.Border border7 = new DocumentFormat.OpenXml.Spreadsheet.Border();

            LeftBorder leftBorder7 = new LeftBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color27 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            leftBorder7.Append(color27);

            RightBorder rightBorder7 = new RightBorder() { Style = BorderStyleValues.Medium };
            DocumentFormat.OpenXml.Spreadsheet.Color color28 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            rightBorder7.Append(color28);

            TopBorder topBorder7 = new TopBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color29 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            topBorder7.Append(color29);
            BottomBorder bottomBorder7 = new BottomBorder();
            DiagonalBorder diagonalBorder7 = new DiagonalBorder();

            border7.Append(leftBorder7);
            border7.Append(rightBorder7);
            border7.Append(topBorder7);
            border7.Append(bottomBorder7);
            border7.Append(diagonalBorder7);

            DocumentFormat.OpenXml.Spreadsheet.Border border8 = new DocumentFormat.OpenXml.Spreadsheet.Border();

            LeftBorder leftBorder8 = new LeftBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color30 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            leftBorder8.Append(color30);
            RightBorder rightBorder8 = new RightBorder();

            TopBorder topBorder8 = new TopBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color31 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            topBorder8.Append(color31);

            BottomBorder bottomBorder8 = new BottomBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color32 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            bottomBorder8.Append(color32);
            DiagonalBorder diagonalBorder8 = new DiagonalBorder();

            border8.Append(leftBorder8);
            border8.Append(rightBorder8);
            border8.Append(topBorder8);
            border8.Append(bottomBorder8);
            border8.Append(diagonalBorder8);

            DocumentFormat.OpenXml.Spreadsheet.Border border9 = new DocumentFormat.OpenXml.Spreadsheet.Border();
            LeftBorder leftBorder9 = new LeftBorder();
            RightBorder rightBorder9 = new RightBorder();

            TopBorder topBorder9 = new TopBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color33 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            topBorder9.Append(color33);

            BottomBorder bottomBorder9 = new BottomBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color34 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            bottomBorder9.Append(color34);
            DiagonalBorder diagonalBorder9 = new DiagonalBorder();

            border9.Append(leftBorder9);
            border9.Append(rightBorder9);
            border9.Append(topBorder9);
            border9.Append(bottomBorder9);
            border9.Append(diagonalBorder9);

            DocumentFormat.OpenXml.Spreadsheet.Border border10 = new DocumentFormat.OpenXml.Spreadsheet.Border();

            LeftBorder leftBorder10 = new LeftBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color35 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            leftBorder10.Append(color35);
            RightBorder rightBorder10 = new RightBorder();

            TopBorder topBorder10 = new TopBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color36 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            topBorder10.Append(color36);
            BottomBorder bottomBorder10 = new BottomBorder();
            DiagonalBorder diagonalBorder10 = new DiagonalBorder();

            border10.Append(leftBorder10);
            border10.Append(rightBorder10);
            border10.Append(topBorder10);
            border10.Append(bottomBorder10);
            border10.Append(diagonalBorder10);

            DocumentFormat.OpenXml.Spreadsheet.Border border11 = new DocumentFormat.OpenXml.Spreadsheet.Border();
            LeftBorder leftBorder11 = new LeftBorder();
            RightBorder rightBorder11 = new RightBorder();

            TopBorder topBorder11 = new TopBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color37 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            topBorder11.Append(color37);
            BottomBorder bottomBorder11 = new BottomBorder();
            DiagonalBorder diagonalBorder11 = new DiagonalBorder();

            border11.Append(leftBorder11);
            border11.Append(rightBorder11);
            border11.Append(topBorder11);
            border11.Append(bottomBorder11);
            border11.Append(diagonalBorder11);

            DocumentFormat.OpenXml.Spreadsheet.Border border12 = new DocumentFormat.OpenXml.Spreadsheet.Border();
            LeftBorder leftBorder12 = new LeftBorder();

            RightBorder rightBorder12 = new RightBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color38 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            rightBorder12.Append(color38);

            TopBorder topBorder12 = new TopBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color39 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            topBorder12.Append(color39);

            BottomBorder bottomBorder12 = new BottomBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color40 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            bottomBorder12.Append(color40);
            DiagonalBorder diagonalBorder12 = new DiagonalBorder();

            border12.Append(leftBorder12);
            border12.Append(rightBorder12);
            border12.Append(topBorder12);
            border12.Append(bottomBorder12);
            border12.Append(diagonalBorder12);

            DocumentFormat.OpenXml.Spreadsheet.Border border13 = new DocumentFormat.OpenXml.Spreadsheet.Border();
            LeftBorder leftBorder13 = new LeftBorder();

            RightBorder rightBorder13 = new RightBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color41 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            rightBorder13.Append(color41);

            TopBorder topBorder13 = new TopBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color42 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            topBorder13.Append(color42);
            BottomBorder bottomBorder13 = new BottomBorder();
            DiagonalBorder diagonalBorder13 = new DiagonalBorder();

            border13.Append(leftBorder13);
            border13.Append(rightBorder13);
            border13.Append(topBorder13);
            border13.Append(bottomBorder13);
            border13.Append(diagonalBorder13);

            DocumentFormat.OpenXml.Spreadsheet.Border border14 = new DocumentFormat.OpenXml.Spreadsheet.Border();

            LeftBorder leftBorder14 = new LeftBorder() { Style = BorderStyleValues.Medium };
            DocumentFormat.OpenXml.Spreadsheet.Color color43 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            leftBorder14.Append(color43);

            RightBorder rightBorder14 = new RightBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color44 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            rightBorder14.Append(color44);

            TopBorder topBorder14 = new TopBorder() { Style = BorderStyleValues.Double };
            DocumentFormat.OpenXml.Spreadsheet.Color color45 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            topBorder14.Append(color45);

            BottomBorder bottomBorder14 = new BottomBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color46 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            bottomBorder14.Append(color46);
            DiagonalBorder diagonalBorder14 = new DiagonalBorder();

            border14.Append(leftBorder14);
            border14.Append(rightBorder14);
            border14.Append(topBorder14);
            border14.Append(bottomBorder14);
            border14.Append(diagonalBorder14);

            DocumentFormat.OpenXml.Spreadsheet.Border border15 = new DocumentFormat.OpenXml.Spreadsheet.Border();

            LeftBorder leftBorder15 = new LeftBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color47 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            leftBorder15.Append(color47);

            RightBorder rightBorder15 = new RightBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color48 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            rightBorder15.Append(color48);

            TopBorder topBorder15 = new TopBorder() { Style = BorderStyleValues.Double };
            DocumentFormat.OpenXml.Spreadsheet.Color color49 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            topBorder15.Append(color49);

            BottomBorder bottomBorder15 = new BottomBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color50 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            bottomBorder15.Append(color50);
            DiagonalBorder diagonalBorder15 = new DiagonalBorder();

            border15.Append(leftBorder15);
            border15.Append(rightBorder15);
            border15.Append(topBorder15);
            border15.Append(bottomBorder15);
            border15.Append(diagonalBorder15);

            DocumentFormat.OpenXml.Spreadsheet.Border border16 = new DocumentFormat.OpenXml.Spreadsheet.Border();

            LeftBorder leftBorder16 = new LeftBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color51 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            leftBorder16.Append(color51);

            RightBorder rightBorder16 = new RightBorder() { Style = BorderStyleValues.Medium };
            DocumentFormat.OpenXml.Spreadsheet.Color color52 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            rightBorder16.Append(color52);

            TopBorder topBorder16 = new TopBorder() { Style = BorderStyleValues.Double };
            DocumentFormat.OpenXml.Spreadsheet.Color color53 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            topBorder16.Append(color53);

            BottomBorder bottomBorder16 = new BottomBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color54 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            bottomBorder16.Append(color54);
            DiagonalBorder diagonalBorder16 = new DiagonalBorder();

            border16.Append(leftBorder16);
            border16.Append(rightBorder16);
            border16.Append(topBorder16);
            border16.Append(bottomBorder16);
            border16.Append(diagonalBorder16);

            DocumentFormat.OpenXml.Spreadsheet.Border border17 = new DocumentFormat.OpenXml.Spreadsheet.Border();

            LeftBorder leftBorder17 = new LeftBorder() { Style = BorderStyleValues.Medium };
            DocumentFormat.OpenXml.Spreadsheet.Color color55 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            leftBorder17.Append(color55);

            RightBorder rightBorder17 = new RightBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color56 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            rightBorder17.Append(color56);
            TopBorder topBorder17 = new TopBorder();

            BottomBorder bottomBorder17 = new BottomBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color57 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            bottomBorder17.Append(color57);
            DiagonalBorder diagonalBorder17 = new DiagonalBorder();

            border17.Append(leftBorder17);
            border17.Append(rightBorder17);
            border17.Append(topBorder17);
            border17.Append(bottomBorder17);
            border17.Append(diagonalBorder17);

            DocumentFormat.OpenXml.Spreadsheet.Border border18 = new DocumentFormat.OpenXml.Spreadsheet.Border();

            LeftBorder leftBorder18 = new LeftBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color58 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            leftBorder18.Append(color58);

            RightBorder rightBorder18 = new RightBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color59 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            rightBorder18.Append(color59);
            TopBorder topBorder18 = new TopBorder();

            BottomBorder bottomBorder18 = new BottomBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color60 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            bottomBorder18.Append(color60);
            DiagonalBorder diagonalBorder18 = new DiagonalBorder();

            border18.Append(leftBorder18);
            border18.Append(rightBorder18);
            border18.Append(topBorder18);
            border18.Append(bottomBorder18);
            border18.Append(diagonalBorder18);

            DocumentFormat.OpenXml.Spreadsheet.Border border19 = new DocumentFormat.OpenXml.Spreadsheet.Border();

            LeftBorder leftBorder19 = new LeftBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color61 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            leftBorder19.Append(color61);

            RightBorder rightBorder19 = new RightBorder() { Style = BorderStyleValues.Medium };
            DocumentFormat.OpenXml.Spreadsheet.Color color62 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            rightBorder19.Append(color62);
            TopBorder topBorder19 = new TopBorder();

            BottomBorder bottomBorder19 = new BottomBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color63 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            bottomBorder19.Append(color63);
            DiagonalBorder diagonalBorder19 = new DiagonalBorder();

            border19.Append(leftBorder19);
            border19.Append(rightBorder19);
            border19.Append(topBorder19);
            border19.Append(bottomBorder19);
            border19.Append(diagonalBorder19);

            DocumentFormat.OpenXml.Spreadsheet.Border border20 = new DocumentFormat.OpenXml.Spreadsheet.Border();

            LeftBorder leftBorder20 = new LeftBorder() { Style = BorderStyleValues.Medium };
            DocumentFormat.OpenXml.Spreadsheet.Color color64 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            leftBorder20.Append(color64);

            RightBorder rightBorder20 = new RightBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color65 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            rightBorder20.Append(color65);

            TopBorder topBorder20 = new TopBorder() { Style = BorderStyleValues.Medium };
            DocumentFormat.OpenXml.Spreadsheet.Color color66 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            topBorder20.Append(color66);

            BottomBorder bottomBorder20 = new BottomBorder() { Style = BorderStyleValues.Double };
            DocumentFormat.OpenXml.Spreadsheet.Color color67 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            bottomBorder20.Append(color67);
            DiagonalBorder diagonalBorder20 = new DiagonalBorder();

            border20.Append(leftBorder20);
            border20.Append(rightBorder20);
            border20.Append(topBorder20);
            border20.Append(bottomBorder20);
            border20.Append(diagonalBorder20);

            DocumentFormat.OpenXml.Spreadsheet.Border border21 = new DocumentFormat.OpenXml.Spreadsheet.Border();

            LeftBorder leftBorder21 = new LeftBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color68 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            leftBorder21.Append(color68);

            RightBorder rightBorder21 = new RightBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color69 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            rightBorder21.Append(color69);

            TopBorder topBorder21 = new TopBorder() { Style = BorderStyleValues.Medium };
            DocumentFormat.OpenXml.Spreadsheet.Color color70 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            topBorder21.Append(color70);

            BottomBorder bottomBorder21 = new BottomBorder() { Style = BorderStyleValues.Double };
            DocumentFormat.OpenXml.Spreadsheet.Color color71 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            bottomBorder21.Append(color71);
            DiagonalBorder diagonalBorder21 = new DiagonalBorder();

            border21.Append(leftBorder21);
            border21.Append(rightBorder21);
            border21.Append(topBorder21);
            border21.Append(bottomBorder21);
            border21.Append(diagonalBorder21);

            DocumentFormat.OpenXml.Spreadsheet.Border border22 = new DocumentFormat.OpenXml.Spreadsheet.Border();

            LeftBorder leftBorder22 = new LeftBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color72 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            leftBorder22.Append(color72);

            RightBorder rightBorder22 = new RightBorder() { Style = BorderStyleValues.Medium };
            DocumentFormat.OpenXml.Spreadsheet.Color color73 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            rightBorder22.Append(color73);

            TopBorder topBorder22 = new TopBorder() { Style = BorderStyleValues.Medium };
            DocumentFormat.OpenXml.Spreadsheet.Color color74 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            topBorder22.Append(color74);

            BottomBorder bottomBorder22 = new BottomBorder() { Style = BorderStyleValues.Double };
            DocumentFormat.OpenXml.Spreadsheet.Color color75 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            bottomBorder22.Append(color75);
            DiagonalBorder diagonalBorder22 = new DiagonalBorder();

            border22.Append(leftBorder22);
            border22.Append(rightBorder22);
            border22.Append(topBorder22);
            border22.Append(bottomBorder22);
            border22.Append(diagonalBorder22);

            DocumentFormat.OpenXml.Spreadsheet.Border border23 = new DocumentFormat.OpenXml.Spreadsheet.Border();

            LeftBorder leftBorder23 = new LeftBorder() { Style = BorderStyleValues.Medium };
            DocumentFormat.OpenXml.Spreadsheet.Color color76 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            leftBorder23.Append(color76);
            RightBorder rightBorder23 = new RightBorder();

            TopBorder topBorder23 = new TopBorder() { Style = BorderStyleValues.Double };
            DocumentFormat.OpenXml.Spreadsheet.Color color77 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            topBorder23.Append(color77);

            BottomBorder bottomBorder23 = new BottomBorder() { Style = BorderStyleValues.Medium };
            DocumentFormat.OpenXml.Spreadsheet.Color color78 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            bottomBorder23.Append(color78);
            DiagonalBorder diagonalBorder23 = new DiagonalBorder();

            border23.Append(leftBorder23);
            border23.Append(rightBorder23);
            border23.Append(topBorder23);
            border23.Append(bottomBorder23);
            border23.Append(diagonalBorder23);

            DocumentFormat.OpenXml.Spreadsheet.Border border24 = new DocumentFormat.OpenXml.Spreadsheet.Border();
            LeftBorder leftBorder24 = new LeftBorder();
            RightBorder rightBorder24 = new RightBorder();

            TopBorder topBorder24 = new TopBorder() { Style = BorderStyleValues.Double };
            DocumentFormat.OpenXml.Spreadsheet.Color color79 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            topBorder24.Append(color79);

            BottomBorder bottomBorder24 = new BottomBorder() { Style = BorderStyleValues.Medium };
            DocumentFormat.OpenXml.Spreadsheet.Color color80 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            bottomBorder24.Append(color80);
            DiagonalBorder diagonalBorder24 = new DiagonalBorder();

            border24.Append(leftBorder24);
            border24.Append(rightBorder24);
            border24.Append(topBorder24);
            border24.Append(bottomBorder24);
            border24.Append(diagonalBorder24);

            DocumentFormat.OpenXml.Spreadsheet.Border border25 = new DocumentFormat.OpenXml.Spreadsheet.Border();
            LeftBorder leftBorder25 = new LeftBorder();

            RightBorder rightBorder25 = new RightBorder() { Style = BorderStyleValues.Medium };
            DocumentFormat.OpenXml.Spreadsheet.Color color81 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            rightBorder25.Append(color81);

            TopBorder topBorder25 = new TopBorder() { Style = BorderStyleValues.Double };
            DocumentFormat.OpenXml.Spreadsheet.Color color82 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            topBorder25.Append(color82);

            BottomBorder bottomBorder25 = new BottomBorder() { Style = BorderStyleValues.Medium };
            DocumentFormat.OpenXml.Spreadsheet.Color color83 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

            bottomBorder25.Append(color83);
            DiagonalBorder diagonalBorder25 = new DiagonalBorder();

            border25.Append(leftBorder25);
            border25.Append(rightBorder25);
            border25.Append(topBorder25);
            border25.Append(bottomBorder25);
            border25.Append(diagonalBorder25);

            borders1.Append(border1);
            borders1.Append(border2);
            borders1.Append(border3);
            borders1.Append(border4);
            borders1.Append(border5);
            borders1.Append(border6);
            borders1.Append(border7);
            borders1.Append(border8);
            borders1.Append(border9);
            borders1.Append(border10);
            borders1.Append(border11);
            borders1.Append(border12);
            borders1.Append(border13);
            borders1.Append(border14);
            borders1.Append(border15);
            borders1.Append(border16);
            borders1.Append(border17);
            borders1.Append(border18);
            borders1.Append(border19);
            borders1.Append(border20);
            borders1.Append(border21);
            borders1.Append(border22);
            borders1.Append(border23);
            borders1.Append(border24);
            borders1.Append(border25);

            CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)1U };

            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };
            Alignment alignment1 = new Alignment() { Vertical = VerticalAlignmentValues.Center };

            cellFormat1.Append(alignment1);

            cellStyleFormats1.Append(cellFormat1);

            CellFormats cellFormats1 = new CellFormats() { Count = (UInt32Value)48U };

            CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            Alignment alignment2 = new Alignment() { Vertical = VerticalAlignmentValues.Center };

            cellFormat2.Append(alignment2);

            CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)5U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment3 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

            cellFormat3.Append(alignment3);

            CellFormat cellFormat4 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)5U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)2U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment4 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

            cellFormat4.Append(alignment4);

            CellFormat cellFormat5 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)5U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)3U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment5 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

            cellFormat5.Append(alignment5);

            CellFormat cellFormat6 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment6 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

            cellFormat6.Append(alignment6);

            CellFormat cellFormat7 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)6U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)2U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment7 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

            cellFormat7.Append(alignment7);

            CellFormat cellFormat8 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)2U, FormatId = (UInt32Value)0U, ApplyBorder = true };
            Alignment alignment8 = new Alignment() { Vertical = VerticalAlignmentValues.Center };

            cellFormat8.Append(alignment8);

            CellFormat cellFormat9 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)2U, FormatId = (UInt32Value)0U, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment9 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

            cellFormat9.Append(alignment9);

            CellFormat cellFormat10 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)3U, FormatId = (UInt32Value)0U, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment10 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

            cellFormat10.Append(alignment10);

            CellFormat cellFormat11 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyBorder = true };
            Alignment alignment11 = new Alignment() { Vertical = VerticalAlignmentValues.Center };

            cellFormat11.Append(alignment11);

            CellFormat cellFormat12 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment12 = new Alignment() { Horizontal = HorizontalAlignmentValues.Right, Vertical = VerticalAlignmentValues.Center };

            cellFormat12.Append(alignment12);

            CellFormat cellFormat13 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)4U, FormatId = (UInt32Value)0U, ApplyBorder = true };
            Alignment alignment13 = new Alignment() { Vertical = VerticalAlignmentValues.Center };

            cellFormat13.Append(alignment13);

            CellFormat cellFormat14 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)5U, FormatId = (UInt32Value)0U, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment14 = new Alignment() { Horizontal = HorizontalAlignmentValues.Right, Vertical = VerticalAlignmentValues.Center };

            cellFormat14.Append(alignment14);

            CellFormat cellFormat15 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)5U, FormatId = (UInt32Value)0U, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment15 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

            cellFormat15.Append(alignment15);

            CellFormat cellFormat16 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)5U, FormatId = (UInt32Value)0U, ApplyBorder = true };
            Alignment alignment16 = new Alignment() { Vertical = VerticalAlignmentValues.Center };

            cellFormat16.Append(alignment16);

            CellFormat cellFormat17 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)6U, FormatId = (UInt32Value)0U, ApplyBorder = true };
            Alignment alignment17 = new Alignment() { Vertical = VerticalAlignmentValues.Center };

            cellFormat17.Append(alignment17);

            CellFormat cellFormat18 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)4U, FormatId = (UInt32Value)0U, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment18 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

            cellFormat18.Append(alignment18);

            CellFormat cellFormat19 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)7U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)2U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment19 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center };

            cellFormat19.Append(alignment19);

            CellFormat cellFormat20 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)13U, FormatId = (UInt32Value)0U, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment20 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

            cellFormat20.Append(alignment20);

            CellFormat cellFormat21 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)7U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)5U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment21 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

            cellFormat21.Append(alignment21);

            CellFormat cellFormat22 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)6U, FormatId = (UInt32Value)0U, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment22 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

            cellFormat22.Append(alignment22);

            CellFormat cellFormat23 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)7U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)7U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment23 = new Alignment() { Horizontal = HorizontalAlignmentValues.Right, Vertical = VerticalAlignmentValues.Center };

            cellFormat23.Append(alignment23);

            CellFormat cellFormat24 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)8U, FormatId = (UInt32Value)0U, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment24 = new Alignment() { Vertical = VerticalAlignmentValues.Center };

            cellFormat24.Append(alignment24);

            CellFormat cellFormat25 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)11U, FormatId = (UInt32Value)0U, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment25 = new Alignment() { Vertical = VerticalAlignmentValues.Center };

            cellFormat25.Append(alignment25);

            CellFormat cellFormat26 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)7U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)9U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment26 = new Alignment() { Horizontal = HorizontalAlignmentValues.Right, Vertical = VerticalAlignmentValues.Center };

            cellFormat26.Append(alignment26);

            CellFormat cellFormat27 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)10U, FormatId = (UInt32Value)0U, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment27 = new Alignment() { Vertical = VerticalAlignmentValues.Center };

            cellFormat27.Append(alignment27);

            CellFormat cellFormat28 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)12U, FormatId = (UInt32Value)0U, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment28 = new Alignment() { Vertical = VerticalAlignmentValues.Center };

            cellFormat28.Append(alignment28);

            CellFormat cellFormat29 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyAlignment = true };
            Alignment alignment29 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

            cellFormat29.Append(alignment29);

            CellFormat cellFormat30 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyAlignment = true };
            Alignment alignment30 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

            cellFormat30.Append(alignment30);

            CellFormat cellFormat31 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)4U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment31 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center };

            cellFormat31.Append(alignment31);

            CellFormat cellFormat32 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)6U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)2U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment32 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

            cellFormat32.Append(alignment32);

            CellFormat cellFormat33 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)6U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)14U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment33 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

            cellFormat33.Append(alignment33);

            CellFormat cellFormat34 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)6U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)15U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment34 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

            cellFormat34.Append(alignment34);

            CellFormat cellFormat35 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)2U, FormatId = (UInt32Value)0U, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment35 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

            cellFormat35.Append(alignment35);

            CellFormat cellFormat36 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)3U, FormatId = (UInt32Value)0U, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment36 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

            cellFormat36.Append(alignment36);

            CellFormat cellFormat37 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)7U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)2U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment37 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

            cellFormat37.Append(alignment37);

            CellFormat cellFormat38 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)7U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)3U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment38 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

            cellFormat38.Append(alignment38);

            CellFormat cellFormat39 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)8U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)2U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment39 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

            cellFormat39.Append(alignment39);

            CellFormat cellFormat40 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)16U, FormatId = (UInt32Value)0U, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment40 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

            cellFormat40.Append(alignment40);

            CellFormat cellFormat41 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)6U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)17U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment41 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

            cellFormat41.Append(alignment41);

            CellFormat cellFormat42 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)6U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)18U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment42 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

            cellFormat42.Append(alignment42);

            CellFormat cellFormat43 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)19U, FormatId = (UInt32Value)0U, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment43 = new Alignment() { Horizontal = HorizontalAlignmentValues.Right, Vertical = VerticalAlignmentValues.Center };

            cellFormat43.Append(alignment43);

            CellFormat cellFormat44 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)4U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)20U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment44 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center };

            cellFormat44.Append(alignment44);

            CellFormat cellFormat45 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)20U, FormatId = (UInt32Value)0U, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment45 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center };

            cellFormat45.Append(alignment45);

            CellFormat cellFormat46 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)21U, FormatId = (UInt32Value)0U, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment46 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center };

            cellFormat46.Append(alignment46);

            CellFormat cellFormat47 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)6U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)22U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment47 = new Alignment() { Horizontal = HorizontalAlignmentValues.Right, Vertical = VerticalAlignmentValues.Center };

            cellFormat47.Append(alignment47);

            CellFormat cellFormat48 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)6U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)23U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment48 = new Alignment() { Horizontal = HorizontalAlignmentValues.Right, Vertical = VerticalAlignmentValues.Center };

            cellFormat48.Append(alignment48);

            CellFormat cellFormat49 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)6U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)24U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment49 = new Alignment() { Horizontal = HorizontalAlignmentValues.Right, Vertical = VerticalAlignmentValues.Center };

            cellFormat49.Append(alignment49);

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
            cellFormats1.Append(cellFormat23);
            cellFormats1.Append(cellFormat24);
            cellFormats1.Append(cellFormat25);
            cellFormats1.Append(cellFormat26);
            cellFormats1.Append(cellFormat27);
            cellFormats1.Append(cellFormat28);
            cellFormats1.Append(cellFormat29);
            cellFormats1.Append(cellFormat30);
            cellFormats1.Append(cellFormat31);
            cellFormats1.Append(cellFormat32);
            cellFormats1.Append(cellFormat33);
            cellFormats1.Append(cellFormat34);
            cellFormats1.Append(cellFormat35);
            cellFormats1.Append(cellFormat36);
            cellFormats1.Append(cellFormat37);
            cellFormats1.Append(cellFormat38);
            cellFormats1.Append(cellFormat39);
            cellFormats1.Append(cellFormat40);
            cellFormats1.Append(cellFormat41);
            cellFormats1.Append(cellFormat42);
            cellFormats1.Append(cellFormat43);
            cellFormats1.Append(cellFormat44);
            cellFormats1.Append(cellFormat45);
            cellFormats1.Append(cellFormat46);
            cellFormats1.Append(cellFormat47);
            cellFormats1.Append(cellFormat48);
            cellFormats1.Append(cellFormat49);

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

            A.FontScheme fontScheme10 = new A.FontScheme() { Name = "Office" };

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

            fontScheme10.Append(majorFont1);
            fontScheme10.Append(minorFont1);

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
            themeElements1.Append(fontScheme10);
            themeElements1.Append(formatScheme1);
            A.ObjectDefaults objectDefaults1 = new A.ObjectDefaults();
            A.ExtraColorSchemeList extraColorSchemeList1 = new A.ExtraColorSchemeList();

            theme1.Append(themeElements1);
            theme1.Append(objectDefaults1);
            theme1.Append(extraColorSchemeList1);

            themePart1.Theme = theme1;
        }

        // Generates content of worksheetPart1.
        private void GenerateWorksheetPart3Content(WorksheetPart worksheetPart1)
        {
            Worksheet worksheet1 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            SheetDimension sheetDimension1 = new SheetDimension() { Reference = "A1:F16" };

            SheetViews sheetViews1 = new SheetViews();

            SheetView sheetView1 = new SheetView() { TabSelected = true, WorkbookViewId = (UInt32Value)0U };
            Selection selection1 = new Selection() { SequenceOfReferences = new ListValue<StringValue>() { InnerText = "A1:F1" } };

            sheetView1.Append(selection1);

            sheetViews1.Append(sheetView1);
            SheetFormatProperties sheetFormatProperties1 = new SheetFormatProperties() { DefaultRowHeight = 13.5D, DyDescent = 0.15D };

            Columns columns1 = new Columns();
            Column column1 = new Column() { Min = (UInt32Value)2U, Max = (UInt32Value)2U, Width = 29.75D, CustomWidth = true };
            Column column2 = new Column() { Min = (UInt32Value)3U, Max = (UInt32Value)3U, Width = 12.375D, CustomWidth = true };
            Column column3 = new Column() { Min = (UInt32Value)4U, Max = (UInt32Value)4U, Width = 13.25D, CustomWidth = true };
            Column column4 = new Column() { Min = (UInt32Value)5U, Max = (UInt32Value)5U, Width = 13D, CustomWidth = true };
            Column column5 = new Column() { Min = (UInt32Value)6U, Max = (UInt32Value)6U, Width = 15D, CustomWidth = true };

            columns1.Append(column1);
            columns1.Append(column2);
            columns1.Append(column3);
            columns1.Append(column4);
            columns1.Append(column5);

            SheetData sheetData1 = new SheetData();

            Row row1 = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:6" }, Height = 22.5D, DyDescent = 0.15D };

            Cell cell1 = new Cell() { CellReference = "A1", StyleIndex = (UInt32Value)27U, DataType = CellValues.SharedString };
            CellValue cellValue1 = new CellValue();
            cellValue1.Text = "0";

            cell1.Append(cellValue1);
            Cell cell2 = new Cell() { CellReference = "B1", StyleIndex = (UInt32Value)28U };
            Cell cell3 = new Cell() { CellReference = "C1", StyleIndex = (UInt32Value)28U };
            Cell cell4 = new Cell() { CellReference = "D1", StyleIndex = (UInt32Value)28U };
            Cell cell5 = new Cell() { CellReference = "E1", StyleIndex = (UInt32Value)28U };
            Cell cell6 = new Cell() { CellReference = "F1", StyleIndex = (UInt32Value)28U };

            row1.Append(cell1);
            row1.Append(cell2);
            row1.Append(cell3);
            row1.Append(cell4);
            row1.Append(cell5);
            row1.Append(cell6);

            Row row2 = new Row() { RowIndex = (UInt32Value)2U, Spans = new ListValue<StringValue>() { InnerText = "1:6" }, Height = 24.95D, CustomHeight = true, ThickBot = true, DyDescent = 0.2D };

            Cell cell7 = new Cell() { CellReference = "A2", StyleIndex = (UInt32Value)10U, DataType = CellValues.SharedString };
            CellValue cellValue2 = new CellValue();
            cellValue2.Text = "1";

            cell7.Append(cellValue2);

            Cell cell8 = new Cell() { CellReference = "B2", StyleIndex = (UInt32Value)29U, DataType = CellValues.SharedString };
            CellValue cellValue3 = new CellValue();
            cellValue3.Text = "7";

            cell8.Append(cellValue3);
            Cell cell9 = new Cell() { CellReference = "C2", StyleIndex = (UInt32Value)29U };
            Cell cell10 = new Cell() { CellReference = "D2", StyleIndex = (UInt32Value)29U };
            Cell cell11 = new Cell() { CellReference = "E2", StyleIndex = (UInt32Value)29U };
            Cell cell12 = new Cell() { CellReference = "F2", StyleIndex = (UInt32Value)29U };

            row2.Append(cell7);
            row2.Append(cell8);
            row2.Append(cell9);
            row2.Append(cell10);
            row2.Append(cell11);
            row2.Append(cell12);

            Row row3 = new Row() { RowIndex = (UInt32Value)3U, Spans = new ListValue<StringValue>() { InnerText = "1:6" }, Height = 24.95D, CustomHeight = true, ThickBot = true, DyDescent = 0.2D };

            Cell cell13 = new Cell() { CellReference = "A3", StyleIndex = (UInt32Value)41U, DataType = CellValues.SharedString };
            CellValue cellValue4 = new CellValue();
            cellValue4.Text = "14";

            cell13.Append(cellValue4);

            Cell cell14 = new Cell() { CellReference = "B3", StyleIndex = (UInt32Value)42U };
            CellValue cellValue5 = new CellValue();
            cellValue5.Text = "10000";

            cell14.Append(cellValue5);
            Cell cell15 = new Cell() { CellReference = "C3", StyleIndex = (UInt32Value)43U };
            Cell cell16 = new Cell() { CellReference = "D3", StyleIndex = (UInt32Value)43U };
            Cell cell17 = new Cell() { CellReference = "E3", StyleIndex = (UInt32Value)43U };
            Cell cell18 = new Cell() { CellReference = "F3", StyleIndex = (UInt32Value)44U };

            row3.Append(cell13);
            row3.Append(cell14);
            row3.Append(cell15);
            row3.Append(cell16);
            row3.Append(cell17);
            row3.Append(cell18);

            Row row4 = new Row() { RowIndex = (UInt32Value)4U, Spans = new ListValue<StringValue>() { InnerText = "1:6" }, Height = 24.95D, CustomHeight = true, ThickTop = true, DyDescent = 0.15D };

            Cell cell19 = new Cell() { CellReference = "A4", StyleIndex = (UInt32Value)38U, DataType = CellValues.SharedString };
            CellValue cellValue6 = new CellValue();
            cellValue6.Text = "8";

            cell19.Append(cellValue6);

            Cell cell20 = new Cell() { CellReference = "B4", StyleIndex = (UInt32Value)39U, DataType = CellValues.SharedString };
            CellValue cellValue7 = new CellValue();
            cellValue7.Text = "9";

            cell20.Append(cellValue7);
            Cell cell21 = new Cell() { CellReference = "C4", StyleIndex = (UInt32Value)39U };
            Cell cell22 = new Cell() { CellReference = "D4", StyleIndex = (UInt32Value)39U };
            Cell cell23 = new Cell() { CellReference = "E4", StyleIndex = (UInt32Value)39U };
            Cell cell24 = new Cell() { CellReference = "F4", StyleIndex = (UInt32Value)40U };

            row4.Append(cell19);
            row4.Append(cell20);
            row4.Append(cell21);
            row4.Append(cell22);
            row4.Append(cell23);
            row4.Append(cell24);

            Row row5 = new Row() { RowIndex = (UInt32Value)5U, Spans = new ListValue<StringValue>() { InnerText = "1:6" }, Height = 20.100000000000001D, CustomHeight = true, DyDescent = 0.15D };

            Cell cell25 = new Cell() { CellReference = "A5", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
            CellValue cellValue8 = new CellValue();
            cellValue8.Text = "2";

            cell25.Append(cellValue8);

            Cell cell26 = new Cell() { CellReference = "B5", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue9 = new CellValue();
            cellValue9.Text = "18";

            cell26.Append(cellValue9);

            Cell cell27 = new Cell() { CellReference = "C5", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue10 = new CellValue();
            cellValue10.Text = "3";

            cell27.Append(cellValue10);

            Cell cell28 = new Cell() { CellReference = "D5", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue11 = new CellValue();
            cellValue11.Text = "4";

            cell28.Append(cellValue11);

            Cell cell29 = new Cell() { CellReference = "E5", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue12 = new CellValue();
            cellValue12.Text = "5";

            cell29.Append(cellValue12);

            Cell cell30 = new Cell() { CellReference = "F5", StyleIndex = (UInt32Value)3U, DataType = CellValues.SharedString };
            CellValue cellValue13 = new CellValue();
            cellValue13.Text = "6";

            cell30.Append(cellValue13);

            row5.Append(cell25);
            row5.Append(cell26);
            row5.Append(cell27);
            row5.Append(cell28);
            row5.Append(cell29);
            row5.Append(cell30);

            Row row6 = new Row() { RowIndex = (UInt32Value)6U, Spans = new ListValue<StringValue>() { InnerText = "1:6" }, Height = 20.100000000000001D, CustomHeight = true, DyDescent = 0.15D };

            Cell cell31 = new Cell() { CellReference = "A6", StyleIndex = (UInt32Value)4U };
            CellValue cellValue14 = new CellValue();
            cellValue14.Text = "1";

            cell31.Append(cellValue14);

            Cell cell32 = new Cell() { CellReference = "B6", StyleIndex = (UInt32Value)6U, DataType = CellValues.SharedString };
            CellValue cellValue15 = new CellValue();
            cellValue15.Text = "7";

            cell32.Append(cellValue15);

            Cell cell33 = new Cell() { CellReference = "C6", StyleIndex = (UInt32Value)7U };
            CellValue cellValue16 = new CellValue();
            cellValue16.Text = "10";

            cell33.Append(cellValue16);

            Cell cell34 = new Cell() { CellReference = "D6", StyleIndex = (UInt32Value)7U, DataType = CellValues.SharedString };
            CellValue cellValue17 = new CellValue();
            cellValue17.Text = "10";

            cell34.Append(cellValue17);

            Cell cell35 = new Cell() { CellReference = "E6", StyleIndex = (UInt32Value)7U, DataType = CellValues.SharedString };
            CellValue cellValue18 = new CellValue();
            cellValue18.Text = "11";

            cell35.Append(cellValue18);

            Cell cell36 = new Cell() { CellReference = "F6", StyleIndex = (UInt32Value)8U, DataType = CellValues.SharedString };
            CellValue cellValue19 = new CellValue();
            cellValue19.Text = "12";

            cell36.Append(cellValue19);

            row6.Append(cell31);
            row6.Append(cell32);
            row6.Append(cell33);
            row6.Append(cell34);
            row6.Append(cell35);
            row6.Append(cell36);

            Row row7 = new Row() { RowIndex = (UInt32Value)7U, Spans = new ListValue<StringValue>() { InnerText = "1:6" }, Height = 20.100000000000001D, CustomHeight = true, ThickBot = true, DyDescent = 0.2D };
            Cell cell37 = new Cell() { CellReference = "A7", StyleIndex = (UInt32Value)11U };

            Cell cell38 = new Cell() { CellReference = "B7", StyleIndex = (UInt32Value)12U, DataType = CellValues.SharedString };
            CellValue cellValue20 = new CellValue();
            cellValue20.Text = "13";

            cell38.Append(cellValue20);

            Cell cell39 = new Cell() { CellReference = "C7", StyleIndex = (UInt32Value)13U };
            CellValue cellValue21 = new CellValue();
            cellValue21.Text = "10";

            cell39.Append(cellValue21);
            Cell cell40 = new Cell() { CellReference = "D7", StyleIndex = (UInt32Value)14U };
            Cell cell41 = new Cell() { CellReference = "E7", StyleIndex = (UInt32Value)14U };
            Cell cell42 = new Cell() { CellReference = "F7", StyleIndex = (UInt32Value)15U };

            row7.Append(cell37);
            row7.Append(cell38);
            row7.Append(cell39);
            row7.Append(cell40);
            row7.Append(cell41);
            row7.Append(cell42);

            Row row8 = new Row() { RowIndex = (UInt32Value)8U, Spans = new ListValue<StringValue>() { InnerText = "1:6" }, Height = 20.100000000000001D, CustomHeight = true, ThickTop = true, DyDescent = 0.15D };

            Cell cell43 = new Cell() { CellReference = "A8", StyleIndex = (UInt32Value)18U, DataType = CellValues.SharedString };
            CellValue cellValue22 = new CellValue();
            cellValue22.Text = "15";

            cell43.Append(cellValue22);

            Cell cell44 = new Cell() { CellReference = "B8", StyleIndex = (UInt32Value)31U, DataType = CellValues.SharedString };
            CellValue cellValue23 = new CellValue();
            cellValue23.Text = "16";

            cell44.Append(cellValue23);
            Cell cell45 = new Cell() { CellReference = "C8", StyleIndex = (UInt32Value)31U };
            Cell cell46 = new Cell() { CellReference = "D8", StyleIndex = (UInt32Value)31U };
            Cell cell47 = new Cell() { CellReference = "E8", StyleIndex = (UInt32Value)31U };
            Cell cell48 = new Cell() { CellReference = "F8", StyleIndex = (UInt32Value)32U };

            row8.Append(cell43);
            row8.Append(cell44);
            row8.Append(cell45);
            row8.Append(cell46);
            row8.Append(cell47);
            row8.Append(cell48);

            Row row9 = new Row() { RowIndex = (UInt32Value)9U, Spans = new ListValue<StringValue>() { InnerText = "1:6" }, Height = 20.100000000000001D, CustomHeight = true, DyDescent = 0.15D };
            Cell cell49 = new Cell() { CellReference = "A9", StyleIndex = (UInt32Value)4U };

            Cell cell50 = new Cell() { CellReference = "B9", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue24 = new CellValue();
            cellValue24.Text = "26";

            cell50.Append(cellValue24);

            Cell cell51 = new Cell() { CellReference = "C9", StyleIndex = (UInt32Value)30U, DataType = CellValues.SharedString };
            CellValue cellValue25 = new CellValue();
            cellValue25.Text = "23";

            cell51.Append(cellValue25);
            Cell cell52 = new Cell() { CellReference = "D9", StyleIndex = (UInt32Value)33U };

            Cell cell53 = new Cell() { CellReference = "E9", StyleIndex = (UInt32Value)30U, DataType = CellValues.SharedString };
            CellValue cellValue26 = new CellValue();
            cellValue26.Text = "24";

            cell53.Append(cellValue26);
            Cell cell54 = new Cell() { CellReference = "F9", StyleIndex = (UInt32Value)34U };

            row9.Append(cell49);
            row9.Append(cell50);
            row9.Append(cell51);
            row9.Append(cell52);
            row9.Append(cell53);
            row9.Append(cell54);

            Row row10 = new Row() { RowIndex = (UInt32Value)10U, Spans = new ListValue<StringValue>() { InnerText = "1:6" }, Height = 20.100000000000001D, CustomHeight = true, DyDescent = 0.15D };

            Cell cell55 = new Cell() { CellReference = "A10", StyleIndex = (UInt32Value)4U };
            CellValue cellValue27 = new CellValue();
            cellValue27.Text = "1";

            cell55.Append(cellValue27);

            Cell cell56 = new Cell() { CellReference = "B10", StyleIndex = (UInt32Value)17U, DataType = CellValues.SharedString };
            CellValue cellValue28 = new CellValue();
            cellValue28.Text = "21";

            cell56.Append(cellValue28);

            Cell cell57 = new Cell() { CellReference = "C10", StyleIndex = (UInt32Value)35U, DataType = CellValues.SharedString };
            CellValue cellValue29 = new CellValue();
            cellValue29.Text = "22";

            cell57.Append(cellValue29);
            Cell cell58 = new Cell() { CellReference = "D10", StyleIndex = (UInt32Value)35U };

            Cell cell59 = new Cell() { CellReference = "E10", StyleIndex = (UInt32Value)35U };
            CellValue cellValue30 = new CellValue();
            cellValue30.Text = "10";

            cell59.Append(cellValue30);
            Cell cell60 = new Cell() { CellReference = "F10", StyleIndex = (UInt32Value)36U };

            row10.Append(cell55);
            row10.Append(cell56);
            row10.Append(cell57);
            row10.Append(cell58);
            row10.Append(cell59);
            row10.Append(cell60);

            Row row11 = new Row() { RowIndex = (UInt32Value)11U, Spans = new ListValue<StringValue>() { InnerText = "1:6" }, Height = 20.100000000000001D, CustomHeight = true, DyDescent = 0.15D };
            Cell cell61 = new Cell() { CellReference = "A11", StyleIndex = (UInt32Value)9U };

            Cell cell62 = new Cell() { CellReference = "B11", StyleIndex = (UInt32Value)21U, DataType = CellValues.SharedString };
            CellValue cellValue31 = new CellValue();
            cellValue31.Text = "25";

            cell62.Append(cellValue31);
            Cell cell63 = new Cell() { CellReference = "C11", StyleIndex = (UInt32Value)22U };
            Cell cell64 = new Cell() { CellReference = "D11", StyleIndex = (UInt32Value)23U };

            Cell cell65 = new Cell() { CellReference = "E11", StyleIndex = (UInt32Value)37U };
            CellValue cellValue32 = new CellValue();
            cellValue32.Text = "10";

            cell65.Append(cellValue32);
            Cell cell66 = new Cell() { CellReference = "F11", StyleIndex = (UInt32Value)36U };

            row11.Append(cell61);
            row11.Append(cell62);
            row11.Append(cell63);
            row11.Append(cell64);
            row11.Append(cell65);
            row11.Append(cell66);

            Row row12 = new Row() { RowIndex = (UInt32Value)12U, Spans = new ListValue<StringValue>() { InnerText = "1:6" }, Height = 20.100000000000001D, CustomHeight = true, DyDescent = 0.15D };

            Cell cell67 = new Cell() { CellReference = "A12", StyleIndex = (UInt32Value)4U, DataType = CellValues.SharedString };
            CellValue cellValue33 = new CellValue();
            cellValue33.Text = "17";

            cell67.Append(cellValue33);

            Cell cell68 = new Cell() { CellReference = "B12", StyleIndex = (UInt32Value)30U, DataType = CellValues.SharedString };
            CellValue cellValue34 = new CellValue();
            cellValue34.Text = "20";

            cell68.Append(cellValue34);
            Cell cell69 = new Cell() { CellReference = "C12", StyleIndex = (UInt32Value)33U };
            Cell cell70 = new Cell() { CellReference = "D12", StyleIndex = (UInt32Value)33U };
            Cell cell71 = new Cell() { CellReference = "E12", StyleIndex = (UInt32Value)33U };
            Cell cell72 = new Cell() { CellReference = "F12", StyleIndex = (UInt32Value)34U };

            row12.Append(cell67);
            row12.Append(cell68);
            row12.Append(cell69);
            row12.Append(cell70);
            row12.Append(cell71);
            row12.Append(cell72);

            Row row13 = new Row() { RowIndex = (UInt32Value)13U, Spans = new ListValue<StringValue>() { InnerText = "1:6" }, Height = 20.100000000000001D, CustomHeight = true, DyDescent = 0.15D };
            Cell cell73 = new Cell() { CellReference = "A13", StyleIndex = (UInt32Value)4U };

            Cell cell74 = new Cell() { CellReference = "B13", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue35 = new CellValue();
            cellValue35.Text = "27";

            cell74.Append(cellValue35);

            Cell cell75 = new Cell() { CellReference = "C13", StyleIndex = (UInt32Value)30U, DataType = CellValues.SharedString };
            CellValue cellValue36 = new CellValue();
            cellValue36.Text = "23";

            cell75.Append(cellValue36);
            Cell cell76 = new Cell() { CellReference = "D13", StyleIndex = (UInt32Value)33U };

            Cell cell77 = new Cell() { CellReference = "E13", StyleIndex = (UInt32Value)30U, DataType = CellValues.SharedString };
            CellValue cellValue37 = new CellValue();
            cellValue37.Text = "24";

            cell77.Append(cellValue37);
            Cell cell78 = new Cell() { CellReference = "F13", StyleIndex = (UInt32Value)34U };

            row13.Append(cell73);
            row13.Append(cell74);
            row13.Append(cell75);
            row13.Append(cell76);
            row13.Append(cell77);
            row13.Append(cell78);

            Row row14 = new Row() { RowIndex = (UInt32Value)14U, Spans = new ListValue<StringValue>() { InnerText = "1:6" }, Height = 20.100000000000001D, CustomHeight = true, DyDescent = 0.15D };

            Cell cell79 = new Cell() { CellReference = "A14", StyleIndex = (UInt32Value)4U };
            CellValue cellValue38 = new CellValue();
            cellValue38.Text = "1";

            cell79.Append(cellValue38);

            Cell cell80 = new Cell() { CellReference = "B14", StyleIndex = (UInt32Value)17U, DataType = CellValues.SharedString };
            CellValue cellValue39 = new CellValue();
            cellValue39.Text = "22";

            cell80.Append(cellValue39);

            Cell cell81 = new Cell() { CellReference = "C14", StyleIndex = (UInt32Value)35U, DataType = CellValues.SharedString };
            CellValue cellValue40 = new CellValue();
            cellValue40.Text = "22";

            cell81.Append(cellValue40);
            Cell cell82 = new Cell() { CellReference = "D14", StyleIndex = (UInt32Value)33U };

            Cell cell83 = new Cell() { CellReference = "E14", StyleIndex = (UInt32Value)35U };
            CellValue cellValue41 = new CellValue();
            cellValue41.Text = "10";

            cell83.Append(cellValue41);
            Cell cell84 = new Cell() { CellReference = "F14", StyleIndex = (UInt32Value)34U };

            row14.Append(cell79);
            row14.Append(cell80);
            row14.Append(cell81);
            row14.Append(cell82);
            row14.Append(cell83);
            row14.Append(cell84);

            Row row15 = new Row() { RowIndex = (UInt32Value)15U, Spans = new ListValue<StringValue>() { InnerText = "1:6" }, Height = 20.100000000000001D, CustomHeight = true, ThickBot = true, DyDescent = 0.2D };
            Cell cell85 = new Cell() { CellReference = "A15", StyleIndex = (UInt32Value)16U };

            Cell cell86 = new Cell() { CellReference = "B15", StyleIndex = (UInt32Value)24U, DataType = CellValues.SharedString };
            CellValue cellValue42 = new CellValue();
            cellValue42.Text = "19";

            cell86.Append(cellValue42);
            Cell cell87 = new Cell() { CellReference = "C15", StyleIndex = (UInt32Value)25U };
            Cell cell88 = new Cell() { CellReference = "D15", StyleIndex = (UInt32Value)26U };

            Cell cell89 = new Cell() { CellReference = "E15", StyleIndex = (UInt32Value)19U };
            CellValue cellValue43 = new CellValue();
            cellValue43.Text = "10";

            cell89.Append(cellValue43);
            Cell cell90 = new Cell() { CellReference = "F15", StyleIndex = (UInt32Value)20U };

            row15.Append(cell85);
            row15.Append(cell86);
            row15.Append(cell87);
            row15.Append(cell88);
            row15.Append(cell89);
            row15.Append(cell90);

            Row row16 = new Row() { RowIndex = (UInt32Value)16U, Spans = new ListValue<StringValue>() { InnerText = "1:6" }, Height = 20.100000000000001D, CustomHeight = true, ThickTop = true, ThickBot = true, DyDescent = 0.2D };

            Cell cell91 = new Cell() { CellReference = "A16", StyleIndex = (UInt32Value)45U, DataType = CellValues.SharedString };
            CellValue cellValue44 = new CellValue();
            cellValue44.Text = "28";

            cell91.Append(cellValue44);
            Cell cell92 = new Cell() { CellReference = "B16", StyleIndex = (UInt32Value)46U };
            Cell cell93 = new Cell() { CellReference = "C16", StyleIndex = (UInt32Value)46U };
            Cell cell94 = new Cell() { CellReference = "D16", StyleIndex = (UInt32Value)46U };
            Cell cell95 = new Cell() { CellReference = "E16", StyleIndex = (UInt32Value)46U };
            Cell cell96 = new Cell() { CellReference = "F16", StyleIndex = (UInt32Value)47U };

            row16.Append(cell91);
            row16.Append(cell92);
            row16.Append(cell93);
            row16.Append(cell94);
            row16.Append(cell95);
            row16.Append(cell96);

            sheetData1.Append(row1);
            sheetData1.Append(row2);
            sheetData1.Append(row3);
            sheetData1.Append(row4);
            sheetData1.Append(row5);
            sheetData1.Append(row6);
            sheetData1.Append(row7);
            sheetData1.Append(row8);
            sheetData1.Append(row9);
            sheetData1.Append(row10);
            sheetData1.Append(row11);
            sheetData1.Append(row12);
            sheetData1.Append(row13);
            sheetData1.Append(row14);
            sheetData1.Append(row15);
            sheetData1.Append(row16);

            MergeCells mergeCells1 = new MergeCells() { Count = (UInt32Value)19U };
            MergeCell mergeCell1 = new MergeCell() { Reference = "B3:F3" };
            MergeCell mergeCell2 = new MergeCell() { Reference = "B8:F8" };
            MergeCell mergeCell3 = new MergeCell() { Reference = "B12:F12" };
            MergeCell mergeCell4 = new MergeCell() { Reference = "C9:D9" };
            MergeCell mergeCell5 = new MergeCell() { Reference = "E9:F9" };
            MergeCell mergeCell6 = new MergeCell() { Reference = "C10:D10" };
            MergeCell mergeCell7 = new MergeCell() { Reference = "E10:F10" };
            MergeCell mergeCell8 = new MergeCell() { Reference = "E11:F11" };
            MergeCell mergeCell9 = new MergeCell() { Reference = "C13:D13" };
            MergeCell mergeCell10 = new MergeCell() { Reference = "E13:F13" };
            MergeCell mergeCell11 = new MergeCell() { Reference = "C14:D14" };
            MergeCell mergeCell12 = new MergeCell() { Reference = "E14:F14" };
            MergeCell mergeCell13 = new MergeCell() { Reference = "A16:F16" };
            MergeCell mergeCell14 = new MergeCell() { Reference = "E15:F15" };
            MergeCell mergeCell15 = new MergeCell() { Reference = "B11:D11" };
            MergeCell mergeCell16 = new MergeCell() { Reference = "B15:D15" };
            MergeCell mergeCell17 = new MergeCell() { Reference = "A1:F1" };
            MergeCell mergeCell18 = new MergeCell() { Reference = "B2:F2" };
            MergeCell mergeCell19 = new MergeCell() { Reference = "B4:F4" };

            mergeCells1.Append(mergeCell1);
            mergeCells1.Append(mergeCell2);
            mergeCells1.Append(mergeCell3);
            mergeCells1.Append(mergeCell4);
            mergeCells1.Append(mergeCell5);
            mergeCells1.Append(mergeCell6);
            mergeCells1.Append(mergeCell7);
            mergeCells1.Append(mergeCell8);
            mergeCells1.Append(mergeCell9);
            mergeCells1.Append(mergeCell10);
            mergeCells1.Append(mergeCell11);
            mergeCells1.Append(mergeCell12);
            mergeCells1.Append(mergeCell13);
            mergeCells1.Append(mergeCell14);
            mergeCells1.Append(mergeCell15);
            mergeCells1.Append(mergeCell16);
            mergeCells1.Append(mergeCell17);
            mergeCells1.Append(mergeCell18);
            mergeCells1.Append(mergeCell19);
            PhoneticProperties phoneticProperties1 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };
            PageMargins pageMargins1 = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            PageSetup pageSetup1 = new PageSetup() { PaperSize = (UInt32Value)9U, Orientation = OrientationValues.Landscape, Id = "rId1" };

            worksheet1.Append(sheetDimension1);
            worksheet1.Append(sheetViews1);
            worksheet1.Append(sheetFormatProperties1);
            worksheet1.Append(columns1);
            worksheet1.Append(sheetData1);
            worksheet1.Append(mergeCells1);
            worksheet1.Append(phoneticProperties1);
            worksheet1.Append(pageMargins1);
            worksheet1.Append(pageSetup1);

            worksheetPart1.Worksheet = worksheet1;
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
            SharedStringTable sharedStringTable1 = new SharedStringTable() { Count = (UInt32Value)34U, UniqueCount = (UInt32Value)29U };

            SharedStringItem sharedStringItem1 = new SharedStringItem();
            Text text1 = new Text();
            text1.Text = "HNCDI院外课题经费明细表";
            PhoneticProperties phoneticProperties2 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

            sharedStringItem1.Append(text1);
            sharedStringItem1.Append(phoneticProperties2);

            SharedStringItem sharedStringItem2 = new SharedStringItem();
            Text text2 = new Text();
            text2.Text = "项目：";
            PhoneticProperties phoneticProperties3 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

            sharedStringItem2.Append(text2);
            sharedStringItem2.Append(phoneticProperties3);

            SharedStringItem sharedStringItem3 = new SharedStringItem();
            Text text3 = new Text();
            text3.Text = "序号";
            PhoneticProperties phoneticProperties4 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

            sharedStringItem3.Append(text3);
            sharedStringItem3.Append(phoneticProperties4);

            SharedStringItem sharedStringItem4 = new SharedStringItem();
            Text text4 = new Text();
            text4.Text = "金额(万元)";
            PhoneticProperties phoneticProperties5 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

            sharedStringItem4.Append(text4);
            sharedStringItem4.Append(phoneticProperties5);

            SharedStringItem sharedStringItem5 = new SharedStringItem();
            Text text5 = new Text();
            text5.Text = "日期";
            PhoneticProperties phoneticProperties6 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

            sharedStringItem5.Append(text5);
            sharedStringItem5.Append(phoneticProperties6);

            SharedStringItem sharedStringItem6 = new SharedStringItem();
            Text text6 = new Text();
            text6.Text = "经手人";
            PhoneticProperties phoneticProperties7 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

            sharedStringItem6.Append(text6);
            sharedStringItem6.Append(phoneticProperties7);

            SharedStringItem sharedStringItem7 = new SharedStringItem();
            Text text7 = new Text();
            text7.Text = "经手项目负责人";
            PhoneticProperties phoneticProperties8 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

            sharedStringItem7.Append(text7);
            sharedStringItem7.Append(phoneticProperties8);

            SharedStringItem sharedStringItem8 = new SharedStringItem();
            Text text8 = new Text();
            text8.Text = "ADBC";
            PhoneticProperties phoneticProperties9 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

            sharedStringItem8.Append(text8);
            sharedStringItem8.Append(phoneticProperties9);

            SharedStringItem sharedStringItem9 = new SharedStringItem();
            Text text9 = new Text();
            text9.Text = "一";
            PhoneticProperties phoneticProperties10 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

            sharedStringItem9.Append(text9);
            sharedStringItem9.Append(phoneticProperties10);

            SharedStringItem sharedStringItem10 = new SharedStringItem();
            Text text10 = new Text();
            text10.Text = "到账";
            PhoneticProperties phoneticProperties11 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

            sharedStringItem10.Append(text10);
            sharedStringItem10.Append(phoneticProperties11);

            SharedStringItem sharedStringItem11 = new SharedStringItem();
            Text text11 = new Text();
            text11.Text = "2010.1.1";
            PhoneticProperties phoneticProperties12 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

            sharedStringItem11.Append(text11);
            sharedStringItem11.Append(phoneticProperties12);

            SharedStringItem sharedStringItem12 = new SharedStringItem();
            Text text12 = new Text();
            text12.Text = "a";
            PhoneticProperties phoneticProperties13 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

            sharedStringItem12.Append(text12);
            sharedStringItem12.Append(phoneticProperties13);

            SharedStringItem sharedStringItem13 = new SharedStringItem();
            Text text13 = new Text();
            text13.Text = "b";
            PhoneticProperties phoneticProperties14 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

            sharedStringItem13.Append(text13);
            sharedStringItem13.Append(phoneticProperties14);

            SharedStringItem sharedStringItem14 = new SharedStringItem();
            Text text14 = new Text();
            text14.Text = "小计";
            PhoneticProperties phoneticProperties15 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

            sharedStringItem14.Append(text14);
            sharedStringItem14.Append(phoneticProperties15);

            SharedStringItem sharedStringItem15 = new SharedStringItem();
            Text text15 = new Text();
            text15.Text = "合同总额";
            PhoneticProperties phoneticProperties16 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

            sharedStringItem15.Append(text15);
            sharedStringItem15.Append(phoneticProperties16);

            SharedStringItem sharedStringItem16 = new SharedStringItem();
            Text text16 = new Text();
            text16.Text = "二";
            PhoneticProperties phoneticProperties17 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

            sharedStringItem16.Append(text16);
            sharedStringItem16.Append(phoneticProperties17);

            SharedStringItem sharedStringItem17 = new SharedStringItem();
            Text text17 = new Text();
            text17.Text = "收入总览";
            PhoneticProperties phoneticProperties18 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

            sharedStringItem17.Append(text17);
            sharedStringItem17.Append(phoneticProperties18);

            SharedStringItem sharedStringItem18 = new SharedStringItem();
            Text text18 = new Text();
            text18.Text = "三";
            PhoneticProperties phoneticProperties19 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

            sharedStringItem18.Append(text18);
            sharedStringItem18.Append(phoneticProperties19);

            SharedStringItem sharedStringItem19 = new SharedStringItem();
            Text text19 = new Text();
            text19.Text = "甲方/乙方";
            PhoneticProperties phoneticProperties20 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

            sharedStringItem19.Append(text19);
            sharedStringItem19.Append(phoneticProperties20);

            SharedStringItem sharedStringItem20 = new SharedStringItem();
            Text text20 = new Text();
            text20.Text = "小计";
            PhoneticProperties phoneticProperties21 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

            sharedStringItem20.Append(text20);
            sharedStringItem20.Append(phoneticProperties21);

            SharedStringItem sharedStringItem21 = new SharedStringItem();
            Text text21 = new Text();
            text21.Text = "支出总览";
            PhoneticProperties phoneticProperties22 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

            sharedStringItem21.Append(text21);
            sharedStringItem21.Append(phoneticProperties22);

            SharedStringItem sharedStringItem22 = new SharedStringItem();
            Text text22 = new Text();
            text22.Text = "ADBC";
            PhoneticProperties phoneticProperties23 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

            sharedStringItem22.Append(text22);
            sharedStringItem22.Append(phoneticProperties23);

            SharedStringItem sharedStringItem23 = new SharedStringItem();
            Text text23 = new Text();
            text23.Text = "b";
            PhoneticProperties phoneticProperties24 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

            sharedStringItem23.Append(text23);
            sharedStringItem23.Append(phoneticProperties24);

            SharedStringItem sharedStringItem24 = new SharedStringItem();
            Text text24 = new Text();
            text24.Text = "经费类型";
            PhoneticProperties phoneticProperties25 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

            sharedStringItem24.Append(text24);
            sharedStringItem24.Append(phoneticProperties25);

            SharedStringItem sharedStringItem25 = new SharedStringItem();
            Text text25 = new Text();
            text25.Text = "金额";
            PhoneticProperties phoneticProperties26 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

            sharedStringItem25.Append(text25);
            sharedStringItem25.Append(phoneticProperties26);

            SharedStringItem sharedStringItem26 = new SharedStringItem();
            Text text26 = new Text();
            text26.Text = "小计";
            PhoneticProperties phoneticProperties27 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

            sharedStringItem26.Append(text26);
            sharedStringItem26.Append(phoneticProperties27);

            SharedStringItem sharedStringItem27 = new SharedStringItem();
            Text text27 = new Text();
            text27.Text = "甲方";
            PhoneticProperties phoneticProperties28 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

            sharedStringItem27.Append(text27);
            sharedStringItem27.Append(phoneticProperties28);

            SharedStringItem sharedStringItem28 = new SharedStringItem();
            Text text28 = new Text();
            text28.Text = "乙方";
            PhoneticProperties phoneticProperties29 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

            sharedStringItem28.Append(text28);
            sharedStringItem28.Append(phoneticProperties29);

            SharedStringItem sharedStringItem29 = new SharedStringItem();
            Text text29 = new Text();
            text29.Text = "收入合计：10万元，支出合计：5万元，结余：5万元";
            PhoneticProperties phoneticProperties30 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

            sharedStringItem29.Append(text29);
            sharedStringItem29.Append(phoneticProperties30);

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

            sharedStringTablePart1.SharedStringTable = sharedStringTable1;
        }

        private void SetPackageProperties(OpenXmlPackage document)
        {
            document.PackageProperties.Creator = "蒋惠林";
            document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2014-07-24T07:34:22Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2014-08-02T08:13:13Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.LastModifiedBy = "蒋惠林";
            document.PackageProperties.LastPrinted = System.Xml.XmlConvert.ToDateTime("2014-08-01T10:18:12Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
        }

        #region Binary Data

        private string spreadsheetPrinterSettingsPart1Data = "0VMBkPOBIABPAG4AZQBOAG8AdABlACAAMgAwADEAMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEEAwbcACwDAy8BAAIACQCaCzQIZAABAA8AWAICAAEAWAIDAAEAQQA0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAAAAAABAAAAAgAAAAEAAAD/////AAAAAAAAAAAAAAAAAAAAAERJTlUiANAALAMAAMKskFEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABQAAAAAAAQABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADQAAAAU01USgAAAAAQAMAAUwBlAG4AZAAgAFQAbwAgAE0AaQBjAHIAbwBzAG8AZgB0ACAATwBuAGUATgBvAHQAZQAgADIAMAAxADAAIABEAHIAaQB2AGUAcgAAAFJFU0RMTABVbmlyZXNETEwAUGFwZXJTaXplAEE0AE9yaWVudGF0aW9uAExBTkRTQ0FQRV9DQzI3MABSZXNvbHV0aW9uAERQSTYwMABDb2xvck1vZGUAMjRicHAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA";

        private System.IO.Stream GetBinaryDataStream(string base64String)
        {
            return new System.IO.MemoryStream(System.Convert.FromBase64String(base64String));
        }

        #endregion Binary Data
    }
}