using System;
using System.Collections.Generic;
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
using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using X15ac = DocumentFormat.OpenXml.Office2013.ExcelAc;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using A = DocumentFormat.OpenXml.Drawing;
using Thm15 = DocumentFormat.OpenXml.Office2013.Theme;


namespace HNCDIExternalProjectManage
{
	/// <summary>
	/// MoneyDetailYear.xaml 的交互逻辑
	/// </summary>
	public partial class MoneyDetailYear : Window
	{
		private DataClassesProjectClassifyDataContext dataContext;
		private List<int> projectIDs;
		private List<int> projectClassifyIDs;
		private FileInfo fileToCreate; //要创建的文件
		private int _year;
		private int totalRows = 0;
		private int totalMergeCells = 0;
		private string lastCellName;
		private int projectNo = 0; //项目计数
		private DateTime startDate, endDate;
		private decimal toAccountTotal = 0; //入账总计
		private decimal payforTotal = 0; //支付外协总计
		private decimal reimburseTotal = 0; //课题组报支总计
        private string department;
		public MoneyDetailYear()
		{
			this.InitializeComponent();
			
			// 在此点之下插入创建对象所需的代码。
		}

		private void buttonRun_Click(object sender, RoutedEventArgs e)
		{
			if (string.IsNullOrEmpty(textboxYear.Text))
			{
				MessageBox.Show("年度数字不能为空！", "错误");
				return;
			}
			try
			{
				_year = Convert.ToInt32(textboxYear.Text.Trim());
				if (_year < 1960)
				{
					MessageBox.Show("年度太早，无效", "错误");
					return;
				}
				if (_year > DateTime.Now.Year)
				{
					MessageBox.Show("年度不能晚于今年", "错误");
					return;
				}
				startDate = new DateTime(_year, 1, 1);
				endDate = new DateTime(_year, 12, 31);
				if (dataContext == null)
				{
					dataContext = new DataClassesProjectClassifyDataContext();
				}
                department = textboxDepartment.Text.Trim();
				SetProjectIDsAndProjectClassifyIDs();
				SetLaseCellNameAndTotalMergeCells();
				try
				{
					Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog();
					saveFileDialog.Filter = "Execl files (*.xlsx)|*.xlsx";
					saveFileDialog.FilterIndex = 0;
					saveFileDialog.RestoreDirectory = true;
                    if (string.IsNullOrEmpty(department))
                    {
                        saveFileDialog.FileName = _year.ToString() + "年度科研经费收支一览表";
                    }
                    else
                    {
                        saveFileDialog.FileName = department + _year.ToString() + "年度科研经费收支一览表";
                    }
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
				}
				catch (Exception error)
				{
					MessageBox.Show(error.Message, "导出失败 ", MessageBoxButton.OK);
					return;
				}
			}
			catch (FormatException)
			{
				MessageBox.Show("年度数字格式错误，应为四位整数", "错误");
				return;
			}
		}

		private void Window_Loaded(object sender, RoutedEventArgs e)
		{
			dataContext = new DataClassesProjectClassifyDataContext();
		}

		/// <summary>
		/// 设定符合条件的项目类型列表和项目列表
		/// </summary>
		private void SetProjectIDsAndProjectClassifyIDs()
		{
			projectIDs = new List<int>();
			projectClassifyIDs = new List<int>();
            IQueryable<Funds> funds;
            if (String.IsNullOrEmpty(department))
            {
                funds = dataContext.Funds.Where(f => f.FundClassifyID < 4 && f.Date >= startDate && f.Date <= endDate);
            }
            else
            {
                funds = dataContext.Funds.Where(f => f.FundClassifyID < 4 && f.Date >= startDate && f.Date <= endDate && f.ProjectBase.AnchoredDepartment.Contains(department));
            }
			foreach(Funds f in funds)
			{
				if (!projectIDs.Contains((int)f.ProjectID))
				{
					projectIDs.Add((int)f.ProjectID);
				}
				if (!projectClassifyIDs.Contains((int)f.ProjectBase.ProjectClassifyID))
				{
					projectClassifyIDs.Add((int)f.ProjectBase.ProjectClassifyID);
				}
			}
			projectClassifyIDs.Sort();
			projectIDs.Sort();
		}

		public void CreatePackage(string filePath)
		{
			using (SpreadsheetDocument package = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
			{
				CreateParts(package);
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
			Ap.Company company1 = new Ap.Company();
			company1.Text = "";
			Ap.LinksUpToDate linksUpToDate1 = new Ap.LinksUpToDate();
			linksUpToDate1.Text = "false";
			Ap.SharedDocument sharedDocument1 = new Ap.SharedDocument();
			sharedDocument1.Text = "false";
			Ap.HyperlinksChanged hyperlinksChanged1 = new Ap.HyperlinksChanged();
			hyperlinksChanged1.Text = "false";
			Ap.ApplicationVersion applicationVersion1 = new Ap.ApplicationVersion();
			applicationVersion1.Text = "15.0300";

			properties1.Append(application1);
			properties1.Append(documentSecurity1);
			properties1.Append(scaleCrop1);
			properties1.Append(headingPairs1);
			properties1.Append(titlesOfParts1);
			properties1.Append(company1);
			properties1.Append(linksUpToDate1);
			properties1.Append(sharedDocument1);
			properties1.Append(hyperlinksChanged1);
			properties1.Append(applicationVersion1);

			extendedFilePropertiesPart1.Properties = properties1;
		}

		// Generates content of workbookPart1.
		private void GenerateWorkbookPart1Content(WorkbookPart workbookPart1)
		{
			Workbook workbook1 = new Workbook() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x15" } };
			workbook1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
			workbook1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
			workbook1.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
			FileVersion fileVersion1 = new FileVersion() { ApplicationName = "xl", LastEdited = "6", LowestEdited = "6", BuildVersion = "14420" };
			WorkbookProperties workbookProperties1 = new WorkbookProperties() { DefaultThemeVersion = (UInt32Value)153222U };

			AlternateContent alternateContent1 = new AlternateContent();
			alternateContent1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

			AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice() { Requires = "x15" };

			X15ac.AbsolutePath absolutePath1 = new X15ac.AbsolutePath() { Url = "G:\\Users\\jhl.cad\\Visual Studio 2013\\Projects\\HNCDIExternalProjectManage\\HNCDIExternalProjectManage\\" };
			absolutePath1.AddNamespaceDeclaration("x15ac", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac");

			alternateContentChoice1.Append(absolutePath1);

			alternateContent1.Append(alternateContentChoice1);

			BookViews bookViews1 = new BookViews();
			WorkbookView workbookView1 = new WorkbookView() { XWindow = 0, YWindow = 0, WindowWidth = (UInt32Value)25200U, WindowHeight = (UInt32Value)11985U };

			bookViews1.Append(workbookView1);

			Sheets sheets1 = new Sheets();
			Sheet sheet1 = new Sheet() { Name = "Sheet1", SheetId = (UInt32Value)1U, Id = "rId1" };

			sheets1.Append(sheet1);
			CalculationProperties calculationProperties1 = new CalculationProperties() { CalculationId = (UInt32Value)152511U };

			WorkbookExtensionList workbookExtensionList1 = new WorkbookExtensionList();

			WorkbookExtension workbookExtension1 = new WorkbookExtension() { Uri = "{140A7094-0E35-4892-8432-C4D2E57EDEB5}" };
			workbookExtension1.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
			X15.WorkbookProperties workbookProperties2 = new X15.WorkbookProperties() { ChartTrackingReferenceBase = true };

			workbookExtension1.Append(workbookProperties2);

			workbookExtensionList1.Append(workbookExtension1);

			workbook1.Append(fileVersion1);
			workbook1.Append(workbookProperties1);
			workbook1.Append(alternateContent1);
			workbook1.Append(bookViews1);
			workbook1.Append(sheets1);
			workbook1.Append(calculationProperties1);
			workbook1.Append(workbookExtensionList1);

			workbookPart1.Workbook = workbook1;
		}

		// Generates content of workbookStylesPart1.
		private void GenerateWorkbookStylesPart1Content(WorkbookStylesPart workbookStylesPart1)
		{
			Stylesheet stylesheet1 = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
			stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
			stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

			DocumentFormat.OpenXml.Spreadsheet.Fonts fonts1 = new DocumentFormat.OpenXml.Spreadsheet.Fonts() { Count = (UInt32Value)5U, KnownFonts = true };

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
			FontSize fontSize3 = new FontSize() { Val = 20D };
			DocumentFormat.OpenXml.Spreadsheet.Color color2 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Theme = (UInt32Value)1U };
			FontName fontName3 = new FontName() { Val = "宋体" };
			FontFamilyNumbering fontFamilyNumbering3 = new FontFamilyNumbering() { Val = 3 };
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
			FontSize fontSize4 = new FontSize() { Val = 12D };
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

			fonts1.Append(font1);
			fonts1.Append(font2);
			fonts1.Append(font3);
			fonts1.Append(font4);
			fonts1.Append(font5);

			Fills fills1 = new Fills() { Count = (UInt32Value)2U };

			Fill fill1 = new Fill();
			PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };

			fill1.Append(patternFill1);

			Fill fill2 = new Fill();
			PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };

			fill2.Append(patternFill2);

			fills1.Append(fill1);
			fills1.Append(fill2);

			Borders borders1 = new Borders() { Count = (UInt32Value)13U };

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
			DocumentFormat.OpenXml.Spreadsheet.Color color5 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			leftBorder2.Append(color5);

			RightBorder rightBorder2 = new RightBorder() { Style = BorderStyleValues.Thin };
			DocumentFormat.OpenXml.Spreadsheet.Color color6 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			rightBorder2.Append(color6);

			TopBorder topBorder2 = new TopBorder() { Style = BorderStyleValues.Medium };
			DocumentFormat.OpenXml.Spreadsheet.Color color7 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			topBorder2.Append(color7);

			BottomBorder bottomBorder2 = new BottomBorder() { Style = BorderStyleValues.Thin };
			DocumentFormat.OpenXml.Spreadsheet.Color color8 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			bottomBorder2.Append(color8);
			DiagonalBorder diagonalBorder2 = new DiagonalBorder();

			border2.Append(leftBorder2);
			border2.Append(rightBorder2);
			border2.Append(topBorder2);
			border2.Append(bottomBorder2);
			border2.Append(diagonalBorder2);

			DocumentFormat.OpenXml.Spreadsheet.Border border3 = new DocumentFormat.OpenXml.Spreadsheet.Border();

			LeftBorder leftBorder3 = new LeftBorder() { Style = BorderStyleValues.Thin };
			DocumentFormat.OpenXml.Spreadsheet.Color color9 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			leftBorder3.Append(color9);

			RightBorder rightBorder3 = new RightBorder() { Style = BorderStyleValues.Thin };
			DocumentFormat.OpenXml.Spreadsheet.Color color10 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			rightBorder3.Append(color10);

			TopBorder topBorder3 = new TopBorder() { Style = BorderStyleValues.Medium };
			DocumentFormat.OpenXml.Spreadsheet.Color color11 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			topBorder3.Append(color11);

			BottomBorder bottomBorder3 = new BottomBorder() { Style = BorderStyleValues.Thin };
			DocumentFormat.OpenXml.Spreadsheet.Color color12 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			bottomBorder3.Append(color12);
			DiagonalBorder diagonalBorder3 = new DiagonalBorder();

			border3.Append(leftBorder3);
			border3.Append(rightBorder3);
			border3.Append(topBorder3);
			border3.Append(bottomBorder3);
			border3.Append(diagonalBorder3);

			DocumentFormat.OpenXml.Spreadsheet.Border border4 = new DocumentFormat.OpenXml.Spreadsheet.Border();

			LeftBorder leftBorder4 = new LeftBorder() { Style = BorderStyleValues.Thin };
			DocumentFormat.OpenXml.Spreadsheet.Color color13 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			leftBorder4.Append(color13);

			RightBorder rightBorder4 = new RightBorder() { Style = BorderStyleValues.Medium };
			DocumentFormat.OpenXml.Spreadsheet.Color color14 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			rightBorder4.Append(color14);

			TopBorder topBorder4 = new TopBorder() { Style = BorderStyleValues.Medium };
			DocumentFormat.OpenXml.Spreadsheet.Color color15 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			topBorder4.Append(color15);

			BottomBorder bottomBorder4 = new BottomBorder() { Style = BorderStyleValues.Thin };
			DocumentFormat.OpenXml.Spreadsheet.Color color16 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			bottomBorder4.Append(color16);
			DiagonalBorder diagonalBorder4 = new DiagonalBorder();

			border4.Append(leftBorder4);
			border4.Append(rightBorder4);
			border4.Append(topBorder4);
			border4.Append(bottomBorder4);
			border4.Append(diagonalBorder4);

			DocumentFormat.OpenXml.Spreadsheet.Border border5 = new DocumentFormat.OpenXml.Spreadsheet.Border();

			LeftBorder leftBorder5 = new LeftBorder() { Style = BorderStyleValues.Medium };
			DocumentFormat.OpenXml.Spreadsheet.Color color17 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			leftBorder5.Append(color17);

			RightBorder rightBorder5 = new RightBorder() { Style = BorderStyleValues.Thin };
			DocumentFormat.OpenXml.Spreadsheet.Color color18 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			rightBorder5.Append(color18);

			TopBorder topBorder5 = new TopBorder() { Style = BorderStyleValues.Thin };
			DocumentFormat.OpenXml.Spreadsheet.Color color19 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			topBorder5.Append(color19);

			BottomBorder bottomBorder5 = new BottomBorder() { Style = BorderStyleValues.Thin };
			DocumentFormat.OpenXml.Spreadsheet.Color color20 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			bottomBorder5.Append(color20);
			DiagonalBorder diagonalBorder5 = new DiagonalBorder();

			border5.Append(leftBorder5);
			border5.Append(rightBorder5);
			border5.Append(topBorder5);
			border5.Append(bottomBorder5);
			border5.Append(diagonalBorder5);

			DocumentFormat.OpenXml.Spreadsheet.Border border6 = new DocumentFormat.OpenXml.Spreadsheet.Border();

			LeftBorder leftBorder6 = new LeftBorder() { Style = BorderStyleValues.Thin };
			DocumentFormat.OpenXml.Spreadsheet.Color color21 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			leftBorder6.Append(color21);

			RightBorder rightBorder6 = new RightBorder() { Style = BorderStyleValues.Thin };
			DocumentFormat.OpenXml.Spreadsheet.Color color22 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			rightBorder6.Append(color22);

			TopBorder topBorder6 = new TopBorder() { Style = BorderStyleValues.Thin };
			DocumentFormat.OpenXml.Spreadsheet.Color color23 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			topBorder6.Append(color23);

			BottomBorder bottomBorder6 = new BottomBorder() { Style = BorderStyleValues.Thin };
			DocumentFormat.OpenXml.Spreadsheet.Color color24 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			bottomBorder6.Append(color24);
			DiagonalBorder diagonalBorder6 = new DiagonalBorder();

			border6.Append(leftBorder6);
			border6.Append(rightBorder6);
			border6.Append(topBorder6);
			border6.Append(bottomBorder6);
			border6.Append(diagonalBorder6);

			DocumentFormat.OpenXml.Spreadsheet.Border border7 = new DocumentFormat.OpenXml.Spreadsheet.Border();

			LeftBorder leftBorder7 = new LeftBorder() { Style = BorderStyleValues.Thin };
			DocumentFormat.OpenXml.Spreadsheet.Color color25 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			leftBorder7.Append(color25);

			RightBorder rightBorder7 = new RightBorder() { Style = BorderStyleValues.Medium };
			DocumentFormat.OpenXml.Spreadsheet.Color color26 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			rightBorder7.Append(color26);

			TopBorder topBorder7 = new TopBorder() { Style = BorderStyleValues.Thin };
			DocumentFormat.OpenXml.Spreadsheet.Color color27 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			topBorder7.Append(color27);

			BottomBorder bottomBorder7 = new BottomBorder() { Style = BorderStyleValues.Thin };
			DocumentFormat.OpenXml.Spreadsheet.Color color28 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			bottomBorder7.Append(color28);
			DiagonalBorder diagonalBorder7 = new DiagonalBorder();

			border7.Append(leftBorder7);
			border7.Append(rightBorder7);
			border7.Append(topBorder7);
			border7.Append(bottomBorder7);
			border7.Append(diagonalBorder7);

			DocumentFormat.OpenXml.Spreadsheet.Border border8 = new DocumentFormat.OpenXml.Spreadsheet.Border();

			LeftBorder leftBorder8 = new LeftBorder() { Style = BorderStyleValues.Medium };
			DocumentFormat.OpenXml.Spreadsheet.Color color29 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			leftBorder8.Append(color29);

			RightBorder rightBorder8 = new RightBorder() { Style = BorderStyleValues.Thin };
			DocumentFormat.OpenXml.Spreadsheet.Color color30 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			rightBorder8.Append(color30);

			TopBorder topBorder8 = new TopBorder() { Style = BorderStyleValues.Thin };
			DocumentFormat.OpenXml.Spreadsheet.Color color31 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			topBorder8.Append(color31);

			BottomBorder bottomBorder8 = new BottomBorder() { Style = BorderStyleValues.Medium };
			DocumentFormat.OpenXml.Spreadsheet.Color color32 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			bottomBorder8.Append(color32);
			DiagonalBorder diagonalBorder8 = new DiagonalBorder();

			border8.Append(leftBorder8);
			border8.Append(rightBorder8);
			border8.Append(topBorder8);
			border8.Append(bottomBorder8);
			border8.Append(diagonalBorder8);

			DocumentFormat.OpenXml.Spreadsheet.Border border9 = new DocumentFormat.OpenXml.Spreadsheet.Border();

			LeftBorder leftBorder9 = new LeftBorder() { Style = BorderStyleValues.Thin };
			DocumentFormat.OpenXml.Spreadsheet.Color color33 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			leftBorder9.Append(color33);

			RightBorder rightBorder9 = new RightBorder() { Style = BorderStyleValues.Thin };
			DocumentFormat.OpenXml.Spreadsheet.Color color34 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			rightBorder9.Append(color34);

			TopBorder topBorder9 = new TopBorder() { Style = BorderStyleValues.Thin };
			DocumentFormat.OpenXml.Spreadsheet.Color color35 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			topBorder9.Append(color35);

			BottomBorder bottomBorder9 = new BottomBorder() { Style = BorderStyleValues.Medium };
			DocumentFormat.OpenXml.Spreadsheet.Color color36 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			bottomBorder9.Append(color36);
			DiagonalBorder diagonalBorder9 = new DiagonalBorder();

			border9.Append(leftBorder9);
			border9.Append(rightBorder9);
			border9.Append(topBorder9);
			border9.Append(bottomBorder9);
			border9.Append(diagonalBorder9);

			DocumentFormat.OpenXml.Spreadsheet.Border border10 = new DocumentFormat.OpenXml.Spreadsheet.Border();

			LeftBorder leftBorder10 = new LeftBorder() { Style = BorderStyleValues.Thin };
			DocumentFormat.OpenXml.Spreadsheet.Color color37 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			leftBorder10.Append(color37);

			RightBorder rightBorder10 = new RightBorder() { Style = BorderStyleValues.Medium };
			DocumentFormat.OpenXml.Spreadsheet.Color color38 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			rightBorder10.Append(color38);

			TopBorder topBorder10 = new TopBorder() { Style = BorderStyleValues.Thin };
			DocumentFormat.OpenXml.Spreadsheet.Color color39 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			topBorder10.Append(color39);

			BottomBorder bottomBorder10 = new BottomBorder() { Style = BorderStyleValues.Medium };
			DocumentFormat.OpenXml.Spreadsheet.Color color40 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			bottomBorder10.Append(color40);
			DiagonalBorder diagonalBorder10 = new DiagonalBorder();

			border10.Append(leftBorder10);
			border10.Append(rightBorder10);
			border10.Append(topBorder10);
			border10.Append(bottomBorder10);
			border10.Append(diagonalBorder10);

			DocumentFormat.OpenXml.Spreadsheet.Border border11 = new DocumentFormat.OpenXml.Spreadsheet.Border();

			LeftBorder leftBorder11 = new LeftBorder() { Style = BorderStyleValues.Medium };
			DocumentFormat.OpenXml.Spreadsheet.Color color41 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			leftBorder11.Append(color41);

			RightBorder rightBorder11 = new RightBorder() { Style = BorderStyleValues.Thin };
			DocumentFormat.OpenXml.Spreadsheet.Color color42 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			rightBorder11.Append(color42);

			TopBorder topBorder11 = new TopBorder() { Style = BorderStyleValues.Thin };
			DocumentFormat.OpenXml.Spreadsheet.Color color43 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			topBorder11.Append(color43);
			BottomBorder bottomBorder11 = new BottomBorder();
			DiagonalBorder diagonalBorder11 = new DiagonalBorder();

			border11.Append(leftBorder11);
			border11.Append(rightBorder11);
			border11.Append(topBorder11);
			border11.Append(bottomBorder11);
			border11.Append(diagonalBorder11);

			DocumentFormat.OpenXml.Spreadsheet.Border border12 = new DocumentFormat.OpenXml.Spreadsheet.Border();

			LeftBorder leftBorder12 = new LeftBorder() { Style = BorderStyleValues.Thin };
			DocumentFormat.OpenXml.Spreadsheet.Color color44 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			leftBorder12.Append(color44);

			RightBorder rightBorder12 = new RightBorder() { Style = BorderStyleValues.Thin };
			DocumentFormat.OpenXml.Spreadsheet.Color color45 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			rightBorder12.Append(color45);

			TopBorder topBorder12 = new TopBorder() { Style = BorderStyleValues.Thin };
			DocumentFormat.OpenXml.Spreadsheet.Color color46 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			topBorder12.Append(color46);
			BottomBorder bottomBorder12 = new BottomBorder();
			DiagonalBorder diagonalBorder12 = new DiagonalBorder();

			border12.Append(leftBorder12);
			border12.Append(rightBorder12);
			border12.Append(topBorder12);
			border12.Append(bottomBorder12);
			border12.Append(diagonalBorder12);

			DocumentFormat.OpenXml.Spreadsheet.Border border13 = new DocumentFormat.OpenXml.Spreadsheet.Border();

			LeftBorder leftBorder13 = new LeftBorder() { Style = BorderStyleValues.Thin };
			DocumentFormat.OpenXml.Spreadsheet.Color color47 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			leftBorder13.Append(color47);

			RightBorder rightBorder13 = new RightBorder() { Style = BorderStyleValues.Medium };
			DocumentFormat.OpenXml.Spreadsheet.Color color48 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			rightBorder13.Append(color48);

			TopBorder topBorder13 = new TopBorder() { Style = BorderStyleValues.Thin };
			DocumentFormat.OpenXml.Spreadsheet.Color color49 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true };

			topBorder13.Append(color49);
			BottomBorder bottomBorder13 = new BottomBorder();
			DiagonalBorder diagonalBorder13 = new DiagonalBorder();

			border13.Append(leftBorder13);
			border13.Append(rightBorder13);
			border13.Append(topBorder13);
			border13.Append(bottomBorder13);
			border13.Append(diagonalBorder13);

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

			CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)1U };

			CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };
			Alignment alignment1 = new Alignment() { Vertical = VerticalAlignmentValues.Center };

			cellFormat1.Append(alignment1);

			cellStyleFormats1.Append(cellFormat1);

			CellFormats cellFormats1 = new CellFormats() { Count = (UInt32Value)27U };

			CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
			Alignment alignment2 = new Alignment() { Vertical = VerticalAlignmentValues.Center };

			cellFormat2.Append(alignment2);

			CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyAlignment = true };
			Alignment alignment3 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

			cellFormat3.Append(alignment3);

			CellFormat cellFormat4 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
			Alignment alignment4 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

			cellFormat4.Append(alignment4);

			CellFormat cellFormat5 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)2U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
			Alignment alignment5 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

			cellFormat5.Append(alignment5);

			CellFormat cellFormat6 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)3U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
			Alignment alignment6 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

			cellFormat6.Append(alignment6);

			CellFormat cellFormat7 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)4U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
			Alignment alignment7 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

			cellFormat7.Append(alignment7);

			CellFormat cellFormat8 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)5U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
			Alignment alignment8 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

			cellFormat8.Append(alignment8);

			CellFormat cellFormat9 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)5U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
			Alignment alignment9 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

			cellFormat9.Append(alignment9);

			CellFormat cellFormat10 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)6U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
			Alignment alignment10 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

			cellFormat10.Append(alignment10);

			CellFormat cellFormat11 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)5U, FormatId = (UInt32Value)0U, ApplyBorder = true, ApplyAlignment = true };
			Alignment alignment11 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

			cellFormat11.Append(alignment11);

			CellFormat cellFormat12 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)6U, FormatId = (UInt32Value)0U, ApplyBorder = true, ApplyAlignment = true };
			Alignment alignment12 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

			cellFormat12.Append(alignment12);

			CellFormat cellFormat13 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)8U, FormatId = (UInt32Value)0U, ApplyBorder = true, ApplyAlignment = true };
			Alignment alignment13 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

			cellFormat13.Append(alignment13);

			CellFormat cellFormat14 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)9U, FormatId = (UInt32Value)0U, ApplyBorder = true, ApplyAlignment = true };
			Alignment alignment14 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

			cellFormat14.Append(alignment14);

			CellFormat cellFormat15 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)5U, FormatId = (UInt32Value)0U, ApplyBorder = true, ApplyAlignment = true };
			Alignment alignment15 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

			cellFormat15.Append(alignment15);

			CellFormat cellFormat16 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)4U, FormatId = (UInt32Value)0U, ApplyBorder = true, ApplyAlignment = true };
			Alignment alignment16 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

			cellFormat16.Append(alignment16);

			CellFormat cellFormat17 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)10U, FormatId = (UInt32Value)0U, ApplyBorder = true, ApplyAlignment = true };
			Alignment alignment17 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

			cellFormat17.Append(alignment17);

			CellFormat cellFormat18 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)11U, FormatId = (UInt32Value)0U, ApplyBorder = true, ApplyAlignment = true };
			Alignment alignment18 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

			cellFormat18.Append(alignment18);

			CellFormat cellFormat19 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)4U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)5U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
			Alignment alignment19 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

			cellFormat19.Append(alignment19);

			CellFormat cellFormat20 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)4U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)6U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
			Alignment alignment20 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

			cellFormat20.Append(alignment20);

			CellFormat cellFormat21 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)7U, FormatId = (UInt32Value)0U, ApplyBorder = true, ApplyAlignment = true };
			Alignment alignment21 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

			cellFormat21.Append(alignment21);

			CellFormat cellFormat22 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)8U, FormatId = (UInt32Value)0U, ApplyBorder = true, ApplyAlignment = true };
			Alignment alignment22 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

			cellFormat22.Append(alignment22);

			CellFormat cellFormat23 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)5U, FormatId = (UInt32Value)0U, ApplyBorder = true, ApplyAlignment = true };
			Alignment alignment23 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center, WrapText = true };

			cellFormat23.Append(alignment23);

			CellFormat cellFormat24 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)8U, FormatId = (UInt32Value)0U, ApplyBorder = true, ApplyAlignment = true };
			Alignment alignment24 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center, WrapText = true };

			cellFormat24.Append(alignment24);

			CellFormat cellFormat25 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)11U, FormatId = (UInt32Value)0U, ApplyBorder = true, ApplyAlignment = true };
			Alignment alignment25 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center, WrapText = true };

			cellFormat25.Append(alignment25);

			CellFormat cellFormat26 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)11U, FormatId = (UInt32Value)0U, ApplyBorder = true, ApplyAlignment = true };
			Alignment alignment26 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

			cellFormat26.Append(alignment26);

			CellFormat cellFormat27 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)12U, FormatId = (UInt32Value)0U, ApplyBorder = true, ApplyAlignment = true };
			Alignment alignment27 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

			cellFormat27.Append(alignment27);

			CellFormat cellFormat28 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)4U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)4U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
			Alignment alignment28 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

			cellFormat28.Append(alignment28);

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

			StylesheetExtension stylesheetExtension2 = new StylesheetExtension() { Uri = "{9260A510-F301-46a8-8635-F512D64BE5F5}" };
			stylesheetExtension2.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
			X15.TimelineStyles timelineStyles1 = new X15.TimelineStyles() { DefaultTimelineStyle = "TimeSlicerStyleLight1" };

			stylesheetExtension2.Append(timelineStyles1);

			stylesheetExtensionList1.Append(stylesheetExtension1);
			stylesheetExtensionList1.Append(stylesheetExtension2);

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
			A.Theme theme1 = new A.Theme() { Name = "Office 主题" };
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
			A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "44546A" };

			dark2Color1.Append(rgbColorModelHex1);

			A.Light2Color light2Color1 = new A.Light2Color();
			A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "E7E6E6" };

			light2Color1.Append(rgbColorModelHex2);

			A.Accent1Color accent1Color1 = new A.Accent1Color();
			A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "5B9BD5" };

			accent1Color1.Append(rgbColorModelHex3);

			A.Accent2Color accent2Color1 = new A.Accent2Color();
			A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "ED7D31" };

			accent2Color1.Append(rgbColorModelHex4);

			A.Accent3Color accent3Color1 = new A.Accent3Color();
			A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "A5A5A5" };

			accent3Color1.Append(rgbColorModelHex5);

			A.Accent4Color accent4Color1 = new A.Accent4Color();
			A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "FFC000" };

			accent4Color1.Append(rgbColorModelHex6);

			A.Accent5Color accent5Color1 = new A.Accent5Color();
			A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "4472C4" };

			accent5Color1.Append(rgbColorModelHex7);

			A.Accent6Color accent6Color1 = new A.Accent6Color();
			A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "70AD47" };

			accent6Color1.Append(rgbColorModelHex8);

			A.Hyperlink hyperlink1 = new A.Hyperlink();
			A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex() { Val = "0563C1" };

			hyperlink1.Append(rgbColorModelHex9);

			A.FollowedHyperlinkColor followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
			A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex() { Val = "954F72" };

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

			A.FontScheme fontScheme6 = new A.FontScheme() { Name = "Office" };

			A.MajorFont majorFont1 = new A.MajorFont();
			A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "Calibri Light", Panose = "020F0302020204030204" };
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
			A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "Calibri", Panose = "020F0502020204030204" };
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

			fontScheme6.Append(majorFont1);
			fontScheme6.Append(minorFont1);

			A.FormatScheme formatScheme1 = new A.FormatScheme() { Name = "Office" };

			A.FillStyleList fillStyleList1 = new A.FillStyleList();

			A.SolidFill solidFill1 = new A.SolidFill();
			A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

			solidFill1.Append(schemeColor1);

			A.GradientFill gradientFill1 = new A.GradientFill() { RotateWithShape = true };

			A.GradientStopList gradientStopList1 = new A.GradientStopList();

			A.GradientStop gradientStop1 = new A.GradientStop() { Position = 0 };

			A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
			A.LuminanceModulation luminanceModulation1 = new A.LuminanceModulation() { Val = 110000 };
			A.SaturationModulation saturationModulation1 = new A.SaturationModulation() { Val = 105000 };
			A.Tint tint1 = new A.Tint() { Val = 67000 };

			schemeColor2.Append(luminanceModulation1);
			schemeColor2.Append(saturationModulation1);
			schemeColor2.Append(tint1);

			gradientStop1.Append(schemeColor2);

			A.GradientStop gradientStop2 = new A.GradientStop() { Position = 50000 };

			A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
			A.LuminanceModulation luminanceModulation2 = new A.LuminanceModulation() { Val = 105000 };
			A.SaturationModulation saturationModulation2 = new A.SaturationModulation() { Val = 103000 };
			A.Tint tint2 = new A.Tint() { Val = 73000 };

			schemeColor3.Append(luminanceModulation2);
			schemeColor3.Append(saturationModulation2);
			schemeColor3.Append(tint2);

			gradientStop2.Append(schemeColor3);

			A.GradientStop gradientStop3 = new A.GradientStop() { Position = 100000 };

			A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
			A.LuminanceModulation luminanceModulation3 = new A.LuminanceModulation() { Val = 105000 };
			A.SaturationModulation saturationModulation3 = new A.SaturationModulation() { Val = 109000 };
			A.Tint tint3 = new A.Tint() { Val = 81000 };

			schemeColor4.Append(luminanceModulation3);
			schemeColor4.Append(saturationModulation3);
			schemeColor4.Append(tint3);

			gradientStop3.Append(schemeColor4);

			gradientStopList1.Append(gradientStop1);
			gradientStopList1.Append(gradientStop2);
			gradientStopList1.Append(gradientStop3);
			A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

			gradientFill1.Append(gradientStopList1);
			gradientFill1.Append(linearGradientFill1);

			A.GradientFill gradientFill2 = new A.GradientFill() { RotateWithShape = true };

			A.GradientStopList gradientStopList2 = new A.GradientStopList();

			A.GradientStop gradientStop4 = new A.GradientStop() { Position = 0 };

			A.SchemeColor schemeColor5 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
			A.SaturationModulation saturationModulation4 = new A.SaturationModulation() { Val = 103000 };
			A.LuminanceModulation luminanceModulation4 = new A.LuminanceModulation() { Val = 102000 };
			A.Tint tint4 = new A.Tint() { Val = 94000 };

			schemeColor5.Append(saturationModulation4);
			schemeColor5.Append(luminanceModulation4);
			schemeColor5.Append(tint4);

			gradientStop4.Append(schemeColor5);

			A.GradientStop gradientStop5 = new A.GradientStop() { Position = 50000 };

			A.SchemeColor schemeColor6 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
			A.SaturationModulation saturationModulation5 = new A.SaturationModulation() { Val = 110000 };
			A.LuminanceModulation luminanceModulation5 = new A.LuminanceModulation() { Val = 100000 };
			A.Shade shade1 = new A.Shade() { Val = 100000 };

			schemeColor6.Append(saturationModulation5);
			schemeColor6.Append(luminanceModulation5);
			schemeColor6.Append(shade1);

			gradientStop5.Append(schemeColor6);

			A.GradientStop gradientStop6 = new A.GradientStop() { Position = 100000 };

			A.SchemeColor schemeColor7 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
			A.LuminanceModulation luminanceModulation6 = new A.LuminanceModulation() { Val = 99000 };
			A.SaturationModulation saturationModulation6 = new A.SaturationModulation() { Val = 120000 };
			A.Shade shade2 = new A.Shade() { Val = 78000 };

			schemeColor7.Append(luminanceModulation6);
			schemeColor7.Append(saturationModulation6);
			schemeColor7.Append(shade2);

			gradientStop6.Append(schemeColor7);

			gradientStopList2.Append(gradientStop4);
			gradientStopList2.Append(gradientStop5);
			gradientStopList2.Append(gradientStop6);
			A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

			gradientFill2.Append(gradientStopList2);
			gradientFill2.Append(linearGradientFill2);

			fillStyleList1.Append(solidFill1);
			fillStyleList1.Append(gradientFill1);
			fillStyleList1.Append(gradientFill2);

			A.LineStyleList lineStyleList1 = new A.LineStyleList();

			A.Outline outline1 = new A.Outline() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

			A.SolidFill solidFill2 = new A.SolidFill();
			A.SchemeColor schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

			solidFill2.Append(schemeColor8);
			A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
			A.Miter miter1 = new A.Miter() { Limit = 800000 };

			outline1.Append(solidFill2);
			outline1.Append(presetDash1);
			outline1.Append(miter1);

			A.Outline outline2 = new A.Outline() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

			A.SolidFill solidFill3 = new A.SolidFill();
			A.SchemeColor schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

			solidFill3.Append(schemeColor9);
			A.PresetDash presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
			A.Miter miter2 = new A.Miter() { Limit = 800000 };

			outline2.Append(solidFill3);
			outline2.Append(presetDash2);
			outline2.Append(miter2);

			A.Outline outline3 = new A.Outline() { Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

			A.SolidFill solidFill4 = new A.SolidFill();
			A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

			solidFill4.Append(schemeColor10);
			A.PresetDash presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
			A.Miter miter3 = new A.Miter() { Limit = 800000 };

			outline3.Append(solidFill4);
			outline3.Append(presetDash3);
			outline3.Append(miter3);

			lineStyleList1.Append(outline1);
			lineStyleList1.Append(outline2);
			lineStyleList1.Append(outline3);

			A.EffectStyleList effectStyleList1 = new A.EffectStyleList();

			A.EffectStyle effectStyle1 = new A.EffectStyle();
			A.EffectList effectList1 = new A.EffectList();

			effectStyle1.Append(effectList1);

			A.EffectStyle effectStyle2 = new A.EffectStyle();
			A.EffectList effectList2 = new A.EffectList();

			effectStyle2.Append(effectList2);

			A.EffectStyle effectStyle3 = new A.EffectStyle();

			A.EffectList effectList3 = new A.EffectList();

			A.OuterShadow outerShadow1 = new A.OuterShadow() { BlurRadius = 57150L, Distance = 19050L, Direction = 5400000, Alignment = A.RectangleAlignmentValues.Center, RotateWithShape = false };

			A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex() { Val = "000000" };
			A.Alpha alpha1 = new A.Alpha() { Val = 63000 };

			rgbColorModelHex11.Append(alpha1);

			outerShadow1.Append(rgbColorModelHex11);

			effectList3.Append(outerShadow1);

			effectStyle3.Append(effectList3);

			effectStyleList1.Append(effectStyle1);
			effectStyleList1.Append(effectStyle2);
			effectStyleList1.Append(effectStyle3);

			A.BackgroundFillStyleList backgroundFillStyleList1 = new A.BackgroundFillStyleList();

			A.SolidFill solidFill5 = new A.SolidFill();
			A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

			solidFill5.Append(schemeColor11);

			A.SolidFill solidFill6 = new A.SolidFill();

			A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
			A.Tint tint5 = new A.Tint() { Val = 95000 };
			A.SaturationModulation saturationModulation7 = new A.SaturationModulation() { Val = 170000 };

			schemeColor12.Append(tint5);
			schemeColor12.Append(saturationModulation7);

			solidFill6.Append(schemeColor12);

			A.GradientFill gradientFill3 = new A.GradientFill() { RotateWithShape = true };

			A.GradientStopList gradientStopList3 = new A.GradientStopList();

			A.GradientStop gradientStop7 = new A.GradientStop() { Position = 0 };

			A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
			A.Tint tint6 = new A.Tint() { Val = 93000 };
			A.SaturationModulation saturationModulation8 = new A.SaturationModulation() { Val = 150000 };
			A.Shade shade3 = new A.Shade() { Val = 98000 };
			A.LuminanceModulation luminanceModulation7 = new A.LuminanceModulation() { Val = 102000 };

			schemeColor13.Append(tint6);
			schemeColor13.Append(saturationModulation8);
			schemeColor13.Append(shade3);
			schemeColor13.Append(luminanceModulation7);

			gradientStop7.Append(schemeColor13);

			A.GradientStop gradientStop8 = new A.GradientStop() { Position = 50000 };

			A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
			A.Tint tint7 = new A.Tint() { Val = 98000 };
			A.SaturationModulation saturationModulation9 = new A.SaturationModulation() { Val = 130000 };
			A.Shade shade4 = new A.Shade() { Val = 90000 };
			A.LuminanceModulation luminanceModulation8 = new A.LuminanceModulation() { Val = 103000 };

			schemeColor14.Append(tint7);
			schemeColor14.Append(saturationModulation9);
			schemeColor14.Append(shade4);
			schemeColor14.Append(luminanceModulation8);

			gradientStop8.Append(schemeColor14);

			A.GradientStop gradientStop9 = new A.GradientStop() { Position = 100000 };

			A.SchemeColor schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
			A.Shade shade5 = new A.Shade() { Val = 63000 };
			A.SaturationModulation saturationModulation10 = new A.SaturationModulation() { Val = 120000 };

			schemeColor15.Append(shade5);
			schemeColor15.Append(saturationModulation10);

			gradientStop9.Append(schemeColor15);

			gradientStopList3.Append(gradientStop7);
			gradientStopList3.Append(gradientStop8);
			gradientStopList3.Append(gradientStop9);
			A.LinearGradientFill linearGradientFill3 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

			gradientFill3.Append(gradientStopList3);
			gradientFill3.Append(linearGradientFill3);

			backgroundFillStyleList1.Append(solidFill5);
			backgroundFillStyleList1.Append(solidFill6);
			backgroundFillStyleList1.Append(gradientFill3);

			formatScheme1.Append(fillStyleList1);
			formatScheme1.Append(lineStyleList1);
			formatScheme1.Append(effectStyleList1);
			formatScheme1.Append(backgroundFillStyleList1);

			themeElements1.Append(colorScheme1);
			themeElements1.Append(fontScheme6);
			themeElements1.Append(formatScheme1);
			A.ObjectDefaults objectDefaults1 = new A.ObjectDefaults();
			A.ExtraColorSchemeList extraColorSchemeList1 = new A.ExtraColorSchemeList();

			A.OfficeStyleSheetExtensionList officeStyleSheetExtensionList1 = new A.OfficeStyleSheetExtensionList();

			A.OfficeStyleSheetExtension officeStyleSheetExtension1 = new A.OfficeStyleSheetExtension() { Uri = "{05A4C25C-085E-4340-85A3-A5531E510DB2}" };

			Thm15.ThemeFamily themeFamily1 = new Thm15.ThemeFamily() { Name = "Office Theme", Id = "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}", Vid = "{4A3C46E8-61CC-4603-A589-7422A47A8E4A}" };
			themeFamily1.AddNamespaceDeclaration("thm15", "http://schemas.microsoft.com/office/thememl/2012/main");

			officeStyleSheetExtension1.Append(themeFamily1);

			officeStyleSheetExtensionList1.Append(officeStyleSheetExtension1);

			theme1.Append(themeElements1);
			theme1.Append(objectDefaults1);
			theme1.Append(extraColorSchemeList1);
			theme1.Append(officeStyleSheetExtensionList1);

			themePart1.Theme = theme1;
		}

		// Generates content of worksheetPart1.
		private void GenerateWorksheetPart1Content(WorksheetPart worksheetPart1)
		{
			Worksheet worksheet1 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
			worksheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
			worksheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
			worksheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
			SheetDimension sheetDimension1 = new SheetDimension() { Reference = "A1:" + lastCellName };

			SheetViews sheetViews1 = new SheetViews();

			SheetView sheetView1 = new SheetView() { TabSelected = true, WorkbookViewId = (UInt32Value)0U };
			Selection selection1 = new Selection() { SequenceOfReferences = new ListValue<StringValue>() { InnerText = "A1:J1" } };

			sheetView1.Append(selection1);

			sheetViews1.Append(sheetView1);
			SheetFormatProperties sheetFormatProperties1 = new SheetFormatProperties() { DefaultRowHeight = 13.5D, DyDescent = 0.15D };

			Columns columns1 = new Columns();
			Column column1 = new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)1U, Width = 5.75D, CustomWidth = true };
			Column column2 = new Column() { Min = (UInt32Value)2U, Max = (UInt32Value)2U, Width = 9.625D, CustomWidth = true };
			Column column3 = new Column() { Min = (UInt32Value)3U, Max = (UInt32Value)3U, Width = 65D, CustomWidth = true };
			Column column4 = new Column() { Min = (UInt32Value)4U, Max = (UInt32Value)4U, Width = 20D, CustomWidth = true };
			Column column5 = new Column() { Min = (UInt32Value)5U, Max = (UInt32Value)5U, Width = 10D, CustomWidth = true };
			Column column6 = new Column() { Min = (UInt32Value)6U, Max = (UInt32Value)6U, Width = 10D, CustomWidth = true };
			Column column7 = new Column() { Min = (UInt32Value)7U, Max = (UInt32Value)7U, Width = 10D, CustomWidth = true };
			Column column8 = new Column() { Min = (UInt32Value)8U, Max = (UInt32Value)8U, Width = 10D, CustomWidth = true };
			Column column9 = new Column() { Min = (UInt32Value)9U, Max = (UInt32Value)9U, Width = 10D, CustomWidth = true };
			Column column10 = new Column() { Min = (UInt32Value)10U, Max = (UInt32Value)10U, Width = 10D, CustomWidth = true };

			columns1.Append(column1);
			columns1.Append(column2);
			columns1.Append(column3);
			columns1.Append(column4);
			columns1.Append(column5);
			columns1.Append(column6);
			columns1.Append(column7);
			columns1.Append(column8);
			columns1.Append(column9);
			columns1.Append(column10);

			SheetData sheetData1 = new SheetData();
			MergeCells mergeCells1 = new MergeCells() { Count = (UInt32Value)Convert.ToUInt32(totalMergeCells) }; //合并单元格定义

			#region 标题行
			Row row1 = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, Height = 30D, CustomHeight = true, ThickBot = true, DyDescent = 0.2D };

			Cell cell1 = new Cell() { CellReference = "A1", StyleIndex = (UInt32Value)1U, DataType = CellValues.String };
			CellValue cellValue1 = new CellValue();
            if (string.IsNullOrEmpty(department))
            {
                cellValue1.Text = _year.ToString() + "年度科研经费收支一览表";
            }
            else
            {
                cellValue1.Text = department + _year.ToString() + "年度科研经费收支一览表";
            }

			cell1.Append(cellValue1);
			Cell cell2 = new Cell() { CellReference = "B1", StyleIndex = (UInt32Value)1U };
			Cell cell3 = new Cell() { CellReference = "C1", StyleIndex = (UInt32Value)1U };
			Cell cell4 = new Cell() { CellReference = "D1", StyleIndex = (UInt32Value)1U };
			Cell cell5 = new Cell() { CellReference = "E1", StyleIndex = (UInt32Value)1U };
			Cell cell6 = new Cell() { CellReference = "F1", StyleIndex = (UInt32Value)1U };
			Cell cell7 = new Cell() { CellReference = "G1", StyleIndex = (UInt32Value)1U };
			Cell cell8 = new Cell() { CellReference = "H1", StyleIndex = (UInt32Value)1U };
			Cell cell9 = new Cell() { CellReference = "I1", StyleIndex = (UInt32Value)1U };
			Cell cell10 = new Cell() { CellReference = "J1", DataType = CellValues.String };
            CellValue cellValue = new CellValue();
            cellValue.Text = "单位：万元";
            cell10.Append(cellValue);

			row1.Append(cell1);
			row1.Append(cell2);
			row1.Append(cell3);
			row1.Append(cell4);
			row1.Append(cell5);
			row1.Append(cell6);
			row1.Append(cell7);
			row1.Append(cell8);
			row1.Append(cell9);
			row1.Append(cell10);

			sheetData1.Append(row1);
			MergeCell mergeCell1 = new MergeCell() { Reference = "A1:I1" };
			mergeCells1.Append(mergeCell1);

			#endregion

			#region 表头行
			Row row2 = new Row() { RowIndex = (UInt32Value)2U, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, Height = 20.100000000000001D, CustomHeight = true, DyDescent = 0.15D };

			Cell cell11 = new Cell() { CellReference = "A2", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
			CellValue cellValue2 = new CellValue();
			cellValue2.Text = "1";

			cell11.Append(cellValue2);

			Cell cell12 = new Cell() { CellReference = "B2", StyleIndex = (UInt32Value)3U, DataType = CellValues.SharedString };
			CellValue cellValue3 = new CellValue();
			cellValue3.Text = "2";

			cell12.Append(cellValue3);

			Cell cell13 = new Cell() { CellReference = "C2", StyleIndex = (UInt32Value)3U, DataType = CellValues.SharedString };
			CellValue cellValue4 = new CellValue();
			cellValue4.Text = "3";

			cell13.Append(cellValue4);

			Cell cell14 = new Cell() { CellReference = "D2", StyleIndex = (UInt32Value)3U, DataType = CellValues.SharedString };
			CellValue cellValue5 = new CellValue();
			cellValue5.Text = "4";

			cell14.Append(cellValue5);

			Cell cell15 = new Cell() { CellReference = "E2", StyleIndex = (UInt32Value)3U, DataType = CellValues.SharedString };
			CellValue cellValue6 = new CellValue();
			cellValue6.Text = "5";

			cell15.Append(cellValue6);
			Cell cell16 = new Cell() { CellReference = "F2", StyleIndex = (UInt32Value)3U };

			Cell cell17 = new Cell() { CellReference = "G2", StyleIndex = (UInt32Value)3U, DataType = CellValues.SharedString };
			CellValue cellValue7 = new CellValue();
			cellValue7.Text = "6";

			cell17.Append(cellValue7);
			Cell cell18 = new Cell() { CellReference = "H2", StyleIndex = (UInt32Value)3U };

			Cell cell19 = new Cell() { CellReference = "I2", StyleIndex = (UInt32Value)3U, DataType = CellValues.SharedString };
			CellValue cellValue8 = new CellValue();
			cellValue8.Text = "7";

			cell19.Append(cellValue8);
			Cell cell20 = new Cell() { CellReference = "J2", StyleIndex = (UInt32Value)4U };

			row2.Append(cell11);
			row2.Append(cell12);
			row2.Append(cell13);
			row2.Append(cell14);
			row2.Append(cell15);
			row2.Append(cell16);
			row2.Append(cell17);
			row2.Append(cell18);
			row2.Append(cell19);
			row2.Append(cell20);
			sheetData1.Append(row2);

			Row row3 = new Row() { RowIndex = (UInt32Value)3U, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, Height = 20.100000000000001D, CustomHeight = true, DyDescent = 0.15D };
			Cell cell21 = new Cell() { CellReference = "A3", StyleIndex = (UInt32Value)5U };
			Cell cell22 = new Cell() { CellReference = "B3", StyleIndex = (UInt32Value)6U };
			Cell cell23 = new Cell() { CellReference = "C3", StyleIndex = (UInt32Value)6U };
			Cell cell24 = new Cell() { CellReference = "D3", StyleIndex = (UInt32Value)6U };

			Cell cell25 = new Cell() { CellReference = "E3", StyleIndex = (UInt32Value)7U, DataType = CellValues.SharedString };
			CellValue cellValue9 = new CellValue();
			cellValue9.Text = "8";

			cell25.Append(cellValue9);

			Cell cell26 = new Cell() { CellReference = "F3", StyleIndex = (UInt32Value)7U, DataType = CellValues.String };
			CellValue cellValue10 = new CellValue();
			cellValue10.Text = "金额";

			cell26.Append(cellValue10);

			Cell cell27 = new Cell() { CellReference = "G3", StyleIndex = (UInt32Value)7U, DataType = CellValues.SharedString };
			CellValue cellValue11 = new CellValue();
			cellValue11.Text = "8";

			cell27.Append(cellValue11);

			Cell cell28 = new Cell() { CellReference = "H3", StyleIndex = (UInt32Value)7U, DataType = CellValues.String };
			CellValue cellValue12 = new CellValue();
			cellValue12.Text = "金额";

			cell28.Append(cellValue12);

			Cell cell29 = new Cell() { CellReference = "I3", StyleIndex = (UInt32Value)7U, DataType = CellValues.SharedString };
			CellValue cellValue13 = new CellValue();
			cellValue13.Text = "9";

			cell29.Append(cellValue13);

			Cell cell30 = new Cell() { CellReference = "J3", StyleIndex = (UInt32Value)8U, DataType = CellValues.String };
			CellValue cellValue14 = new CellValue();
			cellValue14.Text = "金额";

			cell30.Append(cellValue14);

			row3.Append(cell21);
			row3.Append(cell22);
			row3.Append(cell23);
			row3.Append(cell24);
			row3.Append(cell25);
			row3.Append(cell26);
			row3.Append(cell27);
			row3.Append(cell28);
			row3.Append(cell29);
			row3.Append(cell30);
			sheetData1.Append(row3);

			MergeCell mergeCell2 = new MergeCell() { Reference = "D2:D3" };
			MergeCell mergeCell8 = new MergeCell() { Reference = "E2:F2" };
			MergeCell mergeCell9 = new MergeCell() { Reference = "G2:H2" };
			MergeCell mergeCell10 = new MergeCell() { Reference = "I2:J2" };
			MergeCell mergeCell11 = new MergeCell() { Reference = "A2:A3" };
			MergeCell mergeCell12 = new MergeCell() { Reference = "B2:B3" };
			MergeCell mergeCell13 = new MergeCell() { Reference = "C2:C3" };

			mergeCells1.Append(mergeCell2);
			mergeCells1.Append(mergeCell8);
			mergeCells1.Append(mergeCell9);
			mergeCells1.Append(mergeCell10);
			mergeCells1.Append(mergeCell11);
			mergeCells1.Append(mergeCell12);
			mergeCells1.Append(mergeCell13);
			#endregion

			toAccountTotal = 0;
			payforTotal = 0;
			reimburseTotal = 0;

			// 项目类型行
			int projectClassifyNo = 0; //项目类型序号
			int sheetRows = 4; //行号，从4开始
			int classNo = 0;

			foreach (int pcid in projectClassifyIDs)
			{
				classNo += 1;
				int classProjects = 0;
				var prc = dataContext.ProjectClassify.Single(p => p.ClassifyId.Equals(pcid));
				string projectClassifyName = prc.ProjectClassify1;
				//添加项目类别标题
				projectClassifyNo += 1;
				#region 项目类型行
				Row row4 = new Row() { RowIndex = (UInt32Value)Convert.ToUInt32(sheetRows), Spans = new ListValue<StringValue>() { InnerText = "1:10" }, Height = 20.100000000000001D, CustomHeight = true, DyDescent = 0.15D };

				//序号
				DigitToChnText dtt = new DigitToChnText();
				Cell cell31 = new Cell() { CellReference = "A" + sheetRows.ToString(), StyleIndex = (UInt32Value)26U, DataType = CellValues.String };
				CellValue cellValue15 = new CellValue();
				cellValue15.Text = dtt.Convert(projectClassifyNo.ToString(), false);

				cell31.Append(cellValue15);

				//项目类型
				Cell cell32 = new Cell() { CellReference = "B" + sheetRows.ToString(), StyleIndex = (UInt32Value)17U, DataType = CellValues.String };
				CellValue cellValue16 = new CellValue();
				cellValue16.Text = projectClassifyName;

				cell32.Append(cellValue16);
				Cell cell33 = new Cell() { CellReference = "C" + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
				Cell cell34 = new Cell() { CellReference = "D" + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
				Cell cell35 = new Cell() { CellReference = "E" + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
				Cell cell36 = new Cell() { CellReference = "F" + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
				Cell cell37 = new Cell() { CellReference = "G" + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
				Cell cell38 = new Cell() { CellReference = "H" + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
				Cell cell39 = new Cell() { CellReference = "I" + sheetRows.ToString(), StyleIndex = (UInt32Value)17U };
				Cell cell40 = new Cell() { CellReference = "J" + sheetRows.ToString(), StyleIndex = (UInt32Value)18U };

				row4.Append(cell31);
				row4.Append(cell32);
				row4.Append(cell33);
				row4.Append(cell34);
				row4.Append(cell35);
				row4.Append(cell36);
				row4.Append(cell37);
				row4.Append(cell38);
				row4.Append(cell39);
				row4.Append(cell40);

				sheetData1.Append(row4);
				MergeCell mergeCell3 = new MergeCell() { Reference = "B" + sheetRows.ToString() + ":J" + sheetRows.ToString() };
				mergeCells1.Append(mergeCell3);

				#endregion
				sheetRows += 1;
				//添加项目
				var pbs = dataContext.ProjectBase.Where(p => p.ProjectClassifyID.Equals(pcid)).OrderBy(p => p.ProjectId);
				foreach (var pb in pbs)
				{
					if (!projectIDs.Contains(pb.ProjectId))
					{
						continue;
					}
					classProjects += 1;
					//添加项目信息及第一行经费
					//获取经费数据
					List<int> toAccountID = new List<int>(); //到账经费ID列表
					List<int> payforID = new List<int>(); //支付外协经费ID列表
					List<int> reimburseID = new List<int>(); //课题组报支经费ID列表
					decimal totalToAccount = 0; //合计
					decimal totalPayfor = 0;
					decimal totalReimburse = 0;
					var toAccount = dataContext.Funds.Where(f => f.ProjectID.Equals(pb.ProjectId) && f.FundClassifyID == 1 && f.Date >= startDate && f.Date <= endDate).OrderBy(f => f.Date);
					foreach (var fund in toAccount)
					{
						toAccountID.Add(fund.Id);
					}
					var payfor = dataContext.Funds.Where(f => f.ProjectID.Equals(pb.ProjectId) && f.FundClassifyID == 2 && f.Date >= startDate && f.Date <= endDate).OrderBy(f => f.Date);
					foreach (var fund in payfor)
					{
						payforID.Add(fund.Id);
					}
					var reimburse = dataContext.Funds.Where(f => f.ProjectID.Equals(pb.ProjectId) && f.FundClassifyID == 3 && f.Date >= startDate && f.Date <= endDate).OrderBy(f => f.Date);
					foreach (var fund in reimburse)
					{
						reimburseID.Add(fund.Id);
					}
					int maxrow = toAccountID.Count >= payforID.Count && toAccountID.Count >= reimburseID.Count ? toAccountID.Count : payforID.Count >= reimburseID.Count ? payforID.Count : reimburseID.Count;
					int startRow = sheetRows;
					Row row5 = new Row() { RowIndex = (UInt32Value)Convert.ToUInt32(sheetRows), Spans = new ListValue<StringValue>() { InnerText = "1:10" }, Height = 20.100000000000001D, CustomHeight = true, DyDescent = 0.15D };

					if (sheetRows != totalRows - 1)
					{
						//不是最后一行
						//序号
						Cell cell41 = new Cell() { CellReference = "A" + sheetRows.ToString(), StyleIndex = (UInt32Value)14U, DataType = CellValues.String };
						CellValue cellValue17 = new CellValue();
						cellValue17.Text = classProjects.ToString();

						cell41.Append(cellValue17);

						//编号
						Cell cell42 = new Cell() { CellReference = "B" + sheetRows.ToString(), StyleIndex = (UInt32Value)13U, DataType = CellValues.String };
						CellValue cellValue18 = new CellValue();
						cellValue18.Text = pb.ProjectNo;

						cell42.Append(cellValue18);

						//项目名称
						Cell cell43 = new Cell() { CellReference = "C" + sheetRows.ToString(), StyleIndex = (UInt32Value)21U, DataType = CellValues.String };
						CellValue cellValue19 = new CellValue();
						cellValue19.Text = pb.ProjectName;

						cell43.Append(cellValue19);

						//负责人
						Cell cell44 = new Cell() { CellReference = "D" + sheetRows.ToString(), StyleIndex = (UInt32Value)13U, DataType = CellValues.String };
						CellValue cellValue20 = new CellValue();
						cellValue20.Text = pb.Principal;

						cell44.Append(cellValue20);
						

						//第一笔到账
						//时间

						Cell cell45 = new Cell() { CellReference = "E" + sheetRows.ToString(), StyleIndex = (UInt32Value)9U, DataType = CellValues.String };
						CellValue cellValue21 = new CellValue();
						if (toAccountID.Count > 0)
						{
							Funds funds = dataContext.Funds.SingleOrDefault(f => f.Id.Equals(toAccountID[0]));
							cellValue21.Text = ((DateTime)funds.Date).ToString("yyyyMMdd");
						}
						else
						{
							cellValue21.Text = "";
						}

						cell45.Append(cellValue21);

						//金额
						Cell cell46 = new Cell() { CellReference = "F" + sheetRows.ToString(), StyleIndex = (UInt32Value)9U, DataType = CellValues.String };
						CellValue cellValue22 = new CellValue();
						if (toAccountID.Count > 0)
						{
							Funds funds = dataContext.Funds.SingleOrDefault(f => f.Id.Equals(toAccountID[0]));
                            cellValue22.Text = string.Format("{0:N2}", funds.Money);
							totalToAccount += (decimal)funds.Money;
							toAccountTotal += (decimal)funds.Money;
						}
						else
						{
							cellValue22.Text = "";
						}

						cell46.Append(cellValue22);

						//第一笔支付外协
						//时间

						Cell cell47 = new Cell() { CellReference = "G" + sheetRows.ToString(), StyleIndex = (UInt32Value)9U, DataType = CellValues.String };
						CellValue cellValue23 = new CellValue();
						if (payforID.Count > 0)
						{
							Funds funds = dataContext.Funds.SingleOrDefault(f => f.Id.Equals(payforID[0]));
							cellValue23.Text = ((DateTime)funds.Date).ToString("yyyyMMdd");
						}
						else
						{
							cellValue23.Text = "";
						}

						cell47.Append(cellValue23);

						//金额
						Cell cell48 = new Cell() { CellReference = "H" + sheetRows.ToString(), StyleIndex = (UInt32Value)9U, DataType = CellValues.String };
						CellValue cellValue24 = new CellValue();
						if (payforID.Count > 0)
						{
							Funds funds = dataContext.Funds.SingleOrDefault(f => f.Id.Equals(payforID[0]));
                            cellValue24.Text = string.Format("{0:N2}", funds.Money);
							totalPayfor += (decimal)funds.Money;
							payforTotal += (decimal)funds.Money;
						}
						else
						{
							cellValue24.Text = "";
						}

						cell48.Append(cellValue24);

						//第一笔项目组报支
						//时间
						Cell cell49 = new Cell() { CellReference = "I" + sheetRows.ToString(), StyleIndex = (UInt32Value)9U, DataType = CellValues.String };
						CellValue cellValue25 = new CellValue();
						if (reimburseID.Count > 0)
						{
							Funds funds = dataContext.Funds.SingleOrDefault(f => f.Id.Equals(reimburseID[0]));
							cellValue25.Text = ((DateTime)funds.Date).ToString("yyyyMMdd");
						}
						else
						{
							cellValue25.Text = "";
						}

						cell49.Append(cellValue25);

						//金额
						Cell cell50 = new Cell() { CellReference = "J" + sheetRows.ToString(), StyleIndex = (UInt32Value)10U, DataType = CellValues.String };
						CellValue cellValue26 = new CellValue();
						if (reimburseID.Count > 0)
						{
							Funds funds = dataContext.Funds.SingleOrDefault(f => f.Id.Equals(reimburseID[0]));
                            cellValue26.Text = string.Format("{0:N2}", funds.Money);
							totalReimburse += (decimal)funds.Money;
							reimburseTotal += (decimal)funds.Money;
						}
						else
						{
							cellValue26.Text = "";
						}

						cell50.Append(cellValue26);

						row5.Append(cell41);
						row5.Append(cell42);
						row5.Append(cell43);
						row5.Append(cell44);
						row5.Append(cell45);
						row5.Append(cell46);
						row5.Append(cell47);
						row5.Append(cell48);
						row5.Append(cell49);
						row5.Append(cell50);
					}
					else
					{
						//最后一行
						//序号
						Cell cell41 = new Cell() { CellReference = "A" + sheetRows.ToString(), StyleIndex = (UInt32Value)19U, DataType = CellValues.String };
						CellValue cellValue17 = new CellValue();
						cellValue17.Text = classProjects.ToString();

						cell41.Append(cellValue17);

						//编号
						Cell cell42 = new Cell() { CellReference = "B" + sheetRows.ToString(), StyleIndex = (UInt32Value)20U, DataType = CellValues.String };
						CellValue cellValue18 = new CellValue();
						cellValue18.Text = pb.ProjectNo;

						cell42.Append(cellValue18);

						//项目名称
						Cell cell43 = new Cell() { CellReference = "C" + sheetRows.ToString(), StyleIndex = (UInt32Value)22U, DataType = CellValues.String };
						CellValue cellValue19 = new CellValue();
						cellValue19.Text = pb.ProjectName;

						cell43.Append(cellValue19);

						//负责人
						Cell cell44 = new Cell() { CellReference = "D" + sheetRows.ToString(), StyleIndex = (UInt32Value)20U, DataType = CellValues.String };
						CellValue cellValue20 = new CellValue();
						cellValue20.Text = pb.Principal;

						cell44.Append(cellValue20);
						

						//第一笔到账
						//时间

						Cell cell45 = new Cell() { CellReference = "E" + sheetRows.ToString(), StyleIndex = (UInt32Value)11U, DataType = CellValues.String };
						CellValue cellValue21 = new CellValue();
						if (toAccountID.Count > 0)
						{
							Funds funds = dataContext.Funds.SingleOrDefault(f => f.Id.Equals(toAccountID[0]));
							cellValue21.Text = ((DateTime)funds.Date).ToString("yyyyMMdd");
						}
						else
						{
							cellValue21.Text = "";
						}

						cell45.Append(cellValue21);

						//金额
						Cell cell46 = new Cell() { CellReference = "F" + sheetRows.ToString(), StyleIndex = (UInt32Value)11U, DataType = CellValues.String };
						CellValue cellValue22 = new CellValue();
						if (toAccountID.Count > 0)
						{
							Funds funds = dataContext.Funds.SingleOrDefault(f => f.Id.Equals(toAccountID[0]));
                            cellValue22.Text = string.Format("{0:N2}", funds.Money);
							totalToAccount += (decimal)funds.Money;
							toAccountTotal += (decimal)funds.Money;
						}
						else
						{
							cellValue22.Text = "";
						}

						cell46.Append(cellValue22);

						//第一笔支付外协
						//时间

						Cell cell47 = new Cell() { CellReference = "G" + sheetRows.ToString(), StyleIndex = (UInt32Value)11U, DataType = CellValues.String };
						CellValue cellValue23 = new CellValue();
						if (payforID.Count > 0)
						{
							Funds funds = dataContext.Funds.SingleOrDefault(f => f.Id.Equals(payforID[0]));
							cellValue23.Text = ((DateTime)funds.Date).ToString("yyyyMMdd");
						}
						else
						{
							cellValue23.Text = "";
						}

						cell47.Append(cellValue23);

						//金额
						Cell cell48 = new Cell() { CellReference = "H" + sheetRows.ToString(), StyleIndex = (UInt32Value)11U, DataType = CellValues.String };
						CellValue cellValue24 = new CellValue();
						if (payforID.Count > 0)
						{
							Funds funds = dataContext.Funds.SingleOrDefault(f => f.Id.Equals(payforID[0]));
                            cellValue24.Text = string.Format("{0:N2}", funds.Money);
							totalPayfor += (decimal)funds.Money;
							payforTotal += (decimal)funds.Money;
						}
						else
						{
							cellValue24.Text = "";
						}

						cell48.Append(cellValue24);

						//第一笔项目组报支
						//时间
						Cell cell49 = new Cell() { CellReference = "I" + sheetRows.ToString(), StyleIndex = (UInt32Value)11U, DataType = CellValues.String };
						CellValue cellValue25 = new CellValue();
						if (reimburseID.Count > 0)
						{
							Funds funds = dataContext.Funds.SingleOrDefault(f => f.Id.Equals(reimburseID[0]));
							cellValue25.Text = ((DateTime)funds.Date).ToString("yyyyMMdd");
						}
						else
						{
							cellValue25.Text = "";
						}

						cell49.Append(cellValue25);

						//金额
						Cell cell50 = new Cell() { CellReference = "J" + sheetRows.ToString(), StyleIndex = (UInt32Value)12U, DataType = CellValues.String };
						CellValue cellValue26 = new CellValue();
						if (reimburseID.Count > 0)
						{
							Funds funds = dataContext.Funds.SingleOrDefault(f => f.Id.Equals(reimburseID[0]));
                            cellValue26.Text = string.Format("{0:N2}", funds.Money);
							totalReimburse += (decimal)funds.Money;
							reimburseTotal += (decimal)funds.Money;
						}
						else
						{
							cellValue26.Text = "";
						}

						cell50.Append(cellValue26);

						row5.Append(cell41);
						row5.Append(cell42);
						row5.Append(cell43);
						row5.Append(cell44);
						row5.Append(cell45);
						row5.Append(cell46);
						row5.Append(cell47);
						row5.Append(cell48);
						row5.Append(cell49);
						row5.Append(cell50);
					}
					sheetData1.Append(row5);
					sheetRows += 1;
					if (maxrow == 1)
					{
						//只有一行
						continue;
					}
					//不止一行
					for (int i = 1; i < maxrow; i++)
					{
						Row row6 = new Row() { RowIndex = (UInt32Value)Convert.ToUInt32(sheetRows), Spans = new ListValue<StringValue>() { InnerText = "1:10" }, Height = 20.100000000000001D, CustomHeight = true, DyDescent = 0.15D };
						if (sheetRows != totalRows - 1)
						{
							//不是最后一行
							Cell cell51 = new Cell() { CellReference = "A" + sheetRows.ToString(), StyleIndex = (UInt32Value)15U };
							Cell cell52 = new Cell() { CellReference = "B" + sheetRows.ToString(), StyleIndex = (UInt32Value)16U };
							Cell cell53 = new Cell() { CellReference = "C" + sheetRows.ToString(), StyleIndex = (UInt32Value)23U };
							Cell cell54 = new Cell() { CellReference = "D" + sheetRows.ToString(), StyleIndex = (UInt32Value)16U };

							//到账
							//日期
							Cell cell55 = new Cell() { CellReference = "E" + sheetRows.ToString(), StyleIndex = (UInt32Value)24U, DataType = CellValues.String };
							CellValue cellValue27 = new CellValue();
							if (toAccountID.Count > i)
							{
								Funds funds = dataContext.Funds.SingleOrDefault(f => f.Id.Equals(toAccountID[i]));
								cellValue27.Text = ((DateTime)funds.Date).ToString("yyyyMMdd");
							}
							else
							{
								cellValue27.Text = "";
							}

							cell55.Append(cellValue27);

							//金额
							Cell cell56 = new Cell() { CellReference = "F" + sheetRows.ToString(), StyleIndex = (UInt32Value)24U, DataType = CellValues.String };
							CellValue cellValue28 = new CellValue();
							if (toAccountID.Count > i)
							{
								Funds funds = dataContext.Funds.SingleOrDefault(f => f.Id.Equals(toAccountID[i]));
                                cellValue28.Text = string.Format("{0:N2}", funds.Money);
								totalToAccount += (decimal)funds.Money;
								toAccountTotal += (decimal)funds.Money;
							}
							else
							{
								cellValue28.Text = "";
							}

							cell56.Append(cellValue28);

							//支付外协
							//日期
							Cell cell57 = new Cell() { CellReference = "G" + sheetRows.ToString(), StyleIndex = (UInt32Value)24U, DataType = CellValues.String };
							CellValue cellValue29 = new CellValue();
							if (payforID.Count > i)
							{
								Funds funds = dataContext.Funds.SingleOrDefault(f => f.Id.Equals(payforID[i]));
								cellValue29.Text = ((DateTime)funds.Date).ToString("yyyyMMdd");
							}
							else
							{
								cellValue29.Text = "";
							}

							cell57.Append(cellValue29);
							//金额
							Cell cell58 = new Cell() { CellReference = "H" + sheetRows.ToString(), StyleIndex = (UInt32Value)24U, DataType = CellValues.String };
							CellValue cellValue30 = new CellValue();
							if (payforID.Count > i)
							{
								Funds funds = dataContext.Funds.SingleOrDefault(f => f.Id.Equals(payforID[i]));
                                cellValue30.Text = string.Format("{0:N2}", funds.Money);
								totalPayfor += (decimal)funds.Money;
								payforTotal += (decimal)funds.Money;
							}
							else
							{
								cellValue30.Text = "";
							}

							cell58.Append(cellValue30);

							//课题组报支
							//日期
							Cell cell59 = new Cell() { CellReference = "I" + sheetRows.ToString(), StyleIndex = (UInt32Value)24U, DataType = CellValues.String };
							CellValue cellValue31 = new CellValue();
							if (reimburseID.Count > i)
							{
								Funds funds = dataContext.Funds.SingleOrDefault(f => f.Id.Equals(reimburseID[i]));
								cellValue31.Text = ((DateTime)funds.Date).ToString("yyyyMMdd");
							}
							else
							{
								cellValue31.Text = "";
							}

							cell59.Append(cellValue31);
							//金额
							Cell cell60 = new Cell() { CellReference = "J" + sheetRows.ToString(), StyleIndex = (UInt32Value)25U, DataType = CellValues.String };
							CellValue cellValue32 = new CellValue();
							if (reimburseID.Count > i)
							{
								Funds funds = dataContext.Funds.SingleOrDefault(f => f.Id.Equals(reimburseID[i]));
                                cellValue32.Text = string.Format("{0:N2}", funds.Money);
								totalReimburse += (decimal)funds.Money;
								reimburseTotal += (decimal)funds.Money;
							}
							else
							{
								cellValue32.Text = "";
							}

							cell60.Append(cellValue32);

							row6.Append(cell51);
							row6.Append(cell52);
							row6.Append(cell53);
							row6.Append(cell54);
							row6.Append(cell55);
							row6.Append(cell56);
							row6.Append(cell57);
							row6.Append(cell58);
							row6.Append(cell59);
							row6.Append(cell60);
						}
						else
						{
							//最后一行
							Cell cell51 = new Cell() { CellReference = "A" + sheetRows.ToString(), StyleIndex = (UInt32Value)19U };
							Cell cell52 = new Cell() { CellReference = "B" + sheetRows.ToString(), StyleIndex = (UInt32Value)20U };
							Cell cell53 = new Cell() { CellReference = "C" + sheetRows.ToString(), StyleIndex = (UInt32Value)22U };
							Cell cell54 = new Cell() { CellReference = "D" + sheetRows.ToString(), StyleIndex = (UInt32Value)20U };

							//到账
							//日期
							Cell cell55 = new Cell() { CellReference = "E" + sheetRows.ToString(), StyleIndex = (UInt32Value)11U, DataType = CellValues.String };
							CellValue cellValue27 = new CellValue();
							if (toAccountID.Count > i)
							{
								Funds funds = dataContext.Funds.SingleOrDefault(f => f.Id.Equals(toAccountID[i]));
								cellValue27.Text = ((DateTime)funds.Date).ToString("yyyyMMdd");
							}
							else
							{
								cellValue27.Text = "";
							}

							cell55.Append(cellValue27);

							//金额
							Cell cell56 = new Cell() { CellReference = "F" + sheetRows.ToString(), StyleIndex = (UInt32Value)11U, DataType = CellValues.String };
							CellValue cellValue28 = new CellValue();
							if (toAccountID.Count > i)
							{
								Funds funds = dataContext.Funds.SingleOrDefault(f => f.Id.Equals(toAccountID[i]));
                                cellValue28.Text = string.Format("{0:N2}", funds.Money);
								totalToAccount += (decimal)funds.Money;
								toAccountTotal += (decimal)funds.Money;
							}
							else
							{
								cellValue28.Text = "";
							}

							cell56.Append(cellValue28);

							//支付外协
							//日期
							Cell cell57 = new Cell() { CellReference = "G" + sheetRows.ToString(), StyleIndex = (UInt32Value)11U, DataType = CellValues.String };
							CellValue cellValue29 = new CellValue();
							if (payforID.Count > i)
							{
								Funds funds = dataContext.Funds.SingleOrDefault(f => f.Id.Equals(payforID[i]));
								cellValue29.Text = ((DateTime)funds.Date).ToString("yyyyMMdd");
							}
							else
							{
								cellValue29.Text = "";
							}

							cell57.Append(cellValue29);
							//金额
							Cell cell58 = new Cell() { CellReference = "H" + sheetRows.ToString(), StyleIndex = (UInt32Value)11U, DataType = CellValues.String };
							CellValue cellValue30 = new CellValue();
							if (payforID.Count > i)
							{
								Funds funds = dataContext.Funds.SingleOrDefault(f => f.Id.Equals(payforID[i]));
                                cellValue30.Text = string.Format("{0:N2}", funds.Money);
								totalPayfor += (decimal)funds.Money;
								payforTotal += (decimal)funds.Money;
							}
							else
							{
								cellValue30.Text = "";
							}

							cell58.Append(cellValue30);

							//课题组报支
							//日期
							Cell cell59 = new Cell() { CellReference = "I" + sheetRows.ToString(), StyleIndex = (UInt32Value)11U, DataType = CellValues.String };
							CellValue cellValue31 = new CellValue();
							if (reimburseID.Count > i)
							{
								Funds funds = dataContext.Funds.SingleOrDefault(f => f.Id.Equals(reimburseID[i]));
								cellValue31.Text = ((DateTime)funds.Date).ToString("yyyyMMdd");
							}
							else
							{
								cellValue31.Text = "";
							}

							cell59.Append(cellValue31);
							//金额
							Cell cell60 = new Cell() { CellReference = "J" + sheetRows.ToString(), StyleIndex = (UInt32Value)12U, DataType = CellValues.String };
							CellValue cellValue32 = new CellValue();
							if (reimburseID.Count > i)
							{
								Funds funds = dataContext.Funds.SingleOrDefault(f => f.Id.Equals(reimburseID[i]));
                                cellValue32.Text = string.Format("{0:N2}", funds.Money);
								totalReimburse += (decimal)funds.Money;
								reimburseTotal += (decimal)funds.Money;
							}
							else
							{
								cellValue32.Text = "";
							}

							cell60.Append(cellValue32);

							row6.Append(cell51);
							row6.Append(cell52);
							row6.Append(cell53);
							row6.Append(cell54);
							row6.Append(cell55);
							row6.Append(cell56);
							row6.Append(cell57);
							row6.Append(cell58);
							row6.Append(cell59);
							row6.Append(cell60);
						}
						sheetData1.Append(row6);
						sheetRows += 1;
                    }
                    #region 合计行
                    //添加合计行
                    //Row row7 = new Row() { RowIndex = (UInt32Value)Convert.ToUInt32(sheetRows), Spans = new ListValue<StringValue>() { InnerText = "1:10" }, Height = 20.100000000000001D, CustomHeight = true, ThickBot = true, DyDescent = 0.2D };
                    //if (sheetRows != totalRows - 1)
                    //{
                    //    //不是最后一行
                    //    Cell cell61 = new Cell() { CellReference = "A" + sheetRows.ToString(), StyleIndex = (UInt32Value)15U };
                    //    Cell cell62 = new Cell() { CellReference = "B" + sheetRows.ToString(), StyleIndex = (UInt32Value)16U };
                    //    Cell cell63 = new Cell() { CellReference = "C" + sheetRows.ToString(), StyleIndex = (UInt32Value)23U };
                    //    Cell cell64 = new Cell() { CellReference = "D" + sheetRows.ToString(), StyleIndex = (UInt32Value)16U };

                    //    Cell cell65 = new Cell() { CellReference = "E" + sheetRows.ToString(), StyleIndex = (UInt32Value)24U, DataType = CellValues.String };
                    //    CellValue cellValue33 = new CellValue();
                    //    cellValue33.Text = "合计";

                    //    cell65.Append(cellValue33);

                    //    //到账合计
                    //    Cell cell66 = new Cell() { CellReference = "F" + sheetRows.ToString(), StyleIndex = (UInt32Value)24U, DataType = CellValues.Number };
                    //    CellValue cellValue34 = new CellValue();
                    //    cellValue34.Text = totalToAccount.ToString();

                    //    cell66.Append(cellValue34);
                    //    Cell cell67 = new Cell() { CellReference = "G" + sheetRows.ToString(), StyleIndex = (UInt32Value)24U };

                    //    //支付外协合计
                    //    Cell cell68 = new Cell() { CellReference = "H" + sheetRows.ToString(), StyleIndex = (UInt32Value)24U, DataType = CellValues.Number };
                    //    CellValue cellValue35 = new CellValue();
                    //    cellValue35.Text = totalPayfor.ToString();

                    //    cell68.Append(cellValue35);
                    //    Cell cell69 = new Cell() { CellReference = "I" + sheetRows.ToString(), StyleIndex = (UInt32Value)24U };

                    //    //课题组报支合计
                    //    Cell cell70 = new Cell() { CellReference = "J" + sheetRows.ToString(), StyleIndex = (UInt32Value)25U, DataType = CellValues.Number };
                    //    CellValue cellValue36 = new CellValue();
                    //    cellValue36.Text = totalReimburse.ToString();

                    //    cell70.Append(cellValue36);

                    //    row7.Append(cell61);
                    //    row7.Append(cell62);
                    //    row7.Append(cell63);
                    //    row7.Append(cell64);
                    //    row7.Append(cell65);
                    //    row7.Append(cell66);
                    //    row7.Append(cell67);
                    //    row7.Append(cell68);
                    //    row7.Append(cell69);
                    //    row7.Append(cell70);
                    //}
                    //else
                    //{
                    //    //最后一行
                    //    Cell cell61 = new Cell() { CellReference = "A" + sheetRows.ToString(), StyleIndex = (UInt32Value)19U };
                    //    Cell cell62 = new Cell() { CellReference = "B" + sheetRows.ToString(), StyleIndex = (UInt32Value)20U };
                    //    Cell cell63 = new Cell() { CellReference = "C" + sheetRows.ToString(), StyleIndex = (UInt32Value)22U };
                    //    Cell cell64 = new Cell() { CellReference = "D" + sheetRows.ToString(), StyleIndex = (UInt32Value)20U };

                    //    Cell cell65 = new Cell() { CellReference = "E" + sheetRows.ToString(), StyleIndex = (UInt32Value)11U, DataType = CellValues.String };
                    //    CellValue cellValue33 = new CellValue();
                    //    cellValue33.Text = "合计";

                    //    cell65.Append(cellValue33);

                    //    //到账合计
                    //    Cell cell66 = new Cell() { CellReference = "F" + sheetRows.ToString(), StyleIndex = (UInt32Value)11U, DataType = CellValues.Number };
                    //    CellValue cellValue34 = new CellValue();
                    //    cellValue34.Text = totalToAccount.ToString();

                    //    cell66.Append(cellValue34);
                    //    Cell cell67 = new Cell() { CellReference = "G" + sheetRows.ToString(), StyleIndex = (UInt32Value)11U };

                    //    //支付外协合计
                    //    Cell cell68 = new Cell() { CellReference = "H" + sheetRows.ToString(), StyleIndex = (UInt32Value)11U, DataType = CellValues.Number };
                    //    CellValue cellValue35 = new CellValue();
                    //    cellValue35.Text = totalPayfor.ToString();

                    //    cell68.Append(cellValue35);
                    //    Cell cell69 = new Cell() { CellReference = "I" + sheetRows.ToString(), StyleIndex = (UInt32Value)11U };

                    //    //课题组报支合计
                    //    Cell cell70 = new Cell() { CellReference = "J" + sheetRows.ToString(), StyleIndex = (UInt32Value)12U, DataType = CellValues.Number };
                    //    CellValue cellValue36 = new CellValue();
                    //    cellValue36.Text = totalReimburse.ToString();

                    //    cell70.Append(cellValue36);

                    //    row7.Append(cell61);
                    //    row7.Append(cell62);
                    //    row7.Append(cell63);
                    //    row7.Append(cell64);
                    //    row7.Append(cell65);
                    //    row7.Append(cell66);
                    //    row7.Append(cell67);
                    //    row7.Append(cell68);
                    //    row7.Append(cell69);
                    //    row7.Append(cell70);
                    //}
                    //sheetData1.Append(row7);
                    #endregion
                    //添加合并单元格
					MergeCell mergeCell4 = new MergeCell() { Reference = "A" + startRow.ToString() + ":A" + (sheetRows -1).ToString() };
					MergeCell mergeCell5 = new MergeCell() { Reference = "B" + startRow.ToString() + ":B" + (sheetRows -1).ToString() };
					MergeCell mergeCell6 = new MergeCell() { Reference = "C" + startRow.ToString() + ":C" + (sheetRows -1).ToString() };
					MergeCell mergeCell7 = new MergeCell() { Reference = "D" + startRow.ToString() + ":D" + (sheetRows -1).ToString() };

					mergeCells1.Append(mergeCell4);
					mergeCells1.Append(mergeCell5);
					mergeCells1.Append(mergeCell6);
					mergeCells1.Append(mergeCell7);

					//sheetRows += 1;
				}
			}
			
			//添加总计行
			Row row8 = new Row() { RowIndex = (UInt32Value)Convert.ToUInt32(sheetRows), Spans = new ListValue<StringValue>() { InnerText = "1:10" }, Height = 20D, CustomHeight = true, ThickBot = true, DyDescent = 0.2D };

			Cell cell71 = new Cell() { CellReference = "A" + sheetRows.ToString(), StyleIndex = (UInt32Value)19U, DataType = CellValues.String };
			CellValue cellValue37 = new CellValue();
			cellValue37.Text = "总  计";
			cell71.Append(cellValue37);

			Cell cell72 = new Cell() { CellReference = "B" + sheetRows.ToString(), StyleIndex = (UInt32Value)20U };
			Cell cell73 = new Cell() { CellReference = "C" + sheetRows.ToString(), StyleIndex = (UInt32Value)22U };
			Cell cell74 = new Cell() { CellReference = "D" + sheetRows.ToString(), StyleIndex = (UInt32Value)20U };

			//到账总计
			Cell cell75 = new Cell() { CellReference = "E" + sheetRows.ToString(), StyleIndex = (UInt32Value)11U, DataType = CellValues.String };
			CellValue cellValue38 = new CellValue();
            cellValue38.Text = string.Format("{0:N2}", toAccountTotal);

			cell75.Append(cellValue38);

			Cell cell76 = new Cell() { CellReference = "F" + sheetRows.ToString(), StyleIndex = (UInt32Value)11U };

			//支付外协总计
			Cell cell77 = new Cell() { CellReference = "G" + sheetRows.ToString(), StyleIndex = (UInt32Value)11U, DataType = CellValues.String };
			CellValue cellValue39 = new CellValue();
			cellValue39.Text = string.Format("{0:N2}", payforTotal);

			cell77.Append(cellValue39);

			Cell cell78 = new Cell() { CellReference = "H" + sheetRows.ToString(), StyleIndex = (UInt32Value)11U };

			//课题组报支总计
			Cell cell79 = new Cell() { CellReference = "I" + sheetRows.ToString(), StyleIndex = (UInt32Value)11U, DataType = CellValues.String };
			CellValue cellValue40 = new CellValue();
			cellValue40.Text = string.Format("{0:N2}", reimburseTotal);

			cell79.Append(cellValue40);

			Cell cell80 = new Cell() { CellReference = "J" + sheetRows.ToString(), StyleIndex = (UInt32Value)12U };

			row8.Append(cell71);
			row8.Append(cell72);
			row8.Append(cell73);
			row8.Append(cell74);
			row8.Append(cell75);
			row8.Append(cell76);
			row8.Append(cell77);
			row8.Append(cell78);
			row8.Append(cell79);
			row8.Append(cell80);

			sheetData1.Append(row8);

			//添加合并单元格
			MergeCell mergeCell20 = new MergeCell() { Reference = "A" + sheetRows.ToString() + ":D" + sheetRows.ToString() };
			MergeCell mergeCell21 = new MergeCell() { Reference = "E" + sheetRows.ToString() + ":F" + sheetRows.ToString() };
			MergeCell mergeCell22 = new MergeCell() { Reference = "G" + sheetRows.ToString() + ":H" + sheetRows.ToString() };
			MergeCell mergeCell23 = new MergeCell() { Reference = "I" + sheetRows.ToString() + ":J" + sheetRows.ToString() };

			mergeCells1.Append(mergeCell20);
			mergeCells1.Append(mergeCell21);
			mergeCells1.Append(mergeCell22);
			mergeCells1.Append(mergeCell23);

			PhoneticProperties phoneticProperties1 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };
			PageMargins pageMargins1 = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
			PageSetup pageSetup1 = new PageSetup() { PaperSize = (UInt32Value)8U, Orientation = OrientationValues.Landscape, HorizontalDpi = (UInt32Value)4294967295U, VerticalDpi = (UInt32Value)4294967295U, Id = "rId1" };

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
			SharedStringTable sharedStringTable1 = new SharedStringTable() { Count = (UInt32Value)26U, UniqueCount = (UInt32Value)21U };

			SharedStringItem sharedStringItem1 = new SharedStringItem();
			Text text1 = new Text();
			text1.Text = "年度科研经费收支一览表";
			PhoneticProperties phoneticProperties2 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

			sharedStringItem1.Append(text1);
			sharedStringItem1.Append(phoneticProperties2);

			SharedStringItem sharedStringItem2 = new SharedStringItem();
			Text text2 = new Text();
			text2.Text = "序号";
			PhoneticProperties phoneticProperties3 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

			sharedStringItem2.Append(text2);
			sharedStringItem2.Append(phoneticProperties3);

			SharedStringItem sharedStringItem3 = new SharedStringItem();
			Text text3 = new Text();
			text3.Text = "编号";
			PhoneticProperties phoneticProperties4 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

			sharedStringItem3.Append(text3);
			sharedStringItem3.Append(phoneticProperties4);

			SharedStringItem sharedStringItem4 = new SharedStringItem();
			Text text4 = new Text();
			text4.Text = "项目名称";
			PhoneticProperties phoneticProperties5 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

			sharedStringItem4.Append(text4);
			sharedStringItem4.Append(phoneticProperties5);

			SharedStringItem sharedStringItem5 = new SharedStringItem();
			Text text5 = new Text();
			text5.Text = "负责人";
			PhoneticProperties phoneticProperties6 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

			sharedStringItem5.Append(text5);
			sharedStringItem5.Append(phoneticProperties6);

			SharedStringItem sharedStringItem6 = new SharedStringItem();
			Text text6 = new Text();
			text6.Text = "到账";
			PhoneticProperties phoneticProperties7 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

			sharedStringItem6.Append(text6);
			sharedStringItem6.Append(phoneticProperties7);

			SharedStringItem sharedStringItem7 = new SharedStringItem();
			Text text7 = new Text();
			text7.Text = "支付外协";
			PhoneticProperties phoneticProperties8 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

			sharedStringItem7.Append(text7);
			sharedStringItem7.Append(phoneticProperties8);

			SharedStringItem sharedStringItem8 = new SharedStringItem();
			Text text8 = new Text();
			text8.Text = "课题组报支";
			PhoneticProperties phoneticProperties9 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

			sharedStringItem8.Append(text8);
			sharedStringItem8.Append(phoneticProperties9);

			SharedStringItem sharedStringItem9 = new SharedStringItem();
			Text text9 = new Text();
			text9.Text = "时间";
			PhoneticProperties phoneticProperties10 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

			sharedStringItem9.Append(text9);
			sharedStringItem9.Append(phoneticProperties10);

			SharedStringItem sharedStringItem10 = new SharedStringItem();
			Text text10 = new Text();
			text10.Text = "时间";
			PhoneticProperties phoneticProperties11 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

			sharedStringItem10.Append(text10);
			sharedStringItem10.Append(phoneticProperties11);

			SharedStringItem sharedStringItem11 = new SharedStringItem();
			Text text11 = new Text();
			text11.Text = "一、";
			PhoneticProperties phoneticProperties12 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

			sharedStringItem11.Append(text11);
			sharedStringItem11.Append(phoneticProperties12);

			SharedStringItem sharedStringItem12 = new SharedStringItem();
			Text text12 = new Text();
			text12.Text = "西部项目";
			PhoneticProperties phoneticProperties13 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

			sharedStringItem12.Append(text12);
			sharedStringItem12.Append(phoneticProperties13);

			SharedStringItem sharedStringItem13 = new SharedStringItem();
			Text text13 = new Text();
			text13.Text = "A";
			PhoneticProperties phoneticProperties14 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

			sharedStringItem13.Append(text13);
			sharedStringItem13.Append(phoneticProperties14);

			SharedStringItem sharedStringItem14 = new SharedStringItem();
			Text text14 = new Text();
			text14.Text = "B";
			PhoneticProperties phoneticProperties15 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

			sharedStringItem14.Append(text14);
			sharedStringItem14.Append(phoneticProperties15);

			SharedStringItem sharedStringItem15 = new SharedStringItem();
			Text text15 = new Text();
			text15.Text = "金额（万元）";
			PhoneticProperties phoneticProperties16 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

			sharedStringItem15.Append(text15);
			sharedStringItem15.Append(phoneticProperties16);

			SharedStringItem sharedStringItem16 = new SharedStringItem();
			Text text16 = new Text();
			text16.Text = "金额（万元）";
			PhoneticProperties phoneticProperties17 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

			sharedStringItem16.Append(text16);
			sharedStringItem16.Append(phoneticProperties17);

			SharedStringItem sharedStringItem17 = new SharedStringItem();
			Text text17 = new Text();
			text17.Text = "C";
			PhoneticProperties phoneticProperties18 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

			sharedStringItem17.Append(text17);
			sharedStringItem17.Append(phoneticProperties18);

			SharedStringItem sharedStringItem18 = new SharedStringItem();
			Text text18 = new Text();
			text18.Text = "2014.1.1";
			PhoneticProperties phoneticProperties19 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

			sharedStringItem18.Append(text18);
			sharedStringItem18.Append(phoneticProperties19);

			SharedStringItem sharedStringItem19 = new SharedStringItem();
			Text text19 = new Text();
			text19.Text = "2014.1.1";
			PhoneticProperties phoneticProperties20 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

			sharedStringItem19.Append(text19);
			sharedStringItem19.Append(phoneticProperties20);

			SharedStringItem sharedStringItem20 = new SharedStringItem();
			Text text20 = new Text();
			text20.Text = "2014.1.2";
			PhoneticProperties phoneticProperties21 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

			sharedStringItem20.Append(text20);
			sharedStringItem20.Append(phoneticProperties21);

			SharedStringItem sharedStringItem21 = new SharedStringItem();
			Text text21 = new Text();
			text21.Text = "合计";
			PhoneticProperties phoneticProperties22 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };

			sharedStringItem21.Append(text21);
			sharedStringItem21.Append(phoneticProperties22);

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

			sharedStringTablePart1.SharedStringTable = sharedStringTable1;
		}

		private void SetPackageProperties(OpenXmlPackage document)
		{
			document.PackageProperties.Creator = "蒋惠林";
			document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2014-11-03T00:17:32Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
			document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2014-11-03T03:35:27Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
			document.PackageProperties.LastModifiedBy = "蒋惠林";
			document.PackageProperties.LastPrinted = System.Xml.XmlConvert.ToDateTime("2014-11-03T00:39:03Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
		}

		#region Binary Data
		private string spreadsheetPrinterSettingsPart1Data = "SABQACAATABhAHMAZQByAEoAZQB0ACAANQAyADAAMAAgAFAAQwBMADYAIABDAGwAYQBzAHMAIABEAHIAAAAAAAEEAwbcAAgEQ78AAgIACACaCzQIZAABAA8A//8BAAEA//8DAAEAQQA0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAAAAAABAAAAAgAAAEQBAAD/////R0lTNAAAAAAAAAAAAAAAAERJTlUiAJAB7AMcAN55QG4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADQAAAAEAAAAAAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACQAQAAU01USgAAAAAQAIABewAzADgARQA3AEIANwA0ADYALQA0ADYARgBFAC0ANABhADEAZAAtADkANgA3AEQALQA1AEIAQgA2ADEAQwBEAEMARQA3ADQANQB9AAAASW5wdXRCaW4AQXV0b1NlbGVjdABSRVNETEwAVW5pcmVzRExMAFBhcGVyU2l6ZQBMRVRURVIAT3JpZW50YXRpb24AUE9SVFJBSVQATWVkaWFUeXBlAEF1dG8AUmVzb2x1dGlvbgA2MDBEUEkAUGFnZU91dHB1dFF1YWxpdHkATm9ybWFsAENvbG9yTW9kZQBNb25vAERvY3VtZW50TlVwADEAQ29sbGF0ZQBPTgBEdXBsZXgATk9ORQBPdXRwdXRCaW4AQXV0bwBTdGFwbGluZwBOb25lAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAHAAAAFY0RE0BAAAAAAAAAAAAAAAAAAAAAAAAAA==";

		private System.IO.Stream GetBinaryDataStream(string base64String)
		{
			return new System.IO.MemoryStream(System.Convert.FromBase64String(base64String));
		}

		#endregion

		/// <summary>
		/// 计算最后一个单元格名称
		/// </summary>
		private void SetLaseCellNameAndTotalMergeCells()
		{
			lastCellName = "J";
			totalRows = 3 + projectClassifyIDs.Count + 1;
			totalMergeCells = 8 + projectClassifyIDs.Count + 4;
			foreach (int projectID in projectIDs)
			{
				int maxRows = 0;
				for (int i = 1; i < 4; i++)
				{
					var funds = dataContext.Funds.Where(f => f.ProjectID.Equals(projectID) && f.FundClassifyID == i && f.Date >= Convert.ToDateTime(_year.ToString() + "-01-01") && f.Date <= Convert.ToDateTime(_year.ToString() + "-12-31"));
					if (funds.Count() > maxRows)
					{
						maxRows = funds.Count();
					}
				}
				totalRows += maxRows;
				if (maxRows > 1)
				{
					//totalRows += 1; //有合计行
					totalMergeCells += 4;
				}
			}
			totalRows += 1; //总计行
			totalMergeCells += 4; //总计行合并格数
			lastCellName += totalRows.ToString();
		}

	}
}