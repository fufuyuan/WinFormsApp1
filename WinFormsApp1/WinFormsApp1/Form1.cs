namespace WinFormsApp1;

using SpreadsheetLight.Drawing;
using SpreadsheetLight;
using SpreadsheetLight.Charts;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Drawing;


public partial class Form1 : Form
{
    //string inputPath = "C:\\p2_java\\test.xlsx";
    // string outputPath = "C:\\p2_java\\output.xlsx";//保存路径必须是反斜杠
    string inputPath = "C:\\Users\\fu_fu\\Desktop\\github\\WinFormsApp1\\test.xlsx";
    string outputPath = "C:\\Users\\fu_fu\\Desktop\\github\\WinFormsApp1\\output.xlsx";//

    string outputPath1 = "C:\\Users\\fu_fu\\Desktop\\github\\WinFormsApp1\\output1.xlsx";//
    string outputPath2 = "C:\\Users\\fu_fu\\Desktop\\github\\WinFormsApp1\\output2.xlsx";//
    string outputPath3 = "C:\\Users\\fu_fu\\Desktop\\github\\WinFormsApp1\\output3.xlsx";//
    string outputPath4 = "C:\\Users\\fu_fu\\Desktop\\github\\WinFormsApp1\\output4.xlsx";//
    string outputPath5 = "C:\\Users\\fu_fu\\Desktop\\github\\WinFormsApp1\\output5.xlsx";//
    string outputPath6 = "C:\\Users\\fu_fu\\Desktop\\github\\WinFormsApp1\\output6.xlsx";//
    string outputPath7 = "C:\\Users\\fu_fu\\Desktop\\github\\WinFormsApp1\\output7.xlsx";//

    public Form1()
    {
        InitializeComponent();
    }

    private void button1_Click(object sender, EventArgs e)
    {

        FileStream fs = new FileStream(inputPath, FileMode.Open);

        MemoryStream msFirstPass = new MemoryStream();
        SLDocument slFirstPass = new SLDocument(fs, "Sheet1");

        SLChart chart = slFirstPass.CreateChart("C31", "W32");
        chart.SetChartType(SLPieChartType.Pie);
        chart.SetChartPosition(1, 1, 20, 14);//chart的初始行，初始列，结束行，结束列

        SLGroupDataLabelOptions gdloptions = chart.CreateGroupDataLabelOptions();

        gdloptions.SetLabelPosition(DocumentFormat.OpenXml.Drawing.Charts.DataLabelPositionValues.OutsideEnd);//标签设置显示在外侧
        gdloptions.ShowCategoryName = true;// 显示类别名称
        gdloptions.Separator = "\n"; // 设置分隔符为" - "

        gdloptions.ShowValue = true;// 显示数值


        gdloptions.SourceLinked = false;//必须设置这个下面的格式才能生效
        gdloptions.FormatCode = "0.0\"%\"";

        gdloptions.ShowLeaderLines = false; // 是否显示引线，属性设置顺序需要靠后才能生效


        chart.SetGroupDataLabelOptions(1, gdloptions);

        chart.HideChartLegend();//不显示[凡例]

        slFirstPass.InsertChart(chart);
        slFirstPass.SaveAs(outputPath1);

        fs.Close();
    }

    private void button2_Click(object sender, EventArgs e)
    {
        FileStream fs = new FileStream(inputPath, FileMode.Open);

        MemoryStream msFirstPass = new MemoryStream();
        SLDocument sl = new SLDocument(fs, "Sheet1");

        //默认的网格不显示
        SLPageSettings ps = new SLPageSettings();
        ps.ShowGridLines = false;
        sl.SetPageSettings(ps);

        sl.CopyCell("A31", "W42", "A12");

        SLStyle style = sl.CreateStyle();
        style.FormatCode = "0.0";
        sl.SetCellStyle("D13", "W13", style);

        //没用到图表的，给个背景色
        style.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.FromArgb(1, 150, 248), System.Drawing.Color.White);//虽然不懂为啥单独设第二个不行，反正能用就好
        sl.SetCellStyle("D14", "W20", style);

        sl.SetCellValue("A21", "※n=30未満は参考値のため灰色。");
        SLStyle style2 = sl.CreateStyle();
        style.SetFont("Meiryo UI", 9);
        sl.SetCellStyle("A1", style2);

        sl.SetRowHeight(14, 120);
        //插入图片
        //SLPicture pic = new SLPicture("C:\\p2_java\\insert.png");
        //pic.SetPosition(7, 0);
        //sl.InsertPicture(pic);

        SLCreateChartOptions creoptions = new SLCreateChartOptions();
        creoptions.IsStylish = false;//底层代码的锅，为啥要写成这个为true的话SLBarChartOptions就不生效

        SLChart chart = sl.CreateChart("C31", "W32", creoptions);
        chart.Border.SetNoLine();//图表背景透明（为了更贴近下面表单）

        SLFont ft;
        SLRstType rst = sl.CreateRstType();

        ft = sl.CreateFont();
        //ft.SetFontThemeColor(SLThemeColorIndexValues.Accent1Color);
        rst.AppendText("80 80 80 80 80 80 80 80 80 80 80 80 80 80 80 80 80 80 80 80", ft);

        chart.Title.SetTitle(rst);
        chart.ShowChartTitle(false);
        chart.Fill.SetNoFill();

        SLBarChartOptions bcoptions = new SLBarChartOptions();
        bcoptions.GapWidth = 50;//间距越低，柱子越宽

        chart.SetChartType(SLColumnChartType.ClusteredColumn, bcoptions);
        chart.SetChartPosition(1, 2.3, 11.7, 23.2);//chart的初始行，初始列，结束行，结束列（根据个数不同，要和列一一对应可能要一套计算方式）

        chart.HidePrimaryTextAxis(); //隐藏横坐标
        SLGroupDataLabelOptions gdloptions = chart.CreateGroupDataLabelOptions();

        gdloptions.SetLabelPosition(DocumentFormat.OpenXml.Drawing.Charts.DataLabelPositionValues.OutsideEnd);//标签设置显示在外侧
        gdloptions.ShowValue = true;// 显示数值

        gdloptions.SourceLinked = false;
        gdloptions.FormatCode = "0.0";

        chart.SetGroupDataLabelOptions(1, gdloptions);

        chart.HideChartLegend();//不显示[凡例]

        SLValueAxis valueAxis = chart.PrimaryValueAxis;

        valueAxis.ShowMajorGridlines = false;//主要刻度的网格线不显示
        valueAxis.MajorTickMark = TickMarkValues.Inside; // 设置主要刻度标记的位置为内部

        //自定义格式
        valueAxis.SourceLinked = false;
        valueAxis.FormatCode = "0\"%\"";

        // 设置值轴的最小值和最大值
        valueAxis.Minimum = 0;
        valueAxis.Maximum = 30;

        valueAxis.MajorUnit = 10;//step

        sl.InsertChart(chart);
        sl.SaveAs(outputPath2);

        fs.Close();
    }

    private void button3_Click(object sender, EventArgs e)
    {
        FileStream fs = new FileStream(inputPath, FileMode.Open);

        SLDocument sl = new SLDocument(fs, "Sheet1");

        //默认的网格不显示
        SLPageSettings ps = new SLPageSettings();
        ps.ShowGridLines = false;
        sl.SetPageSettings(ps);

        SLCreateChartOptions creoptions = new SLCreateChartOptions();
        creoptions.IsStylish = false;//底层代码的锅，为啥要写成这个为true的话Overlap就不生效

        SLChart chart = sl.CreateChart("D32", "W39", creoptions);

        SLBarChartOptions bcoptions = new SLBarChartOptions();
        bcoptions.Overlap = 0;
        //bcoptions.GapWidth = 0;

        chart.SetChartType(SLColumnChartType.ClusteredColumn, bcoptions);

        chart.SetChartPosition(1, 1, 12, 23);//chart的初始行，初始列，结束行，结束列

        chart.HidePrimaryTextAxis(); //隐藏横坐标
        //chart.HidePrimaryValueAxis(); //隐藏纵坐标

        chart.HideChartLegend();//不显示[凡例]

        SLValueAxis valueAxis = chart.PrimaryValueAxis;

        //自定义格式
        valueAxis.SourceLinked = false;
        valueAxis.FormatCode = "0\"%\"";

        // 设置值轴的最小值和最大值
        valueAxis.Minimum = 0;
        valueAxis.Maximum = 100;

        valueAxis.MajorUnit = 10;//step


        //valueAxis.SetDisplayUnits(DocumentFormat.OpenXml.Drawing.Charts.BuiltInUnitValues.Hundreds, true);

        sl.InsertChart(chart);


        sl.SaveAs(outputPath3);

        fs.Close();
    }

    private void button4_Click(object sender, EventArgs e)
    {
        FileStream fs = new FileStream(inputPath, FileMode.Open);


        // 创建一个包含20个指定颜色的数组
            System.Drawing.Color[] colors = new System.Drawing.Color[]
            {
                System.Drawing.Color.FromArgb(0, 46, 86),    // #002e56 - 深蓝色
                System.Drawing.Color.FromArgb(32, 86, 126),  // #20567e - 中蓝色
                System.Drawing.Color.FromArgb(64, 126, 165), // #407ea5 - 浅蓝色
                System.Drawing.Color.FromArgb(95, 166, 205), // #5fa6cd - 深天蓝色
                System.Drawing.Color.FromArgb(128, 207, 244),// #80cff4 - 浅天蓝色

                System.Drawing.Color.FromArgb(181, 227, 249),// #b5e3f9 - 淡蓝色
                System.Drawing.Color.FromArgb(231, 233, 235),// #e7e9eb - 浅灰色
                System.Drawing.Color.FromArgb(202, 206, 210),// #caced2 - 淡紫色
                System.Drawing.Color.FromArgb(160, 167, 173),// #a0a7ad - 浅灰色
                System.Drawing.Color.FromArgb(107, 118, 128),// #6b7680 - 中灰色

                System.Drawing.Color.FromArgb(71, 83, 93),   // #47535d - 深灰色
                System.Drawing.Color.FromArgb(45, 53, 60),   // #2d353c - 深灰色
                System.Drawing.Color.FromArgb(205, 175, 146),// #cdaf92 - 淡棕色
                System.Drawing.Color.FromArgb(207, 181, 137),// #cfb589 - 淡棕色
                System.Drawing.Color.FromArgb(248, 248, 248),// #f8f8f8 - 白色

                System.Drawing.Color.FromArgb(234, 234, 234),// #eaeaea - 浅灰色
                System.Drawing.Color.FromArgb(221, 221, 221),// #dddddd - 浅灰色
                System.Drawing.Color.FromArgb(178, 178, 178),// #b2b2b2 - 中灰色
                System.Drawing.Color.FromArgb(100, 100, 100),// #646464 - 中灰色
                System.Drawing.Color.FromArgb(70, 70, 70)    // #464646 - 深灰色
            };

        MemoryStream msFirstPass = new MemoryStream();
        SLDocument sl = new SLDocument(fs, "Sheet1");

        SLCreateChartOptions creoptions = new SLCreateChartOptions();
        creoptions.RowsAsDataSeries = false;//指定主轴的数据源是第一列而不是第一行

        SLChart chart = sl.CreateChart("C31", "W39", creoptions);

        SLBarChartOptions bcoptions = new SLBarChartOptions();
        bcoptions.GapWidth = 15;
        bcoptions.Overlap = 100;

        chart.SetChartType(SLBarChartType.StackedBarMax, bcoptions);

        SLDataSeriesOptions[] seriesOptions = new SLDataSeriesOptions[20];
        // 使用循环为数组的每个元素分配颜色
        for (int i = 0; i < colors.Length; i++)
        {
            seriesOptions[i] = new SLDataSeriesOptions(); // 创建新的 SLDataSeriesOptions 实例
            seriesOptions[i].Fill.SetSolidFill(colors[i], 0.5M);
            chart.SetDataSeriesOptions(i + 1, seriesOptions[i]);
        }

        chart.SetChartPosition(2.5, 2.90, 12.5, 23.30);//chart的初始行，初始列，结束行，结束列

        //横竖根据图表指代各有不同
        chart.HidePrimaryTextAxis(); //隐藏横坐标
        chart.HidePrimaryValueAxis(); //隐藏纵坐标

        chart.PrimaryTextAxis.InReverseOrder = true;//主轴数据逆向排序

        SLGroupDataLabelOptions gdloptions = chart.CreateGroupDataLabelOptions();

        //gdloptions.SetLabelPosition(DocumentFormat.OpenXml.Drawing.Charts.DataLabelPositionValues.OutsideEnd);//标签设置显示在外侧
        gdloptions.ShowValue = true;// 显示数值

        gdloptions.SourceLinked = false;
        gdloptions.FormatCode = "0.0";

        chart.SetGroupDataLabelOptions(gdloptions);

        chart.HideChartLegend();//不显示[凡例]


        //valueAxis.SetDisplayUnits(DocumentFormat.OpenXml.Drawing.Charts.BuiltInUnitValues.Hundreds, true);

        sl.InsertChart(chart);


        sl.SaveAs(outputPath);

        fs.Close();




    }

    private void button5_Click(object sender, EventArgs e)
    {
        FileStream fs = new FileStream(inputPath, FileMode.Open);

        SLDocument sl = new SLDocument(fs, "Sheet1");

        //默认的网格不显示
        SLPageSettings ps = new SLPageSettings();
        ps.ShowGridLines = false;
        sl.SetPageSettings(ps);


        sl.CopyCell("A31", "W42", "A12");

        SLStyle style = sl.CreateStyle();
        style.FormatCode = "0.0";
        sl.SetCellStyle("D13", "W13", style);

        //没用到图表的，给个背景色
        style.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.FromArgb(121, 150, 248), System.Drawing.Color.White);//虽然不懂为啥单独设第二个不行，反正能用就好
        sl.SetCellStyle("D14", "W20", style);

        sl.SetCellValue("A21", "※n=30未満は参考値のため灰色。");
        SLStyle style2 = sl.CreateStyle();
        style.SetFont("Meiryo UI", 9);
        sl.SetCellStyle("A1", style2);

        sl.SetRowHeight(14, 120);




        SLChart chart = sl.CreateChart("C31", "W32");

        SLLineChartOptions sLLineChartOptions = new SLLineChartOptions();


        chart.SetChartType(SLBarChartType.ClusteredBar);
        chart.SetChartPosition(1, 1, 21, 14);//chart的初始行，初始列，结束行，结束列

        SLValueAxis valueAxis = chart.PrimaryValueAxis;
        valueAxis.TickLabelPosition = DocumentFormat.OpenXml.Drawing.Charts.TickLabelPositionValues.High;//主轴显示在上方
        // 设置值轴的最小值和最大值
        valueAxis.Minimum = 0;
        valueAxis.Maximum = 100;
        valueAxis.MajorUnit = 20;//step

        SLGroupDataLabelOptions gdloptions = chart.CreateGroupDataLabelOptions();

        gdloptions.SetLabelPosition(DocumentFormat.OpenXml.Drawing.Charts.DataLabelPositionValues.OutsideEnd);//标签设置显示在外侧
        gdloptions.ShowValue = true;// 显示数值

        chart.SetGroupDataLabelOptions(1, gdloptions);

        chart.HideChartLegend();//不显示[凡例]

        sl.InsertChart(chart);
        sl.SaveAs(outputPath5);

        fs.Close();
    }

    private void button6_Click(object sender, EventArgs e)
    {
        FileStream fs = new FileStream(inputPath, FileMode.Open);

        MemoryStream msFirstPass = new MemoryStream();
        SLDocument sl = new SLDocument(fs, "Sheet1");

        SLChart chart = sl.CreateChart("C31", "W39");

        chart.SetChartType(SLLineChartType.LineWithMarkers);

        chart.SetChartPosition(2.5, 2.90, 22.5, 23.30);//chart的初始行，初始列，结束行，结束列

        chart.HidePrimaryTextAxis(); //隐藏横坐标

        chart.HideChartLegend();//不显示[凡例]


        //valueAxis.SetDisplayUnits(DocumentFormat.OpenXml.Drawing.Charts.BuiltInUnitValues.Hundreds, true);

        sl.InsertChart(chart);


        sl.SaveAs(outputPath6);

        fs.Close();
    }

    private void button7_Click(object sender, EventArgs e)
    {
        FileStream fs = new FileStream(inputPath, FileMode.Open);

        MemoryStream msFirstPass = new MemoryStream();
        SLDocument sl = new SLDocument(fs, "Sheet1");

        SLChart chart = sl.CreateChart("C31", "W39");

        SLBarChartOptions bcoptions = new SLBarChartOptions();
        //bcoptions.GapWidth = 0;
        //bcoptions.Overlap = 1;

        chart.SetChartType(SLBarChartType.ClusteredBar, bcoptions);
        chart.SetChartPosition(1, 1, 28, 15);//chart的初始行，初始列，结束行，结束列


        //SLDataSeriesOptions tt = new SLDataSeriesOptions();
        //chart.SetDataSeriesOptions(1, tt);

        chart.HidePrimaryValueAxis(); //隐藏横坐标


        chart.HideChartLegend();//不显示[凡例]

        SLGroupDataLabelOptions gdloptions = chart.CreateGroupDataLabelOptions();

        gdloptions.SetLabelPosition(DocumentFormat.OpenXml.Drawing.Charts.DataLabelPositionValues.OutsideEnd);//标签设置显示在外侧
        gdloptions.ShowValue = true;// 显示数值
        chart.SetGroupDataLabelOptions(gdloptions);

        //valueAxis.SetDisplayUnits(DocumentFormat.OpenXml.Drawing.Charts.BuiltInUnitValues.Hundreds, true);

        sl.InsertChart(chart);


        sl.SaveAs(outputPath7);

        fs.Close();
    }
}