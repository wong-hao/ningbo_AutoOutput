using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Data;
using Steema.TeeChart.Styles;

///Create By liping 20250724
namespace SMGI.Plugin.ThematicChart.TeeChart.PieChart
{
    /// <summary>
    /// 折线图参数
    /// </summary>
    public class LineChartModel
    {
        /// <summary>
        /// 标题
        /// </summary>
        public ChartTile Title;
        /// <summary>
        /// 宽度,厘米为单位
        /// </summary>
        public double Width;
        /// <summary>
        /// 高度,厘米为单位
        /// </summary>
        public double Height;
        /// <summary>
        /// 圆饼半径（厘米问你单位），不参与计算，只做结果展示
        /// </summary>
        public double Radius;
        /// <summary>
        /// 自定义半径（厘米为单位），用于用户输入半径，为0时忽略，将自动计算
        /// </summary>
        public double CustomRadius;

        /// <summary>
        /// 是否3D视图
        /// </summary>
        public bool View3D;
        /// <summary>
        /// 圆环视图
        /// </summary>
        public bool ViewRing;

        /// <summary>
        /// 数据表
        /// </summary>
        public DataTable DataTable;
        /// <summary>
        /// 饼图标注
        /// </summary>
        public LineLabel LineLabel;
        /// <summary>
        /// 扇区颜色
        /// </summary>
        public List<PieSector> Sectors;
        /// <summary>
        /// 扇区范围（只读）
        /// </summary>
        public Rectangle SeriesRect;

        /// <summary>
        /// 圆饼厚度
        /// </summary>
        public int PieDepth;

        /// <summary>
        /// 是否自定义扇区颜色
        /// </summary>
        public bool CustomSectorColor;

        /// <summary>
        /// 图例
        /// </summary>
        public ChartLegend Legend;

        public List<string> AreaTypes;
        public int Index;
        public List<Color> SeriesColorList { get; set; }
        public List<int> LineWidthList { get; set; }

        public bool Smoothed;
        public bool Pointer;
        public List<int> PointerWidthList { get; set; }
        public PointerStyles PointerStyles;
        public BottomTitle BottomTitle;

        // 正常情况下的构造函数
        public LineChartModel(bool reload = false)
        {
            this.Title = new ChartTile();
            this.DataTable = null;
            this.LineLabel = new LineLabel();
            this.Legend = new ChartLegend();
            this.Sectors = new List<PieSector>();
            this.ViewRing = false;
            this.View3D = false;
            //this.Radius = 0;
            this.CustomRadius = 0;
            this.Width = 20;
            this.Height = 15;
            this.CustomSectorColor = false;

            if (!reload)
            {
                SeriesColorList = Enumerable.Repeat(Color.White, 10).ToList();
                LineWidthList = Enumerable.Repeat(1, 10).ToList();

                PointerWidthList = Enumerable.Repeat(1, 10).ToList();
            }

            this.BottomTitle = new BottomTitle();
        }

        // 用于反序列化的构造函数
        public LineChartModel()
        {
            this.Title = new ChartTile();
            this.DataTable = null;
            this.LineLabel = new LineLabel();
            this.Legend = new ChartLegend();
            this.Sectors = new List<PieSector>();
            this.ViewRing = false;
            this.View3D = false;
            //this.Radius = 0;
            this.CustomRadius = 0;
            this.Width = 20;
            this.Height = 15;
            this.CustomSectorColor = false;

            this.BottomTitle = new BottomTitle();

        }
    }
    /// <summary>
    /// 饼图扇形
    /// </summary>
    public class LineSector{
        /// <summary>
        /// 扇区索引
        /// </summary>
        public int Index;
        /// <summary>
        /// 扇区颜色
        /// </summary>
        public Color Color;
        /// <summary>
        /// 扇区名字
        /// </summary>
        public string Name;

        public LineSector() {
            this.Color = Color.Black;
            this.Index = 0;
            this.Name = "";
        }
    }


    /// <summary>
    /// 折线图标注
    /// </summary>
    public class LineLabel {
        /// <summary>
        /// 是否进行标注
        /// </summary>
        public bool Visible;
        /// <summary>
        /// 是否引线标注
        /// </summary>
        public bool LeadLabel;
        /// <summary>
        /// 引出线长度，厘米为单位
        /// </summary>
        public double LeadlineLength;
        /// <summary>
        /// 引线水平线长度，厘米为单位
        /// </summary>
        public double LeadlineHorizonLength;
        /// <summary>
        /// 标注字段
        /// </summary>
        public string LebelField;//
        /// <summary>
        /// 标注字体
        /// </summary>
        public SFont Font;
        /// <summary>
        /// 标注样式
        /// </summary>
        public LabelStyle LabelStyle;


        /// <summary>
        /// 标注项
        /// </summary>
        public List<LabelItem> LabelItems;

        public LineLabel() {
            this.Visible = true;
            this.LeadLabel = false;
            this.LeadlineLength = 0.5;
            this.LeadlineHorizonLength = 0.8;
            this.LebelField = "";
            this.Font = new SFont();
            this.Font.Size = 8;
            this.LabelStyle  = TeeChart.LabelStyle.数值;
            this.LabelItems = new List<LabelItem>();
        }
    }
}
