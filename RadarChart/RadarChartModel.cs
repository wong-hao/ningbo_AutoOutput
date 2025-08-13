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
    /// 雷达图参数
    /// </summary>
    public class RadarChartModel
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
        public RadarLabel RadarLabel;
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
        public List<int> SeriesTransparencyList { get; set; }
        public List<int> LineWidthList { get; set; }
        public List<Color> LineColorList { get; set; }

        // 正常情况下的构造函数
        public RadarChartModel(bool reload = false)
        {
            this.Title = new ChartTile();
            this.DataTable = null;
            this.RadarLabel = new RadarLabel();
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
                SeriesTransparencyList = Enumerable.Repeat(100, 10).ToList();
                LineWidthList = Enumerable.Repeat(1, 10).ToList();
                LineColorList = Enumerable.Repeat(Color.White, 10).ToList();
            }
        }

        // 用于反序列化的构造函数
        public RadarChartModel()
        {
            this.Title = new ChartTile();
            this.DataTable = null;
            this.RadarLabel = new RadarLabel();
            this.Legend = new ChartLegend();
            this.Sectors = new List<PieSector>();
            this.ViewRing = false;
            this.View3D = false;
            //this.Radius = 0;
            this.CustomRadius = 0;
            this.Width = 20;
            this.Height = 15;
            this.CustomSectorColor = false;
        }
    }
    /// <summary>
    /// 饼图扇形
    /// </summary>
    public class RadarSector{
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

        public RadarSector() {
            this.Color = Color.Black;
            this.Index = 0;
            this.Name = "";
        }
    }


    /// <summary>
    /// 雷达图标注
    /// </summary>
    public class RadarLabel {
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

        public RadarLabel() {
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
