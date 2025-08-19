using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Data;
using Steema.TeeChart.Styles;

///Create By liping 20250724
namespace SMGI.Plugin.ThematicChart.TeeChart.LineChart
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
        ///是否3D视图
        /// </summary>
        public bool View3D;
        /// <summary>
        ///是否堆叠显示
        /// </summary>
        public bool Stack;

        /// <summary>
        /// 数据表
        /// </summary>
        public DataTable DataTable;
        /// <summary>
        /// 条图系列标注
        /// </summary>
        public LineLabel LineLabel;
        /// <summary>
        /// 条形图系列
        /// </summary>
        public List<LineSeries> Series;
        /// <summary>
        /// X轴标注（横轴）
        /// </summary>
        public AxisLabel XAxisLabel;
        /// <summary>
        /// Y轴标注（竖轴）
        /// </summary>
        public AxisLabel YAxisLabel;

        /// <summary>
        /// 图形范围（只读）
        /// </summary>
        public Rectangle SeriesRect;

        /// <summary>
        /// 条宽，以厘米为单位
        /// </summary>
        public double LineWidth;


        /// <summary>
        /// 是否自定义扇区颜色
        /// </summary>
        public bool CustomSectorColor;

        /// <summary>
        /// 图例
        /// </summary>
        public ChartLegend Legend;

        public bool Smoothed;
        public bool Pointer;
        public int PointerWidth;
        public PointerStyles PointerStyles;

        public LineChartModel()
        {
            this.Title = new ChartTile();
            this.DataTable = null;
            this.Legend = new ChartLegend();
            this.CustomSectorColor = false;
            this.Series = new List<LineSeries>();
            this.LineLabel = new LineLabel();
            this.XAxisLabel = new AxisLabel();
            this.XAxisLabel.Title.Text = "X轴";
            this.YAxisLabel = new AxisLabel();
            this.YAxisLabel.Title.Text = "Y轴";
            this.Stack = false;
            this.View3D = false;
            this.Width = 20;
            this.Height = 15;
            this.LineWidth = 0.2;
        }
    }
    /// <summary>
    /// 折线图
    /// </summary>
    public class LineSeries{
        /// <summary>
        /// 条形索引
        /// </summary>
        public int Index;
        /// <summary>
        /// 条形颜色
        /// </summary>
        public Color Color;
        /// <summary>
        /// 条形系列名字
        /// </summary>
        public string Name;

        public LineSeries()
        {
            this.Color = Color.Empty;
            this.Index = 0;
            this.Name = "";
        }
    }

    /// <summary>
    /// 坐标轴标注（含标题标注）
    /// </summary>
    public class AxisLabel {
        /// <summary>
        /// 是否可见
        /// </summary>
        public bool Visible;
        /// <summary>
        /// 标注字体
        /// </summary>
        public SFont Font;
        /// <summary>
        /// 坐标轴标注
        /// </summary>
        public AxisTile Title;

        public AxisLabel() {
            this.Visible = true;
            this.Font = new SFont();
            this.Title = new AxisTile();
        }
    }

    /// <summary>
    /// 需要自行绘制
    /// </summary>
    public class AxisTile {
        /// <summary>
        /// 是否可见
        /// </summary>
        public bool Visible;
        /// <summary>
        /// 标注字体
        /// </summary>
        public SFont Font;
        /// <summary>
        /// 文本内容
        /// </summary>
        public string Text;
        /// <summary>
        /// 高度
        /// </summary>
        public int Height;
        /// <summary>
        /// 宽度
        /// </summary>
        public int Width;
        /// <summary>
        ///顶部
        /// </summary>
        public int Top;
        /// <summary>
        /// 左部
        /// </summary>
        public int Left;

        public AxisTile() {
            this.Visible = true;
            this.Font = new SFont();
            this.Text = "示例";
        }
    }




    /// <summary>
    /// 条形图标注
    /// </summary>
    public class LineLabel {
        /// <summary>
        /// 是否进行标注
        /// </summary>
        public bool Visible;
        /// <summary>
        /// 标注字体
        /// </summary>
        public SFont Font;
        /// <summary>
        /// 标注样式
        /// </summary>
        public LabelStyle LabelStyle;

        public LineLabel() {
            this.Visible = true;
            this.Font = new SFont();
            this.Font.Size = 8;
            this.LabelStyle  = TeeChart.LabelStyle.数值;
        }
    }
}
