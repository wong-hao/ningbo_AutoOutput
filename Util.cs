using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ESRI.ArcGIS.Controls;
using ESRI.ArcGIS.Carto;
using Newtonsoft.Json;
using SMGI.Common;

namespace SMGI.Plugin.ThematicChart.TeeChart
{
    /// <summary>
    /// 图表标题
    /// </summary>
    public class ChartTile
    {
        /// <summary>
        /// 字体
        /// </summary>
        public SFont Font;
        /// <summary>
        /// 文本内容
        /// </summary>
        public string Text;
        /// <summary>
        /// 是否可见
        /// </summary>
        public bool Visible;
        /// <summary>
        /// 顶部的距离
        /// </summary>
        public int Top;
        /// <summary>
        /// 左侧的距离，在容器尺寸中居中后计算
        /// </summary>
        public int Left;
        /// <summary>
        /// 宽度
        /// </summary>
        public int Width;
        /// <summary>
        /// 高度
        /// </summary>
        public int Height;


        public ChartTile()
        {
            this.Visible = true;
            this.Font = new SFont();
            this.Font.Size = 15;
            this.Font.Style = FontStyle.Bold;
            this.Text = "示例标题";
            this.Top = 15;

        }
    }

    /// <summary>
    /// X轴标题
    /// </summary>
    public class BottomTitle
    {
        /// <summary>
        /// 字体
        /// </summary>
        public SFont Font;
        /// <summary>
        /// 文本内容
        /// </summary>
        public string Text;
        /// <summary>
        /// 是否可见
        /// </summary>
        public bool Visible;
        /// <summary>
        /// 顶部的距离
        /// </summary>
        public int Top;
        /// <summary>
        /// 左侧的距离，在容器尺寸中居中后计算
        /// </summary>
        public int Left;
        /// <summary>
        /// 宽度
        /// </summary>
        public int Width;
        /// <summary>
        /// 高度
        /// </summary>
        public int Height;


        public BottomTitle()
        {
            this.Visible = true;
            this.Font = new SFont();
            this.Font.Size = 15;
            this.Font.Style = FontStyle.Bold;
            this.Text = "示例X轴标题";
            this.Top = 15;

        }
    }

    /// <summary>
    /// 标注项
    /// </summary>
    public class LabelItem
    {
        /// <summary>
        /// 起点
        /// </summary>
        public Point FromPoint;
        /// <summary>
        /// 中点
        /// </summary>
        public Point ToPoint;
        /// <summary>
        /// 终点
        /// </summary>
        public Point EndPoint;
        /// <summary>
        /// 文本范围
        /// </summary>
        public Rectangle TextRect;
        /// <summary>
        /// 标注文本
        /// </summary>
        public string Text;
    }


    /// <summary>
    /// 自定义图例
    /// </summary>
    public class ChartLegend
    {
        /// <summary>
        /// 左上|中上|右上|左下|左中|左下……类似于九宫格
        /// 对外，实际接口不支持，通过top和left进行计算摆放
        /// </summary>
        public LegendPostion Postion;
        /// <summary>
        /// 不对外使用
        /// </summary>
        public int Top;
        /// <summary>
        /// 不对外使用
        /// </summary>
        public int Left;
        /// <summary>
        /// 是否可见
        /// </summary>
        public bool Visible;
        /// <summary>
        /// 图例列数
        /// </summary>
        public int colNum;
        /// <summary>
        /// 符号文字的间隔
        /// </summary>
        public int SymbolTextGap;
        /// <summary>
        /// 根据内部布局自行计算的
        /// </summary>
        public int Width;
        /// <summary>
        /// 根据内部布局自行计算
        /// </summary>
        public int Height;
        /// <summary>
        /// 不同Item之间的水平间隔
        /// </summary>
        public int HorizonGap;
        /// <summary>
        /// 不同Item之间的高度间隔
        /// </summary>
        public int VerticalGap;

        /// <summary>
        /// 符号的宽度
        /// </summary>
        public int SymbolWidth;
        /// <summary>
        /// 符号的高度，根据文本字体的高度获取
        /// </summary>
        public int SymbolHeight;
        /// <summary>
        /// 图例的集合
        /// </summary>
        public List<LegendItem> LegendItems;
        /// <summary>
        /// 图例字体
        /// </summary>
        public SFont Font;

        public ChartLegend()
        {
            this.Postion = LegendPostion.顶部靠左;
            this.Visible = true;
            this.Font = new SFont();
            this.Font.Size = 8;
            this.LegendItems = new List<LegendItem>();
            this.colNum = 1;
            this.HorizonGap = 8;
            this.VerticalGap = 5;
            this.SymbolWidth = 30;
            this.SymbolTextGap = 8;
        }
    }

    /// <summary>
    /// 自定义图例对象
    /// </summary>
    public class LegendItem
    {
        /// <summary>
        /// 顶侧顶位，通过计算所得
        /// </summary>
        public int Top;
        /// <summary>
        /// 左侧定位，通过计算所得
        /// </summary>
        public int Left;
        /// <summary>
        /// 图例项符号的填充颜色
        /// </summary>
        public Color SymbolColor;
        /// <summary>
        /// 整个图例项的宽度，通过计算所得
        /// </summary>
        public int Width;

        /// <summary>
        /// 整个图例项的高度，通过计算所得
        /// </summary>
        public int Height;
        /// <summary>
        /// 符号的渲染范围
        /// </summary>
        public Rectangle SymbolRect;

        /// <summary>
        /// 图例文本
        /// </summary>
        public string Text;
    }


    /// <summary>
    /// 图例位置
    /// </summary>
    public enum LegendPostion
    {
        顶部靠左 = 0,
        顶部居中 = 1,
        顶部靠右 = 2,
        底部靠左 = 3,
        底部居中 = 4,
        底部靠右 = 5,
        左中 = 6,
        右中 = 7
    }
    /// <summary>
    /// (废弃)图例排列顺序
    /// </summary>
    public enum LegendSort
    {
        水平排列 = 0,
        树直排列 = 1
    }
    /// <summary>
    /// 标注样式，对标TeeChart的MarkStyles
    /// </summary>
    public enum LabelStyle
    {
        数值 = 0,
        数值百分比 = 1,
        标注数值百分比 = 3,
        标注数值 = 4,
    }

    /// <summary>
    /// 自定义字体结构
    /// </summary>
    public class SFont
    {
        private Font _font;
        private Color _color;
        public SFont(Font font = null)
        {
            if (font == null)
                _font = new System.Drawing.Font("宋体", 10);
            else
                _font = font;
            this._color = Color.Black;
        }

        // 代理 Font 的常用属性
        public string Name
        {
            get
            {
                return _font.Name;
            }
            set
            {

                _font = new Font(value, _font.Size, _font.Style);
            }
        }
        // 代理 Font 的常用属性
        public float Size
        {
            get
            {
                return _font.Size;
            }
            set
            {
                // 创建一个新的 Font 对象，保留 Name 和 Style
                _font = new Font(_font.Name, value, _font.Style);
            }
        }
        // 代理 Font 的常用属性
        public FontStyle Style
        {
            get
            {
                return _font.Style;
            }
            set
            {
                _font = new Font(_font.Name, _font.Size, value);
            }
        }

        public Color Color
        {
            get
            {
                return _color;
            }
            set
            {
                _color = value;
            }
        }

        public Font Font
        {
            get
            {
                return _font;
            }
        }

        public override string ToString()
        {
            string styleCn = "常规";
            if (Style == FontStyle.Bold)
            {
                styleCn = "加粗";
            }
            else if (Style == FontStyle.Italic)
            {
                styleCn = "斜体";
            }
            return string.Format("{0}, {1}, {2}", Name, Size, styleCn);
        }
    }

    /// <summary>
    /// 工具类
    /// </summary>
    public static class Util
    {
        /// <summary>
        /// 延时回调方法（回调函数不带返回值和参数）
        /// </summary>
        /// <param name="millisecondsDelay">毫秒延时数</param>
        /// <param name="callback">回调函数，无参无返回值</param>
        public static Task DelayAsync(int millisecondsDelay, Action callback)
        {
            var tcs = new TaskCompletionSource<bool>();
            var timer = new System.Threading.Timer(_ =>
            {
                tcs.SetResult(true);
            }, null, millisecondsDelay, Timeout.Infinite);

            // 任务完成后释放Timer
            return tcs.Task.ContinueWith(t =>
            {
                timer.Dispose();
                callback.Invoke(); //执行回调
                return t;
            }).Unwrap();
        }


        /// <summary>
        /// 像素转厘米
        /// </summary>     
        public static double PixelToCentimeter(double pixel)
        {
            double devicePixelRatio = GetSystemScale();
            return pixel* 25.4 / (96*devicePixelRatio) * 0.1;
        }

        /// <summary>
        /// 厘米转像素
        /// </summary>     
        public static double CentimeterToPixel(double centimeter)
        {
            double devicePixelRatio = GetSystemScale();
            return centimeter * 10 * 96 / 25.4 * devicePixelRatio;
        }


        [DllImport("gdi32.dll")]
        private static extern int GetDeviceCaps(IntPtr hdc, int nIndex);
        /// <summary>
        /// 获取系统放大比例
        /// </summary>
        /// <returns></returns>
        public static double GetSystemScale()
        {
            using (Graphics g = Graphics.FromHwnd(IntPtr.Zero))
            {
                IntPtr desktop = g.GetHdc();
                var physicalScreenHeight = GetDeviceCaps(desktop, (int)DeviceCap.DESKTOPVERTRES);
                var screenScalingFactor = (double)physicalScreenHeight / Screen.PrimaryScreen.Bounds.Height;
                return screenScalingFactor; // 返回缩放倍数（如 1.25 表示 125%）
            }
        }
    }

    public enum DeviceCap
    {
        VERTRES = 10,
        PHYSICALWIDTH = 110,
        SCALINGFACTORX = 114,
        DESKTOPVERTRES = 117
        // http://pinvoke.net/default.aspx/gdi32/GetDeviceCaps.html
    }
    /// <summary>
    /// PageLayout 右键菜单项（仅限统计图）
    /// </summary>
    public static class PageLayoutOnRightClickMenuForChart {

        /// <summary>
        /// 右键菜单
        /// </summary>
        private static ContextMenuStrip contextMenu;
        /// <summary>
        /// 是否初始化
        /// </summary>
        private static bool inited;

        private static GApplication app;

        /// <summary>
        /// 初始化
        /// </summary>
        /// <returns></returns>
        public static void Init(GApplication _app) {
            if (!inited) {
                app = _app;
                contextMenu = new ContextMenuStrip();
                ToolStripMenuItem menuItem1 = new ToolStripMenuItem("修改");
                ToolStripMenuItem menuItem2 = new ToolStripMenuItem("删除");
                contextMenu.Items.Add(menuItem1);
                contextMenu.Items.Add(menuItem2);
                menuItem1.Click += new EventHandler(menuItem_modifyClick);
                menuItem2.Click += new EventHandler(menuItem_deleteClick);
                app.PageLayoutControl.OnMouseUp -= new IPageLayoutControlEvents_Ax_OnMouseUpEventHandler(PageLayoutOnRightClick);
                app.PageLayoutControl.OnMouseUp += new IPageLayoutControlEvents_Ax_OnMouseUpEventHandler(PageLayoutOnRightClick);
                inited = true;
            }
        }


        /// <summary>
        /// 图表元素修改
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private static void menuItem_modifyClick(object sender,EventArgs e) {

            IElement element = contextMenu.Tag as IElement;
            string name = (element as IElementProperties).Name;
            if (name.Contains("Pie"))
            {
                var model = JsonConvert.DeserializeObject<SMGI.Plugin.ThematicChart.TeeChart.PieChart.PieChartModel>((element as IElementProperties).CustomProperty.ToString());
                SMGI.Plugin.ThematicChart.TeeChart.PieChart.PieParameterPanel pp = new SMGI.Plugin.ThematicChart.TeeChart.PieChart.PieParameterPanel(GApplication.Application, model);
                pp.Show();
            }
            else if (name.Contains("Radar"))
            {
                var model = JsonConvert.DeserializeObject<SMGI.Plugin.ThematicChart.TeeChart.PieChart.RadarChartModel>((element as IElementProperties).CustomProperty.ToString());
                SMGI.Plugin.ThematicChart.TeeChart.PieChart.RadarParameterPanel pp = new SMGI.Plugin.ThematicChart.TeeChart.PieChart.RadarParameterPanel(GApplication.Application, model);
                pp.Show();
            }
            //else if()
        }
        /// <summary>
        /// 图表元素删除
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private static void menuItem_deleteClick(object sender, EventArgs e)
        {
            IElement element = contextMenu.Tag as IElement;
            IGraphicsContainer graphicsContainer = (IGraphicsContainer)app.PageLayoutControl.ActiveView;
            graphicsContainer.DeleteElement(element);
            app.PageLayoutControl.ActiveView.Refresh();
        }


        /// <summary>
        /// PageLayout 右键事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private static void PageLayoutOnRightClick(object sender, IPageLayoutControlEvents_OnMouseUpEvent e)
        {
            AxPageLayoutControl pageLayoutControl = sender as AxPageLayoutControl;
            if (pageLayoutControl.CurrentTool is ESRI.ArcGIS.Controls.ControlsSelectTool)
            {
                //右键点击，弹出菜单
                if (e.button == 2)
                {
                    IGraphicsContainerSelect graphicsContainerSelect = (IGraphicsContainerSelect)pageLayoutControl.PageLayout;
                    IEnumElement selectedElements = graphicsContainerSelect.SelectedElements;
                    selectedElements.Reset();
                    IElement element = selectedElements.Next();//默认取第一个
                    string name = (element as IElementProperties).Name;
                    if (element is PictureElement && name.Contains("Chart"))
                    {
                        System.Drawing.Point screenPoint = new System.Drawing.Point(e.x, e.y);
                        contextMenu.Show(pageLayoutControl, screenPoint);
                        contextMenu.Tag = element;
                    }
                }
            }
        }
    }
}
