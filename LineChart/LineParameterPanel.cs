using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SMGI.Common;
using SMGI.Plugin.ThematicChart.TeeChart.Widget;
using Steema.TeeChart;
using Steema.TeeChart.Styles;
using Steema.TeeChart.Drawing;
using System.IO;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using Newtonsoft.Json;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Geometry;
using Svg;
using Svg.Transforms;
using Line = Steema.TeeChart.Styles.Line;

namespace SMGI.Plugin.ThematicChart.TeeChart.LineChart
{
    public partial class LineParameterPanel : Form
    {
        private GApplication m_Application;
        /// <summary>
        /// 折线图设置参数模型
        /// </summary>
        private LineChartModel m_LineChartModel;

        /// <summary>
        /// 是否显示
        /// </summary>
        public bool boolShow;
        /// <summary>
        /// 图表窗体
        /// </summary>
        private ChartForm m_ChartForm;
        /// <summary>
        /// 图表对象
        /// </summary>
        private TChart m_Chart=null;

        public LineParameterPanel(GApplication app, LineChartModel pcm = null)
        {
            InitializeComponent();
            m_Application = app;
            if (pcm == null)
            {
                m_LineChartModel = new LineChartModel();
            }
            else {
                m_LineChartModel = pcm;
            }
            m_ChartForm = new ChartForm();
            m_ChartForm.ShowInTaskbar = false;
            m_ChartForm.StartPosition = FormStartPosition.CenterScreen;
            var hwnd = app.MainForm.Handle;
            var interop = new System.Windows.Forms.NativeWindow();
            interop.AssignHandle(hwnd);
            m_ChartForm.Show(interop);
            this.initUIParam();
            this.Disposed += new EventHandler(ParameterPanel_Disposed);
        }

        /// <summary>
        /// 初始化UI参数
        /// </summary>
        private void initUIParam(){
           //通用设置
           this.num_chartWidth.Value=(decimal)this.m_LineChartModel.Width;
           this.num_chartHeight.Value=(decimal)this.m_LineChartModel.Height;
           this.check_3d.Checked = this.m_LineChartModel.View3D;
           //this.check_stack.Checked = this.m_LineChartModel.Stack;
           //标题设置
           this.check_title.Checked = this.m_LineChartModel.Title.Visible;
           this.check_title_CheckedChanged(this.check_title, EventArgs.Empty);
           this.txt_title.Text = this.m_LineChartModel.Title.Text;
           this.label_titleStyle.Text = this.m_LineChartModel.Title.Font.ToString();
           this.label_titleColor.BackColor = this.m_LineChartModel.Title.Font.Color;
           //条形图标注设置
           this.check_label.Checked = this.m_LineChartModel.LineLabel.Visible;
           this.check_label_CheckedChanged(this.check_label, EventArgs.Empty);
           this.label_labelStyle.Text = this.m_LineChartModel.LineLabel.Font.ToString();
           this.label_labelStyle.ForeColor = this.m_LineChartModel.LineLabel.Font.Color;
           this.label_labelStyle.Text = this.m_LineChartModel.LineLabel.Font.ToString();
           this.label_labelColor.BackColor = this.m_LineChartModel.LineLabel.Font.Color;
           //条形图设置
           if (this.m_LineChartModel.DataTable != null)
           {
               DataTable dt = this.m_LineChartModel.DataTable;
               for (int i = 1; i < dt.Columns.Count; i++)
               {
                   this.checkedComboBox_dataItems.Properties.Items.Add(dt.Columns[i].ColumnName);
               }
           }
           if (this.m_LineChartModel.Series.Count > 0)
           {
               string nameString = "";
               this.m_LineChartModel.Series.ForEach(s =>
               {
                   nameString += s.Name + ",";
                   this.combo_sector.Properties.Items.Add(s.Name);
               });
               nameString= nameString.Substring(0,nameString.Length-1);
               this.checkedComboBox_dataItems.EditValue = nameString;
           }
           this.checkedComboBox_dataItems.EditValueChanged += new System.EventHandler(this.checkedComboBox_dataItems_EditValueChanged);

           this.num_barWidth.Value = (decimal)this.m_LineChartModel.LineWidth;
           //X轴横轴标注设置
           this.check_X_axisLabel.Checked = this.m_LineChartModel.XAxisLabel.Visible;
           this.check_X_axisLabel_CheckedChanged(this.check_X_axisLabel, EventArgs.Empty);
           this.label_X_axisLabelColor.BackColor = this.m_LineChartModel.XAxisLabel.Font.Color;
           this.label_X_axisLabelStyle.Text = this.m_LineChartModel.XAxisLabel.Font.ToString();
           //
           this.check_X_axisTitle.Checked = this.m_LineChartModel.XAxisLabel.Title.Visible;
           this.check_X_axisTitle_CheckedChanged(this.check_X_axisTitle, EventArgs.Empty);
           this.txt_X_title.Text = this.m_LineChartModel.XAxisLabel.Title.Text;
           this.label_X_titleColor.BackColor = this.m_LineChartModel.XAxisLabel.Title.Font.Color;
           this.label_X_titleStyle.Text = this.m_LineChartModel.XAxisLabel.Title.Font.ToString();
           //Y轴纵轴标注设置
           this.check_Y_axisLabel.Checked = this.m_LineChartModel.YAxisLabel.Visible;
           this.check_Y_axisLabel_CheckedChanged(this.check_Y_axisLabel,EventArgs.Empty);
           this.label_Y_axisLabelColor.BackColor = this.m_LineChartModel.YAxisLabel.Font.Color;
           this.label_Y_axisLabelStyle.Text = this.m_LineChartModel.YAxisLabel.Font.ToString();
           //
           this.check_Y_axisTitle.Checked = this.m_LineChartModel.YAxisLabel.Title.Visible;
           this.check_Y_axisTitle_CheckedChanged(this.check_Y_axisTitle, EventArgs.Empty);
           this.txt_Y_title.Text = this.m_LineChartModel.YAxisLabel.Title.Text;
           this.label_Y_titleColor.BackColor = this.m_LineChartModel.YAxisLabel.Title.Font.Color;
           this.label_Y_titleStyle.Text = this.m_LineChartModel.YAxisLabel.Title.Font.ToString();
           ////图例设置
           this.check_legend.Checked = this.m_LineChartModel.Legend.Visible;
           this.check_legend_CheckedChanged(this.check_legend, EventArgs.Empty);
           this.combo_legendPositon.SelectedItem = Enum.GetName(typeof(LegendPostion), this.m_LineChartModel.Legend.Postion);
           this.num_legendCol.Value = this.m_LineChartModel.Legend.colNum;
           this.label_legendStyle.Text = this.m_LineChartModel.Legend.Font.ToString();
           this.label_legendColor.BackColor = this.m_LineChartModel.Legend.Font.Color;
           // //其他设置
           if (this.m_LineChartModel.DataTable != null)
           {
               this.btn_apply.Enabled = true;
               this.btn_chartin.Enabled = true;
               this.btn_export.Enabled = true;
               this.btn_apply_Click(this.btn_apply, EventArgs.Empty);
           }
           else
           {
               this.btn_apply.Enabled = false;
               this.btn_chartin.Enabled = false;
               this.btn_export.Enabled = false;
           }
        }

        /// <summary>
        /// 重写方法
        /// </summary>
        public new void Show() {
            this.m_Application.MainForm.ShowChild2(this.Handle, FormLocation.Right);
            this.boolShow = true;
        }
        /// <summary>
        /// 重写方法
        /// </summary>
        public new void Close() {
            this.m_Application.MainForm.CloseChild(this.Handle);
            this.boolShow = false;
        }

        #region 通用设置
        private void num_chartWidth_EditValueChanged(object sender, EventArgs e)
        {
            this.m_LineChartModel.Width = (double)(sender as DevExpress.XtraEditors.SpinEdit).Value;
        }

        private void num_chartHeight_EditValueChanged(object sender, EventArgs e)
        {
            this.m_LineChartModel.Height = (double)(sender as DevExpress.XtraEditors.SpinEdit).Value;
        }

        private void check_3d_CheckedChanged(object sender, EventArgs e)
        {
            this.m_LineChartModel.View3D = this.check_3d.Checked;
        }

        #endregion

        #region 标题设置
        private void check_title_CheckedChanged(object sender, EventArgs e)
        {
            this.m_LineChartModel.Title.Visible = this.check_title.Checked;
            if (this.check_title.Checked)
            {
                this.txt_title.Enabled = true;
                this.btn_titleStyle.Enabled = true;
            }
            else
            {
                this.txt_title.Enabled = false;
                this.btn_titleStyle.Enabled = false;
            }
        }

        private void txt_title_EditValueChanged(object sender, EventArgs e)
        {
            this.m_LineChartModel.Title.Text = this.txt_title.Text;
        }

        private void btn_titleStyle_Click(object sender, EventArgs e)
        {
            FontColorForm fcf = new FontColorForm(this.m_LineChartModel.Title.Font);
            fcf.ShowInTaskbar = false;
            fcf.ShowIcon = false;
            fcf.StartPosition = FormStartPosition.CenterScreen;
            if (fcf.ShowDialog() == DialogResult.OK)
            {
                this.m_LineChartModel.Title.Font = fcf.sFont;
                this.label_titleStyle.Text = fcf.sFont.ToString();
                this.label_titleColor.BackColor = fcf.sFont.Color;
            }
        }
        #endregion

        #region 条形标注设置
        private void check_label_CheckedChanged(object sender, EventArgs e)
        {
            this.m_LineChartModel.LineLabel.Visible = this.check_label.Checked;
            if (this.check_label.Checked)
            {
                this.btn_labelStyle.Enabled = true;
            }
            else
            {
                this.btn_labelStyle.Enabled = false;
            }
        }

        private void btn_labelStyle_Click(object sender, EventArgs e)
        {
            FontColorForm fcf = new FontColorForm(this.m_LineChartModel.LineLabel.Font);
            fcf.ShowInTaskbar = false;
            fcf.ShowIcon = false;
            fcf.StartPosition = FormStartPosition.CenterScreen;
            if (fcf.ShowDialog() == DialogResult.OK)
            {
                this.m_LineChartModel.LineLabel.Font = fcf.sFont;
                this.label_labelStyle.Text = fcf.sFont.ToString();
                this.label_labelColor.BackColor = fcf.sFont.Color;
            }
        }

        private void combo_sector_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.combo_sector.SelectedItem == null)
            {
                this.btn_sectorColor.Enabled = false;
            }
            else {
                this.btn_sectorColor.Enabled = true;
                this.label_sectorColor.BackColor = this.m_LineChartModel.Series[this.combo_sector.SelectedIndex].Color;
            }
        }
        private void btn_sectorColor_Click(object sender, EventArgs e)
        {
            ColorDialog colorDialog = new ColorDialog();
            // 设置默认颜色（可选）
            colorDialog.Color = this.label_sectorColor.BackColor;
            // 允许自定义颜色（可选）
            colorDialog.AllowFullOpen = true;
            colorDialog.FullOpen = true;
            // 显示对话框并检查用户是否点击“确定”
            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                Color selectedColor = colorDialog.Color;
                this.label_sectorColor.BackColor = selectedColor;
                this.m_LineChartModel.Series[this.combo_sector.SelectedIndex].Color = selectedColor;
                //this.m_BarChartModel.CustomSectorColor = true;
            }
        }
        #endregion

        #region 条形图设置
        private void checkedComboBox_dataItems_EditValueChanged(object sender, EventArgs e)
        {
           //根据选择的列，调整 Series的个数
           var selectValues= this.checkedComboBox_dataItems.Properties.GetItems().GetCheckedValues();
           this.combo_sector.Properties.Items.Clear();
           List<LineSeries> seriesList = new List<LineSeries>();
           for (int i = 0; i < selectValues.Count; i++)
           {
               string val = selectValues[i] as string;
               LineSeries bs = new LineSeries
               {
                   Index=i,
                   Name=val
               };
               var ser = this.m_LineChartModel.Series.Find(s => s.Name == val);
               if (ser != null)
                   bs.Color = ser.Color;
               else
                   bs.Color = Color.Empty;
               seriesList.Add(bs);
               this.combo_sector.Properties.Items.Add(val);
           }
           this.m_LineChartModel.Series.Clear();
           this.m_LineChartModel.Series.AddRange(seriesList);
        }

        private void num_barWidth_EditValueChanged(object sender, EventArgs e)
        {
            this.m_LineChartModel.LineWidth =(double)this.num_barWidth.Value;
        }
        #endregion
         
        #region X轴横轴标注设置
        private void check_X_axisLabel_CheckedChanged(object sender, EventArgs e)
        {
            this.m_LineChartModel.XAxisLabel.Visible = this.check_X_axisLabel.Checked;
            if (this.check_X_axisLabel.Checked)
            {
                this.btn_X_axisLabelStyle.Enabled = true;
            }
            else
            {
                this.btn_X_axisLabelStyle.Enabled = false;
            }
        }
        private void btn_X_axisLabelStyle_Click(object sender, EventArgs e)
        {
            FontColorForm fcf = new FontColorForm(this.m_LineChartModel.XAxisLabel.Font);
            fcf.ShowInTaskbar = false;
            fcf.ShowIcon = false;
            fcf.StartPosition = FormStartPosition.CenterScreen;
            if (fcf.ShowDialog() == DialogResult.OK)
            {
                this.m_LineChartModel.XAxisLabel.Font = fcf.sFont;
                this.label_X_axisLabelStyle.Text = fcf.sFont.ToString();
                this.label_X_axisLabelColor.BackColor = fcf.sFont.Color;
            }
        }
        #endregion

        #region X轴横轴标题设置
        private void check_X_axisTitle_CheckedChanged(object sender, EventArgs e)
        {
            this.m_LineChartModel.XAxisLabel.Title.Visible = this.check_X_axisTitle.Checked;
            if (this.check_X_axisTitle.Checked)
            {
                this.btn_X_titleStyle.Enabled = true;
                this.txt_X_title.Enabled = true;
            }
            else
            {
                this.btn_X_titleStyle.Enabled = false;
                this.txt_X_title.Enabled = false;
            }
        }
        private void txt_X_title_EditValueChanged(object sender, EventArgs e)
        {
            this.m_LineChartModel.XAxisLabel.Title.Text = this.txt_X_title.Text;
        }

        private void btn_X_titleStyle_Click(object sender, EventArgs e)
        {
            FontColorForm fcf = new FontColorForm(this.m_LineChartModel.XAxisLabel.Title.Font);
            fcf.ShowInTaskbar = false;
            fcf.ShowIcon = false;
            fcf.StartPosition = FormStartPosition.CenterScreen;
            if (fcf.ShowDialog() == DialogResult.OK)
            {
                this.m_LineChartModel.XAxisLabel.Title.Font = fcf.sFont;
                this.label_X_titleStyle.Text = fcf.sFont.ToString();
                this.label_X_titleColor.BackColor = fcf.sFont.Color;
            }
        }

        #endregion
        
        #region Y轴纵轴标注设置
        private void check_Y_axisLabel_CheckedChanged(object sender, EventArgs e)
        {
            this.m_LineChartModel.YAxisLabel.Visible = this.check_Y_axisLabel.Checked;
            if (this.check_Y_axisLabel.Checked)
            {
                this.btn_Y_axisLabelStyle.Enabled = true;
            }
            else
            {
                this.btn_Y_axisLabelStyle.Enabled = false;
            }
        }
        private void btn_Y_axisLabelStyle_Click(object sender, EventArgs e)
        {
            FontColorForm fcf = new FontColorForm(this.m_LineChartModel.YAxisLabel.Font);
            fcf.ShowInTaskbar = false;
            fcf.ShowIcon = false;
            fcf.StartPosition = FormStartPosition.CenterScreen;
            if (fcf.ShowDialog() == DialogResult.OK)
            {
                this.m_LineChartModel.YAxisLabel.Font = fcf.sFont;
                this.label_Y_axisLabelStyle.Text = fcf.sFont.ToString();
                this.label_Y_axisLabelColor.BackColor = fcf.sFont.Color;
            }
        }
        private void txt_Y_title_EditValueChanged(object sender, EventArgs e)
        {
            this.m_LineChartModel.YAxisLabel.Title.Text = this.txt_Y_title.Text;
        }
        #endregion

        #region Y轴纵轴标题设置
        private void check_Y_axisTitle_CheckedChanged(object sender, EventArgs e)
        {
            this.m_LineChartModel.YAxisLabel.Title.Visible = this.check_Y_axisTitle.Checked;
            if (this.check_Y_axisTitle.Checked)
            {
                this.btn_Y_titleStyle.Enabled = true;
                this.txt_Y_title.Enabled = true;
            }
            else
            {
                this.btn_Y_titleStyle.Enabled = false;
                this.txt_Y_title.Enabled = false;
            }
        }
        private void btn_Y_titleStyle_Click(object sender, EventArgs e)
        {
            FontColorForm fcf = new FontColorForm(this.m_LineChartModel.YAxisLabel.Title.Font);
            fcf.ShowInTaskbar = false;
            fcf.ShowIcon = false;
            fcf.StartPosition = FormStartPosition.CenterScreen;
            if (fcf.ShowDialog() == DialogResult.OK)
            {
                this.m_LineChartModel.YAxisLabel.Title.Font = fcf.sFont;
                this.label_Y_titleStyle.Text = fcf.sFont.ToString();
                this.label_Y_titleColor.BackColor = fcf.sFont.Color;
            }
        }

        #endregion

        #region 图例设置
        private void check_legend_CheckedChanged(object sender, EventArgs e)
        {
            this.m_LineChartModel.Legend.Visible = this.check_legend.Checked;
            if (this.check_legend.Checked){
                this.combo_legendPositon.Enabled = true;
                this.num_legendCol.Enabled = true;
            }
            else {
                this.combo_legendPositon.Enabled = false;
                this.num_legendCol.Enabled = false;
            }
        }

        private void combo_legendPositon_EditValueChanged(object sender, EventArgs e)
        {
            if (this.combo_legendPositon.SelectedItem==null)
                return;
            this.m_LineChartModel.Legend.Postion = (LegendPostion)Enum.Parse(typeof(LegendPostion), this.combo_legendPositon.SelectedItem.ToString());
        }
        private void num_legendCol_EditValueChanged(object sender, EventArgs e)
        {
            this.m_LineChartModel.Legend.colNum = (int)this.num_legendCol.Value;
        }

        private void btn_legendStyle_Click(object sender, EventArgs e)
        {
            FontColorForm fcf = new FontColorForm(this.m_LineChartModel.Legend.Font);
            fcf.ShowInTaskbar = false;
            fcf.ShowIcon = false;
            fcf.StartPosition = FormStartPosition.CenterScreen;
            if (fcf.ShowDialog() == DialogResult.OK)
            {
                this.m_LineChartModel.Legend.Font = fcf.sFont;
                this.label_legendStyle.Text = fcf.sFont.ToString();
                this.label_legendColor.BackColor = fcf.sFont.Color;
            }
        }

        #endregion

        private void ParameterPanel_Disposed(object sender, EventArgs e)
        {
            this.m_ChartForm.Close();
        }

        private void btn_apply_Click(object sender, EventArgs e)
        {
            //this.m_BarChartModel.Legend.Visible = false;
            //this.m_BarChartModel.XAxisLabel.Title.Visible = true;
            this.Draw(this.m_LineChartModel);
        }

        private void btn_loadData_Click(object sender, EventArgs e)
        {
            if (this.m_LineChartModel.DataTable != null) {
                DialogResult dialogResult = MessageBox.Show("统计图已有数据，是否要重新加载？", "提示", MessageBoxButtons.YesNo);
                if (dialogResult == System.Windows.Forms.DialogResult.No)
                    return;
            }
            OpenFileDialog of = new OpenFileDialog();
            of.Multiselect = false;
            of.Filter = "Excel文件|*.xlsx|Excel文件|*.xls";
            if (of.ShowDialog() != DialogResult.OK)
                return;
            string excelPath = of.FileName;
            DataTable dt = this.ExcelToDataTable(excelPath);
            this.checkedComboBox_dataItems.Properties.Items.Clear();
            for (int i = 1; i < dt.Columns.Count; i++)
            {
                this.checkedComboBox_dataItems.Properties.Items.Add(dt.Columns[i].ColumnName);
            }
            this.checkedComboBox_dataItems.CheckAll();//默认全选
            this.m_LineChartModel.DataTable = dt;
            this.btn_apply_Click(this.btn_apply, EventArgs.Empty);
            this.btn_apply.Enabled = true;
            this.btn_chartin.Enabled = true;
            this.btn_export.Enabled = true;


        }
        /// <summary>
        /// 绘制条形图统计图
        /// </summary>
        /// <param name="model">数据模型</param>
        private void Draw(LineChartModel model=null) {
            if (this.m_Chart != null && !this.m_Chart.IsDisposed) {
                this.m_Chart.Dispose();
            }
            TChart chart = new TChart();
            chart.Dock = DockStyle.Fill;
            this.m_ChartForm.Controls.Add(chart);
            this.m_Chart = chart;
            
            chart.Tag = model;
            chart.Series.Clear();//清除，重绘，避免重新创建
            chart.AutoRepaint = false;//关闭自动绘制
            chart.Panel.MarginUnits = PanelMarginUnits.Pixels;//设置间隔单位，用于控制Pie与图例的位置关系
            chart.Panel.MarginLeft = 0;
            chart.Panel.MarginRight = 0;
            chart.Panel.MarginBottom = 0;
            chart.Panel.MarginTop = 0;
            chart.Panel.Bevel.Inner = BevelStyles.None;
            chart.Panel.Bevel.Outer = BevelStyles.None;
            chart.Chart.CustomChartRect = true;//自定义图表的位置和范围
            chart.Panel.Transparent = true;//图表区透明
            chart.Walls.Back.Transparent = true;//图表区透明
            chart.Legend.MaxNumRows = 0;
            chart.Legend.Alignment = Steema.TeeChart.LegendAlignments.Bottom;
            chart.Axes.Bottom.Grid.Visible = false;
            chart.Axes.Left.Grid.Visible = true;
            chart.Axes.Left.MinorTicks.Visible = false;//关掉小刻度线，其他刻度关不掉

            #region 废弃
            //chart.Axes.Left.Ticks.Visible = false;
            //chart.Axes.Left.TicksInner.Visible = false;
            //chart.Axes.Left.Ticks.Length = 0;
            //chart.Axes.Left.TicksInner.Length = 0;
            //chart.Axes.Left.TickOnLabelsOnly = false;
            //chart.Axes.Left.Ticks.Color = Color.Transparent;
            //chart.Axes.Left.Ticks.DrawingPen.Color = Color.Transparent;
            #endregion
            //通用尺寸
            this.m_ChartForm.Width = (int)Util.CentimeterToPixel(model.Width) + 16;//16为补充form容器的边框宽度，从而保证TChart尺寸的正确性
            this.m_ChartForm.Height = (int)Util.CentimeterToPixel(model.Height) + 39;//39为补充form容器的边框高度，从而保证TChart尺寸的正确性
            chart.Aspect.View3D = model.View3D;
            if(model.View3D)
                chart.Walls.Visible = true;
            else
                chart.Walls.Visible = false;
            //定义个矩形范围，默认按照chart的大小来控制，默认左右上下都留白
            int blankSize = 10;
            //chart.chartRect 仅包含绘图区，不包含坐标轴和X,Y轴的标注，但就算好后，其内部会自适应预留标注的尺寸
            Rectangle chartRect = new Rectangle()
            {
                X = blankSize,
                Y = blankSize,
                Width = chart.Width - 2*blankSize,
                Height = chart.Height - 2*blankSize
            };
            ////========标题设置
            chart.Header.Visible = false;//关闭默认标题，采用自定义渲染
            if (model.Title.Visible)
            {
                SizeF sf = chart.Graphics3D.MeasureString(this.WindowFontToChartFont(model.Title.Font.Font, model.Title.Font.Color), model.Title.Text);
                model.Title.Height = (int)Math.Ceiling(sf.Height);
                model.Title.Width = (int)Math.Ceiling(sf.Width);
                model.Title.Left = (chart.Width - model.Title.Width) / 2;
                //调整 chartRect 的范围
                chartRect.Y = model.Title.Top + model.Title.Height + blankSize;
                chartRect.Height = chart.Height - chartRect.Y - blankSize;
            }
            //========根据标注字段进行赋值，默认DataTable第一列为分类，第二列为数据
            //根据model初始化Series
            if (model.Series.Count >= 0)
            {
                if (model.DataTable == null)
                    return;
                for (int i = 0; i < model.Series.Count; i++)
                {
                    LineSeries barSeries = model.Series[i];
                    Line series = new Line();
                    chart.Series.Add(series);
                    /*
                       series.Pen.Color = Color.Transparent;//关闭掉边框线
                       series.BarStyle = BarStyles.Rectangle;
                       if (model.Stack)
                       series.MultiBar = MultiBars.None;
                       else
                       series.MultiBar = MultiBars.Side;
                       if (model.BarWidth != 0)
                       series.CustomBarWidth = (int)Math.Round(Util.CentimeterToPixel(model.BarWidth));
                     */
                    if(barSeries.Color!=Color.Empty)
                        series.Color = barSeries.Color;
                    for (int j = 0; j < model.DataTable.Rows.Count; j++)
                    {
                        DataRow dr = model.DataTable.Rows[j];
                        series.Add(Convert.ToDouble(dr[barSeries.Name]), dr[0].ToString());
                    }

                    series.Smoothed = false;
                    if (Convert.ToInt32(model.PointerWidth) != 0)
                    {
                        series.Pointer.Visible = model.Pointer;
                    }
                    series.Pointer.VertSize = Convert.ToInt32(model.PointerWidth);
                    series.Pointer.HorizSize = Convert.ToInt32(model.PointerWidth);
                    series.Pointer.Style = model.PointerStyles;
                }
            }
            //=========条形图标注设置
            for (int i = 0; i < chart.Series.Count; i++)
            {
                var series = chart.Series[i];
                series.Marks.Arrow.Visible = false;
                series.Marks.Transparent = true;
                series.Marks.Arrow.Visible = false;
                series.Marks.ArrowLength = 0;
                series.Marks.Visible = model.LineLabel.Visible;
                series.Marks.Font = this.WindowFontToChartFont(model.LineLabel.Font.Font, model.LineLabel.Font.Color);
                series.Marks.Style = (MarksStyles)LabelStyle.数值;

                // 偶数序号系列(0,2,4...)标签在上，奇数序号系列(1,3,5...)标签在下
                if (i % 2 == 0)
                    series.Marks.ArrowLength += 10;
                //else
                    //series.Marks.ArrowLength -= 10;
            }

            //=========图例设置
            chart.Legend.Visible = false;//关闭自带图例，采用手动计算
            if (model.Legend.Visible)
            {
                ChartLegend legend = model.Legend;
                #region 1、计算图例的尺寸（内部元素相对位置）
                legend.LegendItems.Clear();
                //1、获取Series的Label，计算尺寸
                int itemMaxWidth = 0;
                for (int i = 0; i < model.Series.Count; i++)
                {
                    LegendItem li = new LegendItem();
                    var series = model.Series[i];
                    //计算图例项的宽高
                    SizeF sf = chart.Graphics3D.MeasureString(new ChartFont(chart.Chart, model.Legend.Font.Font), series.Name);
                    li.Width = (int)Math.Ceiling(sf.Width) + model.Legend.SymbolTextGap + model.Legend.SymbolWidth;//文本宽度+符号宽度+文本符号间隔宽度
                    li.Height = (int)Math.Ceiling(sf.Height);//文本高度，颜色通过库自动赋值后，获取，先计算尺寸
                    li.Text = series.Name;
                    legend.LegendItems.Add(li);
                    if (li.Width > itemMaxWidth)
                    {
                        itemMaxWidth = li.Width;
                    }
                }
                int colNum = model.Legend.colNum;
                int average = legend.LegendItems.Count / colNum;//计算每列分配多少个item的平均数
                int mod = legend.LegendItems.Count % colNum;//计算平均数后的余数
                int[] colItemNums = new int[colNum];//计算每列item的个数
                for (int i = 0; i < colItemNums.Length; i++)
                {
                    colItemNums[i] = average;
                    if (mod != 0)
                    {
                        colItemNums[i] += 1;
                        mod -= 1;
                    }
                }

                //计算每个item的位置(Left,Top)，相对于Legend的（Left,Top），然后计算Symbol Rect
                int startIndex = 0;
                for (int i = 0; i < colItemNums.Length; i++)
                {
                    int colItemNum = colItemNums[i];
                    int left = 0;
                    if (i != 0)
                        left = i * itemMaxWidth + i * model.Legend.HorizonGap;
                    for (int j = 0; j < colItemNum; j++)
                    {
                        LegendItem li = model.Legend.LegendItems[startIndex + j];
                        int top = 0;
                        if (j != 0)
                            top = j * li.Height + j * model.Legend.VerticalGap;
                        li.Left = left;
                        li.Top = top;
                        //计算Symbol Rect
                        li.SymbolRect = new Rectangle()
                        {
                            X = left,
                            Y = top,
                            Width = model.Legend.SymbolWidth,
                            Height = li.Height
                        };
                    }
                    startIndex += colItemNum;
                }

                //计算整个Legend的长宽
                legend.Width = legend.LegendItems[0].Left + legend.LegendItems[legend.LegendItems.Count - 1].Left + itemMaxWidth;
                legend.Height = legend.LegendItems[0].Top + legend.LegendItems[colItemNums[0] - 1].Top + legend.LegendItems[0].Height;
                #endregion

                #region 2、计算图例的在容器中的位置，同时更新计算：图表（chartRect）的位置和范围
                int headerTop = model.Title.Top;
                int headerHeight = model.Title.Height;
                if (!model.Title.Visible)
                {
                    headerTop = 0;
                    headerHeight = 0;
                }
                switch (model.Legend.Postion)
                {
                    case LegendPostion.顶部靠左:
                        model.Legend.Top = headerTop + headerHeight + blankSize;
                        model.Legend.Left = blankSize;
                        //根据图例的最短边为参考，例如偏树直型，只做X方向限制，Y方向不受图例高度影响，水平型类似计算图表位置
                        if (model.Legend.Width < model.Legend.Height)
                        {
                            chartRect.X = model.Legend.Left + model.Legend.Width;
                            chartRect.Y = model.Title.Top + model.Title.Height + blankSize;
                            chartRect.Width = chart.Width - chartRect.X - blankSize;
                            chartRect.Height = chart.Height - chartRect.Y - blankSize;
                        }
                        else if (model.Legend.Width > model.Legend.Height)//偏水平型，只做Y方向限制，X方向不受图例宽度影响
                        {
                            chartRect.X = blankSize;
                            chartRect.Y = model.Legend.Top + model.Legend.Height;
                            chartRect.Width = chart.Width - chartRect.X - blankSize;
                            chartRect.Height = chart.Height - chartRect.Y - blankSize;
                        }
                        else//都考虑
                        {
                            chartRect.X = model.Legend.Left + model.Legend.Width;
                            chartRect.Y = model.Legend.Top + model.Legend.Height;
                            chartRect.Width = chart.Width - chartRect.X - blankSize;
                            chartRect.Height = chart.Height - chartRect.Y - blankSize;
                        }
                        break;
                    case LegendPostion.顶部靠右:
                        model.Legend.Top = headerTop + headerHeight + blankSize;
                        model.Legend.Left = chart.Width - model.Legend.Width - blankSize;
                        //根据图例的最短边为参考，例如偏树直型，只做X方向限制，Y方向不受图例高度影响，水平型类似计算图表位置
                        if (model.Legend.Width < model.Legend.Height)
                        {
                            chartRect.X = blankSize;
                            chartRect.Y = model.Title.Top + model.Title.Height + blankSize;
                            chartRect.Width = chart.Width - chartRect.X - model.Legend.Width - blankSize;
                            chartRect.Height = chart.Height - chartRect.Y - blankSize;
                        }
                        else if (model.Legend.Width > model.Legend.Height)//偏水平型，只做Y方向限制，X方向不受图例宽度影响
                        {
                            chartRect.X = blankSize;
                            chartRect.Y = model.Legend.Top + model.Legend.Height + blankSize;
                            chartRect.Width = chart.Width - chartRect.X - blankSize;
                            chartRect.Height = chart.Height - chartRect.Y - blankSize;
                        }
                        else
                        {//都考虑
                            chartRect.X = 10;
                            chartRect.Y = model.Legend.Top + model.Legend.Height + blankSize;
                            chartRect.Width = chart.Width - chartRect.X - model.Legend.Width - blankSize;
                            chartRect.Height = chart.Height - chartRect.Y - blankSize;
                        }
                        break;
                    case LegendPostion.顶部居中:
                        model.Legend.Top = headerTop + headerHeight + blankSize;
                        model.Legend.Left = (chart.Width - model.Legend.Width) / 2;
                        //计算图表位置
                        chartRect.X = 10;
                        chartRect.Y = model.Legend.Top + model.Legend.Height;
                        chartRect.Width = chart.Width - blankSize;
                        chartRect.Height = chart.Height - chartRect.Y - blankSize;
                        break;
                    case LegendPostion.底部靠左:
                        model.Legend.Top = chart.Height - model.Legend.Height - blankSize;
                        model.Legend.Left = blankSize;
                        //根据图例的最短边为参考，例如偏树直型，只做X方向限制，Y方向不受图例高度影响，水平型类似计算图表位置
                        if (model.Legend.Width < model.Legend.Height)
                        {
                            chartRect.X = model.Legend.Left + model.Legend.Width;
                            chartRect.Y = model.Title.Top + model.Title.Height + blankSize;
                            chartRect.Width = chart.Width - chartRect.X - blankSize;
                            chartRect.Height = chart.Height - chartRect.Y - 2*blankSize;
                        }
                        else if (model.Legend.Width > model.Legend.Height)//偏水平型，只做Y方向限制，X方向不受图例宽度影响
                        {
                            chartRect.X = blankSize;
                            chartRect.Y = model.Title.Top + model.Title.Height + blankSize;
                            chartRect.Width = chart.Width - chartRect.X - blankSize;
                            chartRect.Height = chart.Height - chartRect.Y - model.Legend.Height - 2*blankSize;
                        }
                        else
                        {//都考虑
                            chartRect.X = blankSize;
                            chartRect.Y = model.Title.Top + model.Title.Height + blankSize;
                            chartRect.Width = chart.Width - chartRect.X - blankSize;
                            chartRect.Height = chart.Height - chartRect.Y - model.Legend.Height - 2*blankSize;
                        }
                        break;
                    case LegendPostion.底部居中:
                        model.Legend.Top = chart.Height - model.Legend.Height - blankSize;
                        model.Legend.Left = (chart.Width - model.Legend.Width) / 2;
                        //计算图表位置
                        chartRect.X = blankSize;
                        chartRect.Y = (model.Title.Top + model.Title.Height) + blankSize;
                        chartRect.Width = chart.Width - blankSize;
                        chartRect.Height = chart.Height - chartRect.Y - model.Legend.Height - 2*blankSize;
                        break;
                    case LegendPostion.底部靠右:
                        model.Legend.Top = chart.Height - model.Legend.Height - blankSize;
                        model.Legend.Left = chart.Width - model.Legend.Width - blankSize;
                        //根据图例的最短边为参考，例如偏树直型，只做X方向限制，Y方向不受图例高度影响，水平型类似计算图表位置
                        if (model.Legend.Width < model.Legend.Height)
                        {
                            chartRect.X = blankSize;
                            chartRect.Y = model.Title.Top + model.Title.Height + blankSize;
                            chartRect.Width = chart.Width - chartRect.X - model.Legend.Width - blankSize;
                            chartRect.Height = chart.Height - chartRect.Y - 2*blankSize;
                        }
                        else if (model.Legend.Width > model.Legend.Height)//偏水平型，只做Y方向限制，X方向不受图例宽度影响
                        {
                            chartRect.X = blankSize;
                            chartRect.Y = model.Title.Top + model.Title.Height + blankSize;
                            chartRect.Width = chart.Width - chartRect.X - blankSize;
                            chartRect.Height = chart.Height - chartRect.Y - model.Legend.Height - 2*blankSize;
                        }
                        else
                        {//都考虑
                            chartRect.X = blankSize;
                            chartRect.Y = model.Title.Top + model.Title.Height + blankSize;
                            chartRect.Width = chart.Width - chartRect.X - model.Legend.Width - blankSize;
                            chartRect.Height = chart.Height - chartRect.Y - model.Legend.Height - 2 * blankSize;
                        }
                        break;
                    case LegendPostion.左中:
                        model.Legend.Top = (chart.Height - model.Legend.Height) / 2;
                        model.Legend.Left = blankSize;
                        //计算图表位置
                        chartRect.X = model.Legend.Left + model.Legend.Width + blankSize;
                        chartRect.Y = (model.Title.Top + model.Title.Height) + blankSize;
                        chartRect.Width = chart.Width - chartRect.X - blankSize;
                        chartRect.Height = chart.Height - chartRect.Y - blankSize;
                        break;
                    case LegendPostion.右中:
                        model.Legend.Left = chart.Width - model.Legend.Width - blankSize;
                        model.Legend.Top = (chart.Height - model.Legend.Height) / 2;
                        //计算图表位置
                        chartRect.X = 10;
                        chartRect.Y = (model.Title.Top + model.Title.Height) + blankSize;
                        chartRect.Width = chart.Width - chartRect.X - model.Legend.Width - blankSize;
                        chartRect.Height = chart.Height - chartRect.Y - blankSize;
                        break;
                }
                #endregion
            }

            //=========X轴横轴标注设置
            chart.Axes.Bottom.Labels.Visible = model.XAxisLabel.Visible;
            if (model.XAxisLabel.Visible)
            {
                chart.Axes.Bottom.Labels.Angle = 0;
                chart.Axes.Bottom.Labels.Font = this.WindowFontToChartFont(model.XAxisLabel.Font.Font, model.XAxisLabel.Font.Color);
                chart.Axes.Bottom.Automatic = true;
                chart.Axes.Bottom.Labels.Separation = 0;
                chart.Axes.Bottom.Labels.CustomSize = 15;
                chart.Axes.Bottom.Labels.Alternate = true;
            }
            if (model.XAxisLabel.Title.Visible)
            {   //Title尺寸计算，后续自定义绘制
                SizeF sf = chart.Graphics3D.MeasureString(this.WindowFontToChartFont(model.XAxisLabel.Title.Font.Font, model.XAxisLabel.Title.Font.Color), model.XAxisLabel.Title.Text);
                model.XAxisLabel.Title.Height = (int)Math.Ceiling(sf.Height);
                model.XAxisLabel.Title.Width = (int)Math.Ceiling(sf.Width);
                //这里推测计算Top和Left，并以此调整图区的范围，参照百度Echarts的标题的摆放设计，进行实现
                //TeeChart的情况，设置了ChartRect后，其会根据设定的范围，内部会进行X,Y轴及标注，图区的范围重新计算，最终计算的出来的ChartRect为图区的范围，并并不包含行X,Y轴及标注范围
                //根据上述情况，进行预先把计算推测计算Top和Left
                int XAxistHeight = 10;//渲染后出来的可以获取的值，这里提前设置
                SizeF sf2 = chart.Graphics3D.MeasureString(this.WindowFontToChartFont(model.XAxisLabel.Font.Font, model.XAxisLabel.Font.Color), "1");//随便给个字，只是为了获取高度
                model.XAxisLabel.Title.Top = chartRect.Top + chartRect.Height - XAxistHeight - (int)Math.Round(sf2.Height) - model.XAxisLabel.Title.Height / 3;//X轴作为其标注的竖直中间位置
                model.XAxisLabel.Title.Left = chartRect.Left + chartRect.Width - model.XAxisLabel.Title.Width;
                //调整 chartRect 的范围
                chartRect.Width -= model.XAxisLabel.Title.Width;
            }
            //=========Y轴横轴标注设置
            chart.Axes.Left.Labels.Visible = model.YAxisLabel.Visible;
            if (model.YAxisLabel.Visible)
            {
                chart.Axes.Left.Labels.Angle = 0;
                chart.Axes.Left.Labels.Font = this.WindowFontToChartFont(model.YAxisLabel.Font.Font, model.YAxisLabel.Font.Color);
            }
            if (model.YAxisLabel.Title.Visible)
            {
                SizeF sf = chart.Graphics3D.MeasureString(this.WindowFontToChartFont(model.YAxisLabel.Title.Font.Font, model.YAxisLabel.Title.Font.Color), model.YAxisLabel.Title.Text);
                model.YAxisLabel.Title.Height = (int)Math.Ceiling(sf.Height);
                model.YAxisLabel.Title.Width = (int)Math.Ceiling(sf.Width);
                //这里推测计算Top和Left，并以此调整图区的范围，参照百度Echarts的标题的摆放设计，进行实现
                //TeeChart的情况，设置了ChartRect后，其会根据设定的范围，内部会进行X,Y轴及标注，图区的范围重新计算，最终计算的出来的ChartRect为图区的范围，并并不包含行X,Y轴及标注范围
                //根据上述情况，进行预先把计算推测计算Top和Left
                int YAxistWidth = 10;//渲染后出来的可以获取的值，这里提前设置
                int YAxisMaxLabel = chart.Axes.Left.MaxLabelsWidth();
                int Separate = chart.Axes.Left.Labels.Separation;
                model.YAxisLabel.Title.Top = chartRect.Top + 10;
                model.YAxisLabel.Title.Left = chartRect.Left + YAxisMaxLabel + YAxistWidth + Separate + 5 -model.YAxisLabel.Title.Width / 2;//Y轴作为其标注的水平中间位置，5是标注到轴的距离
                //调整 chartRect 的范围
                chartRect.Y += model.XAxisLabel.Title.Height + 10;// / 2;
                chartRect.Height -= model.XAxisLabel.Title.Height + 10;// / 2;
            }
            chart.Chart.ChartRect = chartRect;
            chart.BeforeDraw += new PaintChartEventHandler(chart_BeforeDraw);
            chart.AfterDraw += new PaintChartEventHandler(chart_AfterDraw);
        }


        private void chart_BeforeDraw(object sender,Graphics3D g) {
            TChart chart = sender as TChart;
            //绘制图表背景色
            g.Brush.Color = Color.White;//填充色
            g.Pen.Color = Color.White;//覆盖颜色
            g.Rectangle(0, 0, chart.Width, chart.Height);
        }

        /// <summary>
        /// 覆盖的边框线以及其他
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="g"></param>
        private void chart_AfterDraw(object sender, Graphics3D g)
        {
            TChart chart = sender as TChart;
            #region 获取最大值并调大一些（为了避免边缘文字被裁切覆盖，影响效果）和最小值（默认还是0，避免出现负值）
            double maxValue = chart.Axes.Bottom.MaxXValue;
            double labelMax = Math.Ceiling(maxValue / chart.Axes.Bottom.CalcIncrement) * chart.Axes.Bottom.CalcIncrement;//按照渲染后其计算出来的刻度，进行四舍五入
            chart.Axes.Bottom.SetMinMax(chart.Axes.Bottom.Minimum, labelMax);
            #endregion
            chart.AfterDraw -= new PaintChartEventHandler(chart_AfterDraw);
            chart.AfterDraw+=new PaintChartEventHandler(chart_AfterDraw3);
            chart.Refresh();
        }
        /// <summary>
        /// 绘制定义的所有元素
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="g"></param>
        private void chart_AfterDraw3(object sender, Graphics3D g) {
            TChart chart = sender as TChart;
            LineChartModel model = chart.Tag as LineChartModel;
            #region 盖掉图区上面和右侧的边框线，未找到其他关闭方式
            //int _x = chart.Chart.ChartRect.Left; //yAxis.CalcPosValue(yAxis.Minimum);
            //int _y = chart.Chart.ChartRect.Top;
            //int _w = chart.Chart.ChartRect.Width;
            //int _h = chart.Chart.ChartRect.Height+1;
            //g.Pen.Color = Color.White;
            //g.Pen.Width = 3;
            //g.MoveTo(_x, _y);
            //g.LineTo(_x + _w, _y);
            //g.LineTo(_x + _w, _y + _h+1);
            #endregion
            //根据渲染的结果存储series的颜色
            for (int i = 0; i < chart.Series.Count; i++)
            {
                if (model.Series[i].Color == Color.Empty)
                    model.Series[i].Color = chart.Series[i].Color;
            }
            //===========================绘制标题================================
            if (model.Title.Visible)
            {
                int textTop = model.Title.Top;
                int textLeft = model.Title.Left;
                g.Font = this.WindowFontToChartFont(model.Title.Font.Font, model.Title.Font.Color);
                g.TextOut(textLeft, textTop, model.Title.Text);//支持自定义字体,可以导出SVG
            }
            //===========================绘制图例================================
            if (model.Legend.Visible)
            {
                for (int i = 0; i < model.Legend.LegendItems.Count; i++)
                {
                    LegendItem li = model.Legend.LegendItems[i];
                    li.SymbolColor = chart.Series[i].Color;//通过这种方式获取扇区颜色，根据渲染结果更新图例的颜色
                    //model.Legend.LegendItems[i].SymbolColor = series[i].Color;//这种方式只有白色
                    //绘制符号框
                    int absTop = model.Legend.Top + li.Top;
                    int absLeft = model.Legend.Left + li.Left;
                    //绘制填充
                    g.Pen.Color = Color.Black;//边框色
                    g.Pen.Width = 1;//边框宽度
                    g.Brush.Color = li.SymbolColor;//填充色
                    g.Rectangle(absLeft, absTop, absLeft + li.SymbolRect.Width, absTop + li.SymbolRect.Height);
                    //绘制文本
                    int textLeft = absLeft + li.SymbolRect.Width + model.Legend.SymbolTextGap;
                    int textTop = absTop;
                    g.Font = this.WindowFontToChartFont(model.Legend.Font.Font, model.Legend.Font.Color);
                    g.TextOut(textLeft, textTop, li.Text);//支持自定义字体,但可以导出SVG
                }
            }
            //===========================绘制X，Y轴标题================================
            if (model.XAxisLabel.Title.Visible)
            {
                g.Font = this.WindowFontToChartFont(model.XAxisLabel.Title.Font.Font, model.XAxisLabel.Title.Font.Color);
                g.TextOut(model.XAxisLabel.Title.Left, model.XAxisLabel.Title.Top, model.XAxisLabel.Title.Text);//支持自定义字体,但可以导出SVG
            }
            if(model.YAxisLabel.Title.Visible)
            {
                g.Font = this.WindowFontToChartFont(model.YAxisLabel.Title.Font.Font, model.YAxisLabel.Title.Font.Color);
                g.TextOut(model.YAxisLabel.Title.Left, model.YAxisLabel.Title.Top, model.YAxisLabel.Title.Text);//支持自定义字体,但可以导出SVG
            }
        }


        /// <summary>
        /// Excel转DataTable
        /// </summary>
        /// <param name="excelPath"></param>
        private DataTable ExcelToDataTable(string excelPath) {
            IWorkbook workBook = null;
            string fileExt = System.IO.Path.GetExtension(excelPath).ToLower();
            using (FileStream fs = new FileStream(excelPath, FileMode.Open, FileAccess.Read))
            {
                if (fileExt == ".xlsx")
                {
                    workBook = new XSSFWorkbook(fs);
                }
                else if (fileExt == ".xls")
                {
                    workBook = new HSSFWorkbook(fs);
                }
            }

            ISheet workSheet = workBook.GetSheetAt(0);
            DataTable dataTable = new DataTable();
            //获取行标题                             
            for (int i = 0; i < workSheet.GetRow(0).LastCellNum; i++)
            {
                if (workSheet.GetRow(0) == null || workSheet.GetRow(0).GetCell(i) == null || workSheet.GetRow(1).GetCell(i) == null) 
                    continue;
                dataTable.Columns.Add(ParseCellVal(workSheet.GetRow(0).GetCell(i)).ToString());
            }
            //获取数据表
            for (int i = 1; i < workSheet.LastRowNum; i++)
            {
                IRow row = workSheet.GetRow(i);
                if (row == null || row.GetCell(0) == null || row.GetCell(0).ToString() == "") 
                    continue;
                DataRow dr= dataTable.NewRow();
                for (int j = 0; j < row.LastCellNum; j++)
                {
                    dr[j] =ParseCellVal(row.GetCell(j));
                }
                dataTable.Rows.Add(dr);
            }
            return dataTable;
        }
        ///// <summary>
        ///// 矢量单位化
        ///// </summary>
        ///// <param name="p"></param>
        ///// <returns></returns>
        //private Vector2D Normalization(Vector2D p)
        //{
        //    double vectorMod = Math.Sqrt(Math.Pow(p.X, 2) + Math.Pow(p.Y,2));
        //    return new Vector2D(p.X / vectorMod, p.Y / vectorMod);
        //}

        private object ParseCellVal(ICell cell)
        {
            object val = null;
            switch (cell.CellType)
            {
                case CellType.String:
                    val = cell.StringCellValue;
                    break;
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        val = cell.DateCellValue.ToString();
                    }
                    else
                    {
                        val = cell.NumericCellValue;
                    }
                    break;
                case CellType.Boolean:
                    val = cell.BooleanCellValue;
                    break;
            }
            return val;
        }


        private ChartFont WindowFontToChartFont(Font font,Color color)
        {
            ChartFont cfont = new ChartFont();
            if (font.Style == FontStyle.Bold)
            {
                cfont.Bold = true;
                cfont.Italic = false;
            }
            else if (font.Style == FontStyle.Regular)
            {
                cfont.Bold = false;
                cfont.Italic = false;
            }
            else if (font.Style == FontStyle.Italic)
            {
                cfont.Italic = true;
                cfont.Bold = false;
            }
            cfont.Name = font.Name;
            cfont.Size = (int)font.Size;
            cfont.Color = color;
            return cfont;
        }

        /// <summary>
        ///导出
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_export_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "SVG矢量文件|*.svg";
            if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                //序列化对象model
                string modelJson = JsonConvert.SerializeObject(this.m_LineChartModel);
                // 创建内存流来保存SVG数据
                using (MemoryStream stream = new MemoryStream())
                {
                    //将图表导出为SVG并保存到流
                    this.m_Chart.Export.Image.SVG.Save(stream);
                    // 将流位置重置为开始，以便读取
                    stream.Position = 0;
                    SvgDocument svgDoc = SvgDocument.Open(stream);
                    svgDoc.CustomAttributes.Add("smgi_customMetaData", modelJson);
                    svgDoc.Write(sfd.FileName);
                }
                MessageBox.Show("导出完毕！");
            }
        }
        /// <summary>
        /// 导入到地图中
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_chartin_Click(object sender, EventArgs e)
        {
            //序列化对象model
            string modelJson = JsonConvert.SerializeObject(this.m_LineChartModel);
            // 创建内存流来保存SVG数据
            using (MemoryStream stream = new MemoryStream())
            {
                //将图表导出为SVG并保存到流
                this.m_Chart.Export.Image.SVG.Save(stream);
                // 将流位置重置为开始，以便读取
                stream.Position = 0;
                SvgDocument svgDoc = SvgDocument.Open(stream);
                double PixWidth = Util.CentimeterToPixel(this.m_LineChartModel.Width);
                double PixHeight = Util.CentimeterToPixel(this.m_LineChartModel.Height);
                svgDoc.Width = new SvgUnit((float)PixWidth);
                svgDoc.Height = new SvgUnit((float)PixHeight);
                svgDoc.ViewBox = new SvgViewBox(0, 0, svgDoc.Width.Value, svgDoc.Height.Value);
                //Svg 文本字体渲染的起始点坐标不是左上角点，可以理解为左下角点;
                //但栅格图片(Bitmap)渲染默认按照该坐标进行左上角渲染，会出现文字下沉一个字高的问题
                //在渲染前，进行文本块定位的Y值调整；
                for (int i = 0; i < svgDoc.Children.Count; i++)
                {
                    SvgElement svgEle = svgDoc.Children[i];
                    if (svgEle is SvgText)
                    {
                        SvgText svgText = svgEle as SvgText;
                        float fontSize = svgText.FontSize.ToDeviceValue();//这里计算出来的pt值，需要乘以1.3转成像素值
                        float yOffset = fontSize * 1.3f;//估算基线偏移（经验值）96/72约等于1.3
                        svgText.Y += -yOffset;
                    }
                }
                //按照300dpi的分辨率进行放大处理，缩小ViewBox，就是放大
                svgDoc.Ppi = 300;//dpi约等于ppi;
                double multi = 300 / 96;
                SvgViewBox svb = new SvgViewBox();
                svb.MinX = 0;
                svb.MinY = 0;
                svb.Width = (int)(PixWidth / multi);
                svb.Height = (int)(PixHeight / multi);
                svgDoc.ViewBox = svb;

                // 同比放大Bitmap的大小
                int bitmapWidth = (int)(PixWidth * multi);
                int bitmapHeight = (int)(PixHeight * multi);

                using (Bitmap bitmap = new Bitmap(bitmapWidth, bitmapHeight))
                {
                    bitmap.SetResolution(300f, 300f);
                    svgDoc.Draw(bitmap);
                    //// 获取临时目录路径
                    string tempDir = System.IO.Path.GetTempPath();
                    // 生成唯一的文件名
                    string fileName = DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".png";
                    string filePath = System.IO.Path.Combine(tempDir, fileName);
                    //// 保存
                    bitmap.Save(filePath);
                    ////构建ArcGIS图片元素
                    IPictureElement picElement = new PngPictureElementClass();
                    picElement.ImportPictureFromFile(filePath);
                    picElement.MaintainAspectRatio = true;
                    picElement.SavePictureInDocument = true;
                    //计算图片在ArcGIS中的长宽（pt）已经验证：1pt=96/72px;
                    //double picWidth = 0, picHeight = 0;
                    //(picElement as IPictureElement5).QueryIntrinsicSize(ref picWidth, ref picHeight);//不能按照这种方式计算图片的宽高
                    //picWidth = picWidth / 2.8345 * 0.1;//磅转毫米转厘米
                    //picHeight = picHeight / 2.8345 * 0.1;//磅转毫米转厘米
                    IPageLayout pageLayout = this.m_Application.PageLayoutControl.PageLayout;
                    IPoint pageCenterPoint = (this.m_Application.PageLayoutControl.ActiveView.Extent as IArea).Centroid;
                    //pageLayout.Page.Units== ESRI.ArcGIS.esriSystem.esriUnits.esriCentimeters;//这里布局视图为厘米
                    double picWidth = this.m_LineChartModel.Width;
                    double picHeight = this.m_LineChartModel.Height;
                    IGraphicsContainer graphicsContainer = (IGraphicsContainer)pageLayout;
                    IEnvelope envelope = new EnvelopeClass();
                    envelope.PutCoords(pageCenterPoint.X - picWidth / 2, pageCenterPoint.Y - picHeight / 2, pageCenterPoint.X + picWidth / 2, pageCenterPoint.Y + picHeight / 2);
                    (picElement as IElement).Geometry = envelope;
                    (picElement as IElementProperties).CustomProperty = modelJson;
                    string name = "Smgi_Line_Chart" + System.IO.Path.GetFileNameWithoutExtension(fileName);
                    (picElement as IElementProperties).Name = name;
                    graphicsContainer.AddElement(picElement as IElement, 0);
                    this.m_Application.PageLayoutControl.ActiveView.Refresh();
                }
            }
            MessageBox.Show("已经导入到地图布局视图中！");
        }

        private void smoothedCheckEdit_CheckedChanged(object sender, EventArgs e)
        {
            m_LineChartModel.Smoothed = smoothedCheckEdit.Checked;
        }

        private void pointerCheckEdit_CheckedChanged(object sender, EventArgs e)
        {
            m_LineChartModel.Pointer = pointerCheckEdit.Checked;
        }

        private void pointerComboBoxEdit_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (pointerComboBoxEdit.SelectedItem.ToString())
            {
                case "圆形":
                    m_LineChartModel.PointerStyles = PointerStyles.Circle;
                    break;
                case "方形":
                    m_LineChartModel.PointerStyles = PointerStyles.Rectangle;
                    break;
                case "三角":
                    m_LineChartModel.PointerStyles = PointerStyles.Triangle;
                    break;
                case "五角星":
                    m_LineChartModel.PointerStyles = PointerStyles.Star;
                    break;
                case "六角星":
                    m_LineChartModel.PointerStyles = PointerStyles.Hexagon;
                    break;
                case "菱形":
                    m_LineChartModel.PointerStyles = PointerStyles.Diamond;
                    break;
                default:
                    m_LineChartModel.PointerStyles = PointerStyles.Circle;
                    break;
            }
        }

        private void pointerWidthSpinEdit_EditValueChanged(object sender, EventArgs e)
        {
            m_LineChartModel.PointerWidth = (int)(sender as DevExpress.XtraEditors.SpinEdit).Value;
        }
    }
}
