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

namespace SMGI.Plugin.ThematicChart.TeeChart.PieChart
{
    public partial class RadarParameterPanel : Form
    {
        private GApplication m_Application;
        /// <summary>
        /// 雷达图设置参数模型
        /// </summary>
        private RadarChartModel m_RadarChartModel;

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

        TChart chart = new TChart();

        public RadarParameterPanel(GApplication app,RadarChartModel pcm=null)
        {
            InitializeComponent();
            m_Application = app;
            if (pcm == null)
            {
                m_RadarChartModel = new RadarChartModel();

            }
            else {
                m_RadarChartModel = pcm;
            }

            m_ChartForm = new ChartForm();
            m_ChartForm.ShowInTaskbar = false;
            m_ChartForm.StartPosition = FormStartPosition.CenterScreen;
            var hwnd = app.MainForm.Handle;
            var interop = new System.Windows.Forms.NativeWindow();
            interop.AssignHandle(hwnd);
            m_ChartForm.Show(interop);

            this.initUIParam();
            //this.m_Chart = this.m_ChartForm.TeeChart;
            this.Disposed+=new EventHandler(ParameterPanel_Disposed);
        }

        /// <summary>
        /// 初始化UI参数
        /// </summary>
        private void initUIParam(){
           //通用设置
            this.num_chartWidth.Value = (decimal)this.m_RadarChartModel.Width;
            this.num_chartHeight.Value = (decimal)this.m_RadarChartModel.Height;
           this.check_3d.Checked = this.m_RadarChartModel.View3D;
           this.check_ring.Checked = this.m_RadarChartModel.ViewRing;
           this.num_radius.Value = (decimal)this.m_RadarChartModel.CustomRadius;
           //标题设置
           this.check_title.Checked = this.m_RadarChartModel.Title.Visible;
           this.check_title_CheckedChanged(this.check_title, EventArgs.Empty);
           this.txt_title.Text = this.m_RadarChartModel.Title.Text;
           this.label_titleStyle.Text = this.m_RadarChartModel.Title.Font.ToString();
           this.label_titleColor.BackColor = this.m_RadarChartModel.Title.Font.Color;

           //标注设置
           this.check_label.Checked = this.m_RadarChartModel.RadarLabel.Visible;
           this.check_label_CheckedChanged(this.check_label, EventArgs.Empty);
           this.label_labelStyle.Text = this.m_RadarChartModel.RadarLabel.Font.ToString();
           this.label_labelStyle.ForeColor = this.m_RadarChartModel.RadarLabel.Font.Color;
           this.combo_labelStyle.SelectedItem = Enum.GetName(typeof(LabelStyle), this.m_RadarChartModel.RadarLabel.LabelStyle);
           this.label_labelStyle.Text = this.m_RadarChartModel.RadarLabel.Font.ToString();
           this.label_labelColor.BackColor = this.m_RadarChartModel.RadarLabel.Font.Color;
           if (this.m_RadarChartModel.DataTable != null)
           {
               DataTable dt = this.m_RadarChartModel.DataTable;
               for (int i = 1; i < dt.Columns.Count; i++)
               {
                   this.combo_labelField.Properties.Items.Add(dt.Columns[i].ColumnName);
               }
           }
           this.combo_labelField.SelectedItem = this.m_RadarChartModel.RadarLabel.LebelField;
           //引线设置
           this.check_leadline.Checked = this.m_RadarChartModel.RadarLabel.LeadLabel;
           this.check_leadline_CheckedChanged(check_leadline, EventArgs.Empty);
           this.num_leadlineLength.Value = (decimal)this.m_RadarChartModel.RadarLabel.LeadlineLength;
           this.num_leadlineHorizonLength.Value = (decimal)this.m_RadarChartModel.RadarLabel.LeadlineHorizonLength;
           //扇区设置
           List<string> sectorNames = new List<string>();
           for (int i = 0; i < this.m_RadarChartModel.Sectors.Count; i++)
           {
               var sector = this.m_RadarChartModel.Sectors[i];
               sectorNames.Add(sector.Name);
           }
           this.combo_sector.Properties.Items.AddRange(sectorNames);
           this.combo_sector_SelectedIndexChanged(this.combo_sector, EventArgs.Empty);
           //图例设置
           this.check_legend.Checked = this.m_RadarChartModel.Legend.Visible;
           this.check_legend_CheckedChanged(this.check_legend, EventArgs.Empty);
           this.combo_legendPositon.SelectedItem = Enum.GetName(typeof(LegendPostion), this.m_RadarChartModel.Legend.Postion);
           this.num_legendCol.Value = this.m_RadarChartModel.Legend.colNum;
           this.label_legendStyle.Text = this.m_RadarChartModel.Legend.Font.ToString();
           this.label_legendColor.BackColor = this.m_RadarChartModel.Legend.Font.Color;
            //其他设置
           if (this.m_RadarChartModel.DataTable != null)
           {
               this.btn_apply.Enabled = true;
               this.btn_chartin.Enabled = true;
               this.btn_export.Enabled = true;
               this.btn_apply_Click(this.btn_apply,EventArgs.Empty);

           }
           else {
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
            decimal val = (sender as DevExpress.XtraEditors.SpinEdit).Value;
            this.m_RadarChartModel.Width = (double)(sender as DevExpress.XtraEditors.SpinEdit).Value;
            this.num_radius.Value = 0;
        }

        private void num_chartHeight_EditValueChanged(object sender, EventArgs e)
        {
            this.m_RadarChartModel.Height = (double)(sender as DevExpress.XtraEditors.SpinEdit).Value;
            this.num_radius.Value = 0;
        }

        private void check_3d_CheckedChanged(object sender, EventArgs e)
        {
            this.m_RadarChartModel.View3D = this.check_3d.Checked;
        }

        private void check_ring_CheckedChanged(object sender, EventArgs e)
        {
            this.m_RadarChartModel.ViewRing = this.check_ring.Checked;
        }
        #endregion
        #region 标题设置
        private void check_title_CheckedChanged(object sender, EventArgs e)
        {
            this.m_RadarChartModel.Title.Visible = this.check_title.Checked;
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
            this.m_RadarChartModel.Title.Text = this.txt_title.Text;
        }

        private void btn_titleStyle_Click(object sender, EventArgs e)
        {
            FontColorForm fcf = new FontColorForm(this.m_RadarChartModel.Title.Font);
            fcf.ShowInTaskbar = false;
            fcf.ShowIcon = false;
            fcf.StartPosition = FormStartPosition.CenterScreen;
            if (fcf.ShowDialog() == DialogResult.OK)
            {
                this.m_RadarChartModel.Title.Font = fcf.sFont;
                this.label_titleStyle.Text = fcf.sFont.ToString();
                this.label_titleColor.BackColor = fcf.sFont.Color;
            }
        }
        #endregion
        #region 标注设置
        private void check_label_CheckedChanged(object sender, EventArgs e)
        {
            this.m_RadarChartModel.RadarLabel.Visible = this.check_label.Checked;
            if (this.check_label.Checked)
            {
                this.combo_labelField.Enabled = true;
                this.btn_labelStyle.Enabled = true;
                this.combo_labelStyle.Enabled = true;
                this.check_leadline.Enabled = true;
            }
            else
            {
                this.combo_labelField.Enabled = false;
                this.btn_labelStyle.Enabled = false;
                this.combo_labelStyle.Enabled = false;
                this.check_leadline.Enabled = false;
            }
        }

        private void combo_labelField_EditValueChanged(object sender, EventArgs e)
        {
            if (this.combo_labelField.SelectedItem == null)
                return;
            this.m_RadarChartModel.RadarLabel.LebelField = this.combo_labelField.SelectedItem.ToString();
        }

        private void combo_labelStyle_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.combo_labelStyle.SelectedItem == null)
                return;
            this.m_RadarChartModel.RadarLabel.LabelStyle = (LabelStyle)Enum.Parse(typeof(LabelStyle), this.combo_labelStyle.SelectedItem.ToString()); 
        }

        private void btn_labelStyle_Click(object sender, EventArgs e)
        {
            FontColorForm fcf = new FontColorForm(this.m_RadarChartModel.RadarLabel.Font);
            fcf.ShowInTaskbar = false;
            fcf.ShowIcon = false;
            fcf.StartPosition = FormStartPosition.CenterScreen;
            if (fcf.ShowDialog() == DialogResult.OK) {
                this.m_RadarChartModel.RadarLabel.Font = fcf.sFont;
                this.label_labelStyle.Text = fcf.sFont.ToString();
                this.label_labelColor.BackColor = fcf.sFont.Color;
            }
        }


        private void check_leadline_CheckedChanged(object sender, EventArgs e)
        {
            this.m_RadarChartModel.RadarLabel.LeadLabel = this.check_leadline.Checked;
            if (this.check_leadline.Checked)
            {
                this.num_leadlineLength.Enabled = true;
                this.num_leadlineHorizonLength.Enabled = true;
            }
            else
            {
                this.num_leadlineLength.Enabled = false;
                this.num_leadlineHorizonLength.Enabled = false;
            }
        }

        private void num_radius_EditValueChanged(object sender, EventArgs e)
        {
            if ((int)this.num_radius.Value > this.m_RadarChartModel.Radius)
            {
                MessageBox.Show("自定义半径不能大于自适应半径");
                this.num_radius.Value = (decimal)this.m_RadarChartModel.Radius;
                return;
            }
            this.m_RadarChartModel.CustomRadius = (double)this.num_radius.Value;
        }

        private void num_leadlineLength_EditValueChanged(object sender, EventArgs e)
        {
            this.m_RadarChartModel.RadarLabel.LeadlineLength = (double)this.num_leadlineLength.Value;
        }

        private void num_leadlineHorizonLength_EditValueChanged(object sender, EventArgs e)
        {
            this.m_RadarChartModel.RadarLabel.LeadlineHorizonLength = (double)this.num_leadlineHorizonLength.Value;
        }
   
        private void combo_sector_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.combo_sector.SelectedItem == null)
            {
                this.btn_sectorColor.Enabled = false;
                return;
            }
            else {
                this.btn_sectorColor.Enabled = true;
                this.label_sectorColor.BackColor = this.m_RadarChartModel.Sectors[this.combo_sector.SelectedIndex].Color;
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
                this.m_RadarChartModel.Sectors[this.combo_sector.SelectedIndex].Color = selectedColor;
                this.m_RadarChartModel.CustomSectorColor = true;
            }
        }
        #endregion

        #region 图例设置
        private void check_legend_CheckedChanged(object sender, EventArgs e)
        {
            this.m_RadarChartModel.Legend.Visible = this.check_legend.Checked;
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
            this.m_RadarChartModel.Legend.Postion = (LegendPostion)Enum.Parse(typeof(LegendPostion), this.combo_legendPositon.SelectedItem.ToString());
        }
        private void num_legendCol_EditValueChanged(object sender, EventArgs e)
        {
            this.m_RadarChartModel.Legend.colNum = (int)this.num_legendCol.Value;
        }

        private void btn_legendStyle_Click(object sender, EventArgs e)
        {
            FontColorForm fcf = new FontColorForm(this.m_RadarChartModel.Legend.Font);
            fcf.ShowInTaskbar = false;
            fcf.ShowIcon = false;
            fcf.StartPosition = FormStartPosition.CenterScreen;
            if (fcf.ShowDialog() == DialogResult.OK)
            {
                this.m_RadarChartModel.Legend.Font = fcf.sFont;
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
            this.Draw(this.m_RadarChartModel);
        }

        private void btn_loadData_Click(object sender, EventArgs e)
        {
            if (this.m_RadarChartModel.DataTable != null)
            {
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
            List<string> fieldNames = new List<string>();
            for (int i = 1; i < dt.Columns.Count; i++)
            {
                fieldNames.Add(dt.Columns[i].ColumnName);
            }
            this.combo_labelField.Properties.Items.Clear();
            this.combo_labelField.Properties.Items.AddRange(fieldNames);
            this.combo_labelField.SelectedIndex = 0;
            List<string> sectorNames = new List<string>();
            for (int i = 0; i < dt.Rows.Count; i++) {
                sectorNames.Add(dt.Rows[i][0].ToString());
            }
            this.combo_sector.Properties.Items.Clear();
            this.combo_sector.Properties.Items.AddRange(sectorNames);
            this.m_RadarChartModel.DataTable = dt;

            m_RadarChartModel.AreaTypes = fieldNames;

            this.btn_apply_Click(this.btn_apply, EventArgs.Empty);
            this.btn_apply.Enabled = true;
            this.btn_chartin.Enabled = true;
            this.btn_export.Enabled = true;
        }

        private void SetupRadarSeries(
    List<Radar> seriesList,
    List<string> fields,
    List<DataTable> dataTables,
    List<RadarLabel> pieLabels,
    List<Color> fillColors,
    List<Color> lineColors,
    List<int> lineWidths,
    MarksStyles marksStyles
)
        {
            int count = seriesList.Count;

            for (int i = 0; i < count; i++)
            {
                Radar series = seriesList[i];

                // 如果列表只有一个元素，则用这个唯一值
                string field = fields.Count == 1 ? fields[0] : fields[i];
                DataTable dataTable = dataTables.Count == 1 ? dataTables[0] : dataTables[i];
                RadarLabel pieLabel = pieLabels.Count == 1 ? pieLabels[0] : pieLabels[i];
                Color fillColor = fillColors.Count == 1 ? fillColors[0] : fillColors[i];
                Color lineColor = lineColors.Count == 1 ? lineColors[0] : lineColors[i];
                int lineWidth = lineWidths.Count == 1 ? lineWidths[0] : lineWidths[i];

                // 原有逻辑
                series.Title = field;

                series.Circled = true;
                series.ColorEach = false;
                series.RotationAngle = 0;

                series.Marks.Transparent = true;
                series.Marks.Visible = pieLabel.Visible;
                series.Marks.Font = this.WindowFontToChartFont(pieLabel.Font.Font, pieLabel.Font.Color);
                series.Marks.Style = marksStyles;

                series.Clear();
                foreach (DataRow dr in dataTable.Rows)
                {
                    if (!dataTable.Columns.Contains(field)) continue;

                    double value = Convert.ToDouble(dr[field]);
                    string label = dr[0].ToString();
                    series.Add(value, label, fillColor);
                }

                series.Pen.Color = lineColor;
                series.Pen.Width = lineWidth;
            }
        }

        /// <summary>
        /// 根据百分比透明度生成颜色（0 = 全透明, 100 = 不透明）
        /// 忽略 baseColor 自带的 Alpha 值
        /// </summary>
        private Color ColorFromTransparencyAndBaseColor(int transparencyPercent_0_100, Color baseColor)
        {
            // 限制范围
            int percent = Math.Max(0, Math.Min(100, transparencyPercent_0_100));

            // 0% -> 255 (完全透明), 100% -> 0 (完全不透明)
            int alpha = 255 - (percent * 255 / 100);

            return Color.FromArgb(
                alpha,
                baseColor.R,
                baseColor.G,
                baseColor.B
            );
        }

        // 初始化 CheckedComboBoxEdit 候选项
        private void PopulateSeriesCheckList()
        {
            fieldcheckedComboBoxEdit.Properties.Items.Clear();

            foreach (Radar s in chart.Series)
            {
                string displayName = s.Title;

                // 添加到候选框，并设置初始为勾选状态
                fieldcheckedComboBoxEdit.Properties.Items.Add(displayName, true);
            }

            // 注册勾选变化事件
            fieldcheckedComboBoxEdit.EditValueChanged -= fieldcheckedComboBoxEdit_EditValueChanged;
            fieldcheckedComboBoxEdit.EditValueChanged += fieldcheckedComboBoxEdit_EditValueChanged;
        }

        /// <summary>
        /// 绘制雷达统计图
        /// </summary>
        /// <param name="model">数据模型</param>
        private void Draw(RadarChartModel model)
        {
            if (this.m_Chart != null && !this.m_Chart.IsDisposed)
            {
                //this.m_Chart.Dispose();
            }
            chart.Dock = DockStyle.Fill;
            this.m_ChartForm.Controls.Add(chart);
            if (m_Chart == null)
            {
                this.m_Chart = chart;
            }

            chart.Tag = model;
            //chart.Series.Clear();//清除，重绘，避免重新创建
            chart.AutoRepaint = false;//关闭自动绘制

            chart.Legend.Visible = false;//关闭自带图例，采用手动计算

            //
            //通用尺寸
            this.m_ChartForm.Width = (int)Util.CentimeterToPixel(model.Width) + 16;//16为补充form容器的边框宽度，从而保证TChart尺寸的正确性
            this.m_ChartForm.Height = (int)Util.CentimeterToPixel(model.Height) + 39;//39为补充form容器的边框高度，从而保证TChart尺寸的正确性
            chart.Aspect.View3D = model.View3D;

            #region 标题

            //========标题设置
            chart.Header.Visible = model.Title.Visible;
            chart.Header.Text = model.Title.Text;
            chart.Header.Font.Name = model.Title.Font.Font.Name;
            chart.Header.Font.Size = Convert.ToInt32(model.Title.Font.Font.Size);
            if (model.Title.Font.Style == FontStyle.Bold)
            {
                chart.Header.Font.Bold = true;
            }
            else if (model.Title.Font.Style == FontStyle.Italic)
            {
                chart.Header.Font.Italic = true;
            }
            else if (model.Title.Font.Style == FontStyle.Regular)
            {
                chart.Header.Font.Bold = false;
                chart.Header.Font.Italic = false;
            }
            chart.Header.Font.Color = model.Title.Font.Color;
            chart.Header.Alignment = StringAlignment.Center;

            #endregion

            #region 标注

            //========标注设置
            MarksStyles marksStyles = MarksStyles.Value;

            //设置标注样式方式（依赖库本身然后获取）
            switch (model.RadarLabel.LabelStyle)
            {
                case LabelStyle.数值:
                    marksStyles = MarksStyles.Value;
                    break;
                case LabelStyle.数值百分比:
                    marksStyles = MarksStyles.Percent;
                    break;
                case LabelStyle.标注数值:
                    marksStyles = MarksStyles.LabelValue;
                    break;
                case LabelStyle.标注数值百分比:
                    marksStyles = MarksStyles.LabelPercent;
                    break;
            }

            #endregion

            #region 画图

            // 第一次调用时：绘制所有字段
            if (chart.Series.Count == 0)
            {
                // 创建存储多个参数的 List
                List<Radar> seriesList = new List<Radar>();
                List<string> fields = new List<string>();
                List<DataTable> dataTables = new List<DataTable>();
                List<RadarLabel> pieLabels = new List<RadarLabel>();
                List<Color> fillColors = new List<Color>();
                List<Color> lineColors = new List<Color>();
                List<int> lineWidths = new List<int>();

                // 批量收集
                for (int i = 0; i < model.AreaTypes.Count; i++)
                {
                    string field = model.AreaTypes[i].ToString();

                    Radar series = new Radar();
                    chart.Series.Add(series);

                    seriesList.Add(series);
                    fields.Add(field);
                    dataTables.Add(model.DataTable);
                    pieLabels.Add(model.RadarLabel);
                    fillColors.Add(Color.FromArgb(model.SeriesTransparencyList[i], model.SeriesColorList[i]));
                    lineColors.Add(model.LineColorList[i]);
                    lineWidths.Add(model.LineWidthList[i]);
                }

                // 一次性调用批量方法
                SetupRadarSeries(
                    seriesList,
                    fields,
                    dataTables,
                    pieLabels,
                    fillColors,
                    lineColors,
                    lineWidths,
                    marksStyles
                );
            }
            else if (!string.IsNullOrEmpty(model.RadarLabel.LebelField)) // 后续调用时：仅修改指定区域
            {
                foreach (Radar s in chart.Series)
                {
                    if (s.Title == model.RadarLabel.LebelField)
                    {
                        // 准备只有一个元素的 List
                        List<Radar> seriesList = new List<Radar> { s };
                        List<string> fields = new List<string> { model.RadarLabel.LebelField };
                        List<DataTable> dataTables = new List<DataTable> { model.DataTable };
                        List<RadarLabel> pieLabels = new List<RadarLabel> { model.RadarLabel };
                        List<Color> fillColors = new List<Color> 
                        { 
                            ColorFromTransparencyAndBaseColor(model.SeriesTransparencyList[model.Index], model.SeriesColorList[model.Index]) 
                        };
                        List<Color> lineColors = new List<Color> { model.LineColorList[model.Index] };
                        List<int> lineWidths = new List<int> { model.LineWidthList[model.Index] };

                        // 调用批量方法
                        SetupRadarSeries(
                            seriesList,
                            fields,
                            dataTables,
                            pieLabels,
                            fillColors,
                            lineColors,
                            lineWidths,
                            marksStyles
                        );

                        break;
                    }
                }
            }

            PopulateSeriesCheckList();

            chart.Refresh();
            chart.Invalidate();

            #endregion

            #region 图例

            foreach (Radar s in chart.Series)
            {
                //留白的尺寸
                int blankSize = 10;
                //定义个矩形范围，默认按照chart的大小来控制，默认左右上下都留白
                Rectangle chartRect = new Rectangle()
                {
                    X = blankSize,
                    Y = blankSize,
                    Width = chart.Width - blankSize,
                    Height = chart.Height - blankSize
                };
                if (model.CustomRadius != 0)
                {
                    s.CustomXRadius = (int)Util.CentimeterToPixel(model.CustomRadius);
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
                    for (int i = 0; i < chart.Series.Count; i++)
                    {
                        Radar sc = chart.Series[i] as Radar;
                        if (sc == null || sc.Active == false) continue;

                        LegendItem li = new LegendItem();
                        li.Text = sc.Title;
                        //计算图例项的宽高
                        SizeF sf = chart.Graphics3D.MeasureString(new ChartFont(chart.Chart, model.Legend.Font.Font), sc.Title);
                        li.Width = (int)Math.Ceiling(sf.Width) + model.Legend.SymbolTextGap + model.Legend.SymbolWidth;//文本宽度+符号宽度+文本符号间隔宽度
                        li.Height = (int)Math.Ceiling(sf.Height);//文本高度，颜色通过库自动赋值后，获取，先计算尺寸
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
                    if (true)
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
                                chartRect.Height = chart.Height - chartRect.Y - blankSize;
                            }
                            else if (model.Legend.Width > model.Legend.Height)//偏水平型，只做Y方向限制，X方向不受图例宽度影响
                            {
                                chartRect.X = blankSize;
                                chartRect.Y = model.Title.Top + model.Title.Height + blankSize;
                                chartRect.Width = chart.Width - chartRect.X - blankSize;
                                chartRect.Height = chart.Height - chartRect.Y - model.Legend.Height - blankSize;
                            }
                            else
                            {//都考虑
                                chartRect.X = blankSize;
                                chartRect.Y = model.Title.Top + model.Title.Height + blankSize;
                                chartRect.Width = chart.Width - chartRect.X - blankSize;
                                chartRect.Height = chart.Height - chartRect.Y - model.Legend.Height - blankSize;
                            }
                            break;
                        case LegendPostion.底部居中:
                            model.Legend.Top = chart.Height - model.Legend.Height - blankSize;
                            model.Legend.Left = (chart.Width - model.Legend.Width) / 2;
                            //计算图表位置
                            chartRect.X = blankSize;
                            chartRect.Y = (model.Title.Top + model.Title.Height) + blankSize;
                            chartRect.Width = chart.Width - blankSize;
                            chartRect.Height = chart.Height - chartRect.Y - model.Legend.Height - blankSize;
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
                                chartRect.Height = chart.Height - chartRect.Y - blankSize;
                            }
                            else if (model.Legend.Width > model.Legend.Height)//偏水平型，只做Y方向限制，X方向不受图例宽度影响
                            {
                                chartRect.X = blankSize;
                                chartRect.Y = model.Title.Top + model.Title.Height + blankSize;
                                chartRect.Width = chart.Width - chartRect.X - blankSize;
                                chartRect.Height = chart.Height - chartRect.Y - model.Legend.Height - blankSize;
                            }
                            else
                            {//都考虑
                                chartRect.X = blankSize;
                                chartRect.Y = model.Title.Top + model.Title.Height + blankSize;
                                chartRect.Width = chart.Width - chartRect.X - model.Legend.Width - blankSize;
                                chartRect.Height = chart.Height - chartRect.Y - model.Legend.Height - blankSize;
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

                    #region 废弃 Chart自带图例
                    //chart.Legend.Visible = true;
                    //chart.Legend.Inverted = true;
                    //chart.Legend.Shadow.Visible = false;
                    //chart.Legend.Transparent = true;
                    //chart.Legend.ResizeChart = false;
                    //chart.Legend.Symbol.Width = 15;
                    //chart.Legend.Symbol.WidthUnits = LegendSymbolSize.Pixels;
                    //chart.Legend.AutoSize = false;//设置成false，手动设置Legend的长宽，内部items进行自适应
                    //chart.Legend.ColumnWidthAuto = true;
                    //chart.Legend.MaxNumRows = 30;//取消行数限制
                    ////chart.Legend.Items.Custom = true;
                    //if (model.Legend.Sort == LegendSort.树直排列)
                    //    chart.Legend.Alignment = LegendAlignments.Top;//左右时，该库会进行树直排列
                    //if (model.Legend.Sort == LegendSort.水平排列)
                    //    chart.Legend.Alignment = LegendAlignments.Top;//上下时，该库会进行水平排列
                    //chart.Legend.CustomPosition = true;//关闭掉自定义位置
                    //chart.Legend.Items.Custom = true;
                    #endregion
                }


                chart.Chart.ChartRect = chartRect;
                chart.AfterDraw += new PaintChartEventHandler(chart_AfterDraw);

                chart.Refresh();
            }

            #endregion
        }

        private void chart_BeforeDraw(object sender,Graphics3D g) {
            TChart chart = sender as TChart;
            //绘制图表背景色
            g.Brush.Color = Color.White;//填充色
            g.Rectangle(0, 0, chart.Width, chart.Height);
        }

        /// <summary>
        /// 根据渲染好的饼图，计算引线和标注的位置
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="g"></param>
        private void chart_AfterDraw(object sender, Graphics3D g)
        {
            TChart chart = sender as TChart;
            Radar series = chart.Series[0] as Radar;
            RadarChartModel model = chart.Tag as RadarChartModel;
            //保存饼图的扇区颜色
            for (int i = 0; i < series.Count; i++)
            {
                model.Sectors.Add(new PieSector() { 
                     Color=series.ValueColor(i),
                     Name=series.Labels[i]
                });
            }
            //计算饼图的范围（此处范围已经经过标题和图例控制，以前部分判断逻辑将删除）, 后面考虑移动后的标注范围
            Rectangle seriesRect = new Rectangle()
            {
                X = series.CircleRect.X,
                Y = series.CircleRect.Y,
                Width = series.CircleRect.Width,
                Height = series.CircleRect.Height
            };

            //// =========================计算饼图范围与chartRect范围冲突问题，调整饼图的尺寸=======================
            int deltaX = 0;
            int deltaY = 0;
            //容器的范围，这个范围已经避开了图例和标题的问题
            Rectangle chartRect = chart.Chart.ChartRect;
            //判断并计算图表范围与图表容器范围的交集范围
            if (chartRect.IntersectsWith(seriesRect))
            {
                chartRect.Intersect(seriesRect);

                int _deltaX = seriesRect.Width - chartRect.Width;
                int _deltaY = seriesRect.Height -chartRect.Height;
                if (deltaX < _deltaX)
                    deltaX = _deltaX;
                if (deltaY < _deltaY)
                {
                    deltaY = _deltaY;
                }
            }
            int deltaR = deltaX > deltaY ? deltaX : deltaY;
            if (deltaR > 0)
            {
                //调整圆饼图的半径，规避碰撞
                int pieRadius = series.XRadius - deltaR;//记录一下
                series.CustomXRadius = pieRadius;//model.Radius;
                model.Radius = Math.Floor(Util.PixelToCentimeter(pieRadius) * 100) / 100;
                this.label_radius.Text = "自适应半径：" + model.Radius + "厘米";
                chart.Refresh();//触发重绘
            }
            else
            {
                model.Radius = Math.Floor(Util.PixelToCentimeter(series.XRadius) * 100) / 100;
                this.label_radius.Text = "自适应半径：" + model.Radius + "厘米";
                this.m_RadarChartModel.SeriesRect = seriesRect;
            }

            chart.AfterDraw -= new PaintChartEventHandler(chart_AfterDraw);//移除
            chart.AfterDraw -= new PaintChartEventHandler(chart_AfterDraw3);//先移除再添加，避免注册多次
            chart.AfterDraw += new PaintChartEventHandler(chart_AfterDraw3);
            chart.Refresh();//刷新的主要目的 是再次触发chart_AfterDraw3 事件
        }
        /// <summary>
        /// 所有元素的绘制
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="g"></param>
        private void chart_AfterDraw3(object sender, Graphics3D g)
        {
            TChart chart = sender as TChart;
            Radar series = chart.Series[0] as Radar;
            int Z = series.Marks.ZPosition;
            RadarChartModel model = chart.Tag as RadarChartModel;
            //===========================绘制图例================================

            if (model.Legend.Visible)
            {
                for (int i = 0; i < chart.Series.Count; i++)
                {
                    Radar s = chart.Series[i] as Radar;
                    if (s == null || s.Active == false) continue;
                    LegendItem li = model.Legend.LegendItems[i];
                    li.SymbolColor = s.ValueColor(0);
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
                    g.TextOut(textLeft, textTop, s.Title);//支持自定义字体,但可以导出SVG
                }
            }
           //chart.AfterDraw -= new PaintChartEventHandler(chart_AfterDraw3);
        }


        //// 旋转矩阵（绕 X 轴和 Z 轴）
        //public Point3D ApplyRotation(Point3D point, double rotationDeg, double elevationDeg)
        //{
        //    double rotationRad = rotationDeg * Math.PI / 180;
        //    double elevationRad = elevationDeg * Math.PI / 180;

        //    // 绕 Z 轴旋转（Rotation）
        //    double x1 = point.X * Math.Cos(rotationRad) - point.Y * Math.Sin(rotationRad);
        //    double y1 = point.X * Math.Sin(rotationRad) + point.Y * Math.Cos(rotationRad);
        //    double z1 = point.Z;

        //    // 绕 X 轴旋转（Elevation）
        //    double x2 = x1;
        //    double y2 = y1 * Math.Cos(elevationRad) - z1 * Math.Sin(elevationRad);
        //    double z2 = y1 * Math.Sin(elevationRad) + z1 * Math.Cos(elevationRad);

        //    return new Point3D((int)x2, (int)y2, (int)z2);
        //}


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
        /// <summary>
        /// 矢量单位化
        /// </summary>
        /// <param name="p"></param>
        /// <returns></returns>
        private Vector2D Normalization(Vector2D p)
        {
            double vectorMod = Math.Sqrt(Math.Pow(p.X, 2) + Math.Pow(p.Y,2));
            return new Vector2D(p.X / vectorMod, p.Y / vectorMod);
        }

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


        /// <summary>
        /// 绘制文本(弃用)
        /// </summary>
        /// <param name="g"></param>
        /// <param name="exportSVG"></param>
        private void DrawText(TChart chart,int top,int left,string text,Font font,Color color, bool exportSVG=false) {
            Graphics3D g= chart.Graphics3D;
            if (exportSVG)
            {
                g.Font=this.WindowFontToChartFont(font,color);
                g.TextOut(g.Font, left, top, text);//不支持自定义字体,但可以导出SVG
            }
            else
            {
                g.Brush.Color = color;//设置brush的颜色，会修改g.Brush.DrawingBrush中的颜色状态（g.Brush.DrawingBrush本身没有颜色属性）
                g.GDIplusCanvas.DrawString(text, font, g.Brush.DrawingBrush, new PointF(left, top));//不支持自定义字体，但无法导出矢量
            }
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

        private void button1_Click(object sender, EventArgs e)
        {
            Pie pie = m_Chart.Series[0] as Pie;
            //var a= this.m_Chart.Axes.Left.AxisRect();
            //a.Size = new Size(0,0);
            //var b = this.m_Chart.Axes.Left.Labels.Width;
            // MessageBox.Show(this.m_PieChartModel.Legend.Left.ToString() + "，" + (this.m_PieChartModel.Legend.Width).ToString())


            MessageBox.Show(pie.XRadius.ToString());
            //MessageBox.Show(pie.CircleXCenter + "," + pie.CircleYCenter);
            //this.Draw(this.m_PieChartModel, true);
            //this.m_Chart.Export.Image.SVG.Save(@"D:\实验室\宁波多尺度\a111.svg");
            //this.Draw(this.m_PieChartModel, false);
            
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
                string modelJson = JsonConvert.SerializeObject(this.m_RadarChartModel);
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
            string modelJson = JsonConvert.SerializeObject(this.m_RadarChartModel);
            // 创建内存流来保存SVG数据
            using (MemoryStream stream = new MemoryStream())
            {
                //将图表导出为SVG并保存到流
                this.m_Chart.Export.Image.SVG.Save(stream);
                // 将流位置重置为开始，以便读取
                stream.Position = 0;
                SvgDocument svgDoc = SvgDocument.Open(stream);
                double PixWidth = Util.CentimeterToPixel(this.m_RadarChartModel.Width);
                double PixHeight = Util.CentimeterToPixel(this.m_RadarChartModel.Height);
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
                        SvgText svgText=svgEle as SvgText;
                        float fontSize= svgText.FontSize.ToDeviceValue();//这里计算出来的pt值，需要乘以1.3转成像素值
                        float yOffset=fontSize*1.3f;//估算基线偏移（经验值）96/72约等于1.3
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
                    double picWidth = this.m_RadarChartModel.Width;
                    double picHeight = this.m_RadarChartModel.Height;
                    IGraphicsContainer graphicsContainer = (IGraphicsContainer)pageLayout;
                    IEnvelope envelope = new EnvelopeClass();
                    envelope.PutCoords(pageCenterPoint.X - picWidth / 2, pageCenterPoint.Y - picHeight / 2, pageCenterPoint.X + picWidth / 2, pageCenterPoint.Y + picHeight / 2);
                    (picElement as IElement).Geometry = envelope;
                    (picElement as IElementProperties).CustomProperty = modelJson;
                    string name = "Smgi_Radar_Chart" + System.IO.Path.GetFileNameWithoutExtension(fileName);
                    (picElement as IElementProperties).Name = name;
                    graphicsContainer.AddElement(picElement as IElement, 0);

                    this.m_Application.PageLayoutControl.ActiveView.Refresh();
                }
            }
            MessageBox.Show("已经导入到地图布局视图中！");
        }

        private void seriesColorSettingButton_Click(object sender, EventArgs e)
        {
            ColorDialog colorDialog = new ColorDialog();
            // 设置默认颜色（可选）
            colorDialog.Color = this.seriesColorLabelControl.BackColor;
            // 允许自定义颜色（可选）
            colorDialog.AllowFullOpen = true;
            colorDialog.FullOpen = true;
            // 显示对话框并检查用户是否点击“确定”
            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                Color selectedColor = colorDialog.Color;
                this.seriesColorLabelControl.BackColor = selectedColor;

                this.m_RadarChartModel.SeriesColorList[m_RadarChartModel.Index] = selectedColor;
            }
        }

        private void lineColorSettingButton_Click(object sender, EventArgs e)
        {
            ColorDialog colorDialog = new ColorDialog();
            // 设置默认颜色（可选）
            colorDialog.Color = this.lineColorLabelControl.BackColor;
            // 允许自定义颜色（可选）
            colorDialog.AllowFullOpen = true;
            colorDialog.FullOpen = true;
            // 显示对话框并检查用户是否点击“确定”
            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                Color selectedColor = colorDialog.Color;
                this.lineColorLabelControl.BackColor = selectedColor;

                this.m_RadarChartModel.LineColorList[m_RadarChartModel.Index] = selectedColor;
            }
        }

        private void lineWidthSpinEdit_EditValueChanged(object sender, EventArgs e)
        {
            this.m_RadarChartModel.LineWidthList[m_RadarChartModel.Index] = (int)(sender as DevExpress.XtraEditors.SpinEdit).Value;
        }

        private void seriesTransparencySpinEdit_EditValueChanged(object sender, EventArgs e)
        {
            this.m_RadarChartModel.SeriesTransparencyList[m_RadarChartModel.Index] = (int)(sender as DevExpress.XtraEditors.SpinEdit).Value;
        }


        private void combo_labelField_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (combo_labelField.SelectedItem == null) return;

            string selectedField = combo_labelField.SelectedItem.ToString();
            TChart chart = this.m_Chart;

            if (chart == null)
            {
                //初始化第一个series的界面窗体
                lineWidthSpinEdit.EditValue = m_RadarChartModel.LineWidthList[m_RadarChartModel.Index];
                seriesTransparencySpinEdit.EditValue = m_RadarChartModel.SeriesTransparencyList[m_RadarChartModel.Index];
                seriesColorLabelControl.BackColor = m_RadarChartModel.SeriesColorList[m_RadarChartModel.Index];
                lineColorLabelControl.BackColor = m_RadarChartModel.LineColorList[m_RadarChartModel.Index];
                return;
            }

            m_RadarChartModel.Index = m_RadarChartModel.AreaTypes.IndexOf(m_RadarChartModel.RadarLabel.LebelField);

            foreach (Radar s in chart.Series)
            {
                if (s.Title == selectedField)
                {
                    // 系列属性
                    this.m_RadarChartModel.RadarLabel.LebelField = selectedField;

                    // 线条颜色
                    this.lineColorLabelControl.BackColor = s.Pen.Color;
                    this.m_RadarChartModel.LineColorList[m_RadarChartModel.Index] = s.Pen.Color;

                    // 背景网格线
                    this.lineWidthSpinEdit.EditValue = s.Pen.Width;
                    this.m_RadarChartModel.LineWidthList[m_RadarChartModel.Index] = s.Pen.Width;

                    // 标注
                    this.check_label.Checked = s.Marks.Visible;
                    this.m_RadarChartModel.RadarLabel.Visible = s.Marks.Visible;
                    this.label_labelColor.BackColor = s.Marks.Font.Color;
                    this.m_RadarChartModel.RadarLabel.Font.Color = s.Marks.Font.Color;

                    // 颜色
                    if (s.Count > 0)
                    {
                        Color pointColor = s.ValueColor(0);
                        this.seriesColorLabelControl.BackColor = pointColor;
                        this.m_RadarChartModel.SeriesColorList[m_RadarChartModel.Index] = pointColor;

                        this.seriesTransparencySpinEdit.EditValue = (255 - pointColor.A) * 100 / 255; // 透明度
                    };

                    break; // 找到对应系列就退出
                }
            }
        }

        private void fieldcheckedComboBoxEdit_EditValueChanged(object sender, EventArgs e)
        {
            // 获取已勾选的值
            var checkedItems = fieldcheckedComboBoxEdit.Properties.Items
                .Cast<DevExpress.XtraEditors.Controls.CheckedListBoxItem>()
                .Where(item => item.CheckState == CheckState.Checked)
                .Select(item => item.Value.ToString())
                .ToList();

            foreach (Radar s in chart.Series)
            {
                string name = s.Title;

                // 根据勾选结果设置显示/隐藏
                s.Active = checkedItems.Contains(name);
            }

            chart.Refresh();
        }
    }
}
