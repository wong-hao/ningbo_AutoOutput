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

namespace SMGI.Plugin.ThematicChart.TeeChart.PieChart
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

        TChart chart = new TChart();

        public LineParameterPanel(GApplication app,LineChartModel pcm=null)
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
            //this.m_Chart = this.m_ChartForm.TeeChart;
            this.Disposed+=new EventHandler(ParameterPanel_Disposed);
        }

        /// <summary>
        /// 初始化UI参数
        /// </summary>
        private void initUIParam(){
           //通用设置
            this.num_chartWidth.Value = (decimal)this.m_LineChartModel.Width;
            this.num_chartHeight.Value = (decimal)this.m_LineChartModel.Height;
           this.check_3d.Checked = this.m_LineChartModel.View3D;
           this.check_ring.Checked = this.m_LineChartModel.ViewRing;
           this.num_radius.Value = (decimal)this.m_LineChartModel.CustomRadius;
           //标题设置
           this.check_title.Checked = this.m_LineChartModel.Title.Visible;
           this.check_title_CheckedChanged(this.check_title, EventArgs.Empty);
           this.txt_title.Text = this.m_LineChartModel.Title.Text;
           this.label_titleStyle.Text = this.m_LineChartModel.Title.Font.ToString();
           this.label_titleColor.BackColor = this.m_LineChartModel.Title.Font.Color;

           //标注设置
           this.check_label.Checked = this.m_LineChartModel.LineLabel.Visible;
           this.check_label_CheckedChanged(this.check_label, EventArgs.Empty);
           this.label_labelStyle.Text = this.m_LineChartModel.LineLabel.Font.ToString();
           this.label_labelStyle.ForeColor = this.m_LineChartModel.LineLabel.Font.Color;
           this.combo_labelStyle.SelectedItem = Enum.GetName(typeof(LabelStyle), this.m_LineChartModel.LineLabel.LabelStyle);
           this.label_labelStyle.Text = this.m_LineChartModel.LineLabel.Font.ToString();
           this.label_labelColor.BackColor = this.m_LineChartModel.LineLabel.Font.Color;
           if (this.m_LineChartModel.DataTable != null)
           {
               DataTable dt = this.m_LineChartModel.DataTable;
               for (int i = 1; i < dt.Columns.Count; i++)
               {
                   this.combo_labelField.Properties.Items.Add(dt.Columns[i].ColumnName);
               }
           }
           this.combo_labelField.SelectedItem = this.m_LineChartModel.LineLabel.LebelField;
           //引线设置
           this.check_leadline.Checked = this.m_LineChartModel.LineLabel.LeadLabel;
           this.check_leadline_CheckedChanged(check_leadline, EventArgs.Empty);
           this.num_leadlineLength.Value = (decimal)this.m_LineChartModel.LineLabel.LeadlineLength;
           this.num_leadlineHorizonLength.Value = (decimal)this.m_LineChartModel.LineLabel.LeadlineHorizonLength;
           //扇区设置
           List<string> sectorNames = new List<string>();
           for (int i = 0; i < this.m_LineChartModel.Sectors.Count; i++)
           {
               var sector = this.m_LineChartModel.Sectors[i];
               sectorNames.Add(sector.Name);
           }
           //this.combo_sector.Properties.Items.AddRange(sectorNames);
           //this.combo_sector_SelectedIndexChanged(this.combo_sector, EventArgs.Empty);
           //图例设置
           this.check_legend.Checked = this.m_LineChartModel.Legend.Visible;
           this.check_legend_CheckedChanged(this.check_legend, EventArgs.Empty);
           this.combo_legendPositon.SelectedItem = Enum.GetName(typeof(LegendPostion), this.m_LineChartModel.Legend.Postion);
           this.num_legendCol.Value = this.m_LineChartModel.Legend.colNum;
           this.label_legendStyle.Text = this.m_LineChartModel.Legend.Font.ToString();
           this.label_legendColor.BackColor = this.m_LineChartModel.Legend.Font.Color;
            //其他设置
           if (this.m_LineChartModel.DataTable != null)
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
            this.m_LineChartModel.Width = (double)(sender as DevExpress.XtraEditors.SpinEdit).Value;
            this.num_radius.Value = 0;
        }

        private void num_chartHeight_EditValueChanged(object sender, EventArgs e)
        {
            this.m_LineChartModel.Height = (double)(sender as DevExpress.XtraEditors.SpinEdit).Value;
            this.num_radius.Value = 0;
        }

        private void check_3d_CheckedChanged(object sender, EventArgs e)
        {
            this.m_LineChartModel.View3D = this.check_3d.Checked;
        }

        private void check_ring_CheckedChanged(object sender, EventArgs e)
        {
            this.m_LineChartModel.ViewRing = this.check_ring.Checked;
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
        #region 标注设置
        private void check_label_CheckedChanged(object sender, EventArgs e)
        {
            this.m_LineChartModel.LineLabel.Visible = this.check_label.Checked;
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
            this.m_LineChartModel.LineLabel.LebelField = this.combo_labelField.SelectedItem.ToString();
        }

        private void combo_labelStyle_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.combo_labelStyle.SelectedItem == null)
                return;
            this.m_LineChartModel.LineLabel.LabelStyle = (LabelStyle)Enum.Parse(typeof(LabelStyle), this.combo_labelStyle.SelectedItem.ToString()); 
        }

        private void btn_labelStyle_Click(object sender, EventArgs e)
        {
            FontColorForm fcf = new FontColorForm(this.m_LineChartModel.LineLabel.Font);
            fcf.ShowInTaskbar = false;
            fcf.ShowIcon = false;
            fcf.StartPosition = FormStartPosition.CenterScreen;
            if (fcf.ShowDialog() == DialogResult.OK) {
                this.m_LineChartModel.LineLabel.Font = fcf.sFont;
                this.label_labelStyle.Text = fcf.sFont.ToString();
                this.label_labelColor.BackColor = fcf.sFont.Color;
            }
        }


        private void check_leadline_CheckedChanged(object sender, EventArgs e)
        {
            this.m_LineChartModel.LineLabel.LeadLabel = this.check_leadline.Checked;
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
            if ((int)this.num_radius.Value > this.m_LineChartModel.Radius)
            {
                MessageBox.Show("自定义半径不能大于自适应半径");
                this.num_radius.Value = (decimal)this.m_LineChartModel.Radius;
                return;
            }
            this.m_LineChartModel.CustomRadius = (double)this.num_radius.Value;
        }

        private void num_leadlineLength_EditValueChanged(object sender, EventArgs e)
        {
            this.m_LineChartModel.LineLabel.LeadlineLength = (double)this.num_leadlineLength.Value;
        }

        private void num_leadlineHorizonLength_EditValueChanged(object sender, EventArgs e)
        {
            this.m_LineChartModel.LineLabel.LeadlineHorizonLength = (double)this.num_leadlineHorizonLength.Value;
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
            this.Draw(this.m_LineChartModel);
        }

        private void btn_loadData_Click(object sender, EventArgs e)
        {
            if (this.m_LineChartModel.DataTable != null)
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
            //this.combo_sector.Properties.Items.Clear();
            //this.combo_sector.Properties.Items.AddRange(sectorNames);
            this.m_LineChartModel.DataTable = dt;

            m_LineChartModel.AreaTypes = fieldNames;

            this.btn_apply_Click(this.btn_apply, EventArgs.Empty);
            this.btn_apply.Enabled = true;
            this.btn_chartin.Enabled = true;
            this.btn_export.Enabled = true;
        }

        private void SetupLineSeries(
    List<Line> seriesList,
    List<string> fields,
    List<DataTable> dataTables,
    List<LineLabel> pieLabels,
    List<Color> fillColors,
    MarksStyles marksStyles,
    List<int> lineWidths,
    List<int> pointerWidths,
    bool smoothed,
    bool pointer,
    PointerStyles pointerStyles
)
        {
            int count = seriesList.Count;

            for (int i = 0; i < count; i++)
            {
                Line series = seriesList[i];

                // 如果列表只有一个元素，则用这个唯一值
                string field = fields.Count == 1 ? fields[0] : fields[i];
                DataTable dataTable = dataTables.Count == 1 ? dataTables[0] : dataTables[i];
                LineLabel pieLabel = pieLabels.Count == 1 ? pieLabels[0] : pieLabels[i];
                Color fillColor = fillColors.Count == 1 ? fillColors[0] : fillColors[i];
                int lineWidth = lineWidths.Count == 1 ? lineWidths[0] : lineWidths[i];
                int pointerWidth = pointerWidths.Count == 1 ? pointerWidths[0] : pointerWidths[i];

                // 原有逻辑
                series.Title = field;

                series.ColorEach = false;

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

                series.Smoothed = smoothed;
                series.LinePen.Width = lineWidth;
                series.Pointer.Visible = pointer;
                series.Pointer.VertSize = pointerWidth;
                series.Pointer.HorizSize = pointerWidth;
                series.Pointer.Style = pointerStyles;
            }
        }

        // 初始化 CheckedComboBoxEdit 候选项
        private void PopulateSeriesCheckList()
        {
            fieldcheckedComboBoxEdit.Properties.Items.Clear();

            foreach (Line s in chart.Series)
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
        private void Draw(LineChartModel model)
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

            #region 底部标题

            chart.Axes.Bottom.Title.Visible = model.BottomTitle.Visible;
            chart.Axes.Bottom.Title.Text = model.BottomTitle.Text;
            chart.Header.Font.Name = model.BottomTitle.Font.Font.Name;
            chart.Axes.Bottom.Title.Font.Size = Convert.ToInt32(model.BottomTitle.Font.Font.Size);
            if (model.BottomTitle.Font.Style == FontStyle.Bold)
            {
                chart.Axes.Bottom.Title.Font.Bold = true;
            }
            else if (model.BottomTitle.Font.Style == FontStyle.Italic)
            {
                chart.Axes.Bottom.Title.Font.Italic = true;
            }
            else if (model.BottomTitle.Font.Style == FontStyle.Regular)
            {
                chart.Axes.Bottom.Title.Font.Bold = false;
                chart.Axes.Bottom.Title.Font.Italic = false;
            }
            chart.Axes.Bottom.Title.Font.Color = model.BottomTitle.Font.Color;

            #endregion

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
            switch (model.LineLabel.LabelStyle)
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
                List<Line> seriesList = new List<Line>();
                List<string> fields = new List<string>();
                List<DataTable> dataTables = new List<DataTable>();
                List<LineLabel> pieLabels = new List<LineLabel>();
                List<Color> fillColors = new List<Color>();
                List<int> lineWidths = new List<int>();
                List<int> pointerWidths = new List<int>();

                // 批量收集
                for (int i = 0; i < model.AreaTypes.Count; i++)
                {
                    string field = model.AreaTypes[i].ToString();

                    Line series = new Line();
                    chart.Series.Add(series);

                    seriesList.Add(series);
                    fields.Add(field);
                    dataTables.Add(model.DataTable);
                    pieLabels.Add(model.LineLabel);
                    fillColors.Add(Color.FromArgb(100, model.SeriesColorList[i]));
                    lineWidths.Add(model.LineWidthList[i]);
                    pointerWidths.Add(model.PointerWidthList[i]);
                }

                // 一次性调用批量方法
                SetupLineSeries(
                    seriesList,
                    fields,
                    dataTables,
                    pieLabels,
                    fillColors,
                    marksStyles,
                    lineWidths,
                    pointerWidths,
                    model.Smoothed,
                    model.Pointer,
                    model.PointerStyles
                );
            }
            else if (!string.IsNullOrEmpty(model.LineLabel.LebelField)) // 后续调用时：仅修改指定区域
            {
                foreach (Line s in chart.Series)
                {
                    if (s.Title == model.LineLabel.LebelField)
                    {
                        // 准备只有一个元素的 List
                        List<Line> seriesList = new List<Line> { s };
                        List<string> fields = new List<string> { model.LineLabel.LebelField };
                        List<DataTable> dataTables = new List<DataTable> { model.DataTable };
                        List<LineLabel> pieLabels = new List<LineLabel> { model.LineLabel };
                        List<Color> fillColors = new List<Color> 
                        { 
                            model.SeriesColorList[model.Index]
                        };
                        List<int> lineWidths = new List<int> { model.LineWidthList[model.Index] };
                        List<int> pointerWidths = new List<int> { model.PointerWidthList[model.Index] };

                        // 调用批量方法
                        SetupLineSeries(
                            seriesList,
                            fields,
                            dataTables,
                            pieLabels,
                            fillColors,
                            marksStyles,
                            lineWidths,
                            pointerWidths,
                            model.Smoothed,
                            model.Pointer,
                            model.PointerStyles
                        );

                        break;
                    }
                }
            }

            PopulateSeriesCheckList();

            chart.Refresh();
            chart.Invalidate();

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
                this.m_LineChartModel.SeriesRect = seriesRect;
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

                this.m_LineChartModel.SeriesColorList[m_LineChartModel.Index] = selectedColor;
            }
        }

        private void combo_labelField_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (combo_labelField.SelectedItem == null) return;

            string selectedField = combo_labelField.SelectedItem.ToString();
            TChart chart = this.m_Chart;

            if (chart == null)
            {
                //初始化第一个series的界面窗体
                lineWidthSpinEdit.EditValue = m_LineChartModel.LineWidthList[m_LineChartModel.Index];
                seriesColorLabelControl.BackColor = m_LineChartModel.SeriesColorList[m_LineChartModel.Index];
                return;
            }

            m_LineChartModel.Index = m_LineChartModel.AreaTypes.IndexOf(m_LineChartModel.LineLabel.LebelField);

            foreach (Line s in chart.Series)
            {
                if (s.Title == selectedField)
                {
                    // 系列属性
                    this.m_LineChartModel.LineLabel.LebelField = selectedField;

                    // 背景网格线
                    this.lineWidthSpinEdit.EditValue = s.LinePen.Width;
                    this.m_LineChartModel.LineWidthList[m_LineChartModel.Index] = s.LinePen.Width;

                    // 标注
                    this.check_label.Checked = s.Marks.Visible;
                    this.m_LineChartModel.LineLabel.Visible = s.Marks.Visible;
                    this.label_labelColor.BackColor = s.Marks.Font.Color;
                    this.m_LineChartModel.LineLabel.Font.Color = s.Marks.Font.Color;

                    // 颜色
                    if (s.Count > 0)
                    {
                        Color pointColor = s.ValueColor(0);
                        this.seriesColorLabelControl.BackColor = pointColor;
                        this.m_LineChartModel.SeriesColorList[m_LineChartModel.Index] = pointColor;
                    };

                    // 点标注
                    this.pointerWidthSpinEdit.EditValue = s.Pointer.VertSize;
                    this.m_LineChartModel.PointerWidthList[m_LineChartModel.Index] = s.Pointer.VertSize;
                    this.pointerCheckEdit.Checked = s.Pointer.Visible;
                    this.m_LineChartModel.Pointer = s.Pointer.Visible;

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

            foreach (Line s in chart.Series)
            {
                string name = s.Title;

                // 根据勾选结果设置显示/隐藏
                s.Active = checkedItems.Contains(name);
            }

            chart.Refresh();
        }

        private void smoothedCheckEdit_CheckedChanged(object sender, EventArgs e)
        {
            m_LineChartModel.Smoothed = smoothedCheckEdit.Checked;
        }

        private void lineWidthSpinEdit_EditValueChanged(object sender, EventArgs e)
        {
            this.m_LineChartModel.LineWidthList[m_LineChartModel.Index] = (int)(sender as DevExpress.XtraEditors.SpinEdit).Value;
        }

        private void pointerCheckEdit_CheckedChanged(object sender, EventArgs e)
        {
            m_LineChartModel.Pointer = pointerCheckEdit.Checked;

            if (pointerCheckEdit.Checked)
            {
                pointerComboBoxEdit.Enabled = true;
                pointerWidthSpinEdit.Enabled = true;
            }
            else
            {
                pointerComboBoxEdit.Enabled = false;
                pointerWidthSpinEdit.Enabled = false;
            }
        }

        private void pointerWidthSpinEdit_EditValueChanged(object sender, EventArgs e)
        {
            this.m_LineChartModel.PointerWidthList[m_LineChartModel.Index] = (int)(sender as DevExpress.XtraEditors.SpinEdit).Value;
        }

        private void pointerComboBoxEdit_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (pointerComboBoxEdit.SelectedItem.ToString())
            {
                case "圆形":
                    m_LineChartModel.PointerStyles = PointerStyles.Circle;
                    break;
                case "四角":
                    m_LineChartModel.PointerStyles = PointerStyles.Rectangle;
                    break;
                case "三角":
                    m_LineChartModel.PointerStyles = PointerStyles.Triangle;
                    break;
                case "五角形":
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

        private void bottomTitleSimpleButton_Click(object sender, EventArgs e)
        {
            FontColorForm fcf = new FontColorForm(this.m_LineChartModel.BottomTitle.Font);
            fcf.ShowInTaskbar = false;
            fcf.ShowIcon = false;
            fcf.StartPosition = FormStartPosition.CenterScreen;
            if (fcf.ShowDialog() == DialogResult.OK)
            {
                this.m_LineChartModel.BottomTitle.Font = fcf.sFont;
                this.bottomTitleLabelControlText.Text = fcf.sFont.ToString();
                this.bottomTitleLabelControl.BackColor = fcf.sFont.Color;
            }
        }

        private void bottomTittleCheckEdit_CheckedChanged(object sender, EventArgs e)
        {
            this.m_LineChartModel.BottomTitle.Visible = this.bottomTittleCheckEdit.Checked;
            if (this.check_title.Checked)
            {
                this.bottomTitleTextEdit.Enabled = true;
                this.bottomTitleSimpleButton.Enabled = true;
            }
            else
            {
                this.bottomTitleTextEdit.Enabled = false;
                this.bottomTitleSimpleButton.Enabled = false;
            }
        }

        private void bottomTitleTextEdit_EditValueChanged(object sender, EventArgs e)
        {
            this.m_LineChartModel.BottomTitle.Text = this.bottomTitleTextEdit.Text;
        }
    }
}
