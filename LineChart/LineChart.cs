using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SMGI.Common;
using SMGI.Plugin.ThematicChart.TeeChart.PieChart;

namespace SMGI.Plugin.ThematicChart.TeeChart
{
    public class LineChart:SMGICommand
    {
        public LineChart()
        {
            m_caption = "折线图";
        }
        public override bool Enabled
        {
            get
            {
                return true;
            }
        }

        LineParameterPanel pp;

        public override void OnClick()
        {
            PageLayoutOnRightClickMenuForChart.Init(GApplication.Application);
            LineChartModel PCM = new LineChartModel(false);
            pp = new LineParameterPanel(GApplication.Application, PCM);
            pp.Show();
        }
    }
}
