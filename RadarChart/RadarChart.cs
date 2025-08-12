using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SMGI.Common;
using SMGI.Plugin.ThematicChart.TeeChart.PieChart;

namespace SMGI.Plugin.ThematicChart.TeeChart
{
    public class RadarChart:SMGICommand
    {
        public RadarChart()
        {
            m_caption = "雷达图";
        }
        public override bool Enabled
        {
            get
            {
                return true;
            }
        }

        RadarParameterPanel pp;

        public override void OnClick()
        {
            PageLayoutOnRightClickMenuForChart.Init(GApplication.Application);
            RadarChartModel PCM = new RadarChartModel(false);
            pp = new RadarParameterPanel(GApplication.Application, PCM);
            pp.Show();
        }
    }
}
