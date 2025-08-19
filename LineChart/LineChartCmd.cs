using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SMGI.Common;

namespace SMGI.Plugin.ThematicChart.TeeChart.LineChart
{
    public class LineChartCmd : SMGICommand
    {
        public LineChartCmd()
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

        public override void OnClick()
        {
            PageLayoutOnRightClickMenuForChart.Init(GApplication.Application);
            LineChartModel PCM = new LineChartModel();
            LineParameterPanel pp = new LineParameterPanel(GApplication.Application, PCM);
            pp.Show();
        }
    }
}
