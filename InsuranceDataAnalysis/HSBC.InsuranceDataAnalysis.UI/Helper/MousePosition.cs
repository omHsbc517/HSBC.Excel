using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace HSBC.InsuranceDataAnalysis.UI.Helper
{
    public class MousePosition
    {
        [DllImport("User32")]
        public static extern bool GetCursorPos(out POINT pt);
        public struct POINT
        {
            public int X;
            public int Y;
            public POINT(int x, int y)
            {
                this.X = x;
                this.Y = y;
            }
        }
        //鼠标移动并计算坐标
        public void MouseMove(out POINT MousePoint)
        {
            GetCursorPos(out MousePoint);
        }
    }
}
