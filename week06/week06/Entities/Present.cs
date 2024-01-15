using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using week06.Abstractions;

namespace week06.Entities
{
    public class Present : Toy
    {
        protected override void DrawImage(Graphics g)
        {
            g.FillRectangle(BoxColor, 0, 0, Width, Height);
            g.FillEllipse(RibbonColor, 0, 0, Width / 5, Height / 5);
        }
        public SolidBrush BoxColor { get; private set; }
        public SolidBrush RibbonColor { get; private set; }
        public Present(Color color)
        {
            BoxColor = new SolidBrush(color);
            RibbonColor = new SolidBrush(color);
        }
    }
}
