using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace TeamNavigator
{
    class RoundButton : Button
    {
        public RoundButton() : base()
        {
            FlatAppearance.BorderSize = 0;
            FlatStyle = FlatStyle.Flat;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            GraphicsPath grPath = new GraphicsPath();
            grPath.AddEllipse(0, 0, ClientSize.Width, ClientSize.Height);
            Region = new Region(grPath);
            base.OnPaint(e);
        }

    }
}
