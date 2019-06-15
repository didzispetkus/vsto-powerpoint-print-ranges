using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PowerPointPrintSlides
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnPrintRange_Click(object sender, RibbonControlEventArgs e)
        {
            var layout = Globals.ThisAddIn.Application.ActivePresentation.Slides[1].CustomLayout;

            for (int i = 1; i < 5; ++i)
            {
                Globals.ThisAddIn.Application.ActivePresentation.Slides.AddSlide(i, layout);
            }


            var printOptions = Globals.ThisAddIn.Application.ActivePresentation.PrintOptions;
            printOptions.RangeType = PpPrintRangeType.ppPrintSlideRange;

            if (Globals.ThisAddIn.Application.ActivePresentation.PrintOptions.Ranges.Count == 0)
            {
                printOptions.Ranges.Add(1, 1);
                printOptions.Ranges.Add(3, 5);
            }

            // Opens file print dialog
            Globals.ThisAddIn.Application.CommandBars.ExecuteMso("FilePrint");
        }
    }
}
