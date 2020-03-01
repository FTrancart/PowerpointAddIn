using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace PowerPointAddIn1
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            
            var active = Globals.ThisAddIn.Application.ActivePresentation;
            var nbSlides = active.Slides.Count;
            active.Slides.InsertFromFile(@"C:\Users\franc\Documents\ppt-template.pptx", nbSlides, 2, 2);
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            var active = Globals.ThisAddIn.Application.ActivePresentation;
            var nbSlides = active.Slides.Count;
            active.Slides.InsertFromFile(@"C:\Users\franc\Documents\ppt-template.pptx", nbSlides, 1, 1);
        }
    }
}
