using System.Drawing;

namespace Mao.Relatorios.Core.PDF
{
    public class PdfHeaderOptions
    {
        public bool DrawHeaderLine { get; set; }
        public string HeaderTitleText { get; set; }
        public string HeaderSubtitleText { get; set; }
        public Image HeaderImage { get; set; }
    }
}