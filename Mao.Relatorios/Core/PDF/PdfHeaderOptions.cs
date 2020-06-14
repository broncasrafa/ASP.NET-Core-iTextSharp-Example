using System.Drawing;

namespace Mao.Relatorios.Core.PDF
{
    public class PdfHeaderOptions
    {
        public bool DrawHeaderLine { get; set; }
        public string HeaderTitleText { get; set; }
        public string HeaderSubtitleText { get; set; }
        public Image HeaderImageLeft { get; set; }
        public Image HeaderImageRight { get; set; }
        public bool ShowImageRight { get; set; }
    }
}