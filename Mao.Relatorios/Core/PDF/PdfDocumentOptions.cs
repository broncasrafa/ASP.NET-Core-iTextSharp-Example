namespace Mao.Relatorios.Core.PDF
{
    public class PdfDocumentOptions
    {
        public PdfHeaderOptions HeaderOptions { get; set; }
        public PdfFooterOptions FooterOptions { get; set; }
        public PdfDocumentInfo DocumentInfo { get; set; }
        public PdfPageOrientation PageOrientation { get; set; }
        public bool ShowHeader { get; set; }
        public bool ShowFooter { get; set; }
        public bool ShowPagingOnTop { get; set; }
        public bool ShowPagingOnBottom { get; set; }
        public bool ShowPrintDateTime { get; set; }
        
        public PdfDocumentOptions()
        {
            HeaderOptions = new PdfHeaderOptions();
            FooterOptions = new PdfFooterOptions();
            DocumentInfo = new PdfDocumentInfo();
        }
    }
}
