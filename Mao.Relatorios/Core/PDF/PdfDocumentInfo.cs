using System;

namespace Mao.Relatorios.Core.PDF
{
    public class PdfDocumentInfo
    {
        public string AuthorName { get; set; }
        public DateTime? CreatedDate { get; set; }
        public string Keywords { get; set; }
        public string Creator { get; set; }
        public string Subject { get; set; }
        public string Title { get; set; }
    }
}
