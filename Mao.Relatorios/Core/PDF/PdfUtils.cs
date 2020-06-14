using System;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace Mao.Relatorios.Core.PDF
{
    public static class PdfUtils
    {
        public static Rectangle GetPageOrientation(PdfPageOrientation pageOrientation)
        {
            if (pageOrientation == PdfPageOrientation.Retrato)
                return PageSize.A4;

            return PageSize.A4.Rotate();
        }

        public static void GetDocumentInfo(Document document, PdfWriter pdfWriter, PdfDocumentInfo documentInfo)
        {
            bool hasInfo = false;
            if (!String.IsNullOrEmpty(documentInfo.AuthorName))
            {
                document.AddAuthor(documentInfo.AuthorName);
                hasInfo = true;
            }
                
            if (!String.IsNullOrEmpty(documentInfo.Keywords))
            {
                document.AddKeywords(documentInfo.Keywords);
                hasInfo = true;
            }
                
            if (!String.IsNullOrEmpty(documentInfo.Creator))
            {
                document.AddCreator(documentInfo.Creator);
                hasInfo = true;
            }
                
            if (!String.IsNullOrEmpty(documentInfo.Subject))
            {
                document.AddSubject(documentInfo.Subject);
                hasInfo = true;
            }
                
            if (!String.IsNullOrEmpty(documentInfo.Title))
            {
                document.AddTitle(documentInfo.Title);
                hasInfo = true;
            }
                
            if (documentInfo.CreatedDate.HasValue)
            {
                document.AddCreationDate();
                hasInfo = true;
            }

            if (hasInfo)
                pdfWriter.CreateXmpMetadata();
        }
        
    }
}
