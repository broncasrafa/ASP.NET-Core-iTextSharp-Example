using System;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace Mao.Relatorios.Core.PDF
{
    internal class PdfPageEvent : PdfPageEventHelper
    {
        // This is the contentbyte object of the writer Objeto contentbyte do writer (PdfWriter)
        PdfContentByte cb;

        // coloca o numero final das pagina em um template
        PdfTemplate headerTemplate, footerTemplate;

        // BaseFont para usar no header / footer
        BaseFont bf = null;

        // This keeps track of the creation time
        DateTime PrintDateTime = DateTime.Now;

        // opções (configurações) do documento do pdf
        PdfDocumentOptions _documentOptions;

        public PdfPageEvent(PdfDocumentOptions documentOptions)
        {
            _documentOptions = documentOptions;
        }


        public override void OnOpenDocument(PdfWriter writer, Document document)
        {
            base.OnOpenDocument(writer, document);

            try
            {
                PrintDateTime = DateTime.Now;
                bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb = writer.DirectContent;
                headerTemplate = cb.CreateTemplate(100, 100);
                footerTemplate = cb.CreateTemplate(20, 50);
            }
            catch (DocumentException)
            {

            }
            catch (IOException)
            {

            }
        }
        public override void OnEndPage(PdfWriter writer, Document document)
        {
            base.OnEndPage(writer, document);

            String textoPaginacao = writer.PageNumber + " / ";

            #region Adicionando pagina no cabeçalho
            if (_documentOptions.ShowPagingOnTop)
            {
                {
                    cb.BeginText();
                    cb.SetFontAndSize(bf, 8f);
                    cb.SetTextMatrix(document.PageSize.GetRight(60), document.PageSize.GetTop(55));
                    cb.ShowText(textoPaginacao);
                    cb.EndText();
                    float len = bf.GetWidthPoint(textoPaginacao, 8f);
                    //Adds "12" in Page 1 of 12
                    cb.AddTemplate(headerTemplate, document.PageSize.GetRight(60) + len, document.PageSize.GetTop(55));
                }
            }
            #endregion

            #region Adicionando pagina do rodapé
            if (_documentOptions.ShowPagingOnBottom)
            {
                {
                    cb.BeginText();
                    cb.SetFontAndSize(bf, 8f);
                    cb.SetTextMatrix(document.PageSize.GetRight(100), document.PageSize.GetBottom(30));
                    cb.ShowText(textoPaginacao);
                    cb.EndText();
                    float len = bf.GetWidthPoint(textoPaginacao, 8f);


                    #region texto do rodapé
                    if(_documentOptions.FooterOptions.ShowFooterText)
                    {
                        Paragraph paragrafoRodape1 = new Paragraph(new Chunk("Fundação para o Desenvolvimento da Educação", FontFactory.GetFont("Arial", 7f, Font.BOLD, BaseColor.BLACK)));
                        Paragraph paragrafoRodape2 = new Paragraph(new Chunk("Avenida São Luís, 99 - República 01046-001 Tel. (11) 3158-4000", FontFactory.GetFont("Arial", 7f, Font.NORMAL, BaseColor.BLACK)));
                        Paragraph paragrafoRodape3 = new Paragraph(new Chunk("www.fde.sp.gov.br", FontFactory.GetFont("Arial", 6f, Font.NORMAL, BaseColor.BLACK)).SetAnchor("http://www.fde.sp.gov.br/"));

                        float dimInit = 35;
                        ColumnText columnTextRodape1 = new ColumnText(cb);
                        columnTextRodape1.AddText(paragrafoRodape1);
                        if(_documentOptions.PageOrientation == PdfPageOrientation.Retrato)
                            columnTextRodape1.SetSimpleColumn(590, dimInit, 0, 0, 0, Element.ALIGN_CENTER);
                        else
                            columnTextRodape1.SetSimpleColumn(890, dimInit, 0, 0, 0, Element.ALIGN_CENTER);

                        columnTextRodape1.Go();

                        ColumnText columnTextRodape2 = new ColumnText(cb);
                        columnTextRodape2.AddText(paragrafoRodape2);
                        if (_documentOptions.PageOrientation == PdfPageOrientation.Retrato)
                            columnTextRodape2.SetSimpleColumn(590, (dimInit - 10f), 5, 0, 0, Element.ALIGN_CENTER);
                        else
                            columnTextRodape2.SetSimpleColumn(890, (dimInit - 10f), 5, 0, 0, Element.ALIGN_CENTER);

                        columnTextRodape2.Go();

                        ColumnText columnTextRodape3 = new ColumnText(cb);
                        columnTextRodape3.AddText(paragrafoRodape3);
                        if (_documentOptions.PageOrientation == PdfPageOrientation.Retrato)
                            columnTextRodape3.SetSimpleColumn(590, (dimInit - 20f), 0, 0, 0, Element.ALIGN_CENTER);
                        else
                            columnTextRodape3.SetSimpleColumn(890, (dimInit - 20f), 0, 0, 0, Element.ALIGN_CENTER);

                        columnTextRodape3.Go();

                    }
                    #endregion


                    cb.AddTemplate(footerTemplate, document.PageSize.GetRight(100) + len, document.PageSize.GetBottom(30));
                }
            }
            #endregion

            #region Cabeçalho
            if(_documentOptions.ShowHeader)
            {
                float _HeaderHeight = 65f;

                PdfPTable tableHeader = new PdfPTable(1);
                tableHeader.WidthPercentage = 95f;
                tableHeader.DefaultCell.FixedHeight = 20f;
                tableHeader.TotalWidth = document.PageSize.Width - 80f;

                PdfPCell cellHeader = new PdfPCell();
                cellHeader.Border = PdfPCell.NO_BORDER;
                cellHeader.FixedHeight = _HeaderHeight;

                // Retrato
                if (_documentOptions.PageOrientation == PdfPageOrientation.Retrato)
                {
                    Paragraph paragrafoCabecalho1 = new Paragraph(new Chunk(_documentOptions.HeaderOptions?.HeaderTitleText, FontFactory.GetFont("Arial", 12f, Font.BOLD, BaseColor.BLACK)));
                    Paragraph paragrafoCabecalho2 = new Paragraph(new Chunk(_documentOptions.HeaderOptions?.HeaderSubtitleText, FontFactory.GetFont("Arial", 9f, Font.BOLD, BaseColor.BLACK)));

                    ColumnText columnTextCabecalho1 = new ColumnText(cb);
                    columnTextCabecalho1.AddText(paragrafoCabecalho1);
                    columnTextCabecalho1.SetSimpleColumn(34, 815, 540, 217, 10, Element.ALIGN_CENTER);
                    columnTextCabecalho1.Go();

                    ColumnText columnTextCabecalho2 = new ColumnText(cb);
                    columnTextCabecalho2.AddText(paragrafoCabecalho2);
                    columnTextCabecalho2.SetSimpleColumn(34, 800, 540, 217, 10, Element.ALIGN_CENTER);
                    columnTextCabecalho2.Go();

                    #region Imagem LOGO LEFT
                    if (_documentOptions.HeaderOptions != null && _documentOptions.HeaderOptions.HeaderImageLeft != null)
                    {
                        MemoryStream mst = new MemoryStream();
                        _documentOptions.HeaderOptions.HeaderImageLeft.Save(mst, System.Drawing.Imaging.ImageFormat.Jpeg);
                        byte[] b = mst.ToArray();

                        //posicionando a imagem no arquivo
                        iTextSharp.text.Image imageLogo = iTextSharp.text.Image.GetInstance(b);
                        imageLogo.SetAbsolutePosition(document.PageSize.Width - 580f, document.PageSize.Height - 46f);
                        imageLogo.ScalePercent(52f);
                        document.Add(imageLogo);
                    }
                    #endregion

                    #region Imagem LOGO RIGHT
                    if (_documentOptions.HeaderOptions != null && _documentOptions.HeaderOptions.ShowImageRight && _documentOptions.HeaderOptions.HeaderImageRight != null)
                    {
                        MemoryStream mst = new MemoryStream();
                        _documentOptions.HeaderOptions.HeaderImageRight.Save(mst, System.Drawing.Imaging.ImageFormat.Png);
                        byte[] b = mst.ToArray();

                        //posicionando a imagem no arquivo
                        iTextSharp.text.Image imageLogo = iTextSharp.text.Image.GetInstance(b);
                        imageLogo.SetAbsolutePosition(document.PageSize.Width - 125f, document.PageSize.Height - 46f);
                        imageLogo.ScalePercent(35f);
                        document.Add(imageLogo);
                    }
                    #endregion

                    #region Data de impressão
                    if (_documentOptions.ShowPrintDateTime)
                    {
                        Chunk chunkDataImpressaoText = new Chunk("Data de impressão: ", FontFactory.GetFont("Arial", 8f, Font.BOLD, BaseColor.BLACK));
                        Phrase phraseDataImpressaoText = new Phrase();
                        phraseDataImpressaoText.Add(chunkDataImpressaoText);

                        Chunk chunkDataImpressaoData = new Chunk(PrintDateTime.ToString(), FontFactory.GetFont("Arial", 8f, BaseColor.BLACK));
                        Phrase phraseDataImpressaoData = new Phrase();
                        phraseDataImpressaoData.Add(chunkDataImpressaoData);

                        Paragraph paragrafoTextDataImpressao = new Paragraph();
                        paragrafoTextDataImpressao.Add(chunkDataImpressaoText);
                        paragrafoTextDataImpressao.Add(chunkDataImpressaoData);

                        ColumnText columnTextDataImpressao = new ColumnText(cb);
                        columnTextDataImpressao.AddText(paragrafoTextDataImpressao);
                        columnTextDataImpressao.SetSimpleColumn(334, 775, 540, 217, 10, Element.ALIGN_RIGHT);
                        columnTextDataImpressao.Go();
                    }
                    #endregion

                    tableHeader.AddCell(cellHeader);                    
                }
                else // Paisagem
                {
                    Paragraph paragrafoCabecalho1 = new Paragraph(new Chunk(_documentOptions.HeaderOptions?.HeaderTitleText, FontFactory.GetFont("Arial", 12f, Font.BOLD, BaseColor.BLACK)));
                    Paragraph paragrafoCabecalho2 = new Paragraph(new Chunk(_documentOptions.HeaderOptions?.HeaderSubtitleText, FontFactory.GetFont("Arial", 9f, Font.BOLD, BaseColor.BLACK)));
                    
                    ColumnText columnTextCabecalho1 = new ColumnText(cb);
                    columnTextCabecalho1.AddText(paragrafoCabecalho1);
                    columnTextCabecalho1.SetSimpleColumn(790, 578, 10, 0, 10, Element.ALIGN_CENTER);
                    columnTextCabecalho1.Go();

                    ColumnText columnTextCabecalho2 = new ColumnText(cb);
                    columnTextCabecalho2.AddText(paragrafoCabecalho2);
                    columnTextCabecalho2.SetSimpleColumn(790, 558, 10, 0, 10, Element.ALIGN_CENTER);
                    columnTextCabecalho2.Go();

                    #region Imagem LOGO LEFT
                    if (_documentOptions.HeaderOptions != null && _documentOptions.HeaderOptions.HeaderImageLeft != null)
                    {
                        MemoryStream mst = new MemoryStream();
                        _documentOptions.HeaderOptions.HeaderImageLeft.Save(mst, System.Drawing.Imaging.ImageFormat.Jpeg);
                        byte[] b = mst.ToArray();

                        //posicionando a imagem no arquivo
                        iTextSharp.text.Image imageLogo = iTextSharp.text.Image.GetInstance(b);
                        imageLogo.SetAbsolutePosition(document.PageSize.Width - 822f, document.PageSize.Height - 46f);
                        imageLogo.ScalePercent(52f);
                        document.Add(imageLogo);
                    }
                    #endregion

                    #region Imagem LOGO RIGHT
                    if (_documentOptions.HeaderOptions != null && _documentOptions.HeaderOptions.ShowImageRight && _documentOptions.HeaderOptions.HeaderImageRight != null)
                    {
                        MemoryStream mst = new MemoryStream();
                        _documentOptions.HeaderOptions.HeaderImageRight.Save(mst, System.Drawing.Imaging.ImageFormat.Png);
                        byte[] b = mst.ToArray();

                        //posicionando a imagem no arquivo
                        iTextSharp.text.Image imageLogo = iTextSharp.text.Image.GetInstance(b);
                        imageLogo.SetAbsolutePosition(document.PageSize.Width - 125f, document.PageSize.Height - 46f);
                        imageLogo.ScalePercent(35f);
                        document.Add(imageLogo);
                    }
                    #endregion

                    #region Data de impressão
                    if (_documentOptions.ShowPrintDateTime)
                    {
                        Chunk chunkDataImpressaoText = new Chunk("Data de impressão: ", FontFactory.GetFont("Arial", 8f, Font.BOLD, BaseColor.BLACK));
                        Phrase phraseDataImpressaoText = new Phrase();
                        phraseDataImpressaoText.Add(chunkDataImpressaoText);

                        Chunk chunkDataImpressaoData = new Chunk(PrintDateTime.ToString(), FontFactory.GetFont("Arial", 8f, BaseColor.BLACK));
                        Phrase phraseDataImpressaoData = new Phrase();
                        phraseDataImpressaoData.Add(chunkDataImpressaoData);

                        Paragraph paragrafoTextDataImpressao = new Paragraph();
                        paragrafoTextDataImpressao.Add(chunkDataImpressaoText);
                        paragrafoTextDataImpressao.Add(chunkDataImpressaoData);

                        ColumnText columnTextDataImpressao = new ColumnText(cb);
                        columnTextDataImpressao.AddText(paragrafoTextDataImpressao);
                        columnTextDataImpressao.SetSimpleColumn(790, 530, 10, 0, 10, Element.ALIGN_RIGHT);
                        columnTextDataImpressao.Go();
                    }
                    #endregion

                    tableHeader.AddCell(cellHeader);
                }

                //chamar WriteSelectedRows da PdfTable. Este escreve linhas de PdfWriter em PdfTable. 
                //Primeiro parâmetro é começar a linha. -1 indica que não há linha final e todas as linhas a serem incluídos para escrever 
                //Parâmetro terceiro e quarto é x e y posição para começar a escrever
                tableHeader.WriteSelectedRows(0, -1, 40, document.PageSize.Height - 30, writer.DirectContent);
            }
            #endregion

            #region Linha em baixo (cabeçalho e rodapé)
            //Move o pointer e desenhar linha para separar seção de cabeçalho do resto da página
            if (_documentOptions.HeaderOptions.DrawHeaderLine)
            {
                cb.MoveTo(20, document.PageSize.GetTop(60));
                cb.LineTo(document.PageSize.Width - 20, document.PageSize.GetTop(60));
                cb.Stroke();
            }

            //Move o pointer e desenhar linha para separar seção de rodapé do resto da página
            if (_documentOptions.FooterOptions.DrawFooterLine)
            {
                cb.MoveTo(20, document.PageSize.GetBottom(50));
                cb.LineTo(document.PageSize.Width - 20, document.PageSize.GetBottom(50));
                cb.Stroke();
            }
            #endregion
        }
        public override void OnCloseDocument(PdfWriter writer, Document document)
        {
            base.OnCloseDocument(writer, document);

            headerTemplate.BeginText();
            headerTemplate.SetFontAndSize(bf, 8f);
            headerTemplate.SetTextMatrix(0, 0);
            headerTemplate.ShowText((writer.CurrentPageNumber - 1).ToString());
            headerTemplate.EndText();

            footerTemplate.BeginText();
            footerTemplate.SetFontAndSize(bf, 8f);
            footerTemplate.SetTextMatrix(0, 0);
            footerTemplate.ShowText((writer.CurrentPageNumber - 1).ToString());
            footerTemplate.EndText();
        }

    }
}
