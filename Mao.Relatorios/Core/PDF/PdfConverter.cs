using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Mao.Relatorios.Classes;

namespace Mao.Relatorios.Core.PDF
{
    internal class PdfConverter
    {
        public PdfConverter()
        {
            DocumentOptions = new PdfDocumentOptions();
        }

        public PdfDocumentOptions DocumentOptions { get; set; }

        private Document _Document;
        private float MarginLeft = 27f;
        private float MarginRight = 27f;
        private float MarginTop = 115f;
        private float MarginBottom = 75f;

        /// <summary>
        /// Converts the specified HTML file to PDF and returns the rendered PDF document as an array of bytes. Only inline CSS is supported.
        /// </summary>
        /// <param name="html">The HTML string to be converted to PDF.</param>
        /// <returns>An array of bytes containing the binary representation of the PDF document.</returns>
        public byte[] GeneratePdfBytesFromHtmlString(string html)
        {
            try
            {
                return SaveFromHtml(html, null);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Converts the specified HTML file to PDF and returns the rendered PDF document as an array of bytes.
        /// </summary>
        /// <param name="html">The HTML string to be converted to PDF.</param>
        /// <param name="cssStyles">The css styles string used in HTML to be converted to PDF.</param>
        /// <returns>An array of bytes containing the binary representation of the PDF document.</returns>
        public byte[] GeneratePdfBytesFromHtmlString(string html, string cssStyles)
        {
            try
            {
                return SaveFromHtml(html, cssStyles);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
                
        private byte[] SaveFromHtml(string html, string cssStyles)
        {
            try
            {
                byte[] bytes = { };

                using (MemoryStream ms = new MemoryStream())
                {
                    using (_Document = new Document())
                    {
                        _Document.SetMargins(MarginLeft, MarginRight, MarginTop, MarginBottom);
                        _Document.SetPageSize(PdfUtils.GetPageOrientation(DocumentOptions.PageOrientation));                        

                        PdfWriter pdfWriter = PdfWriter.GetInstance(_Document, ms);
                        pdfWriter.PageEvent = new PdfPageEvent(DocumentOptions);

                        PdfUtils.GetDocumentInfo(_Document, pdfWriter, DocumentOptions.DocumentInfo);
                        //pdfWriter.CreateXmpMetadata();

                        _Document.Open();

                        PdfPTable tableMain = new PdfPTable(1);
                        PdfPCell cellMain = new PdfPCell();
                        cellMain.Border = PdfPCell.NO_BORDER;

                        //using (var htmlWorker = new iTextSharp.text.html.simpleparser.HTMLWorker(document))
                        //{
                        //    using (var sr = new StringReader(html))
                        //    {
                        //        htmlWorker.Parse(sr);
                        //    }
                        //}
                        //===============================================================================================

                        //XMLWorker also reads from a TextReader and not directly from a string
                        //using (var srHtml = new StringReader(html))
                        //{
                        //    //Parse the HTML
                        //    iTextSharp.tool.xml.XMLWorkerHelper.GetInstance().ParseXHtml(pdfWriter, _Document, srHtml);
                        //}
                        //===============================================================================================
                                                
                        #region CSS Styles
                        if (!String.IsNullOrEmpty(cssStyles))
                        {
                            //Use the XMLWorker to parse HTML and CSS 
                            //In order to read CSS as a string we need to switch to a different constructor
                            //that takes Streams instead of TextReaders.
                            //Below we convert the strings into UTF8 byte array and wrap those in MemoryStreams
                            using (var msCss = new MemoryStream(Encoding.UTF8.GetBytes(cssStyles)))
                            {
                                using (var msHtml = new MemoryStream(Encoding.UTF8.GetBytes(html)))
                                {
                                    //Parse the HTML
                                    iTextSharp.tool.xml.XMLWorkerHelper.GetInstance().ParseXHtml(pdfWriter, _Document, msHtml, msCss);
                                }
                            }
                        }
                        else
                        {
                            var htmlarraylist = iTextSharp.text.html.simpleparser.HTMLWorker.ParseToList(new StringReader(html), new iTextSharp.text.html.simpleparser.StyleSheet());
                            
                            for (int k = 0; k < htmlarraylist.Count; k++)
                            {
                                var ele = (IElement)htmlarraylist[k];
                                cellMain.AddElement(ele);
                            }
                        }
                        #endregion

                        tableMain.AddCell(cellMain);
                        tableMain.WidthPercentage = 96.7f;
                        _Document.Add(tableMain);

                        _Document.Close();
                        bytes = ms.ToArray();
                    }
                }

                return bytes;
            }
            catch (DocumentException dex)
            {
                throw (dex);
            }
            catch (IOException ioex)
            {
                throw (ioex);
            }
            finally
            {
                _Document.Close();
            }
        }


        /// <summary>
        /// Cria o relatório de frequencia.
        /// </summary>
        /// <param name="listaFrequencias">lista dos dados</param>
        /// <returns>os bytes do relatório</returns>
        public byte[] RelatorioFrequencia(List<Frequencia> listaFrequencias)
        {
            try
            {
                byte[] bytes = { };

                using (MemoryStream ms = new MemoryStream())
                {
                    using (_Document = new Document())
                    {
                        _Document.SetMargins(MarginLeft, MarginRight, MarginTop, MarginBottom);
                        _Document.SetPageSize(PdfUtils.GetPageOrientation(DocumentOptions.PageOrientation));

                        PdfWriter pdfWriter = PdfWriter.GetInstance(_Document, ms);
                        pdfWriter.PageEvent = new PdfPageEvent(DocumentOptions);

                        PdfUtils.GetDocumentInfo(_Document, pdfWriter, DocumentOptions.DocumentInfo);
                        //pdfWriter.CreateXmpMetadata();

                        _Document.Open();

                        // define o nome das colunas no cabeçalho da tabela no pdf
                        string[] header = new string[] {
                            "Data", "Entrada", "Saída", "Total HorasXDia", "Tarefa"
                        };

                        // define a largura das colunas da tabela no pdf
                        float[] headerColumnsWidth = new float[] { 40, 30, 30, 30, 30 };

                        // cria um paragrafo com texto no pdf
                        PdfMethods.AddParagraph(_Document, @"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
                        PdfMethods.AddTableHeader(_Document, header, headerColumnsWidth);
                        PdfMethods.AddTableBodyContent(_Document, header.Count(), listaFrequencias.ToArray(), headerColumnsWidth);
                        PdfMethods.RodapeRelatorioFrequencia(_Document, listaFrequencias.SomarPeriodos(c => c.TotalHorasDia).ToString(), header.Count(), headerColumnsWidth);

                        _Document.Close();
                        bytes = ms.ToArray();

                        return bytes;
                    }
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
