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
                        PdfMethods.AddParagraph(_Document, @"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.", 10f, 10f, Element.ALIGN_LEFT);
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

        public byte[] RelatorioPrestacaoServico(PrestacaoServicos dados)
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

                        


                        PdfPTable tbSP = new PdfPTable(1);
                        PdfPCell cellSP = new PdfPCell(new Phrase(new Chunk($"São Paulo, {AppHelpers.DataMesPorExtenso(DateTime.Now)}", FontFactory.GetFont("Arial", 8f, Font.BOLD, BaseColor.BLACK))));
                        cellSP.HorizontalAlignment = Element.ALIGN_RIGHT;
                        cellSP.Border = PdfPCell.NO_BORDER;

                        tbSP.AddCell(cellSP);
                        tbSP.WidthPercentage = 96.7f;
                        tbSP.DefaultCell.FixedHeight = 15f;
                        _Document.Add(tbSP);

                        // pula uma linha no documento
                        PdfMethods.NewLine(_Document, 10f, false);


                        PdfPTable tbDados1 = new PdfPTable(4);
                        tbDados1.SetTotalWidth(new float[] { 9.5f, 16, 6, 100 });
                        tbDados1.WidthPercentage = 96.7f;
                        tbDados1.DefaultCell.FixedHeight = 15f;

                        PdfPCell cell1 = new PdfPCell(new Phrase(new Chunk(dados.NomeEscola, FontFactory.GetFont("Arial", 8f, Font.BOLD, BaseColor.BLACK))));
                        cell1.HorizontalAlignment = Element.ALIGN_LEFT;
                        cell1.Border = PdfPCell.NO_BORDER;
                        cell1.Colspan = 4;

                        PdfPCell cell2 = new PdfPCell(new Phrase(new Chunk(dados.EnderecoEscola, FontFactory.GetFont("Arial", 8f, BaseColor.BLACK))));
                        cell2.HorizontalAlignment = Element.ALIGN_LEFT;
                        cell2.Border = PdfPCell.NO_BORDER;
                        cell2.Colspan = 4;

                        PdfPCell cell3 = new PdfPCell(new Phrase(new Chunk($"Telefone:", FontFactory.GetFont("Arial", 8f, BaseColor.BLACK))));
                        cell3.HorizontalAlignment = Element.ALIGN_LEFT;
                        cell3.Border = PdfPCell.NO_BORDER;

                        PdfPCell cell3_1 = new PdfPCell(new Phrase(new Chunk(dados.TelefoneEscola, FontFactory.GetFont("Arial", 8f, BaseColor.BLACK))));
                        cell3_1.HorizontalAlignment = Element.ALIGN_LEFT;
                        cell3_1.Border = PdfPCell.NO_BORDER;

                        PdfPCell cell3_2 = new PdfPCell(new Phrase(new Chunk("CEP:", FontFactory.GetFont("Arial", 8f, BaseColor.BLACK))));
                        cell3_2.HorizontalAlignment = Element.ALIGN_LEFT;
                        cell3_2.Border = PdfPCell.NO_BORDER;

                        PdfPCell cell3_3 = new PdfPCell(new Phrase(new Chunk(dados.CepEscola, FontFactory.GetFont("Arial", 8f, BaseColor.BLACK))));
                        cell3_3.HorizontalAlignment = Element.ALIGN_LEFT;
                        cell3_3.Border = PdfPCell.NO_BORDER;


                        PdfPCell cell4 = new PdfPCell(new Phrase(new Chunk("São Paulo - SP", FontFactory.GetFont("Arial", 8f, BaseColor.BLACK))));
                        cell4.HorizontalAlignment = Element.ALIGN_LEFT;
                        cell4.Border = PdfPCell.NO_BORDER;
                        cell4.Colspan = 4;

                        tbDados1.AddCell(cell1);
                        tbDados1.AddCell(cell2);
                        tbDados1.AddCell(cell3);
                        tbDados1.AddCell(cell3_1);
                        tbDados1.AddCell(cell3_2);
                        tbDados1.AddCell(cell3_3);
                        tbDados1.AddCell(cell4);
                        
                        _Document.Add(tbDados1);


                        PdfMethods.AddParagraph(_Document, "Senhor(a) Diretor(a)", 5f, 5f, Element.ALIGN_LEFT);

                        Chunk chk1 = new Chunk(@"Com fundamento no convénio existente entre o Poder Judiciário/Justiça Federal e a Secretaria de Estado da Educação, esta representada pela FDE - Fundação para o Desenvolvimento da Educação, estamos encaminhando(3) candidato(a) abaixo qualificado(a) ", 
                            FontFactory.GetFont("Arial", 8f, BaseColor.BLACK));
                        Chunk chk2 = new Chunk(@"para prestar serviços gratuitos a essa Unidade Escolar de 07 (sete) à 14 (quatorze) horas semanais, ", FontFactory.GetFont("Arial", 8f, Font.BOLD, BaseColor.BLACK)).SetUnderline(0.5f, -1.5f);
                        Chunk chk3 = new Chunk(@"a serem distribuídas de acordo com o interesse da Direção, sem prejuízo da jornada normal de trabalho do(a) prestador(a). Esclarecemos que as tarefas a serem cumpridas ficarão a critério de Vossa Senhoria, de acordo com a aptidão e formação do(a) interessado(a).", FontFactory.GetFont("Arial", 8f, BaseColor.BLACK));


                        Chunk chk4 = new Chunk(@"Ficarão sob a responsabilidade do(a) Senhor(a) Diretor(a) a fiscalização e avaliação dos serviços prestados, a serem informados através do ", FontFactory.GetFont("Arial", 8f, BaseColor.BLACK));
                        Chunk chk5 = new Chunk(@"Atestado de Frequência ", FontFactory.GetFont("Arial", 8f, Font.BOLD, BaseColor.BLACK)).SetUnderline(0.5f, -1.5f);
                        Chunk chk6 = new Chunk(@"anexo que ", FontFactory.GetFont("Arial", 8f, BaseColor.BLACK));
                        Chunk chk7 = new Chunk(@"deverá ser enviado à FDE - Projeto Prestadores de Serviços Gratuitos à Comunidade - P.S.C., Av.: São Luís, 99 - 5° andar - República - SP - CEP 01046-001 via e-mail prestadoresdeservico@fde.sp.dov.br ou fax n° (11) 3158.4287, até o quinto dia útil do mês seguinte.", FontFactory.GetFont("Arial", 8f, Font.BOLD, BaseColor.BLACK)).SetUnderline(0.5f, -1.5f);

                        Chunk chk8 = new Chunk(@"Lembramos que não existe obrigatoriedade em manter o(a) prestador(a) de serviços por todo o tempo abaixo indicado, caso o mesmo não venha
a atender as expectativas. Ocorrendo algum motivo que Vossa Senhoria julgue suficiente, poderá o mesmo ser devolvido à FDE que promoverá a
transferência para outra unidade escolar, ou, devolução ao Poder Judiciário/Justiça Federal.", FontFactory.GetFont("Arial", 8f, BaseColor.BLACK));

                        Phrase ph1 = new Phrase();
                        ph1.Add(chk1);
                        ph1.Add(chk2);
                        ph1.Add(chk3);

                        Phrase ph2 = new Phrase();
                        ph2.Add(chk4);
                        ph2.Add(chk5);
                        ph2.Add(chk6);
                        ph2.Add(chk7);

                        Phrase ph3 = new Phrase();
                        ph3.Add(chk8);

                        // cria um paragrafo com texto no pdf
                        PdfMethods.AddParagraph(_Document, ph1, 0f, 0f, Element.ALIGN_LEFT);
                        // cria um paragrafo com texto no pdf
                        PdfMethods.AddParagraph(_Document, ph2, 0f, 0f, Element.ALIGN_LEFT);
                        // cria um paragrafo com texto no pdf
                        PdfMethods.AddParagraph(_Document, ph3, 0f, 0f, Element.ALIGN_LEFT);


                        Chunk chk9 = new Chunk(@"A interrupção no cumprimento da pena ou abandono por parte do prestador, deverão ser imediatamente informados à equipe da FDE que comunicará o Juízo competente.", FontFactory.GetFont("Arial", 10f, Font.BOLD, BaseColor.BLACK));
                        Phrase ph4 = new Phrase();
                        ph4.Add(chk9);
                        // cria um paragrafo com texto no pdf
                        PdfMethods.AddParagraph(_Document, ph4, 10f, 5f, Element.ALIGN_CENTER);


                        Font FontBold = FontFactory.GetFont("Arial", 8f, Font.BOLD, BaseColor.BLACK);
                        Font FontNormal = FontFactory.GetFont("Arial", 8f, BaseColor.BLACK);

                        // 1ª Linha ------------------------------------------------------------------------------
                        #region Linha 1
                        PdfPTable tbLinha1 = new PdfPTable(6);
                        tbLinha1.SetTotalWidth(new float[] { 5, 40, 5, 10, 6, 20 });
                        tbLinha1.WidthPercentage = 96.7f;
                        tbLinha1.DefaultCell.FixedHeight = 15f;

                        PdfPCell c1 = new PdfPCell(new Phrase(new Chunk("Nome:", FontBold)));
                        c1.HorizontalAlignment = Element.ALIGN_LEFT;
                        c1.Border = PdfPCell.NO_BORDER;

                        PdfPCell c1_vl = new PdfPCell(new Phrase(new Chunk(dados.Nome, FontNormal)));
                        c1_vl.HorizontalAlignment = Element.ALIGN_LEFT;
                        c1_vl.Border = PdfPCell.NO_BORDER;
                        //c1_vl.Colspan = 3;


                        PdfPCell c2 = new PdfPCell(new Phrase(new Chunk("RG nº.", FontBold)));
                        c2.HorizontalAlignment = Element.ALIGN_LEFT;
                        c2.Border = PdfPCell.NO_BORDER;

                        PdfPCell c2_vl = new PdfPCell(new Phrase(new Chunk(dados.Rg, FontNormal)));
                        c2_vl.HorizontalAlignment = Element.ALIGN_LEFT;
                        c2_vl.Border = PdfPCell.NO_BORDER;


                        PdfPCell c3 = new PdfPCell(new Phrase(new Chunk("CPF nº.", FontBold)));
                        c3.HorizontalAlignment = Element.ALIGN_LEFT;
                        c3.Border = PdfPCell.NO_BORDER;

                        PdfPCell c3_vl = new PdfPCell(new Phrase(new Chunk(dados.Cpf, FontNormal)));
                        c3_vl.HorizontalAlignment = Element.ALIGN_LEFT;
                        c3_vl.Border = PdfPCell.NO_BORDER;

                        tbLinha1.AddCell(c1);
                        tbLinha1.AddCell(c1_vl);
                        tbLinha1.AddCell(c2);
                        tbLinha1.AddCell(c2_vl);
                        tbLinha1.AddCell(c3);
                        tbLinha1.AddCell(c3_vl);
                        _Document.Add(tbLinha1);
                        #endregion
                        // ---------------------------------------------------------------------------------------

                        // 2ª Linha ------------------------------------------------------------------------------
                        #region Linha 2
                        PdfPTable tbLinha2 = new PdfPTable(8);
                        tbLinha2.SetTotalWidth(new float[] { 13.5f, 25, 7, 5, 13, 15, 14, 30 });
                        tbLinha2.WidthPercentage = 96.7f;
                        tbLinha2.DefaultCell.FixedHeight = 15f;

                        PdfPCell c4 = new PdfPCell(new Phrase(new Chunk("Naturalidade:", FontBold)));
                        c4.HorizontalAlignment = Element.ALIGN_LEFT;
                        c4.Border = PdfPCell.NO_BORDER;

                        PdfPCell c4_vl = new PdfPCell(new Phrase(new Chunk(dados.Naturalidade, FontNormal)));
                        c4_vl.HorizontalAlignment = Element.ALIGN_LEFT;
                        c4_vl.Border = PdfPCell.NO_BORDER;

                        PdfPCell c5 = new PdfPCell(new Phrase(new Chunk("Idade:", FontBold)));
                        c5.HorizontalAlignment = Element.ALIGN_LEFT;
                        c5.Border = PdfPCell.NO_BORDER;

                        PdfPCell c5_vl = new PdfPCell(new Phrase(new Chunk(dados.Idade, FontNormal)));
                        c5_vl.HorizontalAlignment = Element.ALIGN_LEFT;
                        c5_vl.Border = PdfPCell.NO_BORDER;

                        PdfPCell c6 = new PdfPCell(new Phrase(new Chunk("Estado Civil:", FontBold)));
                        c6.HorizontalAlignment = Element.ALIGN_LEFT;
                        c6.Border = PdfPCell.NO_BORDER;

                        PdfPCell c6_vl = new PdfPCell(new Phrase(new Chunk(dados.EstadoCivil, FontNormal)));
                        c6_vl.HorizontalAlignment = Element.ALIGN_LEFT;
                        c6_vl.Border = PdfPCell.NO_BORDER;

                        PdfPCell c7 = new PdfPCell(new Phrase(new Chunk("Escolaridade:", FontBold)));
                        c7.HorizontalAlignment = Element.ALIGN_LEFT;
                        c7.Border = PdfPCell.NO_BORDER;

                        PdfPCell c7_vl = new PdfPCell(new Phrase(new Chunk(dados.Escolaridade, FontNormal)));
                        c7_vl.HorizontalAlignment = Element.ALIGN_LEFT;
                        c7_vl.Border = PdfPCell.NO_BORDER;

                        tbLinha2.AddCell(c4);
                        tbLinha2.AddCell(c4_vl);
                        tbLinha2.AddCell(c5);
                        tbLinha2.AddCell(c5_vl);
                        tbLinha2.AddCell(c6);
                        tbLinha2.AddCell(c6_vl);
                        tbLinha2.AddCell(c7);
                        tbLinha2.AddCell(c7_vl);

                        _Document.Add(tbLinha2);
                        #endregion
                        // ---------------------------------------------------------------------------------------


                        // 3ª Linha ------------------------------------------------------------------------------
                        #region Linha 3
                        PdfPTable tbLinha3 = new PdfPTable(2);
                        tbLinha3.SetTotalWidth(new float[] { 11, 100 });
                        tbLinha3.WidthPercentage = 96.7f;
                        tbLinha3.DefaultCell.FixedHeight = 15f;

                        PdfPCell c8 = new PdfPCell(new Phrase(new Chunk("Residência:", FontBold)));
                        c8.HorizontalAlignment = Element.ALIGN_LEFT;
                        c8.Border = PdfPCell.NO_BORDER;

                        PdfPCell c8_vl = new PdfPCell(new Phrase(new Chunk(dados.Residencia, FontNormal)));
                        c8_vl.HorizontalAlignment = Element.ALIGN_LEFT;
                        c8_vl.Border = PdfPCell.NO_BORDER;
                        //c8_vl.Colspan = 7;

                        tbLinha3.AddCell(c8);
                        tbLinha3.AddCell(c8_vl);
                        _Document.Add(tbLinha3);
                        #endregion
                        // ---------------------------------------------------------------------------------------

                        // 4ª Linha ------------------------------------------------------------------------------
                        #region Linha 4
                        PdfPTable tbLinha4 = new PdfPTable(6);
                        tbLinha4.SetTotalWidth(new float[] { 9.5f, 14, 10, 25, 17, 40 });
                        tbLinha4.WidthPercentage = 96.7f;
                        tbLinha4.DefaultCell.FixedHeight = 15f;

                        PdfPCell c9 = new PdfPCell(new Phrase(new Chunk("Telefone:", FontBold)));
                        c9.HorizontalAlignment = Element.ALIGN_LEFT;
                        c9.Border = PdfPCell.NO_BORDER;

                        PdfPCell c9_vl = new PdfPCell(new Phrase(new Chunk(dados.Telefone, FontNormal)));
                        c9_vl.HorizontalAlignment = Element.ALIGN_LEFT;
                        c9_vl.Border = PdfPCell.NO_BORDER;

                        PdfPCell c10 = new PdfPCell(new Phrase(new Chunk("Profissão:", FontBold)));
                        c10.HorizontalAlignment = Element.ALIGN_LEFT;
                        c10.Border = PdfPCell.NO_BORDER;

                        PdfPCell c10_vl = new PdfPCell(new Phrase(new Chunk(dados.Profissao, FontNormal)));
                        c10_vl.HorizontalAlignment = Element.ALIGN_LEFT;
                        c10_vl.Border = PdfPCell.NO_BORDER;
                        //c10_vl.Colspan = 2;

                        PdfPCell c11 = new PdfPCell(new Phrase(new Chunk("Órgão de Origem:", FontBold)));
                        c11.HorizontalAlignment = Element.ALIGN_LEFT;
                        c11.Border = PdfPCell.NO_BORDER;

                        PdfPCell c11_vl = new PdfPCell(new Phrase(new Chunk(dados.OrgaoOrigem, FontNormal)));
                        c11_vl.HorizontalAlignment = Element.ALIGN_LEFT;
                        c11_vl.Border = PdfPCell.NO_BORDER;
                        //c11_vl.Colspan = 2;

                        tbLinha4.AddCell(c9);
                        tbLinha4.AddCell(c9_vl);
                        tbLinha4.AddCell(c10);
                        tbLinha4.AddCell(c10_vl);
                        tbLinha4.AddCell(c11);
                        tbLinha4.AddCell(c11_vl);
                        _Document.Add(tbLinha4);
                        #endregion
                        // ---------------------------------------------------------------------------------------

                        // 5ª Linha ------------------------------------------------------------------------------
                        #region Linha 5
                        PdfPTable tbLinha5 = new PdfPTable(2);
                        tbLinha5.SetTotalWidth(new float[] { 34, 100 });
                        tbLinha5.WidthPercentage = 96.7f;
                        tbLinha5.DefaultCell.FixedHeight = 15f;

                        PdfPCell c12 = new PdfPCell(new Phrase(new Chunk("Motivo da prestação de serviços:", FontBold)));
                        c12.HorizontalAlignment = Element.ALIGN_LEFT;
                        c12.Border = PdfPCell.NO_BORDER;
                        //c12.Colspan = 2;

                        PdfPCell c12_vl = new PdfPCell(new Phrase(new Chunk(dados.MotivoPrestacaoServico, FontNormal)));
                        c12_vl.HorizontalAlignment = Element.ALIGN_LEFT;
                        c12_vl.Border = PdfPCell.NO_BORDER;
                        //c12_vl.Colspan = 6;

                        tbLinha5.AddCell(c12);
                        tbLinha5.AddCell(c12_vl);
                        _Document.Add(tbLinha5);
                        #endregion
                        // ---------------------------------------------------------------------------------------

                        // 6ª Linha ------------------------------------------------------------------------------
                        #region Linha 6
                        PdfPTable tbLinha6 = new PdfPTable(2);
                        tbLinha6.SetTotalWidth(new float[] { 39, 100 });
                        tbLinha6.WidthPercentage = 96.7f;
                        tbLinha6.DefaultCell.FixedHeight = 15f;

                        PdfPCell c13 = new PdfPCell(new Phrase(new Chunk("Prazo para a prestação dos serviços:", FontBold)));
                        c13.HorizontalAlignment = Element.ALIGN_LEFT;
                        c13.Border = PdfPCell.NO_BORDER;
                        //c13.Colspan = 2;

                        PdfPCell c13_vl = new PdfPCell(new Phrase(new Chunk(dados.PrazoPrestacaoServico, FontNormal)));
                        c13_vl.HorizontalAlignment = Element.ALIGN_LEFT;
                        c13_vl.Border = PdfPCell.NO_BORDER;
                        //c13_vl.Colspan = 6;

                        tbLinha6.AddCell(c13);
                        tbLinha6.AddCell(c13_vl);
                        _Document.Add(tbLinha6);
                        #endregion
                        // ---------------------------------------------------------------------------------------

                        // 7ª Linha ------------------------------------------------------------------------------
                        #region Linha 7
                        PdfPTable tbLinha7 = new PdfPTable(4);
                        tbLinha7.SetTotalWidth(new float[] { 9, 7, 3, 40 });
                        tbLinha7.WidthPercentage = 96.7f;
                        tbLinha7.DefaultCell.FixedHeight = 15f;

                        PdfPCell c14 = new PdfPCell(new Phrase(new Chunk("Apresentar-se dia:", FontBold)));
                        c14.HorizontalAlignment = Element.ALIGN_LEFT;
                        c14.Border = PdfPCell.NO_BORDER;
                        //c14.Colspan = 2;

                        PdfPCell c14_vl = new PdfPCell(new Phrase(new Chunk(dados.DataApresentacao.ToShortDateString(), FontNormal)));
                        c14_vl.HorizontalAlignment = Element.ALIGN_LEFT;
                        c14_vl.Border = PdfPCell.NO_BORDER;
                        //c14_vl.Colspan = 2;

                        PdfPCell c15 = new PdfPCell(new Phrase(new Chunk("às", FontBold)));
                        c15.HorizontalAlignment = Element.ALIGN_LEFT;
                        c15.Border = PdfPCell.NO_BORDER;

                        PdfPCell c15_vl = new PdfPCell(new Phrase(new Chunk(dados.DataApresentacao.ToShortTimeString(), FontNormal)));
                        c15_vl.HorizontalAlignment = Element.ALIGN_LEFT;
                        c15_vl.Border = PdfPCell.NO_BORDER;

                        //PdfPCell c16 = new PdfPCell(new Phrase(new Chunk("", FontBold)));
                        //c16.HorizontalAlignment = Element.ALIGN_LEFT;
                        //c16.Border = PdfPCell.NO_BORDER;
                        //c16.Colspan = 3;

                        tbLinha7.AddCell(c14);
                        tbLinha7.AddCell(c14_vl);
                        tbLinha7.AddCell(c15);
                        tbLinha7.AddCell(c15_vl);
                        _Document.Add(tbLinha7);
                        #endregion
                        // ---------------------------------------------------------------------------------------


                        Chunk chk10 = new Chunk(@"Para maiores esclarecimentos, entrar em contato com a equipe do Projeto pelos ", FontFactory.GetFont("Arial", 8f, BaseColor.BLACK));
                        Chunk chk11 = new Chunk(@"telefones nº. 3158-4246, 3158-4668, 3158-4250, 3158-4253 ou 3158-4278. ", FontFactory.GetFont("Arial", 8f, Font.BOLD, BaseColor.BLACK)).SetUnderline(0.5f, -1.5f);
                        Phrase ph5 = new Phrase();
                        ph5.Add(chk10);
                        ph5.Add(chk11);
                        // cria um paragrafo com texto no pdf
                        PdfMethods.AddParagraph(_Document, ph5, 5f, 5f, Element.ALIGN_LEFT);


                        // cria um paragrafo com texto no pdf
                        PdfMethods.AddParagraph(_Document, new Phrase(new Chunk("Renovando protestos de estima e consideração, subscrevemo-nos.", FontFactory.GetFont("Arial", 8f, BaseColor.BLACK))), 10f, 10f, Element.ALIGN_LEFT);

                        // cria um paragrafo com texto no pdf
                        PdfMethods.AddParagraph(_Document, new Phrase(new Chunk("Atenciosamente,", FontFactory.GetFont("Arial", 8f, BaseColor.BLACK))), 10f, 0f, Element.ALIGN_CENTER);
                        PdfMethods.AddParagraph(_Document, new Phrase(new Chunk("Nadir de Almeida", FontFactory.GetFont("Arial", 8f, Font.BOLD, BaseColor.BLACK))), 0f, 0f, Element.ALIGN_CENTER);
                        PdfMethods.AddParagraph(_Document, new Phrase(new Chunk("Coordenadora do Projeto", FontFactory.GetFont("Arial", 8f, BaseColor.BLACK))), 0f, 0f, Element.ALIGN_CENTER);
                        PdfMethods.AddParagraph(_Document, new Phrase(new Chunk("Prestadores de Serviços", FontFactory.GetFont("Arial", 8f, BaseColor.BLACK))), 20f, 0f, Element.ALIGN_CENTER);

                        // cria um paragrafo com texto no pdf
                        PdfMethods.AddParagraph(_Document, new Phrase(new Chunk("Recebi a 1ª Via desta Carta para entregar na escola acima.", FontFactory.GetFont("Arial", 8f, BaseColor.BLACK))), 10f, 0f, Element.ALIGN_LEFT);

                        // cria um paragrafo com texto no pdf
                        var text = @"(a)_______________________________________________________________________________  em ___/___/___";
                        PdfMethods.AddParagraph(_Document, new Phrase(new Chunk(text, FontFactory.GetFont("Arial", 8f, BaseColor.BLACK))), 0f, 0f, Element.ALIGN_LEFT);

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
