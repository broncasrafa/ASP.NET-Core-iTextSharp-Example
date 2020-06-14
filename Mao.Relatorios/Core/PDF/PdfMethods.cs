using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace Mao.Relatorios.Core.PDF
{
    public static class PdfMethods
    {
        /// <summary>
        /// Carrega o conteúdo dentro da tag body do html em uma string
        /// </summary>
        /// <param name="templateFilePath">caminho fisico do template junto com o nome do arquivo</param>
        /// <returns>a string (texto) do html </returns>
        public static string LoadHtmlTemplate(string templateFilePath)
        {
            string filename = templateFilePath.Split('/')?.LastOrDefault();

            try
            {
                using (TextReader reader = File.OpenText(templateFilePath))
                {
                    string htmlText = reader.ReadToEnd();

                    if (Regex.IsMatch(htmlText, "</body>"))
                    {
                        try
                        {
                            htmlText = Regex.Split(htmlText, @"<body\s*.*(\w|\d|\s)*(?>>)")[1];
                            htmlText = Regex.Split(htmlText, @"</body>")[0];
                        }
                        catch (Exception ex)
                        {
                            throw new ApplicationException($"Erro no layout do template: {filename}", ex);
                        }
                    }
                    return htmlText;
                }
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Erro ao tentar carregar o template: {filename}", ex);
            }
        }


        /// <summary>
        /// Cria o cabeçalho de uma tabela no documento pdf
        /// </summary>
        /// <param name="document">objeto document</param>
        /// <param name="headerFields">lista de campos do cabeçalho</param>
        /// <param name="totalWidths">lista da largura de cada campo do cabeçalho</param>        
        public static void AddTableHeader(Document document, string[] headerFields, float[] totalWidths)
        {
            int qtdeCells = headerFields.Count();

            PdfPTable table = new PdfPTable(qtdeCells);
            table.SetTotalWidth(totalWidths);

            PdfPCell[] cells = new PdfPCell[qtdeCells];

            string texto = "";

            for (int i = 0; i < qtdeCells; i++)
            {
                texto = headerFields[i].ToString();

                cells[i] = new PdfPCell(new Phrase(new Chunk(texto, FontFactory.GetFont("Arial", 8f, BaseColor.BLACK))));
                cells[i].BackgroundColor = WebColors.GetRGBColor("#B5B5B5");
                cells[i].HorizontalAlignment = Element.ALIGN_LEFT;
                cells[i].Border = PdfPCell.NO_BORDER;

                table.AddCell(cells[i]);
            }

            table.WidthPercentage = 96.7f;
            table.DefaultCell.FixedHeight = 15f;
            document.Add(table);

            //return table;
        }


        /// <summary>
        /// Cria um paragrafo no documento pdf
        /// </summary>
        /// <param name="document">objeto document</param>
        /// <param name="paragraphText">texto do paragrafo</param>
        /// <param name="paddingBottom">distancia do paragrafo para baixo</param>
        /// <param name="paddingTop">distancia do paragrafo para cima</param>
        /// <param name="horizontalAlignment">alinhamento do paragrafo. Usar Element. Por exemplo (Element.ALIGN_CENTER)</param>
        public static void AddParagraph(Document document, string paragraphText, float paddingBottom, float paddingTop, int horizontalAlignment)
        {
            Paragraph paragraph = new Paragraph(new Phrase(new Chunk(paragraphText, FontFactory.GetFont("Arial", 8f, BaseColor.BLACK))));

            PdfPTable table = new PdfPTable(1);

            PdfPCell cell = new PdfPCell();
            cell.Border = PdfPCell.NO_BORDER;
            cell.PaddingBottom = paddingBottom;
            cell.PaddingTop = paddingTop;
            SetParagraphAlignment(paragraph, horizontalAlignment);
            cell.AddElement(paragraph);

            table.AddCell(cell);
            table.WidthPercentage = 96.7f;
            table.DefaultCell.FixedHeight = 15f;
            document.Add(table);
        }
        /// <summary>
        /// Cria um paragrafo no documento pdf
        /// </summary>
        /// <param name="document">objeto document</param>
        /// <param name="phrase">objeto Phrase que consiste em um array de chunks</param>
        /// <param name="paddingBottom">distancia do paragrafo para baixo</param>
        /// <param name="paddingTop">distancia do paragrafo para cima</param>
        /// <param name="horizontalAlignment">alinhamento do paragrafo. Usar Element. Por exemplo (Element.ALIGN_CENTER)</param>
        public static void AddParagraph(Document document, Phrase phrase, float paddingBottom, float paddingTop, int horizontalAlignment)
        {
            Paragraph paragraph = new Paragraph(phrase);
            PdfPTable table = new PdfPTable(1);
            PdfPCell cell = new PdfPCell();
            cell.Border = PdfPCell.NO_BORDER;
            cell.PaddingBottom = paddingBottom;
            cell.PaddingTop = paddingTop;
            SetParagraphAlignment(paragraph, horizontalAlignment);            
            cell.AddElement(paragraph);

            table.AddCell(cell);
            table.WidthPercentage = 96.7f;
            table.DefaultCell.FixedHeight = 15f;
            document.Add(table);
        }


        /// <summary>
        /// Pula uma linha no documento do pdf.
        /// </summary>
        /// <param name="document">objeto document</param>
        /// <param name="height">distancia do pulo da linha</param>
        /// <param name="hasTopBorder">se vai ter uma borda no topo para separação de conteúdo</param>
        public static void NewLine(Document document, float height, bool hasTopBorder)
        {
            PdfPTable table = new PdfPTable(1);

            PdfPCell cell = new PdfPCell();
            cell.FixedHeight = height;
            if (hasTopBorder)
            {
                cell.Border = PdfPCell.TOP_BORDER;
                cell.BorderColorTop = BaseColor.BLACK;
            }
            else
            {
                cell.Border = PdfPCell.NO_BORDER;
            }

            cell.PaddingBottom = 10f;
            cell.PaddingTop = 10f;

            table.AddCell(cell);
            table.WidthPercentage = 96.7f;
            table.DefaultCell.FixedHeight = 15f;
            document.Add(table);
        }


        /// <summary>
        /// Pula uma linha no documento do pdf.
        /// </summary>
        /// <param name="customCell">célula da tabela atual para pular a linha</param>
        /// <param name="height">distancia do pulo da linha</param>
        /// <param name="hasTopBorder">se vai ter uma borda no topo para separação de conteúdo</param>
        public static void NewLine(PdfPCell customCell, float height, bool hasTopBorder)
        {
            PdfPTable table = new PdfPTable(1);

            PdfPCell cell = new PdfPCell();
            cell.FixedHeight = height;
            if (hasTopBorder)
            {
                cell.Border = PdfPCell.TOP_BORDER;
                cell.BorderColorTop = BaseColor.BLACK;
            }
            else
            {
                cell.Border = PdfPCell.NO_BORDER;
            }

            cell.PaddingBottom = 10f;
            cell.PaddingTop = 10f;

            table.AddCell(cell);
            table.WidthPercentage = 96.7f;
            table.DefaultCell.FixedHeight = 15f;
            customCell.AddElement(table);
        }


        /// <summary>
        /// Cria o conteúdo da tabela no documento pdf
        /// </summary>
        /// <param name="document">objeto document</param>
        /// <param name="cellsCount">quantidade de colunas da tabela</param>
        /// <param name="entityValues">array dos dados do objeto</param>
        /// <param name="totalWidths">lista da largura de cada coluna da tabela</param>
        public static void AddTableBodyContent(Document document, int cellsCount, object[] entityValues, float[] totalWidths)
        {
            PdfPTable table = new PdfPTable(cellsCount);
            table.SetTotalWidth(totalWidths);

            PdfPCell[] cells = new PdfPCell[cellsCount];

            foreach (var modelo in entityValues)
            {
                int i = 0;

                foreach (var item in modelo.GetType().GetProperties(BindingFlags.DeclaredOnly | BindingFlags.Instance | BindingFlags.Public))/*.Where(p => p.DeclaringType == typeof(object))*/
                {
                    if (GetValidPropertyType(item))
                    {
                        if (item.GetValue(modelo, null) != null)
                        {
                            cells[i] = new PdfPCell(new Phrase(new Chunk(item.GetValue(modelo, null).ToString(), FontFactory.GetFont("Arial", 8f, BaseColor.BLACK))));
                        }
                        else
                        {
                            cells[i] = new PdfPCell(new Phrase(new Chunk("", FontFactory.GetFont("Arial", 8f, BaseColor.BLACK))));
                        }

                        cells[i].Border = PdfPCell.NO_BORDER;
                        table.AddCell(cells[i]);

                        i++;
                    }
                }
            }

            table.WidthPercentage = 96.7f;
            table.DefaultCell.FixedHeight = 15f;
            document.Add(table);
        }

        public static PdfPTable RodapeRelatorioFrequencia(Document document, string periodoTotal, int qtdeCells, float[] totalWidths)
        {
            PdfPTable table = new PdfPTable(qtdeCells);
            table.SetTotalWidth(totalWidths);

            PdfPCell[] cells = new PdfPCell[qtdeCells];

            for (int i = 0; i < qtdeCells; i++)
            {
                string texto = "";

                if (i == 0)
                {
                    texto = "Total Mês";
                }
                else if (i == 3)
                {
                    texto = periodoTotal;
                }

                cells[i] = new PdfPCell(new Phrase(new Chunk(texto, FontFactory.GetFont("Arial", 8f, BaseColor.BLACK))));
                cells[i].BackgroundColor = WebColors.GetRGBColor("#B5B5B5");
                cells[i].HorizontalAlignment = i == 0 ? Element.ALIGN_CENTER : Element.ALIGN_LEFT;
                cells[i].Border = PdfPCell.NO_BORDER;

                table.AddCell(cells[i]);
            }

            table.WidthPercentage = 96.7f;
            table.DefaultCell.FixedHeight = 15f;
            document.Add(table);

            return table;
        }


        private static void SetParagraphAlignment(Paragraph paragraph, int alignment)
        {
            switch (alignment)
            {
                case Element.ALIGN_UNDEFINED: paragraph.Alignment = Element.ALIGN_UNDEFINED; break;
                case Element.ALIGN_LEFT: paragraph.Alignment = Element.ALIGN_LEFT; break;
                case Element.ALIGN_CENTER: paragraph.Alignment = Element.ALIGN_CENTER; break;
                case Element.ALIGN_RIGHT: paragraph.Alignment = Element.ALIGN_RIGHT; break;
                case Element.ALIGN_JUSTIFIED: paragraph.Alignment = Element.ALIGN_JUSTIFIED; break;
                case Element.ALIGN_TOP: paragraph.Alignment = Element.ALIGN_TOP; break;
                case Element.ALIGN_MIDDLE: paragraph.Alignment = Element.ALIGN_MIDDLE; break;
                case Element.ALIGN_BOTTOM: paragraph.Alignment = Element.ALIGN_BOTTOM; break;
                case Element.ALIGN_BASELINE: paragraph.Alignment = Element.ALIGN_BASELINE; break;
                case Element.ALIGN_JUSTIFIED_ALL: paragraph.Alignment = Element.ALIGN_JUSTIFIED_ALL; break;
                default:
                    paragraph.Alignment = Element.ALIGN_LEFT;
                    break;
            }
        }
        private static bool GetValidPropertyType(PropertyInfo modelo)
        {
            bool _isValid = false;

            if (modelo.PropertyType.Name.Contains("String") ||
                modelo.PropertyType.Name.Contains("string") ||
                modelo.PropertyType.Name.Contains("bool") ||
                modelo.PropertyType.Name.Contains("byte") ||
                modelo.PropertyType.Name.Contains("char") ||
                modelo.PropertyType.Name.Contains("decimal") ||
                modelo.PropertyType.Name.Contains("Decimal") ||
                modelo.PropertyType.Name.Contains("double") ||
                modelo.PropertyType.Name.Contains("float") ||
                modelo.PropertyType.Name.Contains("int") ||
                modelo.PropertyType.Name.Contains("Int16") ||
                modelo.PropertyType.Name.Contains("Int32") ||
                modelo.PropertyType.Name.Contains("Int64") ||
                modelo.PropertyType.Name.Contains("DateTime") ||
                modelo.PropertyType.Name.Contains("TimeSpan") ||
                modelo.PropertyType.Name.Contains("long"))
            {
                _isValid = true;
            }

            return _isValid;
        }
    }
}
