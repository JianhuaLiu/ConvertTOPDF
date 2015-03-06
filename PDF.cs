using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.html.simpleparser;
using System.IO;
using System.Linq;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using TuesPechkin;

using System.Drawing.Printing;

namespace ConvertTOPDF
{
    public class PDF : PdfPageEventHelper
    {

        private string RutaMarcaDeAgua
        {
            get
            {
                return RutaImagen = string.Format(@"MarcaDeAgua.png");
            }
        }

#if DEBUG
        private static readonly IConverter converter = new ThreadSafeConverter(new PdfToolset(new Win32EmbeddedDeployment(new TempFolderDeployment())));
#else
        private static readonly IConverter converter = new ThreadSafeConverter(new PdfToolset(new Win64EmbeddedDeployment(new TempFolderDeployment())));
#endif

        public PDF() { }

        public MemoryStream GenerarPDF(Repeater pControl)
        {
           return  ControlToPDF(pControl);
        }

        public MemoryStream GenerarPDF(string pUrl, string ValorEncabezado = "")
        {
           return UrlToPdf(pUrl, ValorEncabezado);
        }

        private MemoryStream ControlToPDF(Control pControl)
        {
            StringWriter lSWriter = new StringWriter();
            HtmlTextWriter htw = new HtmlTextWriter(lSWriter);
            Page lPagina = new Page();
            HtmlForm lForm = new HtmlForm();
            lPagina.EnableEventValidation = false;
            lPagina.DesignerInitialize();
            lPagina.Controls.Add(lForm);
            lForm.Controls.Add(pControl);
            lPagina.RenderControl(htw);
            StringReader sr = new StringReader(lSWriter.ToString());
            Document pdfDoc = new Document(PageSize.A4, 10f, 10f, 100f, 0f);
            HTMLWorker htmlparser = new HTMLWorker(pdfDoc);
            var bayos = new MemoryStream();
            PdfWriter writer = PdfWriter.GetInstance(pdfDoc, bayos);
            pdfDoc.Open();
            htmlparser.Parse(sr);
            pdfDoc.Close();          

            MemoryStream MemoryS = null;
            try
            {
      
                MemoryS = new MemoryStream();
                using (PdfReader pdfReader = new PdfReader(bayos.ToArray()))
                using (PdfStamper pdfStamper = new PdfStamper(pdfReader, MemoryS))
                {
                    Rectangle pageSize = pdfReader.GetPageSize(1);
                    float y = pageSize.Top - 20;
                    float yWaterMark = pageSize.Top / 3;
                    float xWaterMark = pageSize.Width / 3;
                    iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance(RutaMarcaDeAgua);
                    img.SetAbsolutePosition(xWaterMark, yWaterMark);
                    PdfContentByte waterMark;
                    int PaginasTotal = pdfReader.NumberOfPages;
                    string Paginacion = string.Empty;

                    for (int pageIndex = 1; pageIndex <= PaginasTotal; pageIndex++)
                    {
                        Paginacion = string.Format("Página {0} de {1}", pageIndex, PaginasTotal);
                        waterMark = pdfStamper.GetOverContent(pageIndex);
                        waterMark.AddImage(img);
                        ColumnText.ShowTextAligned(waterMark, Element.ALIGN_LEFT, new Phrase(Paginacion), 10, y, 0);
                    }
                    pdfStamper.FormFlattening = true;
                }
            }
            catch
            {
                MemoryS = null;
                throw;
            }
            return MemoryS;
        }

        private MemoryStream UrlToPdf(string pUrl, string ValorEncabezado = "")
        {
            var document = new HtmlToPdfDocument
            {
                GlobalSettings =
                {
                    ProduceOutline = true,
                    DocumentTitle = "Documento",
                    PaperSize = PaperKind.A4,
                },
                Objects = {
                new ObjectSettings { PageUrl = pUrl}
                }
            };

           

            MemoryStream MemoryS = null;
            try
            {
                byte[] result = converter.Convert(document);
                MemoryS = new MemoryStream();
                using (PdfReader pdfReader = new PdfReader(result.ToArray()))
                using (PdfStamper pdfStamper = new PdfStamper(pdfReader, MemoryS))
                {
                    Rectangle pageSize = pdfReader.GetPageSize(1);
                    float y = pageSize.Top - 20;
                    float yWaterMark = pageSize.Top / 3;
                    float xWaterMark = pageSize.Width / 3;
                    iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance(RutaMarcaDeAgua);
                    img.SetAbsolutePosition(xWaterMark, yWaterMark);
                    PdfContentByte waterMark;
                    int PaginasTotal = pdfReader.NumberOfPages;
                    string Paginacion = string.Empty;

                    for (int pageIndex = 1; pageIndex <= PaginasTotal; pageIndex++)
                    {
                        Paginacion = string.Format("Página {0} de {1} - {2}", pageIndex, PaginasTotal, ValorEncabezado);
                        waterMark = pdfStamper.GetOverContent(pageIndex);
                        waterMark.AddImage(img);
                        ColumnText.ShowTextAligned(waterMark, Element.ALIGN_LEFT, new Phrase(Paginacion), 10, y, 0);
                    }
                    pdfStamper.FormFlattening = true;
                }
            }
            catch
            {
                MemoryS = null;                
                throw;
            }

            return MemoryS;
        }
    }
}