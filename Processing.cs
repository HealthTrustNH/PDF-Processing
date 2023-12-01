using HtmlAgilityPack;
using iText.IO.Font;
using iText.IO.Image;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Action;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Layout;
using iText.Layout.Borders;
using iText.Layout.Element;
using iText.Layout.Properties;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;

namespace healthtrustnh.org_PDF_Processing
{
    public class chunks
    {
        /// <summary>
        /// Cell Contents
        /// </summary>
        public string text { get; set; }
        /// <summary>
        /// Is the cell marked as BOLD
        /// </summary>
        public bool bold { get; set; }
        /// <summary>
        /// Is the cell marked as ITALIC
        /// </summary>
        public bool italic { get; set; }
        /// <summary>
        /// Underline the text
        /// </summary>
        public bool underline { get; set; }

        /// <summary>
        /// Font Size
        /// </summary>
        public float fontSize { get; set; }
        /// <summary>
        /// Cell Font Color
        /// </summary>
        public string fontCol { get; set; }
        /// <summary>
        /// Font Face / Family
        /// </summary>
        public string fontFace { get; set; }
    }

    public class PDFAddressInfo
    {
        public int PageNo { get; set; }
        public string AttnName { get; set; }
        public string Title { get; set; }
        public string MemberName { get; set; }
        public string Address1 { get; set; }
        public string Address2 { get; set; }
        public string CityStZIP { get; set; }
        public byte[] docData { get; set; }
    }
    public class PDFProcessing
    {
        private MemoryStream msOutput;
        private PdfDocument pdf;
        private Document document;
        private PdfWriter writer;
        private Table table;

        public void Cleanup()
        {
            try
            {
                msOutput.Dispose();
            }
            catch
            {
                //Wasn't open but that's OK
            }
            msOutput = null;
            pdf = null;
            document = null;
            writer = null;
            table = null;
        }
       
        #region workers

        public void newDocument(bool suppressLogo = false, bool landscape = false, int leftMargin = 60, int topMargin = 60, int rightMargin = 60, int bottomMargin = 60, float logoScaling = .22f, bool includeFooter = false)
        {
            msOutput = new MemoryStream();
            writer = new PdfWriter(msOutput);
            pdf = new PdfDocument(writer);
            if (landscape)
                document = new Document(pdf, PageSize.LETTER.Rotate());
            else
                document = new Document(pdf, PageSize.LETTER);
            document.SetMargins(topMargin, rightMargin, bottomMargin, leftMargin);

            //Now add the HealthTrust Logo
            if (!suppressLogo)
            {
                if (landscape)
                    pageLogo(pageheight: PageSize.LETTER.Rotate().GetHeight(), scale: logoScaling);
                else
                    pageLogo(pageheight: PageSize.LETTER.GetHeight(), scale: logoScaling);
            }

            if (includeFooter && !landscape)
                setPageFooter();
        }

        public PdfFont getFont(string fontName)
        {
            PdfFont output = null;
            switch (fontName.ToLower())
            {
                case "courier":
                    output = PdfFontFactory.CreateFont(FontConstants.COURIER);
                    break;
                case "times":
                case "times new roman":
                    output = PdfFontFactory.CreateFont(FontConstants.TIMES_ROMAN);
                    break;
                case "helvetica":
                    output = PdfFontFactory.CreateFont(FontConstants.HELVETICA);
                    break;
                case "calibri":
                case "arial narrow":
                    output = PdfFontFactory.CreateFont(Properties.Resources.calibri, true);
                    break;
                case "":
                    break;
                default:
                    throw new Exception("Unknown font detected for PDF Processing: " + ((fontName ?? "") == "" ? "[No Font Specified]" : fontName));
            }
            return output;
        }
        public void setFont(string fontName)
        {
            document.SetFont(getFont(fontName));         
        }

        public void setPageFooter()
        {
            var greenFont = WebColors.GetRGBColor("#006A4E");
            var whiteFont = WebColors.GetRGBColor("#000000");
            var times = PdfFontFactory.CreateFont(FontConstants.TIMES_ROMAN);

            var p = new Paragraph();
            p.SetFixedPosition(20, 55, PageSize.LETTER.GetWidth() - 40);
            var b = new SolidBorder(1);
            b.SetColor(greenFont);
            p.SetBorderBottom(b);
            document.Add(p);

            p = new Paragraph("PO Box 617 * Concord, NH 03302-0617 * Tel. 603.226.2861 * Toll Free 800.527.5001 * Fax: 603.226.2988");
            p.SetFont(times);
            p.SetFontSize(9);
            p.SetFontColor(greenFont);
            p.SetFixedPosition(125, 35, PageSize.LETTER.GetWidth() - 100);
            document.Add(p);

            p = new Paragraph("Email: info@healthtrustnh.org * Website: ");
            var t = new Text("www.healthtrustnh.org");
            t.SetItalic();
            t.SetHorizontalAlignment(HorizontalAlignment.CENTER);
            p.Add(t);

            p.SetFont(times);
            p.SetFontSize(9);
            p.SetFontColor(greenFont);
            p.SetHorizontalAlignment(HorizontalAlignment.CENTER);
            p.SetFixedPosition(175, 25, PageSize.LETTER.GetWidth() - 300);
            document.Add(p);
        }

        public void pageLogo(float hPos = 0, float vPos = 0, float pageheight = 0, float scale = .22f)
        {
            using (var imgStream = new MemoryStream())
            {
                Properties.Resources.ht_logo.Save(imgStream, System.Drawing.Imaging.ImageFormat.Png);
                imgStream.ToArray();
                iText.Layout.Element.Image logo = new iText.Layout.Element.Image(ImageDataFactory.Create(imgStream.ToArray()));
                logo.Scale(scale, scale);
                if (hPos != 0 || vPos != 0)
                {
                    var c = new Paragraph();
                    c.Add(logo);
                    c.SetFixedPosition(hPos, vPos, logo.GetImageWidth());
                    document.Add(c);
                }
                else if (pageheight == 0)
                {
                    document.Add(logo);
                }
                else
                {
                    var c = new Paragraph();
                    c.Add(logo);
                    c.SetFixedPosition(20, pageheight - (float)logo.GetImageScaledHeight() - 20, logo.GetImageScaledWidth());
                    document.Add(c);
                }
            }
        }

        public void image(byte[] image, float scalePercent = 1.0f)
        {
            iText.Layout.Element.Image img = new iText.Layout.Element.Image(ImageDataFactory.Create(image));
            img.Scale(scalePercent, scalePercent);
            document.Add(img);
        }

        public void separator()
        {
            table = new Table(new float[] { 1 });
            table.SetWidth(UnitValue.CreatePercentValue(100));
            var c = new Cell();
            c.SetBorder(Border.NO_BORDER);
            var b = new SolidBorder(1);
            c.SetBorderBottom(b);
            table.AddCell(c);
            document.Add(table);
        }

        public void newTable(double[] cellWidths, bool fullWidth = true)
        {
            List<float> f = new List<float>();
            foreach (var itm in cellWidths)
            {
                f.Add((float)itm);
            }
            newTable(f.ToArray(), fullWidth: fullWidth);
        }

        public void newTable(float[] cellwidths, bool fullWidth = true, bool Center = false)
        {
            table = new Table(cellwidths);
            if (fullWidth)
                table.SetWidth(UnitValue.CreatePercentValue(100));

            if (Center)
                table.SetHorizontalAlignment(HorizontalAlignment.CENTER);
        }

        public void addCell(byte[] image, float scalePercent = 1.0f)
        {
            var c = new Cell();
            c.SetKeepTogether(true);
            c.SetBorder(Border.NO_BORDER);
            iText.Layout.Element.Image img = new iText.Layout.Element.Image(ImageDataFactory.Create(image));
            img.Scale(scalePercent, scalePercent);
            c.Add(img);
            table.AddCell(c);
        }

        public void addCell(string txt, string FontName = "", float FontSize = 10, int LeadingSize = -1, bool Bold = false, bool Italic = false, bool Underline = false, bool KeepWithNext = false,
            bool Center = false, bool RightAlign = false, bool AlignBottom = false, string shade = "", float cellWidth = 0, int colspan = 0, int rowspan = 0, string fontColor = "",
            string vertAlign = "", float cellHeight = 0, string leftBorder = "", string topBorder = "", string rightBorder = "", string bottomBorder = "")
        {
            var chunks = new List<chunks>();
            chunks.Add(new chunks
            {
                text = txt,
                bold = Bold,
                italic = Italic,
                underline = Underline,
                fontSize = FontSize,
                fontCol = fontColor,
                fontFace = FontName
            });

            addCell(chunks, LeadingSize, KeepWithNext, Center, RightAlign, AlignBottom, shade, cellWidth, colspan, rowspan, vertAlign, cellHeight, 
                leftBorder, topBorder, rightBorder, bottomBorder);
        }


        public void addCell(List<chunks> chunks, int LeadingSize = -1, bool KeepWithNext = false,
        bool Center = false, bool RightAlign = false, bool AlignBottom = false, string shade = "", float cellWidth = 0, int colspan = 0, int rowspan = 0,
        string vertAlign = "", float cellHeight = 0, string leftBorder = "", string topBorder = "", string rightBorder = "", string bottomBorder = "")
        {

            if (document == null) newDocument();

            var para = new Paragraph();
            foreach (var cnk in chunks)
            {
                if (cnk.fontFace != "" && cnk.text != "")
                    setFont(cnk.fontFace);
                else if (cnk.text != "")
                    setFont("Times New Roman");


                var c = new Text(cnk.text);

                c.SetFontSize(cnk.fontSize);

                if (cnk.bold) c.SetBold();
                if (cnk.italic) c.SetItalic();
                if (cnk.underline) c.SetUnderline();


                if (cnk.fontCol != "")
                {
                    var fontCol = WebColors.GetRGBColor("#" + cnk.fontCol);
                    c.SetFontColor(fontCol);
                }

                if (cnk.fontFace!= "")
                    c.SetFont(getFont(cnk.fontFace));

                para.Add("\u0000").Add(c);
            }

            if (Center)
                para.SetTextAlignment(TextAlignment.CENTER);

            if (RightAlign)
                para.SetTextAlignment(TextAlignment.RIGHT);
            
            para.SetKeepTogether(KeepWithNext);
            para.SetKeepWithNext(KeepWithNext);

            if (LeadingSize == -1)
                para.SetFixedLeading(chunks.Max(c=>c.fontSize));
            else
                para.SetFixedLeading(LeadingSize);

            Cell cl = new Cell();
            if (colspan != 0 || rowspan != 0)
                cl = new Cell((rowspan == 0 ? 1 : rowspan), (colspan == 0 ? 1 : colspan));

            cl.Add(para);
            if (AlignBottom)
                cl.SetVerticalAlignment(VerticalAlignment.BOTTOM);
            else
                cl.SetVerticalAlignment(VerticalAlignment.TOP);

            if ((shade ?? "") != "")
            {
                var shadeCol = WebColors.GetRGBColor("#" + shade);
                cl.SetBackgroundColor(shadeCol);
            }

            if (cellWidth != 0)
            {
                cl.SetMinWidth(cellWidth);
                cl.SetMaxWidth(cellWidth);
            }

            if (vertAlign.Trim().ToLower() == "top")
                cl.SetVerticalAlignment(VerticalAlignment.TOP);
            if (vertAlign.Trim().ToLower() == "center")
                cl.SetVerticalAlignment(VerticalAlignment.MIDDLE);
            if (vertAlign.Trim().ToLower() == "bottom")
                cl.SetVerticalAlignment(VerticalAlignment.BOTTOM);


            if (leftBorder != "" || topBorder != "" || rightBorder != "" || bottomBorder != "")
            {
                var db = new DashedBorder(ColorConstants.BLACK, 0.5f);
                var sb = new SolidBorder(ColorConstants.BLACK, 1);
                var bb = new DoubleBorder(ColorConstants.BLACK, 1);
                cl.SetBorder(Border.NO_BORDER);
                if ((leftBorder ?? "").ToLower() == "thin")
                    cl.SetBorderLeft(sb);
                if ((leftBorder ?? "").ToLower() == "dotted")
                    cl.SetBorderLeft(db);
                if ((leftBorder ?? "").ToLower() == "double")
                    cl.SetBorderLeft(bb);

                if ((topBorder ?? "").ToLower() == "thin")
                    cl.SetBorderTop(sb);
                if ((topBorder ?? "").ToLower() == "dotted")
                    cl.SetBorderTop(db);
                if ((topBorder ?? "").ToLower() == "double")
                    cl.SetBorderTop(bb);

                if ((rightBorder ?? "").ToLower() == "thin")
                    cl.SetBorderRight(sb);
                if ((rightBorder ?? "").ToLower() == "dotted")
                    cl.SetBorderRight(db);
                if ((rightBorder ?? "").ToLower() == "double")
                    cl.SetBorderRight(bb);

                if ((bottomBorder ?? "").ToLower() == "thin")
                    cl.SetBorderBottom(sb);
                if ((bottomBorder ?? "").ToLower() == "dotted")
                    cl.SetBorderBottom(db);
                if ((bottomBorder ?? "").ToLower() == "double")
                    cl.SetBorderBottom(bb);
            }

            if (cellHeight != 0)
            {
                cl.SetHeight(cellHeight);
                cl.SetMarginTop(0);
                cl.SetMarginBottom(0);
                cl.SetPadding(0);
            }



            table.AddCell(cl);
        }

        public void benPacketCell(byte[] Thumbnail, string header, string detail, string url, string documentType)
        {
            var font = PdfFontFactory.CreateFont(FontConstants.HELVETICA);

            if (document == null) throw new Exception("No document set.  Unable to process.");
            if (header == "" && detail == "")
            {
                addCell("", colspan: 2);
            }
            else
            {
                //add the thumbnail
                addCell(Thumbnail, scalePercent: .25f);

                //header
                var para = new Paragraph(header);
                var fontCol = WebColors.GetRGBColor("#006a4e");
                para.SetFontColor(fontCol);
                para.SetFontSize(14);
                //para.SetBold();
                para.SetFont(font);
                para.SetFixedLeading(14f);

                para.SetKeepTogether(true);
                para.SetKeepWithNext(true);

                Cell cl = new Cell();
                cl.Add(para);

                //detail
                para = new Paragraph();
                para.SetFont(font);
                para.Add(detail);
                para.SetFontSize(10);
                para.SetFixedLeading(10f);
                cl.Add(para);

                var relativePath = "~/img/packets/" + documentType + ".jpg";
                var absolutePath = HttpContext.Current.Server.MapPath(relativePath);
                var p = new Paragraph();
                if (!File.Exists(absolutePath))
                {
                    relativePath = "~/img/packets/default.png";
                    absolutePath = HttpContext.Current.Server.MapPath(relativePath);
                }

                var img = System.Drawing.Image.FromFile(absolutePath);

                using (var imgStream = new MemoryStream())
                {

                    img.Save(imgStream, System.Drawing.Imaging.ImageFormat.Png);
                    imgStream.ToArray();
                    iText.Layout.Element.Image btn = new iText.Layout.Element.Image(ImageDataFactory.Create(imgStream.ToArray()));
                    btn.Scale(.5f, .5f);

                    para = new Paragraph();
                    para.Add(btn);
                    para.SetAction(PdfAction.CreateURI(url));
                    para.SetFixedLeading(40);

                    para.SetVerticalAlignment(VerticalAlignment.BOTTOM);
                    cl.Add(para);
                }

                cl.SetBorder(Border.NO_BORDER);

                cl.SetPaddingBottom(5);

                table.AddCell(cl);
            }
        }


        public void finishTable()
        {
            document.Add(table);
        }

        public void newPage(bool landscape = false)
        {

            if (landscape)
                document.Add(new AreaBreak(PageSize.LETTER.Rotate()));
            else
                document.Add(new AreaBreak(PageSize.LETTER));
        }

        public byte[] getDocument()
        {
            if (document == null) throw new Exception("You do not appear to have created a document");
            document.Close();
            return msOutput.ToArray();
        }

        public byte[] getDocument(bool evenPages)
        {
            return getDocument(footerLeft: "", footerCenter: "", footerRight: "", evenPages: evenPages);
        }

        public byte[] getDocument(string footerLeft, string footerCenter, string footerRight, bool evenPages = false, bool showLogo = false, string footerFont = "Times")
        {
            if (document == null) throw new Exception("You do not appear to have created a document");
            document.Close();

            PdfFont font;
            if (footerFont.ToLower().Contains("roman"))
                font = PdfFontFactory.CreateFont(FontConstants.TIMES_ROMAN);
            else
                font = PdfFontFactory.CreateFont(Properties.Resources.calibri, true);

            footerLeft = (footerLeft ?? "").Split('"').Last();
            footerCenter = (footerCenter ?? "").Split('"').Last();
            footerRight = (footerRight ?? "").Split('"').Last();

            var docBytes = msOutput.ToArray();
            msOutput.Dispose();
            msOutput = new MemoryStream();

            var msInput = new MemoryStream(docBytes);
            PdfDocument pageDoc = new PdfDocument(new PdfReader(msInput), new PdfWriter(msOutput));
            Document doc = new Document(pageDoc);
            int numPages = pageDoc.GetNumberOfPages();
            for (int i = 1; i <= numPages; i++)
            {
                var lTxt = footerLeft.Replace("&P", i.ToString()).Replace("&N", numPages.ToString());
                var cTxt = footerCenter.Replace("&P", i.ToString()).Replace("&N", numPages.ToString());
                var rTxt = footerRight.Replace("&P", i.ToString()).Replace("&N", numPages.ToString());

                doc.ShowTextAligned(p: new Paragraph(""), x: 0, y: 0, pageNumber: i, textAlign: TextAlignment.LEFT, vertAlign: VerticalAlignment.BOTTOM, radAngle: 0).SetFontSize(8);
                if ((footerLeft ?? "") != "")
                {
                    var lp = new Paragraph(lTxt);
                    lp.SetFont(font);
                    doc.ShowTextAligned(p: lp, x: 15, y: 12, pageNumber: i, textAlign: TextAlignment.LEFT, vertAlign: VerticalAlignment.BOTTOM, radAngle: 0).SetFontSize(8);
                }
                if ((footerCenter ?? "") != "")
                {
                    var cp = new Paragraph(cTxt);
                    cp.SetFont(font);
                    doc.ShowTextAligned(p: cp, x: pageDoc.GetPage(i).GetCropBox().GetWidth() / 2, y: 12, pageNumber: i, textAlign: TextAlignment.CENTER, vertAlign: VerticalAlignment.BOTTOM, radAngle: 0).SetFontSize(8);
                }
                if ((footerRight ?? "") != "")
                {
                    var rp = new Paragraph(rTxt);
                    rp.SetFont(font);
                    doc.ShowTextAligned(p: rp, x: pageDoc.GetPage(i).GetCropBox().GetWidth() - 20, y: 12, pageNumber: i, textAlign: TextAlignment.RIGHT, vertAlign: VerticalAlignment.BOTTOM, radAngle: 0).SetFontSize(8);
                }

                if (footerLeft != "")
                {
                    var lp = new Paragraph(lTxt);
                    lp.SetFont(font);
                    doc.ShowTextAligned(p: lp, x: 15, y: 12, pageNumber: i, textAlign: TextAlignment.LEFT, vertAlign: VerticalAlignment.BOTTOM, radAngle: 0).SetFontSize(8);
                }

                if (showLogo)
                {
                    using (var imgStream = new MemoryStream())
                    {
                        Properties.Resources.ht_logo.Save(imgStream, System.Drawing.Imaging.ImageFormat.Png);
                        imgStream.ToArray();
                        iText.Layout.Element.Image logo = new iText.Layout.Element.Image(ImageDataFactory.Create(imgStream.ToArray()));
                        logo.Scale(0.15F, 0.15F);
                        var c = new Paragraph();
                        c.Add(logo);
                        c.SetFixedPosition(i, 20, (float)pageDoc.GetPage(i).GetCropBox().GetHeight() - (float)logo.GetImageScaledHeight() - 20, logo.GetImageScaledWidth());
                        doc.Add(c);
                    }
                }

            }

            if (numPages % 2 == 1 && evenPages)
            {
                pageDoc.AddNewPage(pageDoc.GetDefaultPageSize());
            }


            doc.Close();
            pageDoc.Close();
            return msOutput.ToArray();

        }

        public void fromHTML(string HTML)
        {
            if (document == null)
                newDocument(suppressLogo: true);

            var docIn = new HtmlDocument();
            HTML.Replace("\r", "").Replace("\n", "");
            docIn.LoadHtml(HTML);
            var paras = new List<string>();

            processHTMLNode(ref paras, docIn.DocumentNode);
        }

        public void processHTMLNode(ref List<string> paragraphs, HtmlNode node, bool bold = false, bool italic = false)
        {
            if (node.HasChildNodes)
            {
                foreach (var n in node.ChildNodes)
                {
                    if (node.Name != "style" && node.Name != "a")
                        processHTMLNode(ref paragraphs, n, bold: bold || node.Name == "strong", italic: italic || node.Name == "em");
                }
                if (node.Name == "hr")
                    separator();

                if ((node.Name == "p" || node.Name == "div") && paragraphs.Count != 0)
                {
                    int fontSz = 10;
                    if (node.Name == "div") fontSz = 9;
                    paragraph(string.Join("\n", paragraphs), FontSize: fontSz, LeadingSize: (fontSz + 2), italicMarker: "~i~", boldMarker: "~b~", emphasisMarker: "~e~");
                    paragraphs = new List<string>();
                }


                if ((node.Name == "h1"))
                {
                    int fontSz = 12;
                    paragraph(string.Join("\n", paragraphs), FontSize: fontSz, LeadingSize: (fontSz + 2), Bold: true);
                    paragraphs = new List<string>();
                }
                if ((node.Name == "h2"))
                {
                    int fontSz = 11;
                    paragraph(string.Join("\n", paragraphs), FontSize: fontSz, LeadingSize: (fontSz + 2), Bold: true);
                    paragraphs = new List<string>();
                }
            }
            else
            {
                if (node.Name == "hr")
                    separator();

                var strOut = "";
                if (bold && italic) strOut += "~e~";
                else if (bold) strOut += "~b~";
                else if (italic) strOut += "~i~";

                if (node.Name == "br")
                {
                    strOut += Environment.NewLine;
                }
                else if (node.Name == "#text")
                {
                    if (node.InnerText.ToLower().Contains("[signature]"))
                    {
                        paragraphs = new List<string>();
                        var me = HttpContext.Current.User.Identity.Name.Split('\\');
                        var myName = me[me.Length - 1].ToLower();

                        var relativePath = "~/img/sig/" + myName + ".png";
                        var absolutePath = HttpContext.Current.Server.MapPath(relativePath);
                        var p = new Paragraph();
                        if (!File.Exists(absolutePath))
                        {
                            relativePath = "~/img/sig/default.png";                            
                            absolutePath = HttpContext.Current.Server.MapPath(relativePath);
                        }

                        var img = System.Drawing.Image.FromFile(absolutePath);
                        var originalW = img.Width;
                        var origH = img.Height;

                        using (var imgStream = new MemoryStream())
                        {

                            img.Save(imgStream, ImageFormat.Png);
                            imgStream.ToArray();
                            iText.Layout.Element.Image sig = new iText.Layout.Element.Image(ImageDataFactory.Create(imgStream.ToArray()));
                            sig.Scale(.75F, .75F);
                            
                            document.Add(sig);
                        }
                    }
                    else
                    {
                        strOut += node.InnerText.Replace("&nbsp;", " ").Replace("&#39;", "'").Replace("\n", "").Replace("\r", "");
                    }
                }
                if (strOut != "")
                    paragraphs.Add(strOut);
            }
        }

        internal static System.Drawing.Image ScaleBySize(System.Drawing.Image imgPhoto, int Width, int Height)
        {
            float sourceWidth = imgPhoto.Width;
            float sourceHeight = imgPhoto.Height;
            int sourceX = 0;
            int sourceY = 0;
            int destX = 0;
            int destY = 0;

            Bitmap bmPhoto = new Bitmap((int)Width, (int)Height,
                                        PixelFormat.Format32bppPArgb);
            bmPhoto.SetResolution(imgPhoto.HorizontalResolution, imgPhoto.VerticalResolution);

            Graphics grPhoto = Graphics.FromImage(bmPhoto);
            grPhoto.CompositingMode = CompositingMode.SourceCopy;
            grPhoto.PixelOffsetMode = PixelOffsetMode.Half;
            grPhoto.InterpolationMode = InterpolationMode.HighQualityBicubic;
            grPhoto.PixelOffsetMode = PixelOffsetMode.HighQuality;

            //using (var ia = new ImageAttributes())
            //{
            //                ia.SetWrapMode(WrapMode.TileFlipXY);
            grPhoto.DrawImage(imgPhoto,
                new System.Drawing.Rectangle(destX, destY, (int)Width, (int)Height),
                new System.Drawing.Rectangle(sourceX, sourceY, (int)sourceWidth, (int)sourceHeight),
                GraphicsUnit.Pixel);

            grPhoto.Dispose();

            return bmPhoto;
        }

        public void paragraph(string paraText, int FontSize = 10, int LeadingSize = -1, bool Bold = false, bool Italic = false, bool Underline = false, bool KeepWithNext = false, string italicMarker = "", string boldMarker = "", string emphasisMarker = "", bool Center = false, bool RightAlign = false, string fontColor = "")
        {

            if (document == null) newDocument();
            var txt = paraText.Split('\n');
            var para = new Paragraph();

            foreach (var itm in txt)
            {
                var c = new Text(itm);
                if (italicMarker != "" && itm.StartsWith(italicMarker))
                {
                    c = new Text(itm.Substring(italicMarker.Length));
                    c.SetItalic();
                }

                if (boldMarker != "" && itm.StartsWith(boldMarker))
                {
                    c = new Text(itm.Substring(boldMarker.Length));
                    c.SetBold();
                }

                if (emphasisMarker != "" && itm.StartsWith(emphasisMarker))
                {
                    c = new Text(itm.Substring(emphasisMarker.Length));
                    c.SetItalic().SetBold();
                }
                para.Add("\u0000").Add(c);
            }

            if (Center)
                para.SetTextAlignment(TextAlignment.CENTER);

            if (RightAlign)
                para.SetTextAlignment(TextAlignment.RIGHT);


            //var para = new Paragraph(paraText + "\n");
            para.SetFontSize(FontSize);
            if(fontColor != "")
            {
                var fontCol = WebColors.GetRGBColor("#" + fontColor);
                para.SetFontColor(fontCol);
            }
            if(Bold) para.SetBold();
            if(Italic) para.SetItalic();
            if (Underline) para.SetUnderline();

            para.SetKeepTogether(KeepWithNext);
            para.SetKeepWithNext(KeepWithNext);

            if (LeadingSize == -1)
                para.SetFixedLeading(FontSize);
            else
                para.SetFixedLeading(LeadingSize);

            document.Add(para);
        }

        public void setPageFooter(string leftTxt, string centerTxt, string rightTxt)
        {
            document.ShowTextAligned("", 0, 0, TextAlignment.LEFT, VerticalAlignment.BOTTOM, 0).SetFontSize(8);

            for (var i = 1; i <= pdf.GetNumberOfPages(); i++)
            {
                
                if (leftTxt != "")
                    document.ShowTextAligned(leftTxt, 15, 20, TextAlignment.LEFT, VerticalAlignment.BOTTOM, 0).SetFontSize(8);

                if (centerTxt != "")
                    document.ShowTextAligned(centerTxt, 300, 20, TextAlignment.CENTER, VerticalAlignment.BOTTOM, 0).SetFontSize(8);

                if (rightTxt != "")
                    document.ShowTextAligned(rightTxt, 600, 20, TextAlignment.RIGHT, VerticalAlignment.BOTTOM, 0).SetFontSize(8);

                if (leftTxt != "")
                    document.ShowTextAligned(leftTxt, 15, 20, TextAlignment.LEFT, VerticalAlignment.BOTTOM, 0).SetFontSize(8);
            }
        }

        /// <summary>
        /// Appends a given PDF to the current PDF Document object (if none is set will create one for you)
        /// </summary>
        /// <param name="docToAppend">New Document to extract pages from</param>
        /// <param name="pagesToAppend">(Optional) List of pages to append (if NULL will append all)</param>
        /// <param name="padToEvenPages">Pad the output to an even number of pages</param>
        public void appendDocument(byte[] docToAppend, List<int> pagesToAppend = null, bool padToEvenPages = false)
        {
            //try
            //{
            if (msOutput == null)
                msOutput = new MemoryStream();
            if (writer == null)
                writer = new PdfWriter(msOutput);
            if (pdf == null)
                pdf = new PdfDocument(writer);
            if (document == null)
                document = new Document(pdf);

            //Let's append the new document data
            var msPrint = new MemoryStream(docToAppend);
            var src = new PdfReader(msPrint);

            var srcDoc = new PdfDocument(src);
            if (pagesToAppend == null)
            {
                pagesToAppend = new List<int>();
                for (var i = 1; i <= srcDoc.GetNumberOfPages(); i++)
                {
                    pagesToAppend.Add(i);
                }
            }
            srcDoc.CopyPagesTo(pagesToAppend, pdf);

            if (srcDoc.GetNumberOfPages() % 2 == 1 && padToEvenPages)
            {
                pdf.AddNewPage(new PageSize(srcDoc.GetFirstPage().GetPageSizeWithRotation()));
            }

            //srcDoc.Close();
            //srcDoc = null;
            src.Close();
            //src = null;
            //msRead.Dispose();
            //msRead = null;
            //}
            //catch(Exception e)
            //{
            //    //I don't know what to do here - hopefully your document is still OK.
            //}
        }

        /// <summary>
        /// Extracts page text from given PDF File
        /// </summary>
        /// <param name="pdfFile">File to process</param>
        /// <param name="pageNumber">Page to extract</param>
        /// <returns></returns>
        public List<string> getPageText(byte[] pdfFile, int pageNumber)
        {
            var filein = new MemoryStream(pdfFile);

            PdfReader reader = new PdfReader(filein);

            PdfDocument doc = new PdfDocument(reader);
            ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
            return PdfTextExtractor.GetTextFromPage(doc.GetPage(pageNumber)).Split('\n').ToList();
        }

        /// <summary>
        /// Extracts all page text from the given PDF
        /// </summary>
        /// <param name="pdfFile">File to process</param>
        /// <returns></returns>
        public List<string> getPageText(byte[] pdfFile)
        {
            List<string> output = new List<string>();
            var filein = new MemoryStream(pdfFile);

            PdfReader reader = new PdfReader(filein);

            PdfDocument doc = new PdfDocument(reader);
            ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
            for (var i = 1; i <= doc.GetNumberOfPages(); i++)
            {
                output.Add(PdfTextExtractor.GetTextFromPage(doc.GetPage(i)));
            }
            return output;
        }

        /// <summary>
        /// Given the standard HealthTrust address format (*blank line*blank line*date*blank line*name*title*member*address1*address2*citystatezip*blank line*Dear...:) will attempt to extract address date
        /// </summary>
        /// <param name="pdfFile">File to extract data from</param>
        /// <returns></returns>
        public List<PDFAddressInfo> extractAddressInfo(byte[] pdfFile)
        {
            var output = new List<PDFAddressInfo>();
            var filein = new MemoryStream(pdfFile);

            PdfReader reader = new PdfReader(filein);

            PdfDocument doc = new PdfDocument(reader);
            ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
            for (var i = 1; i <= doc.GetNumberOfPages(); i++)
            {
                var pageTxt = PdfTextExtractor.GetTextFromPage(doc.GetPage(i));

                bool hasDate = false;
                var addBlock = new List<string>();
                foreach (var ln in pageTxt.Split('\n'))
                {
                    if (ln.ToLower().StartsWith("dear") && ln.Trim().EndsWith(":") && hasDate && addBlock.Count != 0)
                    {
                        //Log the output since this is apparently a real address page
                        if (addBlock.Count == 5)
                        {
                            var adTxt = addBlock.ToArray();
                            output.Add(new PDFAddressInfo { PageNo = i, AttnName = adTxt[0], Title = adTxt[1], MemberName = adTxt[2].Replace(Convert.ToChar(160).ToString(), " ").Replace(Convert.ToChar(8208).ToString(), "-"), Address1 = adTxt[3].Replace(Convert.ToChar(160).ToString(), " ").Replace(Convert.ToChar(8208).ToString(), "-"), CityStZIP = adTxt[4].Replace(Convert.ToChar(160).ToString(), " ").Replace(Convert.ToChar(8208).ToString(), "-") });
                            break;
                        }
                        if (addBlock.Count == 6)
                        {
                            var adTxt = addBlock.ToArray();
                            output.Add(new PDFAddressInfo { PageNo = i, AttnName = adTxt[0].Replace(Convert.ToChar(160).ToString(), " ").Replace(Convert.ToChar(8208).ToString(), "-"), Title = adTxt[1].Replace(Convert.ToChar(160).ToString(), " ").Replace(Convert.ToChar(8208).ToString(), "-"), MemberName = adTxt[2].Replace(Convert.ToChar(160).ToString(), " "), Address1 = adTxt[3].Replace(Convert.ToChar(160).ToString(), " ").Replace(Convert.ToChar(8208).ToString(), "-"), Address2 = adTxt[4].Replace(Convert.ToChar(160).ToString(), " ").Replace(Convert.ToChar(8208).ToString(), "-"), CityStZIP = adTxt[5].Replace(Convert.ToChar(160).ToString(), " ").Replace(Convert.ToChar(8208).ToString(), "-") });
                            break;
                        }
                    }
                    else if (ln.Trim() != "")
                    {
                        var dt = new DateTime();
                        //Check to see if the line can be parsed as a date
                        if (DateTime.TryParse(ln, out dt))
                            hasDate = true;
                        else
                        {
                            addBlock.Add(ln);
                        }
                    }
                }
            }

            foreach (var l in output)
            {
                var pgFrom = l.PageNo;
                int pgTo = 0;
                try
                {
                    pgTo = output.Where(x => x.PageNo > l.PageNo).Min(x => x.PageNo) - 1;
                }
                catch
                {
                    pgTo = doc.GetNumberOfPages();
                }

                //Clear any existing document
                msOutput = new MemoryStream();
                writer = new PdfWriter(msOutput);
                pdf = new PdfDocument(writer);
                document = new Document(pdf);

                //add the pages from this letter to a new PDF document and stitch that onto the output object
                List<int> pages = new List<int>();
                for (var i = pgFrom; i <= pgTo; i++)
                {
                    pages.Add(i);
                }
                var appender = new PDFProcessing();
                appender.appendDocument(pdfFile, pages);
                l.docData = appender.getDocument();
                appender.Cleanup();
                appender = null;
            }

            return output;
        }
        #endregion
    }
}
