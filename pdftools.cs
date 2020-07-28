using System;
// For the  Server integration
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.IO;


namespace pdfTools
{
    class pdftools
    {
        static void Main(string[] args)
        {

            String err = "";
            if (args.Length == 0)
            {
                Console.WriteLine("Please choose function mode: setpwd, setwatermark or split");
            }
            else if (args[0].ToLower() == "setpwd")
            {
                if (args.Length < 4)
                {
                    Console.WriteLine("Please complete input filename, output filename and password");
                }
                else
                {
                    SetPwd(args[1], args[2], args[3]);
                }
            }
            else if (args[0].ToLower() == "setwatermark")
            {
                if (args.Length < 4)
                {
                    Console.WriteLine("Please complete input filename, output filename and watermark text.\n\nOptions: \n-fa:x x=0:left;1:center;2:right 1=default\n-fs:x x=font size 20=default\n-fc:x x=font color BLACK;GRAY(default);RED;BLUE;YELLOW;GREEN\n-rd:x x=rotation 0 to 360 0=default\n-of:x x=opacity fill 0.0 to 1.0 1.0=solid 0.35=default\n-os:x x=opacity stroke 0.0 to 1.0 1.0=solid 0.8=default\n-xp:x x=coordinate 0.0 to 1.0 0.0=left 0.5=default\n-yp:x 0.0 to 1.0 0.0=top 0.5=default\n");
                }
                else
                {
                    BaseColor color = BaseColor.GRAY;
                    int align = Element.ALIGN_CENTER;
                    int rotationD = 45;
                    int fontSize = 20;
                    float opFill = 0.35F;
                    float opStroke = 0.8F;
                    float xP = 0.5F;
                    float yP = 0.5F;
                    if (args.Length >= 4)
                    {
                        for (int i = 4; i < args.Length; i++)
                        {
                            String c = "";
                            if (args[i].IndexOf("-fc:") >= 0)
                            {
                                c = args[i].Split(':')[1];
                                if (c.ToUpper() == "BLACK") color = BaseColor.BLACK;
                                else if (c.ToUpper() == "GRAY") color = BaseColor.GRAY;
                                else if (c.ToUpper() == "RED") color = BaseColor.RED;
                                else if (c.ToUpper() == "BLUE") color = BaseColor.BLUE;
                                else if (c.ToUpper() == "GREEN") color = BaseColor.GREEN;
                                else if (c.ToUpper() == "YELLOW") color = BaseColor.YELLOW;
                                //Console.WriteLine(c);

                            }
                            if (args[i].IndexOf("-fa:") >= 0) align = int.Parse(args[i].Split(':')[1]);
                            if (args[i].IndexOf("-fs:") >= 0) fontSize = int.Parse(args[i].Split(':')[1]);
                            if (args[i].IndexOf("-rd:") >= 0) rotationD = int.Parse(args[i].Split(':')[1]);
                            if (args[i].IndexOf("-of:") >= 0) opFill = float.Parse(args[i].Split(':')[1]);
                            if (args[i].IndexOf("-os:") >= 0) opStroke = float.Parse(args[i].Split(':')[1]);
                            if (args[i].IndexOf("-xp:") >= 0) xP = float.Parse(args[i].Split(':')[1]);
                            if (args[i].IndexOf("-yp:") >= 0) yP = float.Parse(args[i].Split(':')[1]);
                        }
                    }
                    String txt = args[3];//.Replace("*", "\n");
                    //Console.Write(txt);
                    SetWatermark(args[1], args[2], txt, color, align, rotationD, fontSize, opFill, opStroke, xP, yP);

                }
            }

            else if (args[0].ToLower() == "split")
            {
                if (args.Length < 3)
                {
                    Console.WriteLine("Please complete input filename, output folder and password (optional)");
                }
                else
                {
                    err = Split(args[1], args[2], args[3]);
                    if (err != "") Console.WriteLine(err);
                }
            }
        }

        public static String Split(String input, String outputFolder, String pwdText)
        {
            String r = "";
            if (File.Exists(input))
            {
                FileInfo file = new FileInfo(input.ToString());
                string name = file.Name.Substring(0, file.Name.LastIndexOf("."));
                string path = file.DirectoryName;

                using (PdfReader reader = new PdfReader(input.ToString()))
                {

                    for (int pagenumber = 1; pagenumber <= reader.NumberOfPages; pagenumber++)
                    {
                        if (!Directory.Exists(path + '\\' + outputFolder.ToString()))
                        {
                            Console.Write("Creating folder:" + path + '\\' + outputFolder.ToString());
                            Directory.CreateDirectory(path + '\\' + outputFolder.ToString());
                        }
                        string filename = path + '\\' + outputFolder.ToString() + '\\' + name + '_' + pagenumber.ToString() + ".pdf";
                        string filenameTemp = filename;

                        if (pwdText.ToString() != "")
                            filenameTemp = path + '\\' + outputFolder.ToString() + '\\' + name + '_' + pagenumber.ToString() + "_temp.pdf";
                        if (!Directory.Exists(path + '\\' + outputFolder.ToString()))
                            Directory.CreateDirectory(path + '\\' + outputFolder.ToString());

                        Document document = new Document();
                        PdfCopy copy = new PdfCopy(document, new FileStream(filenameTemp, FileMode.Create));

                        document.Open();

                        copy.AddPage(copy.GetImportedPage(reader, pagenumber));
                        document.Close();
                        if (pwdText.ToString() != "")
                        {
                            SetPwd(filenameTemp, filename, pwdText.ToString());
                            File.Delete(filenameTemp);
                        }

                    }
                }
            }
            else
            {
                r = "File not found";
            }
            return r;
        }



        public static Boolean SetPwd(String inputFile, String outputFile, String pwd)
        {
            Boolean r = false;
            string WorkingFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string InputFile = inputFile.ToString();
            string OutputFile = outputFile.ToString();

            using (Stream input = new FileStream(InputFile, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                using (Stream output = new FileStream(OutputFile, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    PdfReader reader = new PdfReader(input);
                    PdfEncryptor.Encrypt(reader, output, true, pwd.ToString(), pwd.ToString(), PdfWriter.ALLOW_SCREENREADERS);
                }
            }

            return r;
        }

        public static Boolean SetWatermark(String inputFile, String outputFile, String watermarkText, BaseColor color, int align, int rotationD, int fontSize, float opFill, float opStroke, float xP, float yP)
        {
            Boolean r = false;
            var bytes = File.ReadAllBytes(inputFile.ToString());
            byte[] rbytes = AddWatermark(bytes, BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, false), watermarkText.ToString(), color, align, rotationD, fontSize, opFill, opStroke, xP, yP);

            File.WriteAllBytes(outputFile.ToString(), rbytes);

            r = true;
            return r;
        }

        private static byte[] AddWatermark(byte[] bytes, BaseFont baseFont, string watermarkText, BaseColor color, int align, int rotationD = 45, int fontSize = 50, float opFill = 0.35F, float opStroke = 0.3F, float xP = 0.5F, float yP = 0.5F)
        {
            using (var ms = new MemoryStream(10 * 1024))
            {
                using (var reader = new PdfReader(bytes))
                using (var stamper = new PdfStamper(reader, ms))
                {
                    var pages = reader.NumberOfPages;
                    for (var i = 1; i <= pages; i++)
                    {
                        var dc = stamper.GetOverContent(i);
                        AddWaterMarkText(dc, watermarkText, baseFont, fontSize, rotationD, color,
                            reader.GetPageSizeWithRotation(i), opFill, opStroke, xP, yP, align
                            );
                    }
                    stamper.Close();
                }
                return ms.ToArray();
            }
        }
        private static void AddWaterMarkText(PdfContentByte pdfData, string watermarkText, BaseFont font, float fontSize, float angle, BaseColor color, iTextSharp.text.Rectangle realPageSize, float opFill, float opStroke, float xP, float yP, int align)
        {
            //var gstate = new PdfGState { FillOpacity = 0.35f, StrokeOpacity = 0.3f };
            var gstate = new PdfGState { FillOpacity = opFill, StrokeOpacity = opStroke };
            pdfData.SaveState();
            pdfData.SetGState(gstate);
            pdfData.SetColorFill(color);
            pdfData.BeginText();
            pdfData.SetFontAndSize(font, fontSize);
            //var x = (realPageSize.Right + realPageSize.Left) / 2;
            //var y = (realPageSize.Bottom + realPageSize.Top) / 2;
            var x = realPageSize.Right * (xP);
            var y = realPageSize.Top * (1-yP);
            //pdfData.ShowTextAligned(Element.ALIGN_LEFT, watermarkText, x, y, angle);
            pdfData.ShowTextAligned(align, watermarkText, x, y, angle);
            pdfData.EndText();
            pdfData.RestoreState();
        }

    }
}
