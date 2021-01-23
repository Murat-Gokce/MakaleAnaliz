using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using MakaleAnalizWebApp.Models;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;

namespace MakaleAnalizWebApp.Service
{
    public class Analiysis
    {
        public Analiysis(string path)
        {
            pdfPath1 = path;
        }
        string pdfPath1 = "";
       public List<Result> results = new List<Result>();
        List<string> referances = new List<string>();
        List<string> figures = new List<string>();
        List<string> tables = new List<string>();
        string saltText = "";
        string x;
        public string ReadFile(string pdfPath)
        {
            var text = new StringBuilder();
            if (pdfPath.Contains(".pdf"))
            {

                PdfReader reader = new PdfReader(pdfPath);
                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                    string currentText = PdfTextExtractor.GetTextFromPage(reader, i, strategy);

                    currentText = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(currentText)));
                    text.Append(currentText);

                }

            }
            else
            {
                Application word = new Application();
                Microsoft.Office.Interop.Word.Document doc = new Microsoft.Office.Interop.Word.Document();

                object fileName = pdfPath;
                // Define an object to pass to the API for missing parameters
                object missing = System.Type.Missing;
                doc = word.Documents.Open(ref fileName,
                        ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing);

                String read = string.Empty;
                read = doc.Content.Text;
                //for (int i =  0; i < doc.Paragraphs.Count; i++)
                //{
                //    string temp = doc.Paragraphs[i + 1].Range.Text.Trim();
                //    if (temp != string.Empty)
                //        text.Append(temp);
                //}
                text.Append(read);
((_Document)doc).Close();
                ((_Application)word).Quit();

            }

            return text.ToString();
        }
        void addLog(string message, bool isSuccess)
        {
            this.results.Add(new Result
            {
                id = results.Count + 1,
                message = message,
                isSuccess = isSuccess
            });
        }

        public void checkFile()
        {
            var text = ReadFile(pdfPath1);
            if (text.Length > 0)
                addLog("Makale Başarı İle okundu", true);
            saltText = text;
            checkReferences(text);
            checkFigure();
            checkTable();
            checkQuotes();
        }
        //kaynaklarin kontrolu
        public void checkReferences(string text)
        {
            //kaynakca sayısı
            int count = 0, nextCount = 0, index = 1;
            bool isValid = true;
            do
            {
                count = text.LastIndexOf("[" + index + "]");
                
                if (count > 0)
                {

                    text = text.Substring(count);
                    index++;
                }
                //nextCount>0 ise son kayit degildir
                if(!checkReferencesIsValid(text, index - 1, nextCount))
                {
                    addLog("[" + index + "]. Kaynak Uygun formatta değil.(Makale içerisinde kullanılmamış)",false );
                }

            } while (count > 0);
            if (index > 1)
            {
                addLog(index - 1 + " Adet Kaynak Bulundu.", true);

            }
            else
            {
                addLog("Kaynaklar Bölümü bulunamadı.", false);
            }

        }
        public bool checkReferencesIsValid(string text, int index, int lastCount)
        {
            int count = 0;
            bool isValid = true;
            text = text.Substring(text.IndexOf(']') + 1);
            text = text.Trim();
            if (lastCount > 0)
            {
                int a = text.IndexOf("[" + (index + 1) + "]");
                if (a > 0)
                    text = text.Substring(0, text.IndexOf("[" + (index + 1) + "]"));
                else
                    text = "";
            }
            else
            {
                int a = text.IndexOf("\n");
                if(a>-1)
                text = text.Substring(0, a);
            }
            referances.Add(text.Trim());

            //referansa atıfta bulunulmus mu?
            isValid = areThereAnyReferences(saltText, index);
            return isValid;

        }
        public int kacdefakullanildi(string text,string aranan)
        {
            // "[" + index.ToString() + "]"

            int index = 0, count = 0;
            do
            {
                index = text.IndexOf(aranan);
                if (index > -1)
                {
                    count++;
                    text = text.Substring(index+2);
                }
                else
                    return count;
            } while (true);
        }
        public bool areThereAnyReferences(string text, int index)
        {
            int count = kacdefakullanildi(text, "[" + index.ToString() + "]");
            if ( count <2   )
            {
                Console.WriteLine(index + "- atıf yapılmayan kaynak");
            }
            return (count > 1);
        }

        public bool checkFigure()
        {
            this.figures = new List<string>();
            bool isValid = false;
            //sekil 1.2 -> sekil a.b
            int a = 1, b = 1, c = 0, x = 0;
            int index = 0;
            do
            {
                do
                {
                    string figure = "Şekil " + a + "." + b + ".";
                    index = saltText.IndexOf(figure);
                    if (index > -1)
                    {
                        b++;
                        c = 0;
                        this.figures.Add(figure);
                        Console.WriteLine(figure);
                    }
                    else
                    {
                        if (x > 3)
                        {
                            x = 0;
                            c++;
                            break;
                        }
                        x++;

                    }
                } while (true);
                if (c > 10)
                    break;
                a++;
                b = 1;

            } while (true);
            addLog(this.figures.Count + " Adet Şekil Bulundu.", true);

            return isValid;
        }
        public bool checkTable()
        {
            this.tables = new List<string>();
            bool isValid = false;
            //sekil 1.2 -> sekil a.b
            int a = 1, b = 1, c = 0, x = 0;
            int index = 0;
            do
            {
                do
                {
                    string table = "Tablo " + a + "." + b + ".";
                    index = saltText.IndexOf(table);
                    if (index > -1)
                    {
                        b++;
                        c = 0;
                        this.tables.Add(table);
                        Console.WriteLine(table);
                    }
                    else
                    {
                        if (x > 3)
                        {
                            x = 0;
                            c++;
                            break;
                        }
                        x++;

                    }
                } while (true);
                if (c > 10)
                    break;
                a++;
                b = 1;

            } while (true);


            addLog(this.tables.Count + " Adet Tablo Bulundu.", true);
            return isValid;
        }

        public bool checkQuotes()
        {
            int quotes = kacdefakullanildi(saltText, "“");
            if (quotes > 50)
            {
                addLog(quotes + " Adet Alıntı bulundu. 50 adetten fazla olamaz",false);
            }
            else
            {
                addLog(quotes + " Adet Alıntı bulundu.", true);


            }
            return (quotes > 50);
        }
        //onsozun ilk paragrafinda tesekkur olmamali
        public bool preface_thank()
        {

            //var preface = this.saltText.Split("ÖNSÖZ")[1];
            return true;
        }
        //public bool checkFigure1()
        //{
        //    bool isValid = false;
        //    //sekil 1.2 -> sekil a.b
        //    int a = 0, b = 1;
        //    int index = 0;
        //    Regex r = new Regex(@"^[Şekil ]* [0-9]+.[0-9]+.$");

        //    string figureList = saltText.Split("ŞEKİLLER LİSTESİ")[2];
        //    index = figureList.IndexOf("Şekil ");
        //    figureList = figureList.Substring(index);
        //    figureList = figureList.Split("\n\n")[0];
        //    string[] lines = figureList.Split("\n");
        //    foreach (var item in lines)
        //    {
        //        index = item.IndexOf("Şekil ");
        //        if (index > -1)
        //        {
        //            if (r.IsMatch(item))
        //                figures.Add(item);
        //            a = 0;
        //        }
        //        else
        //        {
        //            if (a > 10)
        //                break;
        //            a++;

        //        }
        //    }


        //    return isValid;
        //}
    }
}