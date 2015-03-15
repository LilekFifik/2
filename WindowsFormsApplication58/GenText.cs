using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml;
using System.Windows;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using S = DocumentFormat.OpenXml.Spreadsheet.Sheets;
using E = DocumentFormat.OpenXml.OpenXmlElement;
using A = DocumentFormat.OpenXml.OpenXmlAttribute;
using System.Text.RegularExpressions;

namespace WindowsFormsApplication58
{
    class GenText
    {
        public static void GentText(Dictionary<string, string> dict,List<List<string>> List1, string path2, string path3)
        {
            List<string> Tables = new List<string>();
            int flazhok = 0;
            int fcount = 0;
            int k = 0;
            byte[] byteArray = File.ReadAllBytes(path3);
            List<string> M = new List<string>();
            int u = 0;
            int m = 0;

            using (MemoryStream mem = new MemoryStream())
            {
                mem.Write(byteArray, 0, (int)byteArray.Length);




                using (WordprocessingDocument outDoc = WordprocessingDocument.Open(mem, true))
                {


                    var doc = outDoc.MainDocumentPart.Document;

                    MainDocumentPart mainPart = outDoc.MainDocumentPart;


                    DocDefaults defaults = doc.MainDocumentPart.StyleDefinitionsPart.Styles.Descendants<DocDefaults>().FirstOrDefault();

                    RunFonts runFont = defaults.RunPropertiesDefault.RunPropertiesBaseStyle.RunFonts;
                    runFont.Ascii = "Calibri";
                    //runFont.AsciiTheme = "Times New Roman";
                    string font = runFont.Ascii;
                    FontSize fs = new FontSize();
                    fs.Val = "16";
                    string[] tblTag = new string[0];//Табличные теги
                    bool flag = true;
                    for (int f1 = 0; f1 < List1.Count; f1++)
                    {
                        flazhok = 0;
                        Array.Resize(ref tblTag, tblTag.Length + 1);

                        fcount = 0;
                        int i = 0;



                        tblTag[i] = List1[k][fcount].ToString();
                        m = k;
                        fcount++;
                        i++;





                        var ccWithTable1 = mainPart.Document.Body.Descendants<SdtElement>();



                        SdtBlock ccWithTable = mainPart.Document.Body.Descendants<SdtBlock>().FirstOrDefault();

                        int index = 0;




                        foreach (var tt in ccWithTable1)
                        {

                            if (flag != false)
                            {

                                if (tt.SdtProperties.GetFirstChild<Tag>().Val == tblTag[index])
                                {
                                    ccWithTable = mainPart.Document.Body.Descendants<SdtBlock>().Where
                        (r => r.SdtProperties.GetFirstChild<Tag>().Val == tblTag[index]).Single();

                                    Tables.Add(tblTag[index]);
                                    flazhok = 1;
                                    flag = false;


                                }
                            }

                        }
                        int count = 0;
                        int row = 0;
                        int r2 = 3; int n = 0;
                        int u2 = 1;
                    
                        if (flazhok == 1)
                        {
                            Table theTable = ccWithTable.Descendants<Table>().FirstOrDefault();


                            for (row = 4; row < theTable.Elements<TableRow>().Count(); row++)
                            {
                                count = 0; bool p = true;
                                int countp = 0;
                                TableRow row8 = theTable.Elements<TableRow>().ElementAt(row);


                                for (int yacheika2 = 0; yacheika2 < theTable.Elements<TableRow>().ElementAt(row).Elements<TableCell>().Count(); yacheika2++)
                                {
                                    if (row == 4)
                                    {


                                        TableCell cell1 = row8.Elements<TableCell>().ElementAt(yacheika2);
                                        int b3 = 0;
                                        int[] gridSpan1 = null;
                                        gridSpan1 = new int[] { 1, 7 };
                                        TableCellProperties tcp3 = new TableCellProperties(new GridSpan() { Val = gridSpan1[b3] }); b3++;

                                        TableCell cell23 = new TableCell(tcp3, new Paragraph(new Run(new Text(List1[m][u2]))));

                                        TableCell cell2 = row8.Elements<TableCell>().ElementAt(yacheika2);
                                        cell2 = cell23;

                                        var sdts12 = mainPart.Document.Descendants<SdtElement>();

                                        //TableRow rowCopy = (TableRow)theRow.CloneNode(true);

                                        row8.Descendants<TableCell>().ElementAt(yacheika2).Append(new Paragraph
                                            (new Run(new Text(List1[m][u2]))));

                                        //row8.Elements<TableCell>().ElementAt(yacheika).InnerText = List1[m][u2];
                                        Paragraph p3 = row8.Descendants<TableCell>().ElementAt(yacheika2).Elements<Paragraph>().First();
                                        Run t2 = p3.Elements<Run>().First();


                                        RunProperties rPr2 = new RunProperties(
                    new RunFonts()
                    {
                        Ascii = font,
                        HighAnsi = font
                    },

                                   new FontSize()
                                   {
                                       Val = fs.Val
                                   });
                                        t2.PrependChild<RunProperties>(rPr2);

                                        if (t2.Count() > 1)
                                        {
                                            t2.LastChild.Remove();
                                        }

                                      
                                        if (row8.Descendants<TableCell>().ElementAt(yacheika2).Elements<Paragraph>().Count() > 1)
                                        {
                                            foreach (var t3 in row8.Descendants<TableCell>().ElementAt(yacheika2).Elements<Paragraph>())
                                            {
                                                if (countp == 0)
                                                {
                                                    t3.Remove();
                                                    break;
                                                }

                                            }
                                        }
                                        foreach (var t3 in row8.Descendants<TableCell>().ElementAt(yacheika2).Elements<Paragraph>())
                                        {
                                            if (t3.Elements<ParagraphProperties>().Count() > 1)
                                            {
                                                for (int i7 = 0; i7 < t3.Elements<ParagraphProperties>().Count(); i7++)
                                                {
                                                    t3.Elements<ParagraphProperties>().ElementAt(i7).Remove();
                                                }
                                            }

                                            u2++;



                                            TableCellProperties tcp5 = new TableCellProperties(

                    new TableCellVerticalAlignment()
                    {
                        Val = TableVerticalAlignmentValues.Center
                    });



                                            ParagraphProperties pp = new ParagraphProperties(new Justification() { Val = JustificationValues.Center });

                                            t3.PrependChild<ParagraphProperties>(pp);

                                            t3.PrependChild<TableCellProperties>(tcp5);


                                        }



                                    }

                                   if(row != 4)
                                   {
                                       countp = 0;
                                       
                                        if (theTable.Elements<TableRow>().Count() <= row)
                                        {
                                            break;
                                        }
                                       
                                    
                                        TableRow row9 = theTable.Elements<TableRow>().ElementAt(row);
                                        row8 = theTable.Elements<TableRow>().ElementAt(row - 1);
                                     
                                      
                                           // row9 = theTable.Elements<TableRow>().ElementAt(row);
                                            if (yacheika2 == 15)
                                            {
                                                break;
                                            }
                                            if (count == 0)
                                            {

                                                if (u2 <= List1[m].Count() - 1)
                                                {
                                                    if (List1[m][u2 - 15] != List1[m][u2])
                                                    {
                                                        p = false;
                                                    }

                                                }

                                                else
                                                {
                                                    break;
                                                }

                                            }
                                                    if (p == false)
                                                    {




                                                      

                                       




                                                            TableCell cell1 = row8.Elements<TableCell>().ElementAt(yacheika2);
                                                            int b3 = 0;
                                                            int[] gridSpan1 = null;
                                                            gridSpan1 = new int[] { 1, 7 };
                                                            TableCellProperties tcp3 = new TableCellProperties(new GridSpan() { Val = gridSpan1[b3] }); b3++;

                                                            // TableCell cell23 = new TableCell(tcp3, new Paragraph(new Run(new Text(List1[m][u2]))));

                                                            // TableCell cell2 = row9.Elements<TableCell>().ElementAt(yacheika2);
                                                            // cell2 = cell23;

                                                            //  var sdts12 = mainPart.Document.Descendants<SdtElement>();

                                                            //TableRow rowCopy = (TableRow)theRow.CloneNode(true);

                                                            row9.Descendants<TableCell>().ElementAt(yacheika2).Append(new Paragraph
                                                                (new Run(new Text(List1[m][u2]))));

                                                            //row8.Elements<TableCell>().ElementAt(yacheika).InnerText = List1[m][u2];
                                                            Paragraph p3 = row9.Descendants<TableCell>().ElementAt(yacheika2).Elements<Paragraph>().First();
                                                            Run t2 = p3.Elements<Run>().First();

                                                            if (row9.Descendants<TableCell>().ElementAt(yacheika2).Elements<Paragraph>().Count() > 1)
                                                            {
                                                                foreach (var t3 in row9.Descendants<TableCell>().ElementAt(yacheika2).Elements<Paragraph>())
                                                                {
                                                                    if (countp == 0)
                                                                    {
                                                                        t3.Remove();
                                                                        break;
                                                                    }

                                                                }
                                                            }
                                                            RunProperties rPr2 = new RunProperties(
                                        new RunFonts()
                                        {
                                            Ascii = font,
                                            HighAnsi = font
                                        },

                                                       new FontSize()
                                                       {
                                                           Val = fs.Val
                                                       });
                                                            t2.PrependChild<RunProperties>(rPr2);
                                                            if (t2.Count() > 1)
                                                            {
                                                                t2.LastChild.Remove();
                                                            }

                                                        
                                                            foreach (var t3 in row9.Descendants<TableCell>().ElementAt(yacheika2).Elements<Paragraph>())
                                                            {
                                                                if (t3.Elements<ParagraphProperties>().Count() > 1)
                                                                {
                                                                    for (int i7 = 0; i7 < t3.Elements<ParagraphProperties>().Count(); i7++)
                                                                    {
                                                                        t3.Elements<ParagraphProperties>().ElementAt(i7).Remove();
                                                                    }
                                                                }





                                                                TableCellProperties tcp5 = new TableCellProperties(

                                        new TableCellVerticalAlignment()
                                        {
                                            Val = TableVerticalAlignmentValues.Center
                                        });



                                                                ParagraphProperties pp = new ParagraphProperties(new Justification() { Val = JustificationValues.Center });

                                                                t3.PrependChild<ParagraphProperties>(pp);

                                                                t3.PrependChild<TableCellProperties>(tcp5);


                                                            }

                                                        }
                                                    
                                                    else
                                                    {
                                                        countp = 0;
                                                        if (List1[m][u2 - 15] != List1[m][u2])
                                                        {
                                                           // row9 = theTable.Elements<TableRow>().ElementAt(row - 1);




                                                            TableCell cell1 = row8.Elements<TableCell>().ElementAt(yacheika2);
                                                            int b3 = 0;
                                                            int[] gridSpan1 = null;
                                                            gridSpan1 = new int[] { 1, 7 };
                                                            TableCellProperties tcp3 = new TableCellProperties(new GridSpan() { Val = gridSpan1[b3] }); b3++;

                                                            TableCell cell23 = new TableCell(tcp3, new Paragraph(new Run(new Text(List1[m][u2]))));

                                                            TableCell cell2 = row8.Elements<TableCell>().ElementAt(yacheika2);
                                                            cell2 = cell23;

                                                            var sdts12 = mainPart.Document.Descendants<SdtElement>();

                                                            //TableRow rowCopy = (TableRow)theRow.CloneNode(true);

                                                            row9.Descendants<TableCell>().ElementAt(yacheika2).Append(new Paragraph
                                                                (new Run(new Text(List1[m][u2]))));

                                                            //row8.Elements<TableCell>().ElementAt(yacheika).InnerText = List1[m][u2];
                                                            Paragraph p3 = row9.Descendants<TableCell>().ElementAt(yacheika2).Elements<Paragraph>().First();
                                                            Run t2 = p3.Elements<Run>().First();

                                                            if (row9.Descendants<TableCell>().ElementAt(yacheika2).Elements<Paragraph>().Count() > 1)
                                                            {
                                                                foreach (var t3 in row9.Descendants<TableCell>().ElementAt(yacheika2).Elements<Paragraph>())
                                                                {
                                                                    if (countp == 0)
                                                                    {
                                                                        t3.Remove();
                                                                        break;
                                                                    }

                                                                }
                                                            }

                                                            RunProperties rPr2 = new RunProperties(
                                        new RunFonts()
                                        {
                                            Ascii = font,
                                            HighAnsi = font
                                        },

                                                       new FontSize()
                                                       {
                                                           Val = fs.Val
                                                       });
                                                            t2.PrependChild<RunProperties>(rPr2);
                                                            if (t2.Count() > 1)
                                                            {
                                                                t2.LastChild.Remove();
                                                            }

                                                       
                                                        


                                                            foreach (var t3 in row9.Descendants<TableCell>().ElementAt(yacheika2).Elements<Paragraph>())
                                                            {
                                                                if (t3.Elements<ParagraphProperties>().Count() > 1)
                                                                {
                                                                    for (int i7 = 0; i7 < t3.Elements<ParagraphProperties>().Count(); i7++)
                                                                    {
                                                                        t3.Elements<ParagraphProperties>().ElementAt(i7).Remove();
                                                                    }
                                                                }





                                                                TableCellProperties tcp5 = new TableCellProperties(

                                        new TableCellVerticalAlignment()
                                        {
                                            Val = TableVerticalAlignmentValues.Center
                                        });



                                                                ParagraphProperties pp = new ParagraphProperties(new Justification() { Val = JustificationValues.Center });

                                                                t3.PrependChild<ParagraphProperties>(pp);

                                                                t3.PrependChild<TableCellProperties>(tcp5);


                                                            }
                                                        }
                                                    }


                                                    //string row11 = theTable.Elements<TableRow>().ElementAt(row1).Elements<TableCell>().ElementAt(yacheika1).InnerText;
                                                
                                                    count++;
                                                    u2++;
                                                }


                                     
                                             
                                            }

                                        
                                    }

//
                            int y = theTable.Elements<TableRow>().Count();

                            for (row = 4; row < theTable.Elements<TableRow>().Count(); row++)
                            {
                                TableRow row8 = theTable.Elements<TableRow>().ElementAt(row);

                                for (int yacheika2 = 0; yacheika2 < theTable.Elements<TableRow>().ElementAt(row).Elements<TableCell>().Count(); yacheika2++)
                                {
                                    foreach (var t3 in row8.Descendants<TableCell>().ElementAt(yacheika2).Elements<Paragraph>())
                                    {
                                        Paragraph p3 = row8.Descendants<TableCell>().ElementAt(yacheika2).Elements<Paragraph>().First();
                                        Run t2 = p3.Elements<Run>().First();


                                        RunProperties rPr2 = new RunProperties(
                    new RunFonts()
                    {
                        Ascii = font,
                        HighAnsi = font
                    },

                                   new FontSize()
                                   {
                                       Val = fs.Val
                                   });
                                        t2.PrependChild<RunProperties>(rPr2);
                                       
                                    }
                                }
                            }

                            }

                         var sdts123 = mainPart.Document.Descendants<SdtElement>();


                         foreach (var sdt3 in sdts123)
                         {

                             Tag ff = sdt3.SdtProperties.GetFirstChild<Tag>();
                             string old_text = sdt3.SdtProperties.InnerText;


                             if (ff != null)
                             {

                                 if (dict.ContainsKey((ff.Val)))
                                 {
                                     string value = dict[ff.Val]; ;
                                     sdt3.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().FirstOrDefault().Text = value;


                                 }
                             }
                         }


                        }

                    }
                

                try
                {


                    File.Delete(path2);


                    using (FileStream fileStream = new FileStream(path2,
                   System.IO.FileMode.CreateNew))
                    {
                        mem.WriteTo(fileStream);
                    }

                }


                catch (Exception ex)
                {

                    throw ex;


                }



            }
            
        }
    }
}

    
