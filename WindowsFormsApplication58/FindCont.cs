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
    class FindCont
    {
        public static void Find_table(List<List<string>> List1, string path1, string path2)
        {

            List<string> Tables = new List<string>();
            int flazhok = 0;
            int fcount = 0;
            int k = 0;
            byte[] byteArray = File.ReadAllBytes(path1);
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
                    fs.Val = "8";
                    string[] tblTag = new string[0];//Табличные теги

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


                        bool b = true;
                        SdtBlock ccWithTable = mainPart.Document.Body.Descendants<SdtBlock>().FirstOrDefault();

                        int index = 0;




                        foreach (var tt in ccWithTable1)
                        {




                            if (tt.SdtProperties.GetFirstChild<Tag>().Val == tblTag[index])
                            {
                                ccWithTable = mainPart.Document.Body.Descendants<SdtBlock>().Where
                    (r => r.SdtProperties.GetFirstChild<Tag>().Val == tblTag[index]).Single();

                                Tables.Add(tblTag[index]);
                                flazhok = 1;
                                break;

                            }

                        }
                        int count = 0;
                        int countt = 0;
                        int r2 = 3; int n = 0;
                        var tr2 = new TableRow(); var tr3 = new TableRow();
                        if (flazhok == 1)
                        {
                            int struct2 = 0;
                            for (int u23 = 1; u23 < List1[k].Count; u23++)
                            {
                                if (u23 >= 15 + struct2)
                                {
                                    countt++;
                                    struct2 += 15;
                                }
                            }

                            int f5 = 0; int f6 = 0;
                            Table theTable = ccWithTable.Descendants<Table>().FirstOrDefault();

                            TableRow row8 = theTable.Elements<TableRow>().ElementAt(r2);



                            theTable.InsertAfter<TableRow>(tr2, row8);
                            int county = 0;
                            int u2 = 1;
                            int u233 = 1;
                            int u234 = 1;

                            fs.Val = "16";

                            for (int u23 = 0; u23 < countt; u23++)
                            {
                                if (u23 != 0)
                                {
                                    u233 = u2 + 1;
                                    u234 = u233 - 1;
                                }
                                county = 0;
                                for (u2 = 1; u2 < List1[0].Count(); u2++)
                                {

                                    u2 = u233;
                                    if (u2 > 15)
                                        if (county == 0)
                                        {
                                            {
                                                u233 = u233 - 1;
                                            }
                                        }
                                    county++;
                                    u233++;
                                    b = true;

                                    if (u2 >= u234 + 15)
                                    {
                                        break;
                                    }

                                    else
                                    {
                                        b = false;




                                        int i3 = 0;
                                        int i8 = 3;
                                        int y = 0;
                                        int b3 = 0;

                                        int c = 0;
                                        M.Clear();
                                        //TableRow row2 = theTable.Elements<TableRow>().ElementAt(i8);

                                        TableCell cell = tr2.Elements<TableCell>().FirstOrDefault();



                                        int[] gridSpan1 = null;
                                        gridSpan1 = new int[] { 1, 7 };
                                        TableCellProperties tcp3 = new TableCellProperties(new GridSpan() { Val = gridSpan1[b3] }); b3++;



                                        TableCell cell2 = new TableCell();

                                        cell2.Append(new Paragraph
                                          (new Run(new Text(""))));
                                        tr2.AppendChild(cell2);
                                   /*
                                        Paragraph p3 = cell2.Elements<Paragraph>().First();
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

                                      */


                                        
                                        if (b == false)
                                        {


                                            if (u2 >= u234 + 15 - 1 && u2!= List1[0].Count-1)
                                            {
                                                n = 14;
                                                tr2 = new TableRow();
                                                r2++;
                                                row8 = theTable.Elements<TableRow>().ElementAt(r2);

                                                theTable.InsertAfter<TableRow>(tr2, row8);




                                            }
                                        }
                                        else
                                        {
                                            if (u2 != List1[0].Count() - 1)
                                            {

                                                if (u2 >= 16 + n)
                                                {
                                                    n = n + 15;
                                                    tr2 = new TableRow();
                                                    r2++;
                                                    row8 = theTable.Elements<TableRow>().ElementAt(r2);

                                                    theTable.InsertAfter<TableRow>(tr2, row8);




                                                }
                                            }
                                        }
                                    }
                                }
                            }
                             }
                        int n18=0;
                        int tyacheika = 1;int count7=0;
                        Dictionary<int, int> d3 = new Dictionary<int, int>();
                   
                            for (tyacheika = 1; tyacheika < List1[0].Count(); tyacheika++)
                            {
                                d3.Add(tyacheika + n18, n18 + 15 + tyacheika - 1);
                            
                              n18 += 14;
                                
                            }

                        
                        Table theTable2 = ccWithTable.Descendants<Table>().FirstOrDefault();
                                 Dictionary<int, int> d = new Dictionary<int, int>();
                                 int g = 0;
                            TableCellProperties tcp1 = new TableCellProperties();
                                  List<int> l = new List<int>();
                                int n15=0;
                            
                                int yy = 1;
                                int prt = 1;
                                int lcount=1;
                        int t16=0;
                        int t17 = 0;
                  
                        bool flag = true;
                        int ff = 0;  int ff2 = 0;
                                int yacheika1 = 1;
                                int perv = 0;
                                int count55 = 0;
                                int yperv = 0;
                                bool flazhok2 = true;
                                    prt = yacheika1;
                                    int ccount = 0;
                                    prt = yy;
                                    bool nk = true;
                                    int nor2 = 0;
                                    int nor = 0;
                                    for (int yacheika12 = 1; yacheika12 < List1[0].Count(); yacheika12++)
                                    {

                                        if (perv >= 16)
                                        {
                                            break;
                                        }
                                        g = 0;
                                        t16 = 0;
                                        if (flazhok2 == false)
                                        {
                                            g = ff;
                                            yacheika12 = ff2;
                                            n15 = nor;
                                            t16 = t17;
                                        }


                                        if (flag == false)
                                        {
                                            perv++;
                                            yacheika12 = perv;
                                            count55 = 0;
                                        }

                                        for (yacheika1 = yacheika12; yacheika1 < List1[0].Count(); yacheika1++)
                                        {
                                            yacheika12++;
                                            if (yacheika1 == 10)
                                            {

                                            }
                                            if (count55 == 0)
                                            {
                                                perv = yacheika1;

                                            }
                                            count55++;
                                            flag = true;
                                            if (g > 15)
                                            {
                                                break;
                                            }
                                            flazhok2 = true;

                                            t16++;
                                            t17 = t16;
                                            ff = g;

                                            ff2 = yacheika1;

                                            if (n15 + 15 + yacheika1 >= List1[0].Count())
                                            {
                                                flag = false;
                                                break;
                                            }
                                            string k4 = List1[0][yacheika1 + n15];
                                            string k5 = List1[0][n15 + 15 + yacheika1];
                                            string oneone = ""; string twotwo = "";
                                            string one1 = List1[0][yacheika1 + n15];
                                            string one2 = List1[0][n15 + 15 + yacheika1 - 1];
                                            bool ka = true;








                                            if (n15 + 15 + yacheika1 <= List1[0].Count() - 1)
                                            {

                                                if (List1[0][yacheika1 + n15] != List1[0][n15 + 15 + yacheika1])
                                                {
                                                    n15 += 15;
                                                    nor = n15;
                                                    nk = false;
                                                    flazhok2 = false;
                                                    yperv = perv;
                                                    break;

                                                }


                                                g++;


                                                if (n15 + 15 + yacheika1 <= List1[0].Count() - 1)
                                                {




                                                    foreach (var pair in d3)
                                                    {

                                                        if (pair.Key <= yacheika1 + n15 && yacheika1 + n15 <= pair.Value)
                                                        {

                                                            oneone = pair.Key.ToString();
                                                            break;
                                                        }
                                                    }


                                                    foreach (var pair1 in d3)
                                                    {
                                                        if (pair1.Key <= n15 + 15 + yacheika1 - 1 && n15 + 15 + yacheika1 - 1 <= pair1.Value)
                                                        {
                                                            twotwo = pair1.Key.ToString();
                                                            break;
                                                        }
                                                    }

                                                    bool flag33 = true;
                                                    if (List1[0][Convert.ToInt32(oneone)] != List1[0][Convert.ToInt32(twotwo)])
                                                    {


                                                        flag33 = false;


                                                        nk = true;

                                                        n15 += 15;
                                                        nor = n15;

                                                        flazhok2 = false;

                                                        break;


                                                    }

                                                    if (List1[0][yacheika1 + n15] == List1[0][n15 + 15 + yacheika1])
                                                    {
                                                        if (flag33 != false)
                                                        {

                                                            l.Add(t16 - 1 + 4);


                                                            l.Add(t16 + 5 - 1);
                                                            n15 += 14;

                                                            nor = n15;
                                                            nk = true;
                                                        }

                                                    }






                                                }

                                            }






                                        }


                                        n15 = 0;

                                        for (int i1 = 0; i1 < l.Count; i1++)
                                        {

                                            if (!d.ContainsKey(l[i1]))
                                            {
                                                d.Add(l[i1], l[i1]);

                                            }
                                        }

                                        bool kas = true;
                                        List<int> list = new List<int>(d.Keys);
                                        bool yp = true;
                                        if (list.Count != 0)
                                        {
                                            for (int i1 = 0; i1 < list.Count(); i1++)
                                            {
                                                if (list[i1] >= 18)
                                                {
                                                  

                                                    if (list[i1] >= 18)
                                                    {/*
                                                        list.Remove(list[i1]);
                                                      * */
                                                        yp = false;
                                                    }

                                                }
                                            }
                                         
                                            /*
                                            kas = false;
                                            l.Clear();
                                            list.Clear();
                                            d.Clear();
                                             * */



                                            if (kas == true)
                                            {
                                                if (perv < 16)
                                                {
                                                    int yacheika = 0;
                                                    //  List<int> list = new List<int>(d.Keys);
                                                    yacheika = yacheika1;


                                                    /*
                                                    if (yp == false)
                                                    {*/
                                                    
                                                        tcp1 = new TableCellProperties(

                                                                                            new VerticalMerge()
                                                                                            {
                                                                                                Val = MergedCellValues.Restart
                                                                                            }
                                                                                             );
                                                        theTable2.Elements<TableRow>().ElementAt(list[0]).ElementAt(perv - 1).Append(tcp1);

                                                        for (int t = 1; t < list.Count; t++)
                                                        {
                                                            TableCellProperties tcp11 = new TableCellProperties(

                                                            new VerticalMerge()
                                                            {
                                                                Val = MergedCellValues.Continue
                                                            }
                                                             );


                                                            theTable2.Elements<TableRow>().ElementAt(list[t]).ElementAt(perv - 1).Append(tcp11);
                                                        }
                                                    }
                                                    /*
                                                    else
                                                    {
                                                     
                                                        for (int t = 1; t < list.Count; t++)
                                                        {
                                                            TableCellProperties tcp11 = new TableCellProperties(

                                                            new VerticalMerge()
                                                            {
                                                                Val = MergedCellValues.Continue
                                                            }
                                                             );
                                                            theTable2.Elements<TableRow>().ElementAt(list[t]).ElementAt(perv - 1).Append(tcp11);
                                                        
                                                        }
                                                   // }
                                                    /*
                                                    

                                                                                                        for (int t = 1; t < list.Count - 1; t++)
                                                        {
                                                            TableCellProperties tcp11 = new TableCellProperties(

                                                            new VerticalMerge()
                                                            {
                                                                Val = MergedCellValues.Continue
                                                            }
                                                             );

                                                    */
                                                           // theTable2.Elements<TableRow>().ElementAt(list[t]).ElementAt(perv - 1).Append(tcp11);

                                                        
                                                    
                                                }
                                              

                                                    
                                                
                                                l.Clear();

                                                d.Clear();
                                            }
                                        
                                    }

                                  //  theTable2.Elements<TableRow>().ElementAt(theTable2.Elements<TableRow>().Count() - 1).Remove();
                       
                    
                                    }







                               
                                    
                                }
                            
                           

                            /*
                            TableCellProperties tcp10 = new TableCellProperties(

new VerticalMerge()
{
  Val = MergedCellValues.Restart
}
);
                            int t = d[l[0]];
                            theTable.Elements<TableRow>().ElementAt(d[l[1]]).Elements<TableCell>().ElementAt(yacheika).Append(tcp10);


                            */

                            //TableRow r23 = theTable.Elements<TableRow>().ElementAt(row1);
                            /*
                                        if (flag != false)
                                        {
                                            if (list.Count != 0)
                                            {
                                                TableCellProperties tcp1 = new TableCellProperties(

                               new VerticalMerge()
                               {
                                   Val = MergedCellValues.Restart
                               }
                                );
                                                theTable.Elements<TableRow>().ElementAt(list[0]).ElementAt(yacheika1).Append(tcp1);

                                                for (int i11 = 1; i11 < d.Count; i11++)
                                                {
                                                    TableCellProperties tcp11 = new TableCellProperties(

                                   new VerticalMerge()
                                   {
                                       Val = MergedCellValues.Continue
                                   }
                                    );


                                                    theTable.Elements<TableRow>().ElementAt(list[i11]).ElementAt(yacheika1).Append(tcp11);

                                                }
                                            }

                                        }
                                        d.Clear();

                                    }
                            
                                }


                                /*
           TableCellProperties tcp11 = new TableCellProperties(

  new VerticalMerge()
  {
      Val = MergedCellValues.Restart
  }
   );
                                theTable.Elements<TableRow>().ElementAt(d[l[1]]).Elements<TableCell>().ElementAt(yacheika).Append(tcp11);



                                TableCellProperties tcp12 = new TableCellProperties(

               new VerticalMerge()
               {
                   Val = MergedCellValues.Continue
               }
                );
                                theTable.Elements<TableRow>().ElementAt(d[l[2]]).ElementAt(yacheika).Append(tcp12);
                            }
*/





                /*
                                
                                     for (int u2 = 1; u2 < List1[k].Count; u2++)
                                  {

                                    if (theTable.Elements<TableRow>().ElementAt(u2-1) == theTable.Elements<TableRow>().ElementAt(u2))
                                        {
                                            TableCellProperties tcp10 = new TableCellProperties(

     new VerticalMerge()
     {
         Val = MergedCellValues.Restart
     }
      );

                                            TableCell cl = new TableCell();
                                            cl = tr2.Elements<TableCell>().ElementAt(u2);

                                            theTable.Elements<TableRow>().ElementAt(u2).Elements<TableCell>().ElementAt(u2).Append(tcp10);
                                           cl= theTable.Elements<TableRow>().ElementAt(u2).Elements<TableCell>().ElementAt(u2);
                                        }
                                    }

                                }

                                    

                */





                /*

for (int u3 = 5; u3 < theTable.Elements<TableRow>().Count(); u3++)
                {
                    TableRow tr5 = theTable.Elements<TableRow>().ElementAt(u3);
                   string d= theTable.Elements<TableRow>().ElementAt(u3).InnerText;
                   for (int u4 = 0; u4 < tr5.Elements<TableCell>().Count(); u4++)
                   {

                       TableRow tr = theTable.Elements<TableRow>().ElementAt(u3);
                       TableCell tc12 = tr.Elements<TableCell>().ElementAt(u4);
                       string text = theTable.Elements<TableRow>().ElementAt(u3).ElementAt(u4).InnerText;
                       string text2 = theTable.Elements<TableRow>().ElementAt(u3 - 1).ElementAt(u4).InnerText;

                       if (text == text2)
                       {

                           count++;
                       }}}






                int schet = 0;
                for (int u3 = 5; u3 < theTable.Elements<TableRow>().Count(); u3++)
                {
                    TableRow tr5 = theTable.Elements<TableRow>().ElementAt(u3);
                   string d= theTable.Elements<TableRow>().ElementAt(u3).InnerText;
                   for (int u4 = 0; u4 < tr5.Elements<TableCell>().Count(); u4++)
                   {

                       TableRow tr = theTable.Elements<TableRow>().ElementAt(u3);
                       TableCell tc12 = tr.Elements<TableCell>().ElementAt(u4);
                       string text = theTable.Elements<TableRow>().ElementAt(u3).ElementAt(u4).InnerText;
                       string text2 = theTable.Elements<TableRow>().ElementAt(u3 - 1).ElementAt(u4).InnerText;

                       if (text == text2)
                       {

                           count++;


                           string c1 = theTable.Elements<TableRow>().ElementAt(u3).Elements<TableCell>().ElementAt(u4).InnerText;










                           TableCellProperties tcp10 = new TableCellProperties(

   new VerticalMerge()
   {
       Val = MergedCellValues.Restart
   }
    );


                           theTable.Elements<TableRow>().ElementAt(u3 - 2).ElementAt(u4).Append(tcp10);






                           TableCellProperties tcp1 = new TableCellProperties(

           new VerticalMerge()
           {
               Val = MergedCellValues.Continue
           }
            );
                           theTable.Elements<TableRow>().ElementAt(u3).ElementAt(u4).Append(tcp1);
                           // tc12 = theTable.Elements<TableRow>().ElementAt(u3).Elements<TableCell>().ElementAt(u4);


                       }
                   }

                                       

                   */







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

