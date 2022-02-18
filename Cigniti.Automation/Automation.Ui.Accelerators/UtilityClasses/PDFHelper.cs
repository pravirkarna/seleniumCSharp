using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Document = Microsoft.Office.Interop.Word.Document;
using Image = System.Drawing.Image;
using System.Linq;

using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Microsoft.Office.Interop.Word;
using System.Configuration;
using Automation.Ui.Accelerators.ReportingClassess;

namespace Automation.Ui.Accelerators.UtilityClasses
{
    public static class PDFHelper
    {
        /// <summary>
        /// Reads Pdf File and gets text between starting and ending text.
        /// </summary>
        /// <param name="filename"></param>
        /// <param name="page"></param>
        /// <param name="startText"></param>
        /// <param name="endText"></param>
        /// <returns></returns>
        public static string ReadPdfFile(string filename, int page, string startText, string endText)
        {
            string expected = string.Empty;
            string code = string.Empty;
            PdfReader reader = null;
            string tempTextFromPage = string.Empty;

            try
            {
                using (reader = new PdfReader(filename))
                {
                    ITextExtractionStrategy its = new SimpleTextExtractionStrategy();
                    string textFromPage = PdfTextExtractor.GetTextFromPage(reader, page) + "";

                    textFromPage = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(textFromPage)));

                    tempTextFromPage = textFromPage;
                    String finalCode = string.Empty;
                    int lines = 0;
                    foreach (var myString in tempTextFromPage.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        if (myString != " ")
                        {
                            lines++;
                            finalCode = finalCode + "\n" + myString;
                        }
                    }
                    Console.WriteLine(finalCode);
                    if (endText == string.Empty)
                        expected = textFromPage.Substring(textFromPage.IndexOf(startText) + startText.Length);
                    else
                        expected = GetBetween(finalCode, startText, endText);
                }

            }
            catch (Exception ex)
            {
                reader.Close();
                reader.Dispose();
                ThrowErrorMessage(ex);
            }
            return expected.Trim();
        }

        /// <summary>
        /// Read Pdf File
        /// </summary>
        /// <param name="Filename"></param>
        /// <returns></returns>
        public static string ReadPdfFile(string fileName)
        {
            string strText = string.Empty;
            PdfReader reader = null;

            try
            {
                using (reader = new PdfReader(fileName))
                {
                    int numberOfPages = reader.NumberOfPages;

                    for (int page = 1; page <= numberOfPages; page++)
                    {
                        ITextExtractionStrategy its = new iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy();
                        string s = PdfTextExtractor.GetTextFromPage(reader, page, its);
                        s = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(s)));
                        strText = strText + s;
                    }
                }
            }
            catch (Exception ex)
            {
                reader.Close();
                reader.Dispose();
                ThrowErrorMessage(ex);
            }

            return strText;
        }

        /// <summary>
        /// Read Pdf File content as a List
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static List<string> ReadPdfFileContentAsAList(string fileName)
        {
            List<string> strText = null; ;
            if (string.IsNullOrEmpty(fileName))
            {
                return strText;
            }
            strText = new List<string>();
            PdfReader reader = null;
            try
            {
                using (reader = new PdfReader(fileName))
                {
                    int numberOfPages = reader.NumberOfPages;

                    for (int page = 1; page <= numberOfPages; page++)
                    {
                        ITextExtractionStrategy its = new iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy();
                        string s = PdfTextExtractor.GetTextFromPage(reader, page, its);
                        s = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(s)));
                        strText.AddRange(s.Split('\n'));
                    }
                }
            }
            catch (Exception ex)
            {
                ThrowErrorMessage(ex);
            }
            finally
            {
                reader.Close();
                reader.Dispose();
            }

            return strText;
        }

        /// <summary>
        /// Get Pdf file Table Count
        /// </summary>
        /// <param name="pdfFilePath">pdf file location</param>
        /// <returns>Table count</returns>
        public static int GetPdfTableCount(string pdfFilePath)
        {
            if (string.IsNullOrEmpty(pdfFilePath))
            {
                return -1;
            }
            int tableCount = -1;
            Application app = null;
            Documents docs = null;
            Document doc = null;
            try
            {
                app = new Application();
                docs = app.Documents;
                doc = docs.Open(pdfFilePath, ReadOnly: true);
                Console.WriteLine(string.Format("Opening {0} file...\n", pdfFilePath));

                tableCount = doc.Tables.Count;
                Console.WriteLine("Total tables found in the file : " + tableCount + "\n");

            }
            catch (Exception ex)
            {
                Console.WriteLine("Message: {0}", ex.Message);
                Console.WriteLine("InnerException: {0}", ex.InnerException);
                ThrowErrorMessage(ex);
            }
            finally
            {
                doc.Close(false);
                app.Quit(false);
                Marshal.ReleaseComObject(doc);
                Marshal.ReleaseComObject(docs);
                Marshal.ReleaseComObject(app);
            }
            return tableCount;
        }

        /// <summary>
        /// Get Table Content From PDF file.
        /// Note: Fetching of table depends on table format (No.of columns and rows).
        ///       If table format is not correct, method will throw exception.
        /// </summary>
        /// <param name="pdfFilePat"></param>
        /// <returns></returns>
        public static DataSet GetTableContentFromPDF(string pdfFilePath, string tableName)
        {
            if (string.IsNullOrEmpty(pdfFilePath))
            {
                return null;
            }
            Application app = null;
            Documents docs = null;
            Document doc = null;
            Table table = null;
            Range range = null;
            DataSet dataSet = null;
            Cells cells = null;
            System.Data.DataTable dls = null;
            DataColumn dataColumn = null;
            DataRow dataRow = null;

            int tableCount = -1;
            try
            {
                app = new Application();
                docs = app.Documents;
                Console.WriteLine(string.Format("Opening {0} file...\n", pdfFilePath));
                doc = docs.Open(pdfFilePath, ReadOnly: true);
                tableCount = doc.Tables.Count;
                if (tableCount <= 0)
                {
                    return null;
                }

                Console.WriteLine("Total tables found in the file : " + tableCount + "\n");
                dataSet = new DataSet();
                for (int t = 1; t <= tableCount; t++)
                {
                    dls = new System.Data.DataTable(tableName + "Table" + t);

                    table = doc.Tables[t];

                    Console.WriteLine("Getting content from table : " + t + "\n");
                    range = table.Range;
                    cells = range.Cells;

                    Console.WriteLine("Row count {0}, Column count {1}", table.Rows.Count, table.Columns.Count);
                    string[] columnNames = new string[table.Columns.Count];
                    //Adding columns to the table.
                    for (int c = 1; c <= table.Columns.Count; c++)
                    {
                        dataColumn = new System.Data.DataColumn();
                        string str = table.Rows[1].Cells[c].Range.Text;
                        if (!string.IsNullOrEmpty(str) && c == 1 && str.Equals("\r\a"))
                        {
                            dataColumn.ColumnName = "Status";
                            columnNames[c - 1] = dataColumn.ColumnName;
                            dls.Columns.Add(dataColumn);
                            Console.WriteLine("\t" + dataColumn.ColumnName + "\t");
                        }
                        else
                        {
                            str = str.TrimEnd('\a');
                            str = str.TrimEnd('\r');
                            if (!string.IsNullOrEmpty(str) && !str.Contains("\r\a"))
                            {
                                dataColumn.ColumnName = str.Replace("\r", "").Replace(" ", "");
                                columnNames[c - 1] = dataColumn.ColumnName;
                                dls.Columns.Add(dataColumn);
                                Console.WriteLine("\t" + dataColumn.ColumnName + "\t");
                            }
                        }
                    }

                    dataSet.Tables.Add(dls);

                    for (int row = 2; row <= table.Rows.Count; row++)
                    {
                        dataRow = dls.NewRow();
                        Console.WriteLine("");
                        for (int c = 1; c <= table.Columns.Count; c++)
                        {
                            if (table.Columns.Count.Equals(table.Rows[row].Cells.Count))
                            {
                                string str = table.Rows[row].Cells[c].Range.Text;
                                str = str.TrimEnd('\a');
                                str = str.TrimEnd('\r');
                                if (!string.IsNullOrEmpty(str) && !str.Contains("\r\a"))
                                {
                                    dataRow[columnNames[c - 1]] = str.Replace("\r", "");
                                    Console.WriteLine("\t" + str + "\t");
                                }
                            }
                            else
                            {
                                break;
                            }
                        }
                        dls.Rows.Add(dataRow);
                    }
                }
            }
            catch (Exception ex)
            {
                ThrowErrorMessage(ex);

            }
            finally
            {
                if (doc != null && app != null)
                {
                    doc.Close(false);
                    app.Quit(false);
                    Marshal.ReleaseComObject(cells);
                    Marshal.ReleaseComObject(range);
                    Marshal.ReleaseComObject(table);
                    Marshal.ReleaseComObject(doc);
                    Marshal.ReleaseComObject(docs);
                    Marshal.ReleaseComObject(app);
                }
            }
            return dataSet;
        }

        /// <summary>
        /// Get Between text from string.
        /// </summary>
        /// <param name="content"></param>
        /// <param name="startString"></param>
        /// <param name="endString"></param>
        /// <returns></returns>
        private static string GetBetween(string content, string startString, string endString)
        {
            int Start = 0, End = 0;
            if (content.Contains(startString) && content.Contains(endString))
            {
                Start = content.IndexOf(startString, 0) + startString.Length;
                End = content.IndexOf(endString, Start);
                return content.Substring(Start, End - Start);
            }
            else
                return string.Empty;
        }

        /// <summary>
        /// Get Diff
        /// </summary>
        /// <param name="l1"></param>
        /// <param name="l2"></param>
        /// <returns></returns>
        public static List<string> GetDiff(List<string> l1, List<string> l2)
        {
            List<string> same = new List<string>();
            List<string> diff = new List<string>();
            var leftComp = l1.Count > l2.Count ? l1 : l2;
            var rightComp = l2.Count < l1.Count ? l2 : l1;

            for (int i = 0; i < leftComp.Count; i++)
            {
                if (rightComp.Count > i)
                {
                    if (leftComp[i].Equals(rightComp[i]))
                    {
                        continue;
                    }
                    else
                    {
                        string[] l1Split = leftComp[i].Split(' ');
                        string[] l2Split = rightComp[i].Split(' ');


                        for (int jl1 = 0; jl1 < l1Split.Length; jl1++)
                        {
                            if (!string.IsNullOrEmpty(l1Split[jl1]))
                            {
                                if (l2Split.Length > jl1)
                                {
                                    if (l1Split[jl1].Equals(l2Split[jl1]))
                                    {
                                        same.Add(l1Split[jl1]);
                                    }
                                    else
                                    {
                                        diff.Add(l1Split[jl1]);

                                    }
                                }
                                else
                                {
                                    diff.Add(l1Split[jl1]);
                                }
                            }
                        }
                    }
                }
                else
                {
                    string[] l1Split = leftComp[i].Split(' ');
                    for (int jl1 = 0; jl1 < l1Split.Length; jl1++)
                    {
                        diff.Add(l1Split[jl1]);
                    }
                }
            }
            return diff;
        }
        /// <summary>
        /// Genarate PDF Report
        /// </summary>
        /// <param name="expected"> Expected File path</param>
        /// <param name="actual">Actual File path</param>
        public static string GenaratePDFReport(string expected, string actual)
        {
            string finalOutputPath = string.Empty;

            try
            {
                string guid = "DiffPDF" + Utility.GetGUID();
                string relativePath = string.Concat("DiffPdfs", string.Format(@"\{0}.html", guid));
                var pdfPath = System.IO.Path.Combine(Engine.PathOfReport, relativePath);
                var command = ConfigurationManager.AppSettings.Get("BComparerExecutablePath").ToString();
                var arguments = string.Format("\"@{0}\\Script.txt\" \"{1}\" \"{2}\" \"{3}\" /silent", Directory.GetCurrentDirectory(), expected, actual, pdfPath);
                var exitCode = Utility.ExecuteExternalProgram(command, arguments);


                if (exitCode.Equals(0))
                {
                    finalOutputPath = relativePath;
                }

            }
            catch (Exception ex)
            {
                ThrowErrorMessage(ex);
            }

            return finalOutputPath;
        }

        /// <summary>
        /// Throws error message to console.
        /// </summary>
        /// <param name="ex"></param>
        private static void ThrowErrorMessage(Exception ex)
        {
            Console.WriteLine("Message: {0}", ex.Message);
            Console.WriteLine("InnerException: {0}", ex.InnerException);
            throw new Exception(string.Format("ErrorMessage: {0}. InnerExceptionMessage: {1}", ex.Message, ex.InnerException.Message));
        }
    }

    public class MyLocationTextExtractionStrategy : LocationTextExtractionStrategy
    {

        public float UndercontentCharacterSpacing { get; set; }
        public float UndercontentHorizontalScaling { get; set; }
        private SortedList<string, DocumentFont> ThisPdfDocFonts = new SortedList<string, DocumentFont>();
        private List<TextChunk> locationalResult = new List<TextChunk>();

        private bool StartsWithSpace(String inText)
        {
            if (string.IsNullOrEmpty(inText))
            {
                return false;
            }
            if (inText.StartsWith(" "))
            {
                return true;
            }
            return false;
        }

        private bool EndsWithSpace(String inText)
        {
            if (string.IsNullOrEmpty(inText))
            {
                return false;
            }
            if (inText.EndsWith(" "))
            {
                return true;
            }
            return false;
        }

        public override string GetResultantText()
        {
            locationalResult.Sort();

            StringBuilder sb = new StringBuilder();
            TextChunk lastChunk = null;

            foreach (var chunk in locationalResult)
            {
                if (lastChunk == null)
                {
                    sb.Append(chunk.text);
                }
                else
                {
                    if (chunk.SameLine(lastChunk))
                    {
                        float dist = chunk.DistanceFromEndOf(lastChunk);
                        if (dist < -chunk.charSpaceWidth)
                        {
                            sb.Append(" ");
                        }
                        else if (dist > chunk.charSpaceWidth / 2.0F && !StartsWithSpace(chunk.text) && !EndsWithSpace(lastChunk.text))
                        {
                            sb.Append(" ");
                        }
                        sb.Append(chunk.text);
                    }
                    else
                    {
                        sb.Append("\n");
                        sb.Append(chunk.text);
                    }
                }
                lastChunk = chunk;
            }
            return sb.ToString();
        }

        public List<iTextSharp.text.Rectangle> GetTextLocations(string pSearchString, System.StringComparison pStrComp)
        {
            List<iTextSharp.text.Rectangle> FoundMatches = new List<iTextSharp.text.Rectangle>();
            StringBuilder sb = new StringBuilder();
            List<TextChunk> ThisLineChunks = new List<TextChunk>();
            bool bStart = false;
            bool bEnd = false;
            TextChunk FirstChunk = null;
            TextChunk LastChunk = null;
            string sTextInUsedChunks = null;

            foreach (var chunk in locationalResult)
            {
                if (ThisLineChunks.Count > 0 && !chunk.SameLine(ThisLineChunks.Last()))
                {
                    if (sb.ToString().IndexOf(pSearchString, pStrComp) > -1)
                    {
                        string sLine = sb.ToString();

                        int iCount = 0;
                        int lPos = 0;
                        lPos = sLine.IndexOf(pSearchString, 0, pStrComp);
                        while (lPos > -1)
                        {
                            iCount++;
                            if (lPos + pSearchString.Length > sLine.Length)
                            {
                                break;
                            }
                            else
                            {
                                lPos = lPos + pSearchString.Length;
                            }
                            lPos = sLine.IndexOf(pSearchString, lPos, pStrComp);
                        }

                        int curPos = 0;
                        for (int i = 1; i <= iCount; i++)
                        {
                            string sCurrentText;
                            int iFromChar;
                            int iToChar;

                            iFromChar = sLine.IndexOf(pSearchString, curPos, pStrComp);
                            curPos = iFromChar;
                            iToChar = iFromChar + pSearchString.Length - 1;
                            sCurrentText = null;
                            sTextInUsedChunks = null;
                            FirstChunk = null;
                            LastChunk = null;

                            foreach (var chk in ThisLineChunks)
                            {
                                sCurrentText = sCurrentText + chk.text;

                                if (!bStart && sCurrentText.Length - 1 >= iFromChar)
                                {
                                    FirstChunk = chk;
                                    bStart = true;
                                }

                                if (bStart && !bEnd)
                                {
                                    sTextInUsedChunks = sTextInUsedChunks + chk.text;
                                }

                                if (!bEnd && sCurrentText.Length - 1 >= iToChar)
                                {
                                    LastChunk = chk;
                                    bEnd = true;
                                }
                                if (bStart && bEnd)
                                {
                                    FoundMatches.Add(GetRectangleFromText(FirstChunk, LastChunk, pSearchString, sTextInUsedChunks, iFromChar, iToChar, pStrComp));
                                    curPos = curPos + pSearchString.Length;
                                    bStart = false;
                                    bEnd = false;
                                    break;
                                }
                            }
                        }
                    }
                    sb.Clear();
                    ThisLineChunks.Clear();
                }
                ThisLineChunks.Add(chunk);
                sb.Append(chunk.text);
            }
            return FoundMatches;
        }

        private iTextSharp.text.Rectangle GetRectangleFromText(TextChunk FirstChunk, TextChunk LastChunk, string pSearchString,
          string sTextinChunks, int iFromChar, int iToChar, System.StringComparison pStrComp)
        {
            float LineRealWidth = LastChunk.PosRight - FirstChunk.PosLeft;

            float LineTextWidth = GetStringWidth(sTextinChunks, LastChunk.curFontSize,
                                                         LastChunk.charSpaceWidth,
                                                         ThisPdfDocFonts.ElementAt(LastChunk.FontIndex).Value);

            float TransformationValue = LineRealWidth / LineTextWidth;

            int iStart = sTextinChunks.IndexOf(pSearchString, pStrComp);

            int iEnd = iStart + pSearchString.Length - 1;

            string sLeft;
            if (iStart == 0)
            {
                sLeft = null;
            }
            else
            {
                sLeft = sTextinChunks.Substring(0, iStart);
            }

            string sRight;
            if (iEnd == sTextinChunks.Length - 1)
            {
                sRight = null;
            }
            else
            {
                sRight = sTextinChunks.Substring(iEnd + 1, sTextinChunks.Length - iEnd - 1);
            }

            float LeftWidth = 0;
            if (iStart > 0)
            {
                LeftWidth = GetStringWidth(sLeft, LastChunk.curFontSize,
                                                  LastChunk.charSpaceWidth,
                                                  ThisPdfDocFonts.Values.ElementAt(LastChunk.FontIndex));
                LeftWidth = LeftWidth * TransformationValue;
            }

            float RightWidth = 0;
            if (iEnd < sTextinChunks.Length - 1)
            {
                RightWidth = GetStringWidth(sRight, LastChunk.curFontSize,
                                                    LastChunk.charSpaceWidth,
                                                    ThisPdfDocFonts.Values.ElementAt(LastChunk.FontIndex));
                RightWidth = RightWidth * TransformationValue;
            }

            float LeftOffset = FirstChunk.distParallelStart + LeftWidth;
            float RightOffset = LastChunk.distParallelEnd - RightWidth;
            return new iTextSharp.text.Rectangle(LeftOffset, FirstChunk.PosBottom, RightOffset, FirstChunk.PosTop);
        }


        private float GetStringWidth(string str, float curFontSize, float pSingleSpaceWidth, DocumentFont pFont)
        {

            char[] chars = str.ToCharArray();
            float totalWidth = 0;
            float w = 0;

            foreach (Char c in chars)
            {
                w = pFont.GetWidth(c) / 1000;
                totalWidth += (w * curFontSize + this.UndercontentCharacterSpacing) * this.UndercontentHorizontalScaling / 100;
            }

            return totalWidth;
        }

        public override void RenderText(TextRenderInfo renderInfo)
        {
            LineSegment segment = renderInfo.GetBaseline();
            TextChunk location = new TextChunk(renderInfo.GetText(), segment.GetStartPoint(), segment.GetEndPoint(), renderInfo.GetSingleSpaceWidth());

            location.PosLeft = renderInfo.GetDescentLine().GetStartPoint()[Vector.I1];
            location.PosRight = renderInfo.GetAscentLine().GetEndPoint()[Vector.I1];
            location.PosBottom = renderInfo.GetDescentLine().GetStartPoint()[Vector.I2];
            location.PosTop = renderInfo.GetAscentLine().GetEndPoint()[Vector.I2];
            location.curFontSize = location.PosTop - segment.GetStartPoint()[Vector.I2];

            string StrKey = renderInfo.GetFont().PostscriptFontName + location.curFontSize.ToString();
            if (!ThisPdfDocFonts.ContainsKey(StrKey))
            {
                ThisPdfDocFonts.Add(StrKey, renderInfo.GetFont());
            }
            location.FontIndex = ThisPdfDocFonts.IndexOfKey(StrKey);
            locationalResult.Add(location);
        }


        private class TextChunk : IComparable<TextChunk>
        {
            public string text { get; set; }
            public Vector startLocation { get; set; }
            public Vector endLocation { get; set; }
            public Vector orientationVector { get; set; }
            public int orientationMagnitude { get; set; }
            public int distPerpendicular { get; set; }
            public float distParallelStart { get; set; }
            public float distParallelEnd { get; set; }
            public float charSpaceWidth { get; set; }
            public float PosLeft { get; set; }
            public float PosRight { get; set; }
            public float PosTop { get; set; }
            public float PosBottom { get; set; }
            public float curFontSize { get; set; }
            public int FontIndex { get; set; }


            public TextChunk(string str, Vector startLocation, Vector endLocation, float charSpaceWidth)
            {
                this.text = str;
                this.startLocation = startLocation;
                this.endLocation = endLocation;
                this.charSpaceWidth = charSpaceWidth;

                Vector oVector = endLocation.Subtract(startLocation);
                if (oVector.Length == 0)
                {
                    oVector = new Vector(1, 0, 0);
                }
                orientationVector = oVector.Normalize();
                orientationMagnitude = (int)(Math.Truncate(Math.Atan2(orientationVector[Vector.I2], orientationVector[Vector.I1]) * 1000));

                Vector origin = new Vector(0, 0, 1);
                distPerpendicular = (int)((startLocation.Subtract(origin)).Cross(orientationVector)[Vector.I3]);

                distParallelStart = orientationVector.Dot(startLocation);
                distParallelEnd = orientationVector.Dot(endLocation);
            }

            public bool SameLine(TextChunk a)
            {
                if (orientationMagnitude != a.orientationMagnitude)
                {
                    return false;
                }
                if (distPerpendicular != a.distPerpendicular)
                {
                    return false;
                }
                return true;
            }

            public float DistanceFromEndOf(TextChunk other)
            {
                float distance = distParallelStart - other.distParallelEnd;
                return distance;
            }

            int IComparable<TextChunk>.CompareTo(TextChunk rhs)
            {
                if (this == rhs)
                {
                    return 0;
                }

                int rslt;
                rslt = orientationMagnitude.CompareTo(rhs.orientationMagnitude);
                if (rslt != 0)
                {
                    return rslt;
                }

                rslt = distPerpendicular.CompareTo(rhs.distPerpendicular);
                if (rslt != 0)
                {
                    return rslt;
                }
                rslt = (distParallelStart < rhs.distParallelStart ? -1 : 1);

                return rslt;
            }
        }
    }
}