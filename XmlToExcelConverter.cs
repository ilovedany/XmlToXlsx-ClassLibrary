using Spire.Xls;
using System.Xml;
using System.Text;
using System.Diagnostics;

namespace XmlToXlsx
{
    public class XmlToExcelConverter
    {
        public void ConvertXmlToXlsxPath(string pathToXml,string pathToSave, int delOrNo, int openOrNo)
        {

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            XmlDocument xDoc = new XmlDocument();

            xDoc.Load(pathToXml); 

            XmlElement? xRoot = xDoc.DocumentElement;
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            WorksheetColumns worksheetColumns = new WorksheetColumns();
            FileInfo xmlFile = new FileInfo(pathToXml);

            string[] columns = xRoot.GetAttribute("headers").Split(" ");
            int countTags = columns.Length;

            int count_G_List = 0;
            int countLine = 0;

            List<string> listObject = new List<string>();
            List<int> lenCount = new List<int>();
            HashSet<string> tags = new HashSet<string> { };
            HashSet<string> temp = new HashSet<string> { };

            string[] tags_array;

            string[] strings;
            var strLen = 0;

            int firstStartPosition;
            int startPosition;

            int maxLen = 0;

            countLine++;
            worksheetColumns.WorkingWithXml(workbook, countLine, columns);
            worksheet.Range[1, 1, 1, countTags].Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            worksheet.Range[1, 1, 1, countTags].Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            worksheet.Range[1,1,1,countTags].Style.Font.IsBold = true;
            worksheet.Range[1,1,1,countTags].Style.Font.Size = 11;

            foreach (XmlElement xnode in xRoot)
            {
                foreach (XmlNode userNode in xnode.ChildNodes)
                {

                    foreach (XmlNode profList in userNode.ChildNodes)
                    {
                        if (columns.Contains(profList.Name))
                        {
                            temp.Add(profList.Name);
                        }
                        foreach (XmlNode profList2 in profList.ChildNodes)
                        {
                            if (profList2.NodeType == XmlNodeType.Element)
                            {
                                tags.Add(profList2.Name);
                            }
                        }
                    }
                }
            }
            startPosition = temp.Count;
            tags_array = tags.ToArray();

            foreach (XmlNode xnode in xRoot)
            {
                foreach (XmlNode userNode in xnode.ChildNodes)
                {
                    firstStartPosition = startPosition;
                    countLine += maxLen + 1;
                    maxLen = 0;

                    worksheetColumns.WorkingWithXml(userNode, workbook, countLine, columns, 0);

                    foreach (XmlNode profList in userNode.ChildNodes)
                    {
                        foreach (XmlNode profList2 in profList.ChildNodes)
                        {

                            for (int i = 0; i < tags_array.Count(); i++)
                            {
                                if (profList2.Name == tags_array[i])
                                {

                                    foreach (XmlNode userList in profList2.ChildNodes)
                                    {
                                        listObject.Add(userList.Name);
                                    }

                                    strings = listObject.ToArray();
                                    strLen = strings.Length;

                                    worksheetColumns.WorkingWithXml(profList2, workbook, countLine, strings, firstStartPosition);
                                    countLine++;
                                    count_G_List++;

                                    listObject.Clear();
                                }
                            }
                        }
                        lenCount.Add(count_G_List);

                        maxLen = lenCount.Max();
                        countLine -= count_G_List;
                        firstStartPosition += strLen;
                        strLen = 0;

                        count_G_List = 0;
                    }
                    countLine--;

                    worksheetColumns.WorkingWithXml(workbook, countLine + 1, countLine + maxLen, startPosition);

                    worksheet.Range[countLine+1, 1, countLine + maxLen, countTags].Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;

                    worksheet.Range[countLine + maxLen, 1, countLine + maxLen, countTags].Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;

                    lenCount.Clear();
                }

            }
            worksheet.AllocatedRange.AutoFitColumns();

            workbook.SaveToFile(pathToSave, ExcelVersion.Version2016);

            if (openOrNo == 1)
            {
                Process.Start(new ProcessStartInfo(pathToSave) { UseShellExecute = true });
            }
        
            
            if (delOrNo == 1)
            {
                xmlFile.Delete();
            }
            

        }


        public void ConvertXmlToXlsxString(string xmlString,string pathToSave, int openOrNo)
        {

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            XmlDocument xDoc = new XmlDocument();

            xDoc.LoadXml(xmlString);

            XmlElement? xRoot = xDoc.DocumentElement;
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            WorksheetColumns worksheetColumns = new WorksheetColumns();

            string[] columns = xRoot.GetAttribute("headers").Split(" ");
            int countTags = columns.Length;

            int count_G_List = 0;
            int countLine = 0;

            List<string> listObject = new List<string>();
            List<int> lenCount = new List<int>();
            HashSet<string> tags = new HashSet<string> { };
            HashSet<string> temp = new HashSet<string> { };

            string[] tags_array;

            string[] strings;
            var strLen = 0;

            int firstStartPosition;
            int startPosition;

            int maxLen = 0;

            countLine++;
            worksheetColumns.WorkingWithXml(workbook, countLine, columns);
            worksheet.Range[1, 1, 1, countTags].Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            worksheet.Range[1, 1, 1, countTags].Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            worksheet.Range[1,1,1,countTags].Style.Font.IsBold = true;
            worksheet.Range[1,1,1,countTags].Style.Font.Size = 11;

            foreach (XmlElement xnode in xRoot)
            {
                foreach (XmlNode userNode in xnode.ChildNodes)
                {

                    foreach (XmlNode profList in userNode.ChildNodes)
                    {
                        if (columns.Contains(profList.Name))
                        {
                            temp.Add(profList.Name);
                        }
                        foreach (XmlNode profList2 in profList.ChildNodes)
                        {
                            if (profList2.NodeType == XmlNodeType.Element)
                            {
                                tags.Add(profList2.Name);
                            }
                        }
                    }
                }
            }
            startPosition = temp.Count;
            tags_array = tags.ToArray();

            foreach (XmlNode xnode in xRoot)
            {
                foreach (XmlNode userNode in xnode.ChildNodes)
                {
                    firstStartPosition = startPosition;
                    countLine += maxLen + 1;
                    maxLen = 0;

                    worksheetColumns.WorkingWithXml(userNode, workbook, countLine, columns, 0);

                    foreach (XmlNode profList in userNode.ChildNodes)
                    {
                        foreach (XmlNode profList2 in profList.ChildNodes)
                        {

                            for (int i = 0; i < tags_array.Count(); i++)
                            {
                                if (profList2.Name == tags_array[i])
                                {

                                    foreach (XmlNode userList in profList2.ChildNodes)
                                    {
                                        listObject.Add(userList.Name);
                                    }

                                    strings = listObject.ToArray();
                                    strLen = strings.Length;

                                    worksheetColumns.WorkingWithXml(profList2, workbook, countLine, strings, firstStartPosition);
                                    countLine++;
                                    count_G_List++;

                                    listObject.Clear();
                                }
                            }
                        }
                        lenCount.Add(count_G_List);

                        maxLen = lenCount.Max();
                        countLine -= count_G_List;
                        firstStartPosition += strLen;
                        strLen = 0;

                        count_G_List = 0;
                    }
                    countLine--;

                    worksheetColumns.WorkingWithXml(workbook, countLine + 1, countLine + maxLen, startPosition);

                    worksheet.Range[countLine+1, 1, countLine + maxLen, countTags].Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;

                    worksheet.Range[countLine + maxLen, 1, countLine + maxLen, countTags].Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;

                    lenCount.Clear();
                }

            }
            worksheet.AllocatedRange.AutoFitColumns();
            workbook.SaveToFile(pathToSave, ExcelVersion.Version2016);
            if (openOrNo==1)
            {
                Process.Start(new ProcessStartInfo(pathToSave) { UseShellExecute = true });
            }
            
            


        }
    }
}