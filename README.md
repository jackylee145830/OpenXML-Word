# OpenXML-Word
#### Word套版動態多筆


擴充原本MakeDocxFile
```
public string MakeDocxFileTable(string template, string fileName, Dictionary<string, string> fields, List<Dictionary<string, string>> mList)
        {
            template = template + ".docx";
            fileName = fileName + ".docx";
            string tempFile = Path.ChangeExtension(Path.GetTempFileName(), "docx");
            File.Copy(templatePath + template, tempFile);

            using (WordprocessingDocument wd = WordprocessingDocument.Open(tempFile, true))
            {
                
                //Replace document body
                parse(wd.MainDocumentPart, fields);
                foreach (HeaderPart hp in wd.MainDocumentPart.HeaderParts)
                {
                    parse(hp, fields);
                }
                foreach (FooterPart fp in wd.MainDocumentPart.FooterParts)
                {
                    parse(fp, fields);
                }

                parseTable(wd.MainDocumentPart, mList);
            }
            string outputFile = outputPath + Path.GetFileName(fileName);
            File.Copy(tempFile, outputFile, true);
            return outputFile;
        }
```


```
private void parseTable(OpenXmlPart oxp, List<Dictionary<string, string>> mList)
        {
            string xmlString = null;
            using (StreamReader sr = new StreamReader(oxp.GetStream()))
            {
                xmlString = sr.ReadToEnd();
            }

            //找到table位置
            int tempPosition = xmlString.IndexOf("[$" + mList[0].First().Key + "$]");
            if (tempPosition > -1)
            {
                //找出字串
                string firstString = xmlString.Substring(0, tempPosition);
                int startPosition = firstString.LastIndexOf("<w:tr ");
                string secondString = xmlString.Substring(startPosition, xmlString.Length - startPosition - 1);
                int endPosition = secondString.IndexOf("</w:tr>");
                string rowString = secondString.Substring(0, endPosition + 7);
                string newRowString = "";
                for (int i = 0; i < mList.Count; i++)
                {
                    string temp = rowString;
                    foreach (KeyValuePair<string, string> item in mList[i])
                    {
                        temp = temp.Replace("[$" + item.Key + "$]", item.Value);
                    }
                    newRowString = newRowString + temp;
                }

                xmlString = xmlString.Replace(rowString, newRowString);

                using (StreamWriter sw = new StreamWriter(oxp.GetStream(FileMode.Create)))
                {
                    sw.Write(xmlString);
                }
            }
        }
```

#### 應用

```
            Dictionary<string, string> dct = new Dictionary<string, string>();
            dct.Add("txtSN","hello");
           
            //插入table
            List<Dictionary<string, string>> mList = new List<Dictionary<string, string>>();
            for (int i = 0; i < 10; i++)
            {
                Dictionary<string, string> m1 = new Dictionary<string, string>();
                m1.Add("step1", "A"+i);
                m1.Add("job1", "B" + i);
                m1.Add("signp1", "C" + i);
                m1.Add("ss1", "D" + i);
                m1.Add("ssg1", "E" + i);
                m1.Add("sst1", "F" + i);
                mList.Add(m1);
            }

            WordHelper helper = new WordHelper();
            helper.MakeDocxFileTable("品質異常處理單匯出樣板", "Hello", dct, mList);
```
