using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.IO;

namespace Minerva.DAO
{
    class WordDAO
    {
        private string wordPath;
        private Document doc;
        private Dictionary<string, DataTable> tableSet;
        private Dictionary<string, object> fieldSet;

        public WordDAO(string wordPath)
        {
            this.wordPath = wordPath;
            if (File.Exists(wordPath))
            {
                doc = new Document(wordPath);
            }
        }

        public WordDAO SetTableSet(Dictionary<string, DataTable> tableSet)
        {
            this.tableSet = tableSet;
            return this;
        }

        public WordDAO SetFieldSet(Dictionary<string, object> fieldSet)
        {
            this.fieldSet = fieldSet;
            return this;
        }

        public WordDAO Process()
        {
            if (null != tableSet)
            {
                tableSet.ToList()
                    .ForEach(p =>
                    {
                        doc.MailMerge.ExecuteWithRegions(p.Value);
                    });
            }

            if (null != fieldSet)
            {
                doc.MailMerge.Execute(fieldSet.Keys.ToArray(),fieldSet.Values.ToArray());
            }

            return this;
        }

        public WordDAO SaveAs(string newFilePath)
        {
            doc.Save(newFilePath, SaveFormat.Docx);
            return this;
        }

    }
}
