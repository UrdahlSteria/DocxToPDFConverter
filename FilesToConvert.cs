using System.Collections.Generic;
using System.Linq;

namespace docxToPdfConverter
{
    public class FilesToConvert
    {
        public List<string> filePaths { get; set; }

        public FilesToConvert()
        {
            this.filePaths = new List<string>();
        }

        public FilesToConvert(string[] paths)
        {
            this.filePaths = paths.ToList<string>();
        }

        public string FilesToTextBox()
        {
            var temp = "";
            foreach (var path in this.filePaths)
            {
                temp = temp + "\r\n" + path;
            }
            return temp;
        }
    }
}