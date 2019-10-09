using System.Collections.Generic;

namespace MainProject
{
    public class WordInfo
    {
        public string RusWord { get; set; }
        public string EngWord { get; set; }
        public List<WordDateInfo> Dates { get; set; } = new List<WordDateInfo>();

        public override string ToString()
        {
            return RusWord;
        }
    }
}