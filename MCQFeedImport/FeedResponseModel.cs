using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MCQFeedImport
{
   public class FeedResponseModel
    {
        public bool IsFileProcessSuccessful { get; set; }
        public int FeedFileCount { get; set; }
        public string FeedFileName { get; set; }

    }
}
