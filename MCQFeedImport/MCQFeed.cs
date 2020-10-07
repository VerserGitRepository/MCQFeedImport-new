namespace MCQFeedImport
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("MCQFeed")]
    public partial class MCQFeed
    {
        public int ID { get; set; }

        public int? AssetNumber { get; set; }

        public int? MCQ_ID { get; set; }

        public int? ServiceNumber { get; set; }

        [StringLength(10)]
        public string Swap_Assure_Eligible { get; set; }

        [StringLength(50)]
        public string Make { get; set; }

        [StringLength(50)]
        public string Model { get; set; }

        [StringLength(50)]
        public string IMEI { get; set; }

        [StringLength(50)]
        public string Customer_Contract_Start_Date { get; set; }

        public DateTime? CreatedDate { get; set; }
        public DateTime? UpdatedDate { get; set; }

        [StringLength(50)]
        public string FeedVendor { get; set; }
    }
}
