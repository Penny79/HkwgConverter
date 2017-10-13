namespace HkwgConverter.Model
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Transaction")]
    public partial class Transaction
    {
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        public DateTime DeliveryDate { get; set; }

        public int Version { get; set; }

        [Required]
        [StringLength(200)]
        public string CsvFile { get; set; }
        
        [StringLength(200)]
        public string FlexPosFile { get; set; }
        
        [StringLength(200)]
        public string FlexNegFile { get; set; }

        [StringLength(200)]
        public string ConfirmedDealFile { get; set; }

        [Required]      
        public DateTime CreateDate { get; set; }
      
        public DateTime UpdateDate { get; set; }
    }
}
