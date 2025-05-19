using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace API.Models
{
    [Table("otpcodes")]
    public class OtpCode
    {
        [Key]
        [Column("id")]
        public int Id { get; set; }

        [Required]
        [Column("phonenumber", TypeName = "text")]
        public string PhoneNumber { get; set; } = null!;

        [Required]
        [Column("code", TypeName = "text")]
        public string Code { get; set; } = null!;

        [Required]
        [Column("expirytime", TypeName = "timestamp with time zone")]
        public DateTime ExpiryTime { get; set; }
    }
}
