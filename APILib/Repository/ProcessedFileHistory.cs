using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace APILib.Repository;

[Table("processedfilehistory")]
public class ProcessedFileHistory
{
    [Key]
    [Column("id")]
    public Guid Id { get; set; }

    [Required]
    [Column("file_id")]
    public Guid FileId { get; set; }

    [Required]
    [Column("user_id")]
    public string UserId { get; set; }

    [Required]
    [Column("processed_at")]
    public DateTime ProcessedAt { get; set; }

    [Required]
    [Column("operation")]
    public string Operation { get; set; }

    [ForeignKey("FileId")]
    public File File { get; set; }

    [ForeignKey("UserId")]
    public ApplicationUser User { get; set; }
}