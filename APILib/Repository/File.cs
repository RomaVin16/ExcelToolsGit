using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace APILib.Repository;

[Table("files")]
public class File
{
    [Key]
    [Column("id")]
    public Guid Id { get; set; }

    [Required]
    [Column("filename")]
    public string FileName { get; set; }

    [Required]
    [Column("savedate")]
    public DateTime SaveDate { get; set; }

    [Required]
    [Column("state")]
    public FileState State { get; set; }

    [Required]
    [Column("sizebytes")]
    public long SizeBytes { get; set; }

    [Required]
    [Column("operation")]
    public int OperationId { get; set; }

    [ForeignKey("OperationId")]
    public FileOperation Operation { get; set; }

    public enum FileState
    {
        Active = 0,
        Removed = 1
    }
}