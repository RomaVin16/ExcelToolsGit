using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace APILib.Repository;

[Table("fileoperations")]
public class FileOperation
{
    [Key]
    [Column("id")]
    public int Id { get; set; }

    [Required]
    [Column("operationname")]
    public string OperationName { get; set; }
}
public enum FileOperations
{
    Upload = 1,
    Download = 2,
    Clean = 3,
    DuplicateRemove = 4,
    Merge = 5,
    Split = 6,
    SplitColumn = 7,
    Rotate = 8
}