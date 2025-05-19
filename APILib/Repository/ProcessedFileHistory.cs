using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;

namespace APILib.Repository
{
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
		public Files File { get; set; } = null!;

		[ForeignKey("UserId")]
		public ApplicationUser User { get; set; } = null!;
	}
}
