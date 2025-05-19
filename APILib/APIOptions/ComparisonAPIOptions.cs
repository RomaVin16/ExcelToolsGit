using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace APILib.APIOptions
{
	public class ComparisonAPIOptions: APIOptionsBase
	{
		/// <summary>
		/// Имя исходного файла 
		/// </summary>
		public Guid SourceFileId { get; set; }

		/// <summary>
		/// Имя исходного файла 
		/// </summary>
		public Guid ModifiedFileId { get; set; }

		public string[]? Id { get; set; }

		/// <summary>
		/// Количество строк, которые требуется использовать как заголовки 
		/// </summary>
		public int[] HeaderRows { get; set; }
	}
}
