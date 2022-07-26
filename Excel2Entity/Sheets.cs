using System.Collections.Generic;

namespace Excel2Entity
{
	public class Sheets
	{
		/// <summary>取込対象</summary>
		public bool Target { get; set; } = true;

		/// <summary>論理名</summary>
		public string LogicalName { get; set; }

		/// <summary>物理名</summary>
		public string PhysicsName { get; set; }

		/// <summary>クラス名</summary>
		public string ClassName { get; set; }

		/// <summary></summary>
		public List<Columns> ColumnsList { get; set; } = new List<Columns>();
	}
}
