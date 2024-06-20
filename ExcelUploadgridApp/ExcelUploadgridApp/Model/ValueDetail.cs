using System.ComponentModel.DataAnnotations.Schema;

namespace ExcelUploadgridApp.Model
{
    public class ValueDetail
    {
        [Column("id")]
        public int Id { get; set; }

        [Column("column_id")]
        public int ColumnId { get; set; }

        [Column("row_id")]
        public int RowId { get; set; }

        [Column("value")]
        public string Value { get; set; }

      

        public ColumnDetail ColumnDetail { get; set; }
    }
}
