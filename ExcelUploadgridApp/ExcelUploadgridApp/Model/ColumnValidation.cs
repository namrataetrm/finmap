using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;

namespace ExcelUploadgridApp.Model
{
    public class ColumnValidation
    {
        [Column("id")]
        public int Id { get; set; }

        [Column("validation_name")]
        public string ValidationName { get; set; }

        [Column("constraint_expression")]
        public string ConstraintExpression { get; set; }

        // Navigation property to ColumnDetail
        public ICollection<ColumnDetail> ColumnDetails { get; set; }
    }
}
