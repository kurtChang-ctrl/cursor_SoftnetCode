using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BaseWeb.Models
{
    public class SimulationDto
    {
        [Required]
        [StringLength(20)]
        public string Name { get; set; }

        [Required]
        [StringLength(20)]
        public string Account { get; set; }

        [Required]
        [StringLength(32)]
        public string Pwd { get; set; }

        [Required]
        [StringLength(10)]
        public string DeptId { get; set; }
    }
}
