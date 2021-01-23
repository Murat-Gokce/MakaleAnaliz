using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MakaleAnalizWebApp.Models
{
    public class Result
    {
        public int id { get; set; }
        public string message { get; set; }
        public bool isSuccess { get; set; }
        int level;
        public Result subResult { get; set; }
    }
}