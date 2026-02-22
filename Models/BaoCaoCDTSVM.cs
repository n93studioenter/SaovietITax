using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Taxweb.Models
{
    public class BaoCaoCDTSVM
    {
        public List<CDTSVM> CDTS { get; set; }
        public List<QTongHopCTVM> QTongHop { get; set; }
    }
    public class CDTSVM
    {
        public string MaSo { get; set; }
        public double DauNam { get; set; }
        public double CuoiKy { get; set; }
    }
    public class QTongHopCTVM
    {
        public string MaSo { get; set; }
        public double DkNo { get; set; }
        public double DkCo { get; set; }
        public double PsNo { get; set; }
        public double PsCo { get; set; }
        public double CkNo { get; set; }
        public double CkCo
        {
            get; set;
        }
    }
}