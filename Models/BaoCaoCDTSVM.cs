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
        public List<LCTT> LCTTs { get; set; }
        public List<TTLCTT> TTLCTTs { get; set; }=new List<TTLCTT>();
    }
    public class CDTSVM
    {
        public string MaSo { get; set; }
        public double DauNam { get; set; }
        public double CuoiKy { get; set; }
    }
    public class TTLCTT
    {
        public string MaSo { get; set; }    
        public double Namnay { get; set; }  

    }
    public class LCTT
    {
        public string MaSo { get; set; }
        public string TKNo { get; set; }
        public string TKCo { get; set; }
        public double KyTruoc { get; set; }
        public double KyNay { get; set; }
        public string TenE { get; set; }
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