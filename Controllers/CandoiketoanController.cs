using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using WebGrease.Activities;

namespace Taxweb.Controllers
{
    public class YourModel
    {
        public string path { get; set; }
        public string maso { get; set; }
        public string DauNam { get; set; }
        public string CuoiKy { get; set; }
        public int idparent { get; set; } // QUAN TRỌNG: Dùng List<string>
    }
    public class CandoiketoanController : Controller
    {
        // GET: Candoiketoan
        public static string dbPath = "";
        string password, connectionString;
        public ActionResult Index(string path)
        {
            if(!string.IsNullOrEmpty(path))
            {
                dbPath = path;
                ViewBag.path = dbPath;
                OleDbConnection conn = null;
                string password = "1@35^7*9)1";
                connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};Jet OLEDB:Database Password={password};";
                conn = new OleDbConnection(connectionString);
                conn.Open();

                if (!TableExists(conn, "tbCDTS"))
                {
                    CreateTableCDTS(conn, "tbCDTS");
                }
                if (!TableExists(conn, "tbCDTSChild"))
                {
                    CreateTableCDTSchild(conn, "tbCDTSChild");
                }



                string query = "SELECT * FROM CDTS";
                DataTable CDTS = ExecuteQuery(query, null);
                var model= CDTS.AsEnumerable().ToList();

                 query = "SELECT * FROM License";
                DataTable data = ExecuteQuery(query, null);
                ViewBag.NamTC = data.Rows[0]["NamTC"].ToString();
                return View(model);
            }
            return View();
        }
        public bool TableExists(OleDbConnection connection, string tableName)
        {
            try
            {
                // Kiểm tra sự tồn tại của bảng
                System.Data.DataTable schemaTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                foreach (DataRow row in schemaTable.Rows)
                {
                    if (row["TABLE_NAME"].ToString().Equals(tableName, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi khi kiểm tra bảng: {ex.Message}");
            }
            return false;
        }
        static void CreateTableCDTS(OleDbConnection connection, string tableName)
        {
            string createTableQuery = $@"
        CREATE TABLE {tableName} (
            ID AUTOINCREMENT PRIMARY KEY, 
            kyKKhaiTuNgay DATETIME,
            kyKKhaiDenNgay DATETIME,
            ngayKy DATETIME,
            nguoilapbieu TEXT,
            ketoantruong TEXT,
            nguoidaidien TEXT,
            sochungchi TEXT,
            donvicungcap TEXT  ,
            nam NUMBER  
        );";

            using (OleDbCommand command = new OleDbCommand(createTableQuery, connection))
            {
                command.ExecuteNonQuery();
            }
        }
        static void CreateTableCDTSchild(OleDbConnection connection, string tableName)
        {
            string createTableQuery = $@"
        CREATE TABLE {tableName} (
            ID AUTOINCREMENT PRIMARY KEY, 
            MaSo NUMBER,
            DauNam NUMBER,
            CuoiKy NUMBER,
            Idparent NUMBER
        );";

            using (OleDbCommand command = new OleDbCommand(createTableQuery, connection))
            {
                command.ExecuteNonQuery();
            }
        }
     
        [HttpPost]
        public int CreateData(string path,string nlbieu,string giamdoc,string ketoantruong,string chungchi,string dvcc,DateTime ngayky,int nam)
        {
            dbPath = path;
            try
            {
                string qrchk = @"SELECT * FROM tbCDTS where nam= ? ";
                var parameterss = new OleDbParameter[]
                {
                new OleDbParameter("?",nam),
                };
                var chkCDtS = ExecuteQuery(qrchk, parameterss);
                if (chkCDtS.Rows.Count == 0)
                {
                    string query = "INSERT INTO tbCDTS (nguoilapbieu,nguoidaidien,ketoantruong,sochungchi,donvicungcap,ngayKy,nam) VALUES (?,?,?,?,?,?,?)";
                    OleDbParameter[] parameters = new OleDbParameter[]
                    {
                new OleDbParameter("?", nlbieu),
                new OleDbParameter("?", giamdoc),
                new OleDbParameter("?", ketoantruong),
                new OleDbParameter("?", chungchi),
                new OleDbParameter("?", dvcc),
                new OleDbParameter("?", ngayky),
                 new OleDbParameter("?", nam),
                    };
                    var getID = ExecuteQueryResult(query, parameters);
                    return getID;
                }
                else
                {
                    string query = "UPDATE tbCDTS SET nguoilapbieu=?,nguoidaidien=?,ketoantruong=?,sochungchi=?,donvicungcap=?,ngayKy=? WHERE nam=?";
                    OleDbParameter[] parameters = new OleDbParameter[]
                        {
                            new OleDbParameter("?", nlbieu),
                            new OleDbParameter("?", giamdoc),
                            new OleDbParameter("?", ketoantruong),
                            new OleDbParameter("?", chungchi),
                            new OleDbParameter("?", dvcc),
                            new OleDbParameter("?", ngayky),
                             new OleDbParameter("?", nam),
                        };
                     ExecuteQueryResult(query, parameters);
                    var getID = chkCDtS.Rows[0].Field<int>("ID");
                    return getID;

                }
            }
            catch (Exception ex)
            {
                return -1;
            }
        }

        [HttpGet]
        public ActionResult CreateTaxXMLFull(int nam)
        {
            string qrchk = @"SELECT * FROM tbCDTS where nam= ? ";
            var parameterss = new OleDbParameter[]
            {
                new OleDbParameter("?",nam),
            };
            var chkCDtS = ExecuteQuery(qrchk, parameterss);

            qrchk = @"SELECT * FROM tbCDTSChild where Idparent= ? ";
             parameterss = new OleDbParameter[]
            {
                new OleDbParameter("?",chkCDtS.Rows[0]["ID"].ToString()),
            };
             chkCDtS = ExecuteQuery(qrchk, parameterss);

            var tuNgay = new DateTime(nam, 1, 1);
            var denNgay = new DateTime(nam, 12, 31);
            //Lấy thong tin công ty

            string query = "SELECT * FROM License";
            DataTable data = ExecuteQuery(query, null);
            string tencty= data.Rows[0]["TenCty"].ToString();
            string diachi= data.Rows[0]["DiaChi"].ToString();    
            string mst = data.Rows[0]["MaSoThue"].ToString();
            string xmlContent = $@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""no""?><HSoThueDTu xmlns=""http://kekhaithue.gdt.gov.vn/TKhaiThue"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">
  <HSoKhaiThue id=""_NODE_TO_SIGN"">
    <TTinChung>
      <TTinDVu>
        <maDVu>HTKK</maDVu>
        <tenDVu>HỖ TRỢ KÊ KHAI THUẾ</tenDVu>
        <pbanDVu>5.3.1</pbanDVu>
        <ttinNhaCCapDVu>8B7771255E46CCFB75A8DE7D33CA0D12</ttinNhaCCapDVu>
      </TTinDVu>
      <TTinTKhaiThue>
        <TKhaiThue>
          <maTKhai>684</maTKhai>
          <tenTKhai>Bộ báo cáo tài chính (B01b - DNN)(TT133/2016/TT-BTC)</tenTKhai>
          <moTaBMau>(Ban hành theo Thông tư số 133/2016/TT-BTC ngày 26/8/2016 của Bộ Tài chính)</moTaBMau>
          <pbanTKhaiXML>2.3.2</pbanTKhaiXML>
          <loaiTKhai>C</loaiTKhai>
          <soLan>0</soLan>
          <KyKKhaiThue>
            <kieuKy>Y</kieuKy>
            <kyKKhai>2024</kyKKhai>
            <kyKKhaiTuNgay>01/01/2024</kyKKhaiTuNgay>
            <kyKKhaiDenNgay>31/12/2024</kyKKhaiDenNgay>
            <kyKKhaiTuThang/>
            <kyKKhaiDenThang/>
          </KyKKhaiThue>
          <maCQTNoiNop>71701</maCQTNoiNop>
          <tenCQTNoiNop>Tp. Vũng Tàu - Đội Thuế liên huyện Vũng Tàu - Côn Đảo</tenCQTNoiNop>
          <ngayLapTKhai>2025-03-18</ngayLapTKhai>
          <GiaHan>
            <maLyDoGiaHan>
            </maLyDoGiaHan>
            <lyDoGiaHan>
            </lyDoGiaHan>
          </GiaHan>
          <nguoiKy>Nguyễn Trường Giang</nguoiKy>
          <ngayKy>2025-03-18</ngayKy>
          <nganhNgheKD>
          </nganhNgheKD>
        </TKhaiThue>
        <NNT>
          <mst>3500807781</mst>
          <tenNNT>CÔNG TY TNHH THƯƠNG MẠI-XÂY DỰNG ĐỨC THỊNH</tenNNT>
          <dchiNNT>44 Lê Lai - Phường 3</dchiNNT>
          <phuongXa>
          </phuongXa>
          <maHuyenNNT>71701</maHuyenNNT>
          <tenHuyenNNT>Thành phố Vũng Tàu</tenHuyenNNT>
          <maTinhNNT>717</maTinhNNT>
          <tenTinhNNT>Bà Rịa - Vũng Tàu</tenTinhNNT>
          <dthoaiNNT>
          </dthoaiNNT>
          <faxNNT>
          </faxNNT>
          <emailNNT>
          </emailNNT>
        </NNT>
      </TTinTKhaiThue>
    </TTinChung>
    <CTieuTKhaiChinh>
      <bctcDaKiemToan>0</bctcDaKiemToan>
      <maYKienKToan>0</maYKienKToan>
      <tenYKienKToan>
      </tenYKienKToan>
      <ThuyetMinh>
        <ct100>
        </ct100>
        <ct110>
        </ct110>
        <ct120>
        </ct120>
        <ct121>
        </ct121>
        <ct122>
        </ct122>
        <ct123>
        </ct123>
        <ct130>
        </ct130>
        <ct131>
        </ct131>
        <ct132>
        </ct132>
        <ct133>
        </ct133>
        <ct134>
        </ct134>
        <ct135>
        </ct135>
        <ct140>
        </ct140>
        <ct141>
        </ct141>
        <ct142>
        </ct142>
        <ct150>
        </ct150>
        <ct151>
        </ct151>
        <ct152>
        </ct152>
        <ct200>
        </ct200>
        <ct210>
        </ct210>
        <ct211>
        </ct211>
        <ct212>
        </ct212>
        <ct213>
        </ct213>
        <ct214>
        </ct214>
        <ct215>
        </ct215>
        <ct220>
        </ct220>
        <ct221>
        </ct221>
        <ct222>
        </ct222>
        <ct230>
        </ct230>
        <ct231>
        </ct231>
        <ct232>
        </ct232>
        <ct240>
        </ct240>
        <ct250>
        </ct250>
        <ct251>
        </ct251>
        <ct252>
        </ct252>
        <ct253>
        </ct253>
        <ct260>
        </ct260>
        <ct300>
        </ct300>
        <ct400>
        </ct400>
        <ct410>
        </ct410>
        <ct411>
        </ct411>
        <ct412>
        </ct412>
        <ct413>
        </ct413>
        <ct414>
        </ct414>
        <ct415>
        </ct415>
        <ct416>
        </ct416>
        <ct417>
        </ct417>
        <ct418>
        </ct418>
        <ct420>
        </ct420>
        <ct421>
        </ct421>
        <ct422>
        </ct422>
        <ct423>
        </ct423>
        <ct424>
        </ct424>
        <ct425>
        </ct425>
        <ct426>
        </ct426>
        <ct427>
        </ct427>
        <ct500>
        </ct500>
        <ct511>
        </ct511>
        <ct512>
        </ct512>
        <ct513>
        </ct513>
        <ct514>
        </ct514>
        <ct515>
        </ct515>
        <ct516>
        </ct516>
        <ct517>
        </ct517>
        <ct600>
        </ct600>
      </ThuyetMinh>
      <SoCuoiNam>
        <ct100>10525938376</ct100>
        <ct110>1643352088</ct110>
        <ct120>0</ct120>
        <ct121>0</ct121>
        <ct122>0</ct122>
        <ct123>0</ct123>
        <ct130>300115794</ct130>
        <ct131>300052500</ct131>
        <ct132>63294</ct132>
        <ct133>0</ct133>
        <ct134>0</ct134>
        <ct135>0</ct135>
        <ct140>8343365072</ct140>
        <ct141>8343365072</ct141>
        <ct142>0</ct142>
        <ct150>239105422</ct150>
        <ct151>239105422</ct151>
        <ct152>0</ct152>
        <ct200>10067340</ct200>
        <ct210>0</ct210>
        <ct211>0</ct211>
        <ct212>0</ct212>
        <ct213>0</ct213>
        <ct214>0</ct214>
        <ct215>0</ct215>
        <ct220>8400660</ct220>
        <ct221>1089445454</ct221>
        <ct222>-1081044794</ct222>
        <ct230>0</ct230>
        <ct231>0</ct231>
        <ct232>0</ct232>
        <ct240>0</ct240>
        <ct250>0</ct250>
        <ct251>0</ct251>
        <ct252>0</ct252>
        <ct253>0</ct253>
        <ct260>1666680</ct260>
        <ct300>10536005716</ct300>
        <ct400>9868699</ct400>
        <ct410>9868699</ct410>
        <ct411>0</ct411>
        <ct412>0</ct412>
        <ct413>9868699</ct413>
        <ct414>0</ct414>
        <ct415>0</ct415>
        <ct416>0</ct416>
        <ct417>0</ct417>
        <ct418>0</ct418>
        <ct420>0</ct420>
        <ct421>0</ct421>
        <ct422>0</ct422>
        <ct423>0</ct423>
        <ct424>0</ct424>
        <ct425>0</ct425>
        <ct426>0</ct426>
        <ct427>0</ct427>
        <ct500>10526137017</ct500>
        <ct511>10000000000</ct511>
        <ct512>0</ct512>
        <ct513>0</ct513>
        <ct514>0</ct514>
        <ct515>0</ct515>
        <ct516>0</ct516>
        <ct517>526137017</ct517>
        <ct600>10536005716</ct600>
      </SoCuoiNam>
      <SoDauNam>
        <ct100>10735807976</ct100>
        <ct110>1006410597</ct110>
        <ct120>0</ct120>
        <ct121>0</ct121>
        <ct122>0</ct122>
        <ct123>0</ct123>
        <ct130>300052500</ct130>
        <ct131>300052500</ct131>
        <ct132>0</ct132>
        <ct133>0</ct133>
        <ct134>0</ct134>
        <ct135>0</ct135>
        <ct140>9064217231</ct140>
        <ct141>9064217231</ct141>
        <ct142>0</ct142>
        <ct150>365127648</ct150>
        <ct151>365127648</ct151>
        <ct152>0</ct152>
        <ct200>115875516</ct200>
        <ct210>0</ct210>
        <ct211>0</ct211>
        <ct212>0</ct212>
        <ct213>0</ct213>
        <ct214>0</ct214>
        <ct215>0</ct215>
        <ct220>109208844</ct220>
        <ct221>1089445454</ct221>
        <ct222>-980236610</ct222>
        <ct230>0</ct230>
        <ct231>0</ct231>
        <ct232>0</ct232>
        <ct240>0</ct240>
        <ct250>0</ct250>
        <ct251>0</ct251>
        <ct252>0</ct252>
        <ct253>0</ct253>
        <ct260>6666672</ct260>
        <ct300>10851683492</ct300>
        <ct400>325352842</ct400>
        <ct410>325352842</ct410>
        <ct411>314559628</ct411>
        <ct412>0</ct412>
        <ct413>10793214</ct413>
        <ct414>0</ct414>
        <ct415>0</ct415>
        <ct416>0</ct416>
        <ct417>0</ct417>
        <ct418>0</ct418>
        <ct420>0</ct420>
        <ct421>0</ct421>
        <ct422>0</ct422>
        <ct423>0</ct423>
        <ct424>0</ct424>
        <ct425>0</ct425>
        <ct426>0</ct426>
        <ct427>0</ct427>
        <ct500>10526330650</ct500>
        <ct511>10000000000</ct511>
        <ct512>0</ct512>
        <ct513>0</ct513>
        <ct514>0</ct514>
        <ct515>0</ct515>
        <ct516>0</ct516>
        <ct517>526330650</ct517>
        <ct600>10851683492</ct600>
      </SoDauNam>
      <nguoiLapBieu>
      </nguoiLapBieu>
      <keToanTruong>
      </keToanTruong>
      <ngayLap>2025-02-14</ngayLap>
      <nguoiDaiDienTheoPhapLuat>Nguyễn Trường Giang</nguoiDaiDienTheoPhapLuat>
    </CTieuTKhaiChinh>
    <PLuc>
      <PL_KQHDSXKD>
        <ThuyetMinh>
          <ct01/>
          <ct02/>
          <ct10/>
          <ct11/>
          <ct20/>
          <ct21/>
          <ct22/>
          <ct23/>
          <ct24/>
          <ct30/>
          <ct31/>
          <ct32/>
          <ct40/>
          <ct50/>
          <ct51/>
          <ct60/>
        </ThuyetMinh>
        <NamNay>
          <ct01>3800823318</ct01>
          <ct02>0</ct02>
          <ct10>3800823318</ct10>
          <ct11>3485810206</ct11>
          <ct20>315013112</ct20>
          <ct21>40448614</ct21>
          <ct22>0</ct22>
          <ct23>0</ct23>
          <ct24>537303654</ct24>
          <ct30>-181841928</ct30>
          <ct31>232219355</ct31>
          <ct32>40495575</ct32>
          <ct40>191723780</ct40>
          <ct50>9881852</ct50>
          <ct51>10075485</ct51>
          <ct60>-193633</ct60>
        </NamNay>
        <NamTruoc>
          <ct01>5334959897</ct01>
          <ct02>0</ct02>
          <ct10>5334959897</ct10>
          <ct11>4967482914</ct11>
          <ct20>367476983</ct20>
          <ct21>88313116</ct21>
          <ct22>0</ct22>
          <ct23>0</ct23>
          <ct24>699480366</ct24>
          <ct30>-243690267</ct30>
          <ct31>297656339</ct31>
          <ct32>0</ct32>
          <ct40>297656339</ct40>
          <ct50>53966072</ct50>
          <ct51>10793214</ct51>
          <ct60>43172858</ct60>
        </NamTruoc>
      </PL_KQHDSXKD>
      <PL_LCTTTT>
        <ThuyetMinh>
          <ct01/>
          <ct02/>
          <ct03/>
          <ct04/>
          <ct05/>
          <ct06/>
          <ct07/>
          <ct20/>
          <ct21/>
          <ct22/>
          <ct23/>
          <ct24/>
          <ct25/>
          <ct30/>
          <ct31/>
          <ct32/>
          <ct33/>
          <ct34/>
          <ct35/>
          <ct40/>
          <ct50/>
          <ct60/>
          <ct61/>
          <ct70/>
        </ThuyetMinh>
        <NamNay>
          <ct01>4219468120</ct01>
          <ct02>-3115911325</ct02>
          <ct03>-422930000</ct03>
          <ct04>0</ct04>
          <ct05>-24696000</ct05>
          <ct06>44465000</ct06>
          <ct07>-63498375</ct07>
          <ct20>636897420</ct20>
          <ct21>0</ct21>
          <ct22>0</ct22>
          <ct23>0</ct23>
          <ct24>0</ct24>
          <ct25>44071</ct25>
          <ct30>44071</ct30>
          <ct31>0</ct31>
          <ct32>0</ct32>
          <ct33>0</ct33>
          <ct34>0</ct34>
          <ct35>0</ct35>
          <ct40>0</ct40>
          <ct50>636941491</ct50>
          <ct60>1006410597</ct60>
          <ct61>0</ct61>
          <ct70>1643352088</ct70>
        </NamNay>
        <NamTruoc>
          <ct01>5898402426</ct01>
          <ct02>-3566934779</ct02>
          <ct03>-586650000</ct03>
          <ct04>0</ct04>
          <ct05>-2463806</ct05>
          <ct06>1090975000</ct06>
          <ct07>-8452000000</ct07>
          <ct20>-5618671159</ct20>
          <ct21>0</ct21>
          <ct22>0</ct22>
          <ct23>0</ct23>
          <ct24>0</ct24>
          <ct25>96399</ct25>
          <ct30>96399</ct30>
          <ct31>6500000000</ct31>
          <ct32>0</ct32>
          <ct33>0</ct33>
          <ct34>0</ct34>
          <ct35>0</ct35>
          <ct40>6500000000</ct40>
          <ct50>881425240</ct50>
          <ct60>124985357</ct60>
          <ct61>0</ct61>
          <ct70>1006410597</ct70>
        </NamTruoc>
      </PL_LCTTTT>
      <PL_CDTK>
        <SoDuDauKy>
          <No>
            <ct111>917519211</ct111>
            <ct1111>917519211</ct1111>
            <ct1112>0</ct1112>
            <ct112>88891386</ct112>
            <ct1121>88891386</ct1121>
            <ct1122>0</ct1122>
            <ct121>0</ct121>
            <ct128>0</ct128>
            <ct1281>0</ct1281>
            <ct1288>0</ct1288>
            <ct131>300052500</ct131>
            <ct133>365127648</ct133>
            <ct1331>365127648</ct1331>
            <ct1332>0</ct1332>
            <ct136>0</ct136>
            <ct1361>0</ct1361>
            <ct1368>0</ct1368>
            <ct138>0</ct138>
            <ct1381>0</ct1381>
            <ct1386>0</ct1386>
            <ct1388>0</ct1388>
            <ct141>0</ct141>
            <ct151>0</ct151>
            <ct152>0</ct152>
            <ct153>0</ct153>
            <ct154>0</ct154>
            <ct155>0</ct155>
            <ct156>9064217231</ct156>
            <ct157>0</ct157>
            <ct211>1089445454</ct211>
            <ct2111>1089445454</ct2111>
            <ct2112>0</ct2112>
            <ct2113>0</ct2113>
            <ct214>0</ct214>
            <ct2141>0</ct2141>
            <ct2142>0</ct2142>
            <ct2143>0</ct2143>
            <ct2147>0</ct2147>
            <ct217>0</ct217>
            <ct228>0</ct228>
            <ct2281>0</ct2281>
            <ct2288>0</ct2288>
            <ct229>0</ct229>
            <ct2291>0</ct2291>
            <ct2292>0</ct2292>
            <ct2293>0</ct2293>
            <ct2294>0</ct2294>
            <ct241>0</ct241>
            <ct2411>0</ct2411>
            <ct2412>0</ct2412>
            <ct2413>0</ct2413>
            <ct242>6666672</ct242>
            <ct331>0</ct331>
            <ct333>0</ct333>
            <ct3331>0</ct3331>
            <ct33311>0</ct33311>
            <ct33312>0</ct33312>
            <ct3332>0</ct3332>
            <ct3333>0</ct3333>
            <ct3334>0</ct3334>
            <ct3335>0</ct3335>
            <ct3336>0</ct3336>
            <ct3337>0</ct3337>
            <ct3338>0</ct3338>
            <ct33381>0</ct33381>
            <ct33382>0</ct33382>
            <ct3339>0</ct3339>
            <ct334>0</ct334>
            <ct335>0</ct335>
            <ct336>0</ct336>
            <ct3361>0</ct3361>
            <ct3368>0</ct3368>
            <ct338>0</ct338>
            <ct3381>0</ct3381>
            <ct3382>0</ct3382>
            <ct3383>0</ct3383>
            <ct3384>0</ct3384>
            <ct3385>0</ct3385>
            <ct3386>0</ct3386>
            <ct3387>0</ct3387>
            <ct3388>0</ct3388>
            <ct341>0</ct341>
            <ct3411>0</ct3411>
            <ct3412>0</ct3412>
            <ct352>0</ct352>
            <ct3521>0</ct3521>
            <ct3522>0</ct3522>
            <ct3524>0</ct3524>
            <ct353>0</ct353>
            <ct3531>0</ct3531>
            <ct3532>0</ct3532>
            <ct3533>0</ct3533>
            <ct3534>0</ct3534>
            <ct356>0</ct356>
            <ct3561>0</ct3561>
            <ct3562>0</ct3562>
            <ct411>0</ct411>
            <ct4111>0</ct4111>
            <ct4112>0</ct4112>
            <ct4118>0</ct4118>
            <ct413>0</ct413>
            <ct418>0</ct418>
            <ct419>0</ct419>
            <ct421>0</ct421>
            <ct4211>0</ct4211>
            <ct4212>0</ct4212>
            <ct511>0</ct511>
            <ct5111>0</ct5111>
            <ct5112>0</ct5112>
            <ct5113>0</ct5113>
            <ct5118>0</ct5118>
            <ct515>0</ct515>
            <ct611>0</ct611>
            <ct631>0</ct631>
            <ct632>0</ct632>
            <ct635>0</ct635>
            <ct642>0</ct642>
            <ct6421>0</ct6421>
            <ct6422>0</ct6422>
            <ct711>0</ct711>
            <ct811>0</ct811>
            <ct821>0</ct821>
            <ct911>0</ct911>
            <tongCong>11831920102</tongCong>
          </No>
          <Co>
            <ct111>0</ct111>
            <ct1111>0</ct1111>
            <ct1112>0</ct1112>
            <ct112>0</ct112>
            <ct1121>0</ct1121>
            <ct1122>0</ct1122>
            <ct121>0</ct121>
            <ct128>0</ct128>
            <ct1281>0</ct1281>
            <ct1288>0</ct1288>
            <ct131>0</ct131>
            <ct133>0</ct133>
            <ct1331>0</ct1331>
            <ct1332>0</ct1332>
            <ct136>0</ct136>
            <ct1361>0</ct1361>
            <ct1368>0</ct1368>
            <ct138>0</ct138>
            <ct1381>0</ct1381>
            <ct1386>0</ct1386>
            <ct1388>0</ct1388>
            <ct141>0</ct141>
            <ct151>0</ct151>
            <ct152>0</ct152>
            <ct153>0</ct153>
            <ct154>0</ct154>
            <ct155>0</ct155>
            <ct156>0</ct156>
            <ct157>0</ct157>
            <ct211>0</ct211>
            <ct2111>0</ct2111>
            <ct2112>0</ct2112>
            <ct2113>0</ct2113>
            <ct214>980236610</ct214>
            <ct2141>980236610</ct2141>
            <ct2142>0</ct2142>
            <ct2143>0</ct2143>
            <ct2147>0</ct2147>
            <ct217>0</ct217>
            <ct228>0</ct228>
            <ct2281>0</ct2281>
            <ct2288>0</ct2288>
            <ct229>0</ct229>
            <ct2291>0</ct2291>
            <ct2292>0</ct2292>
            <ct2293>0</ct2293>
            <ct2294>0</ct2294>
            <ct241>0</ct241>
            <ct2411>0</ct2411>
            <ct2412>0</ct2412>
            <ct2413>0</ct2413>
            <ct242>0</ct242>
            <ct331>314559628</ct331>
            <ct333>10793214</ct333>
            <ct3331>0</ct3331>
            <ct33311>0</ct33311>
            <ct33312>0</ct33312>
            <ct3332>0</ct3332>
            <ct3333>0</ct3333>
            <ct3334>10793214</ct3334>
            <ct3335>0</ct3335>
            <ct3336>0</ct3336>
            <ct3337>0</ct3337>
            <ct3338>0</ct3338>
            <ct33381>0</ct33381>
            <ct33382>0</ct33382>
            <ct3339>0</ct3339>
            <ct334>0</ct334>
            <ct335>0</ct335>
            <ct336>0</ct336>
            <ct3361>0</ct3361>
            <ct3368>0</ct3368>
            <ct338>0</ct338>
            <ct3381>0</ct3381>
            <ct3382>0</ct3382>
            <ct3383>0</ct3383>
            <ct3384>0</ct3384>
            <ct3385>0</ct3385>
            <ct3386>0</ct3386>
            <ct3387>0</ct3387>
            <ct3388>0</ct3388>
            <ct341>0</ct341>
            <ct3411>0</ct3411>
            <ct3412>0</ct3412>
            <ct352>0</ct352>
            <ct3521>0</ct3521>
            <ct3522>0</ct3522>
            <ct3524>0</ct3524>
            <ct353>0</ct353>
            <ct3531>0</ct3531>
            <ct3532>0</ct3532>
            <ct3533>0</ct3533>
            <ct3534>0</ct3534>
            <ct356>0</ct356>
            <ct3561>0</ct3561>
            <ct3562>0</ct3562>
            <ct411>10000000000</ct411>
            <ct4111>10000000000</ct4111>
            <ct4112>0</ct4112>
            <ct4118>0</ct4118>
            <ct413>0</ct413>
            <ct418>0</ct418>
            <ct419>0</ct419>
            <ct421>526330650</ct421>
            <ct4211>483157792</ct4211>
            <ct4212>43172858</ct4212>
            <ct511>0</ct511>
            <ct5111>0</ct5111>
            <ct5112>0</ct5112>
            <ct5113>0</ct5113>
            <ct5118>0</ct5118>
            <ct515>0</ct515>
            <ct611>0</ct611>
            <ct631>0</ct631>
            <ct632>0</ct632>
            <ct635>0</ct635>
            <ct642>0</ct642>
            <ct6421>0</ct6421>
            <ct6422>0</ct6422>
            <ct711>0</ct711>
            <ct811>0</ct811>
            <ct821>0</ct821>
            <ct911>0</ct911>
            <tongCong>11831920102</tongCong>
          </Co>
        </SoDuDauKy>
        <SoPhatSinhTrongKy>
          <No>
            <ct111>4107723520</ct111>
            <ct1111>4107723520</ct1111>
            <ct1112>0</ct1112>
            <ct112>3097253671</ct112>
            <ct1121>3097253671</ct1121>
            <ct1122>0</ct1122>
            <ct121>0</ct121>
            <ct128>0</ct128>
            <ct1281>0</ct1281>
            <ct1288>0</ct1288>
            <ct131>154709600</ct131>
            <ct133>257923776</ct133>
            <ct1331>257923776</ct1331>
            <ct1332>0</ct1332>
            <ct136>0</ct136>
            <ct1361>0</ct1361>
            <ct1368>0</ct1368>
            <ct138>0</ct138>
            <ct1381>0</ct1381>
            <ct1386>0</ct1386>
            <ct1388>0</ct1388>
            <ct141>0</ct141>
            <ct151>0</ct151>
            <ct152>0</ct152>
            <ct153>0</ct153>
            <ct154>0</ct154>
            <ct155>0</ct155>
            <ct156>2764958047</ct156>
            <ct157>0</ct157>
            <ct211>0</ct211>
            <ct2111>0</ct2111>
            <ct2112>0</ct2112>
            <ct2113>0</ct2113>
            <ct214>0</ct214>
            <ct2141>0</ct2141>
            <ct2142>0</ct2142>
            <ct2143>0</ct2143>
            <ct2147>0</ct2147>
            <ct217>0</ct217>
            <ct228>0</ct228>
            <ct2281>0</ct2281>
            <ct2288>0</ct2288>
            <ct229>0</ct229>
            <ct2291>0</ct2291>
            <ct2292>0</ct2292>
            <ct2293>0</ct2293>
            <ct2294>0</ct2294>
            <ct241>0</ct241>
            <ct2411>0</ct2411>
            <ct2412>0</ct2412>
            <ct2413>0</ct2413>
            <ct242>0</ct242>
            <ct331>3336647949</ct331>
            <ct333>413759002</ct333>
            <ct3331>381663002</ct3331>
            <ct33311>381663002</ct33311>
            <ct33312>0</ct33312>
            <ct3332>0</ct3332>
            <ct3333>0</ct3333>
            <ct3334>24696000</ct3334>
            <ct3335>5400000</ct3335>
            <ct3336>0</ct3336>
            <ct3337>0</ct3337>
            <ct3338>0</ct3338>
            <ct33381>0</ct33381>
            <ct33382>0</ct33382>
            <ct3339>2000000</ct3339>
            <ct334>422930000</ct334>
            <ct335>0</ct335>
            <ct336>0</ct336>
            <ct3361>0</ct3361>
            <ct3368>0</ct3368>
            <ct338>0</ct338>
            <ct3381>0</ct3381>
            <ct3382>0</ct3382>
            <ct3383>0</ct3383>
            <ct3384>0</ct3384>
            <ct3385>0</ct3385>
            <ct3386>0</ct3386>
            <ct3387>0</ct3387>
            <ct3388>0</ct3388>
            <ct341>0</ct341>
            <ct3411>0</ct3411>
            <ct3412>0</ct3412>
            <ct352>0</ct352>
            <ct3521>0</ct3521>
            <ct3522>0</ct3522>
            <ct3524>0</ct3524>
            <ct353>0</ct353>
            <ct3531>0</ct3531>
            <ct3532>0</ct3532>
            <ct3533>0</ct3533>
            <ct3534>0</ct3534>
            <ct356>0</ct356>
            <ct3561>0</ct3561>
            <ct3562>0</ct3562>
            <ct411>0</ct411>
            <ct4111>0</ct4111>
            <ct4112>0</ct4112>
            <ct4118>0</ct4118>
            <ct413>0</ct413>
            <ct418>0</ct418>
            <ct419>0</ct419>
            <ct421>43366491</ct421>
            <ct4211>0</ct4211>
            <ct4212>43366491</ct4212>
            <ct511>3800823318</ct511>
            <ct5111>3800823318</ct5111>
            <ct5112>0</ct5112>
            <ct5113>0</ct5113>
            <ct5118>0</ct5118>
            <ct515>40448614</ct515>
            <ct611>0</ct611>
            <ct631>0</ct631>
            <ct632>3485810206</ct632>
            <ct635>0</ct635>
            <ct642>537303654</ct642>
            <ct6421>233520000</ct6421>
            <ct6422>303783654</ct6422>
            <ct711>232219355</ct711>
            <ct811>40495575</ct811>
            <ct821>10075485</ct821>
            <ct911>4073684920</ct911>
            <tongCong>26820133183</tongCong>
          </No>
          <Co>
            <ct111>3406877298</ct111>
            <ct1111>3406877298</ct1111>
            <ct1112>0</ct1112>
            <ct112>3161158402</ct112>
            <ct1121>3161158402</ct1121>
            <ct1122>0</ct1122>
            <ct121>0</ct121>
            <ct128>0</ct128>
            <ct1281>0</ct1281>
            <ct1288>0</ct1288>
            <ct131>154709600</ct131>
            <ct133>383946002</ct133>
            <ct1331>383946002</ct1331>
            <ct1332>0</ct1332>
            <ct136>0</ct136>
            <ct1361>0</ct1361>
            <ct1368>0</ct1368>
            <ct138>0</ct138>
            <ct1381>0</ct1381>
            <ct1386>0</ct1386>
            <ct1388>0</ct1388>
            <ct141>0</ct141>
            <ct151>0</ct151>
            <ct152>0</ct152>
            <ct153>0</ct153>
            <ct154>0</ct154>
            <ct155>0</ct155>
            <ct156>3485810206</ct156>
            <ct157>0</ct157>
            <ct211>0</ct211>
            <ct2111>0</ct2111>
            <ct2112>0</ct2112>
            <ct2113>0</ct2113>
            <ct214>100808184</ct214>
            <ct2141>100808184</ct2141>
            <ct2142>0</ct2142>
            <ct2143>0</ct2143>
            <ct2147>0</ct2147>
            <ct217>0</ct217>
            <ct228>0</ct228>
            <ct2281>0</ct2281>
            <ct2288>0</ct2288>
            <ct229>0</ct229>
            <ct2291>0</ct2291>
            <ct2292>0</ct2292>
            <ct2293>0</ct2293>
            <ct2294>0</ct2294>
            <ct241>0</ct241>
            <ct2411>0</ct2411>
            <ct2412>0</ct2412>
            <ct2413>0</ct2413>
            <ct242>4999992</ct242>
            <ct331>3022025027</ct331>
            <ct333>412834487</ct333>
            <ct3331>381663002</ct3331>
            <ct33311>381663002</ct33311>
            <ct33312>0</ct33312>
            <ct3332>0</ct3332>
            <ct3333>0</ct3333>
            <ct3334>23771485</ct3334>
            <ct3335>5400000</ct3335>
            <ct3336>0</ct3336>
            <ct3337>0</ct3337>
            <ct3338>0</ct3338>
            <ct33381>0</ct33381>
            <ct33382>0</ct33382>
            <ct3339>2000000</ct3339>
            <ct334>422930000</ct334>
            <ct335>0</ct335>
            <ct336>0</ct336>
            <ct3361>0</ct3361>
            <ct3368>0</ct3368>
            <ct338>0</ct338>
            <ct3381>0</ct3381>
            <ct3382>0</ct3382>
            <ct3383>0</ct3383>
            <ct3384>0</ct3384>
            <ct3385>0</ct3385>
            <ct3386>0</ct3386>
            <ct3387>0</ct3387>
            <ct3388>0</ct3388>
            <ct341>0</ct341>
            <ct3411>0</ct3411>
            <ct3412>0</ct3412>
            <ct352>0</ct352>
            <ct3521>0</ct3521>
            <ct3522>0</ct3522>
            <ct3524>0</ct3524>
            <ct353>0</ct353>
            <ct3531>0</ct3531>
            <ct3532>0</ct3532>
            <ct3533>0</ct3533>
            <ct3534>0</ct3534>
            <ct356>0</ct356>
            <ct3561>0</ct3561>
            <ct3562>0</ct3562>
            <ct411>0</ct411>
            <ct4111>0</ct4111>
            <ct4112>0</ct4112>
            <ct4118>0</ct4118>
            <ct413>0</ct413>
            <ct418>0</ct418>
            <ct419>0</ct419>
            <ct421>43172858</ct421>
            <ct4211>43172858</ct4211>
            <ct4212>0</ct4212>
            <ct511>3800823318</ct511>
            <ct5111>3800823318</ct5111>
            <ct5112>0</ct5112>
            <ct5113>0</ct5113>
            <ct5118>0</ct5118>
            <ct515>40448614</ct515>
            <ct611>0</ct611>
            <ct631>0</ct631>
            <ct632>3485810206</ct632>
            <ct635>0</ct635>
            <ct642>537303654</ct642>
            <ct6421>233520000</ct6421>
            <ct6422>303783654</ct6422>
            <ct711>232219355</ct711>
            <ct811>40495575</ct811>
            <ct821>10075485</ct821>
            <ct911>4073684920</ct911>
            <tongCong>26820133183</tongCong>
          </Co>
        </SoPhatSinhTrongKy>
        <SoDuCuoiKy>
          <No>
            <ct111>1618365433</ct111>
            <ct1111>1618365433</ct1111>
            <ct1112>0</ct1112>
            <ct112>24986655</ct112>
            <ct1121>24986655</ct1121>
            <ct1122>0</ct1122>
            <ct121>0</ct121>
            <ct128>0</ct128>
            <ct1281>0</ct1281>
            <ct1288>0</ct1288>
            <ct131>300052500</ct131>
            <ct133>239105422</ct133>
            <ct1331>239105422</ct1331>
            <ct1332>0</ct1332>
            <ct136>0</ct136>
            <ct1361>0</ct1361>
            <ct1368>0</ct1368>
            <ct138>0</ct138>
            <ct1381>0</ct1381>
            <ct1386>0</ct1386>
            <ct1388>0</ct1388>
            <ct141>0</ct141>
            <ct151>0</ct151>
            <ct152>0</ct152>
            <ct153>0</ct153>
            <ct154>0</ct154>
            <ct155>0</ct155>
            <ct156>8343365072</ct156>
            <ct157>0</ct157>
            <ct211>1089445454</ct211>
            <ct2111>1089445454</ct2111>
            <ct2112>0</ct2112>
            <ct2113>0</ct2113>
            <ct214>0</ct214>
            <ct2141>0</ct2141>
            <ct2142>0</ct2142>
            <ct2143>0</ct2143>
            <ct2147>0</ct2147>
            <ct217>0</ct217>
            <ct228>0</ct228>
            <ct2281>0</ct2281>
            <ct2288>0</ct2288>
            <ct229>0</ct229>
            <ct2291>0</ct2291>
            <ct2292>0</ct2292>
            <ct2293>0</ct2293>
            <ct2294>0</ct2294>
            <ct241>0</ct241>
            <ct2411>0</ct2411>
            <ct2412>0</ct2412>
            <ct2413>0</ct2413>
            <ct242>1666680</ct242>
            <ct331>63294</ct331>
            <ct333>0</ct333>
            <ct3331>0</ct3331>
            <ct33311>0</ct33311>
            <ct33312>0</ct33312>
            <ct3332>0</ct3332>
            <ct3333>0</ct3333>
            <ct3334>0</ct3334>
            <ct3335>0</ct3335>
            <ct3336>0</ct3336>
            <ct3337>0</ct3337>
            <ct3338>0</ct3338>
            <ct33381>0</ct33381>
            <ct33382>0</ct33382>
            <ct3339>0</ct3339>
            <ct334>0</ct334>
            <ct335>0</ct335>
            <ct336>0</ct336>
            <ct3361>0</ct3361>
            <ct3368>0</ct3368>
            <ct338>0</ct338>
            <ct3381>0</ct3381>
            <ct3382>0</ct3382>
            <ct3383>0</ct3383>
            <ct3384>0</ct3384>
            <ct3385>0</ct3385>
            <ct3386>0</ct3386>
            <ct3387>0</ct3387>
            <ct3388>0</ct3388>
            <ct341>0</ct341>
            <ct3411>0</ct3411>
            <ct3412>0</ct3412>
            <ct352>0</ct352>
            <ct3521>0</ct3521>
            <ct3522>0</ct3522>
            <ct3524>0</ct3524>
            <ct353>0</ct353>
            <ct3531>0</ct3531>
            <ct3532>0</ct3532>
            <ct3533>0</ct3533>
            <ct3534>0</ct3534>
            <ct356>0</ct356>
            <ct3561>0</ct3561>
            <ct3562>0</ct3562>
            <ct411>0</ct411>
            <ct4111>0</ct4111>
            <ct4112>0</ct4112>
            <ct4118>0</ct4118>
            <ct413>0</ct413>
            <ct418>0</ct418>
            <ct419>0</ct419>
            <ct421>193633</ct421>
            <ct4211>0</ct4211>
            <ct4212>193633</ct4212>
            <ct511>0</ct511>
            <ct5111>0</ct5111>
            <ct5112>0</ct5112>
            <ct5113>0</ct5113>
            <ct5118>0</ct5118>
            <ct515>0</ct515>
            <ct611>0</ct611>
            <ct631>0</ct631>
            <ct632>0</ct632>
            <ct635>0</ct635>
            <ct642>0</ct642>
            <ct6421>0</ct6421>
            <ct6422>0</ct6422>
            <ct711>0</ct711>
            <ct811>0</ct811>
            <ct821>0</ct821>
            <ct911>0</ct911>
            <tongCong>11617244143</tongCong>
          </No>
          <Co>
            <ct111>0</ct111>
            <ct1111>0</ct1111>
            <ct1112>0</ct1112>
            <ct112>0</ct112>
            <ct1121>0</ct1121>
            <ct1122>0</ct1122>
            <ct121>0</ct121>
            <ct128>0</ct128>
            <ct1281>0</ct1281>
            <ct1288>0</ct1288>
            <ct131>0</ct131>
            <ct133>0</ct133>
            <ct1331>0</ct1331>
            <ct1332>0</ct1332>
            <ct136>0</ct136>
            <ct1361>0</ct1361>
            <ct1368>0</ct1368>
            <ct138>0</ct138>
            <ct1381>0</ct1381>
            <ct1386>0</ct1386>
            <ct1388>0</ct1388>
            <ct141>0</ct141>
            <ct151>0</ct151>
            <ct152>0</ct152>
            <ct153>0</ct153>
            <ct154>0</ct154>
            <ct155>0</ct155>
            <ct156>0</ct156>
            <ct157>0</ct157>
            <ct211>0</ct211>
            <ct2111>0</ct2111>
            <ct2112>0</ct2112>
            <ct2113>0</ct2113>
            <ct214>1081044794</ct214>
            <ct2141>1081044794</ct2141>
            <ct2142>0</ct2142>
            <ct2143>0</ct2143>
            <ct2147>0</ct2147>
            <ct217>0</ct217>
            <ct228>0</ct228>
            <ct2281>0</ct2281>
            <ct2288>0</ct2288>
            <ct229>0</ct229>
            <ct2291>0</ct2291>
            <ct2292>0</ct2292>
            <ct2293>0</ct2293>
            <ct2294>0</ct2294>
            <ct241>0</ct241>
            <ct2411>0</ct2411>
            <ct2412>0</ct2412>
            <ct2413>0</ct2413>
            <ct242>0</ct242>
            <ct331>0</ct331>
            <ct333>9868699</ct333>
            <ct3331>0</ct3331>
            <ct33311>0</ct33311>
            <ct33312>0</ct33312>
            <ct3332>0</ct3332>
            <ct3333>0</ct3333>
            <ct3334>9868699</ct3334>
            <ct3335>0</ct3335>
            <ct3336>0</ct3336>
            <ct3337>0</ct3337>
            <ct3338>0</ct3338>
            <ct33381>0</ct33381>
            <ct33382>0</ct33382>
            <ct3339>0</ct3339>
            <ct334>0</ct334>
            <ct335>0</ct335>
            <ct336>0</ct336>
            <ct3361>0</ct3361>
            <ct3368>0</ct3368>
            <ct338>0</ct338>
            <ct3381>0</ct3381>
            <ct3382>0</ct3382>
            <ct3383>0</ct3383>
            <ct3384>0</ct3384>
            <ct3385>0</ct3385>
            <ct3386>0</ct3386>
            <ct3387>0</ct3387>
            <ct3388>0</ct3388>
            <ct341>0</ct341>
            <ct3411>0</ct3411>
            <ct3412>0</ct3412>
            <ct352>0</ct352>
            <ct3521>0</ct3521>
            <ct3522>0</ct3522>
            <ct3524>0</ct3524>
            <ct353>0</ct353>
            <ct3531>0</ct3531>
            <ct3532>0</ct3532>
            <ct3533>0</ct3533>
            <ct3534>0</ct3534>
            <ct356>0</ct356>
            <ct3561>0</ct3561>
            <ct3562>0</ct3562>
            <ct411>10000000000</ct411>
            <ct4111>10000000000</ct4111>
            <ct4112>0</ct4112>
            <ct4118>0</ct4118>
            <ct413>0</ct413>
            <ct418>0</ct418>
            <ct419>0</ct419>
            <ct421>526330650</ct421>
            <ct4211>526330650</ct4211>
            <ct4212>0</ct4212>
            <ct511>0</ct511>
            <ct5111>0</ct5111>
            <ct5112>0</ct5112>
            <ct5113>0</ct5113>
            <ct5118>0</ct5118>
            <ct515>0</ct515>
            <ct611>0</ct611>
            <ct631>0</ct631>
            <ct632>0</ct632>
            <ct635>0</ct635>
            <ct642>0</ct642>
            <ct6421>0</ct6421>
            <ct6422>0</ct6422>
            <ct711>0</ct711>
            <ct811>0</ct811>
            <ct821>0</ct821>
            <ct911>0</ct911>
            <tongCong>11617244143</tongCong>
          </Co>
        </SoDuCuoiKy>
      </PL_CDTK>
    </PLuc>
  </HSoKhaiThue>
<CKyDTu><ds:Signature xmlns:ds=""http://www.w3.org/2000/09/xmldsig#""><ds:SignedInfo><ds:CanonicalizationMethod Algorithm=""http://www.w3.org/TR/2001/REC-xml-c14n-20010315#WithComments""/><ds:SignatureMethod Algorithm=""http://www.w3.org/2001/04/xmldsig-more#rsa-sha256""/><ds:Reference URI=""#_NODE_TO_SIGN""><ds:Transforms><ds:Transform Algorithm=""http://www.w3.org/2000/09/xmldsig#enveloped-signature""/></ds:Transforms><ds:DigestMethod Algorithm=""http://www.w3.org/2001/04/xmlenc#sha256""/><ds:DigestValue>n09vu16A/eeA7CnRQECUvXXjcYmdRNJH9wAhUfqk988=</ds:DigestValue></ds:Reference></ds:SignedInfo><ds:SignatureValue>dPcSGAE/KUTg+srnBCrcgyICM1h/AekglfsDGPTYHmaudcZ8od0uaoO8yPNRR2mQ+o2AxlpJhitF
aivPPBD5zVN1WDrixNi+mtXepR9Tthc2GX9Nun5Fs7/6GfdX+i+uJ8FPyb5lFtN8WyQro19WWp4V
Ir2uSQ1ZiszeD+tL8IU=</ds:SignatureValue><ds:KeyInfo><ds:KeyValue><ds:RSAKeyValue><ds:Modulus>3V9xpMb9yfC29nG1VtHDBDUKP//IvPM49+m61uBF8urGiacWs1NgH8anXUG6xybV28wTwmgR/W4w
IpZXUL8UfFB/wav9w02gjmglWIbOjQFZUoeJNg6/Vv4H59j67ItXV2zX+gsHxEPJuHnD56OjAga7
BKF6oKpLacQyJWPvHrU=</ds:Modulus><ds:Exponent>AQAB</ds:Exponent></ds:RSAKeyValue></ds:KeyValue><ds:X509Data><ds:X509SubjectName>UID=MST:3500807781,CN=CÔNG TY TRÁCH NHIỆM HỮU HẠN THƯƠNG MẠI - XÂY DỰNG ĐỨC THỊNH,L=BÀ RỊA VŨNG TÀU,C=VN</ds:X509SubjectName><ds:X509Certificate>MIIEEjCCAvqgAwIBAgIQVAT//rcDP7MW1nIgG9QA7TANBgkqhkiG9w0BAQsFADBCMQswCQYDVQQG
EwJWTjEWMBQGA1UECgwNVmlldHRlbCBHcm91cDEbMBkGA1UEAwwSVmlldHRlbC1DQSBTSEEtMjU2
MB4XDTI0MDUyMjA0MjAzM1oXDTI1MDkyNDEwNDExMVowgaYxCzAJBgNVBAYTAlZOMR0wGwYDVQQH
DBRCw4AgUuG7ikEgVsWoTkcgVMOAVTFYMFYGA1UEAwxPQ8OUTkcgVFkgVFLDgUNIIE5ISeG7hk0g
SOG7rlUgSOG6oE4gVEjGr8agTkcgTeG6oEkgLSBYw4JZIEThu7BORyDEkOG7qEMgVEjhu4pOSDEe
MBwGCgmSJomT8ixkAQEMDk1TVDozNTAwODA3NzgxMIGfMA0GCSqGSIb3DQEBAQUAA4GNADCBiQKB
gQDdX3Gkxv3J8Lb2cbVW0cMENQo//8i88zj36brW4EXy6saJpxazU2AfxqddQbrHJtXbzBPCaBH9
bjAilldQvxR8UH/Bq/3DTaCOaCVYhs6NAVlSh4k2Dr9W/gfn2Prsi1dXbNf6CwfEQ8m4ecPno6MC
BrsEoXqgqktpxDIlY+8etQIDAQABo4IBITCCAR0wNQYIKwYBBQUHAQEEKTAnMCUGCCsGAQUFBzAB
hhlodHRwOi8vb2NzcC52aWV0dGVsLWNhLnZuMB0GA1UdDgQWBBTZk7GYQNCfMlmNqkqBB6rBxdE7
gDAMBgNVHRMBAf8EAjAAMB8GA1UdIwQYMBaAFLpfG+l5A3440l7+9Js/agjkLnvhMIGFBgNVHR8E
fjB8MHqgMKAuhixodHRwOi8vY3JsLnZpZXR0ZWwtY2Eudm4vVmlldHRlbC1DQS1TSEEyLmNybKJG
pEQwQjEbMBkGA1UEAwwSVmlldHRlbC1DQSBTSEEtMjU2MRYwFAYDVQQKDA1WaWV0dGVsIEdyb3Vw
MQswCQYDVQQGEwJWTjAOBgNVHQ8BAf8EBAMCBeAwDQYJKoZIhvcNAQELBQADggEBAD70oMVAo5D3
203DXrmzlGELujX6N0Tau7GxLJHMRKMaLV1q8aipWchDOorqVEgWJAdYa/fh7JSm8HOt6yFrRh9S
p/+Ch5XT911DLXrcLybFTVMNtepnveDvWIfNe986//SqXOkaWR7occhTiQqnE7Y23EHsLPYG4yVn
vf2JqtrQkMQ/GYFQOJSA41nXSgP5De4KLVx5FRfa2FCJ4tbmBO1H3YqYmYX1uwM8XMXMtRJw4NUU
Fr9Mu0jm/E+UAq6fNTCXoo7TcM43QitMlsVHWEj2cgbhkaO4OB0n7yuQddn00ONV82sSrgTbN//l
K2BD4WNep8Mug+G9ruJB/VoRzyo=</ds:X509Certificate></ds:X509Data></ds:KeyInfo></ds:Signature></CKyDTu></HSoThueDTu>";
            byte[] fileBytes = Encoding.UTF8.GetBytes(xmlContent);
            return File(fileBytes, "application /xml", "HSoThueDTu.xml");
        }

       [HttpPost]
        public int CreateDataChild(List<YourModel> items, string path, int nam, int idparent)
        {
            try
            {
                var qrdl = "delete from tbCDTSChild  WHERE Idparent=?";
                var parameters = new OleDbParameter[]
                {
                                                new OleDbParameter("?", idparent),
                };
                ExecuteQueryResult(qrdl, parameters);


                // 2. Lọc items hợp lệ
                var validItems = items?
                    .Where(x => x.maso != "0" && !string.IsNullOrWhiteSpace(x.maso))
                    .ToList();

                if (validItems == null || validItems.Count == 0)
                    return 0;

                // 3. Dùng TRANSACTION để tối ưu
                int insertedCount = ExecuteWithTransaction(validItems, idparent);

                return insertedCount;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                return -1;
            }
        }

        private bool CheckIfExists(int idparent)
        {
            string query = "SELECT COUNT(*) FROM tbCDTSChild WHERE Idparent = ?";
            var count = ExecuteScalar(query, new OleDbParameter("?", idparent));
            return Convert.ToInt32(count) > 0;
        }

        private int ExecuteWithTransaction(List<YourModel> items, int idparent)
        {
            string password = "1@35^7*9)1";
            string connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};Jet OLEDB:Database Password={password};";

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();

                using (OleDbTransaction transaction = connection.BeginTransaction())
                {
                    try
                    {
                        int insertedCount = 0;

                        // Dùng Prepared Statement (chỉ chuẩn bị 1 lần)
                        string insertQuery = "INSERT INTO tbCDTSChild (MaSo, CuoiKy, DauNam, Idparent) VALUES (?, ?, ?, ?)";

                        using (OleDbCommand command = new OleDbCommand(insertQuery, connection, transaction))
                        {
                            // Chuẩn bị parameters trước
                            command.Parameters.Add("?", OleDbType.VarChar, 50);
                            command.Parameters.Add("?", OleDbType.VarChar, 50);
                            command.Parameters.Add("?", OleDbType.VarChar, 50);
                            command.Parameters.Add("?", OleDbType.Integer);

                            // Thực hiện từng INSERT
                            foreach (var item in items)
                            {
                                // Gán giá trị cho parameters
                                command.Parameters[0].Value = item.maso;
                                command.Parameters[1].Value = item.CuoiKy?.Replace(".", "") ?? string.Empty;
                                command.Parameters[2].Value = item.DauNam?.Replace(".", "") ?? string.Empty;
                                command.Parameters[3].Value = idparent;

                                // Thực thi
                                command.ExecuteNonQuery();
                                insertedCount++;
                            }
                        }

                        transaction.Commit();
                        return insertedCount;
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        throw new Exception($"Transaction failed: {ex.Message}", ex);
                    }
                }
            }
        }

        // Hàm ExecuteScalar helper
        public object ExecuteScalar(string query, params OleDbParameter[] parameters)
        {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                using (OleDbCommand command = new OleDbCommand(query, connection))
                {
                    if (parameters != null)
                        command.Parameters.AddRange(parameters);

                    return command.ExecuteScalar();
                }
            }
        }
        //[HttpPost]
        //public int CreateDataChild(string path,string maso,string DauNam, string CuoiKy,int idparent)
        //{
        //    dbPath = path;
        //    if (maso == "0")
        //        return 0;
        //    string qrchk = @"SELECT * FROM tbCDTSChild where Idparent = ? and MaSo=?";
        //    var parameterss = new OleDbParameter[]
        //    {
        //  new OleDbParameter("?",idparent),
        //  new OleDbParameter("?",maso)
        //    };
        //    var chkCDtS = ExecuteQuery(qrchk, parameterss);
        //    //Thêm mới
        //    if(chkCDtS.Rows.Count == 0)
        //    {
        //        string query = "INSERT INTO tbCDTSChild (MaSo,CuoiKy,DauNam,Idparent) VALUES (?,?,?,?)";
        //        OleDbParameter[] parameters = new OleDbParameter[]
        //        {
        //         new OleDbParameter("?", maso),
        //         new OleDbParameter("?", DauNam.Replace(".","")) ,
        //         new OleDbParameter("?", CuoiKy.Replace(".","")) ,
        //         new OleDbParameter("?", idparent),
        //        };
        //        var getID = ExecuteQueryResult(query, parameters);
        //        return getID;
        //    }
        //    //Update
        //    else
        //    {
        //        string query = "UPDATE tbCDTSChild SET DauNam=?,CuoiKy=? WHERE Idparent=? and MaSo=?";
        //        OleDbParameter[] parameters = new OleDbParameter[]
        //        {
        //         new OleDbParameter("?", DauNam.Replace(".","")) ,
        //         new OleDbParameter("?", CuoiKy.Replace(".","")) ,
        //         new OleDbParameter("?", idparent),
        //         new OleDbParameter("?", maso),
        //        };
        //        var getID=ExecuteQueryResult(query, parameters);
        //    }

        //        return 1;
        //}
        public int ExecuteQueryResult(string query, params OleDbParameter[] parameters)
        {
            string password = "1@35^7*9)1";
            connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};Jet OLEDB:Database Password={password};";
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                Console.WriteLine("Kết nối đến cơ sở dữ liệu thành công!");

                using (OleDbCommand command = new OleDbCommand(query, connection))
                {
                    // Thêm tham số
                    if (parameters != null)
                        command.Parameters.AddRange(parameters);

                    // Thực thi INSERT, UPDATE, DELETE
                    command.ExecuteNonQuery();
                }

                // Lấy ID vừa thêm bằng @@IDENTITY
                using (OleDbCommand idCommand = new OleDbCommand("SELECT @@IDENTITY", connection))
                {
                    object result = idCommand.ExecuteScalar();
                    return Convert.ToInt32(result);
                }
            }
        }
        public System.Data.DataTable ExecuteQuery(string query, params OleDbParameter[] parameters)
        {

            string password = "1@35^7*9)1";
            connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};Jet OLEDB:Database Password={password};";

            System.Data.DataTable dataTable = new System.Data.DataTable();

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                try
                {
                    connection.Open();

                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        // Thêm các tham số vào command
                        if (parameters != null)
                        {
                            command.Parameters.AddRange(parameters);
                        }

                        using (OleDbDataAdapter dataAdapter = new OleDbDataAdapter(command))
                        {
                            dataAdapter.Fill(dataTable);
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }

            }

            return dataTable; // Trả về DataTable chứa dữ liệu
        }
    }
}