using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace _058_현대
{
    public class CONVERT
    {
        //기본 인코딩 설정
        private static string strEncoding = "ks_c_5601-1987";
        private static char chCSV = ';';
        private static string strCardTypeID = "058";
        private static string strCardTypeName = "현대";

        //현 DLL의 카드 타입 코드 반환
        public static string GetCardTypeID()
        {
            return strCardTypeID;
        }

        //현 DLL의 카드 타입명 반환
        public static string GetCardTypeName()
        {
            return strCardTypeName;
        }

        //제휴사코드 반환
        public static string GetCardType(string path)
        {
            System.Text.Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            StreamReader _sr = null;
            byte[] _byteAry = null;
            string _strReturn = "";
            string _strLine = "";

            //파일 읽기 Stream과 오류시 저장할 쓰기 Stream 생성
            _sr = new StreamReader(path, _encoding);
            _strLine = _sr.ReadLine();
            try
            {
                if (_strLine.Trim() != "")
                {
                    _strReturn = _strLine.Substring(_strLine.Length - 7, 7);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }

            return _strReturn;
        }

        //등록 자료 생성
        //public static string ConvertRegister(string path, string xmlZipcodePath, string xmlZipcodeAreaPath, string xmlPath)
        public static string ConvertRegister(string path, string xmlZipcodePath, string xmlZipcodeAreaPath, string xmlZipcodePath_new, string xmlZipcodeAreaPath_new, string xmlPath)
        {
            System.Text.Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            //FileInfo _fi = null;
            StreamReader _sr = null;													//파일 읽기 스트림
            StreamWriter _swError = null;   											//파일 쓰기 스트림
            DataSet _dsetZipcode = null, _dsetZipcdeArea = null;						//우편번호 관련 DataSet
            DataSet _dsetZipcode_new = null, _dsetZipcdeArea_new = null;				//우편번호 관련 DataSet
            DataTable _dtable = null;																				//마스터 저장 테이블
            DataRow _dr = null;
            DataRow[] _drs = null;

            //정렬변경
            DataTable _dtable_HD = null;
            DataRow _dr_HD = null;
            DataRow[] _drs_HD = null;

            string[] _strAry = null;
            string _strReturn = "";
            string _strLine = "";
            string _strZipcode = "", _strAreaType = "", _strAreaGroup = "", _strBranch = "", strCard_type_detail = "";
            string _strDeliveryPlaceType = "", _strOwner_only = "";
            int _iSeq = 1, _iErrorCount = 0;
            Branches _branches = new Branches();
            try
            {
                _dtable = new DataTable("Convert");
                //기본 컬럼
                _dtable.Columns.Add("degree_arrange_number");
                _dtable.Columns.Add("card_area_type");
                _dtable.Columns.Add("card_branch");
                _dtable.Columns.Add("card_area_group");
                _dtable.Columns.Add("area_arrange_number");
                //세부 컬럼
                _dtable = new DataTable("Convert");
                //기본 컬럼
                _dtable.Columns.Add("degree_arrange_number");
                _dtable.Columns.Add("card_area_type");
                _dtable.Columns.Add("card_branch");
                _dtable.Columns.Add("card_area_group");
                _dtable.Columns.Add("area_arrange_number");
                //세부 컬럼
                _dtable.Columns.Add("client_send_number");             // dr[5] 발송번호
                _dtable.Columns.Add("card_number");                    // 앞1자리 뺀 발송번호 16자리 
                _dtable.Columns.Add("card_issue_type_code");           // 카드신청구분코드 : 1 = 신규, 2 = 재발급, 3 = 갱신
                _dtable.Columns.Add("client_express_code");            // 업체코드 B002
                _dtable.Columns.Add("client_enterprise_code");         // 제작지점코드
                _dtable.Columns.Add("card_count");                     // dr[10] 카드매수
                _dtable.Columns.Add("client_register_type");           // 발송구분코드 : 10=당사교부, 11=특별관리, 20=일반등기 등등
                _dtable.Columns.Add("card_product_name");              // 카드상품코드명
                _dtable.Columns.Add("card_brand_code");                // 카드브랜드코드명
                _dtable.Columns.Add("card_level_code");                // 카드등급코드명
                _dtable.Columns.Add("customer_name");                  // dr[15] 고객명
                _dtable.Columns.Add("customer_SSN");                   // 고객주민번호
                _dtable.Columns.Add("family_name");                    // 가족명
                _dtable.Columns.Add("family_SSN");                     // 가족주민번호
                _dtable.Columns.Add("family_name2");                   // 가족명2
                _dtable.Columns.Add("family_SSN2");                    // dr[20] 가족주민번호2
                _dtable.Columns.Add("customer_type_code");             // 동봉카드 내역 : 00=본인, 01=본인+가족1, 02=가족1, 03=가족1+가족2
                _dtable.Columns.Add("card_family_code");               // 분리(가족)배송여부 : Y/N
                _dtable.Columns.Add("family_relation");                // 신청인관계코드명
                _dtable.Columns.Add("card_zipcode");                   // 우편번호
                _dtable.Columns.Add("card_address_local");             // dr[25] 자택주소 동이상
                _dtable.Columns.Add("card_address_detail");            // 자택주소 동이하
                _dtable.Columns.Add("card_tel1");                      // 자택전화번호
                _dtable.Columns.Add("card_mobile_tel");                // 핸드폰번호
                _dtable.Columns.Add("card_zipcode2");                  // 직장우편번호
                _dtable.Columns.Add("card_address2_local");            // dr[30] 직장주소 동이상
                _dtable.Columns.Add("card_address2_detail");           // 직장주소 동이하
                _dtable.Columns.Add("card_tel2");                      // 직장전화번호
                _dtable.Columns.Add("customer_office");                // 직장명
                _dtable.Columns.Add("card_zipcode3");                  // 기타우편번호
                _dtable.Columns.Add("card_address3");                  // dr[35] 기타주소
                _dtable.Columns.Add("card_tel3");                      // 기타전화번호
                _dtable.Columns.Add("customer_email_ID");              // EMAIL ID
                _dtable.Columns.Add("customer_email_domain");          // EMAIL 도매인
                _dtable.Columns.Add("family_tel1");                    // 가족자택전화번호
                _dtable.Columns.Add("family_tel2");                    // dr[40] 가족직장번호
                _dtable.Columns.Add("family_mobile_tel");              // 가족핸드폰번호
                _dtable.Columns.Add("card_bill_place_type");           // 청구지구분코드
                _dtable.Columns.Add("card_delivery_place_code");       // 수령지구분코드
                _dtable.Columns.Add("card_bank_account_name");         // 자동이체은행코드
                _dtable.Columns.Add("card_bank_account_no");           //  dr[45] 계좌번호
                _dtable.Columns.Add("card_payment_day");               // 결제일
                _dtable.Columns.Add("card_cooperation1");              // 제휴사명
                _dtable.Columns.Add("card_limit");                     // 희망한도
                _dtable.Columns.Add("card_agree_code");                // 동의서식별코드 : A=일반, B=MSC, C=웅진코웨이, D=동양증권, W=웅진코웨이(세이브)
                _dtable.Columns.Add("card_terminal_issue");            // dr[50] 단말기여부 Y/N
                _dtable.Columns.Add("card_bill_way");                  // 청구서발송방법코드
                _dtable.Columns.Add("card_vip_code");                  // VIP
                _dtable.Columns.Add("card_pt_code");                   // 백화점회원등급코드
                _dtable.Columns.Add("card_register_type");             // 동의서 출력구분
                //2013.01.28 태희철 추가
                _dtable.Columns.Add("choice_agree3");                  // dr[55] 이용권유동의
                _dtable.Columns.Add("save_agreement");                 // dr[56] 세이브다이렉트 여부 : "00" 일반, "01" 다이렉트
                _dtable.Columns.Add("card_urgency_code");              // dr[57] 무료보험
                _dtable.Columns.Add("card_cooperation_code");          // dr[58] web신청구분
                _dtable.Columns.Add("card_cooperation2");              // dr[59] 선택체크 구분

                _dtable.Columns.Add("card_design_code");               // dr[60] 표준동의코드
                _dtable.Columns.Add("card_product_code");              // 상품동의코드
                _dtable.Columns.Add("text1");                          // 제휴동의코드
                _dtable.Columns.Add("card_cost_code");                 // 신청경로
                _dtable.Columns.Add("text2");                          // 동의서코드
                _dtable.Columns.Add("card_is_for_owner_only");         // dr[65] 본인배송코드
                _dtable.Columns.Add("BC_Part");                        // dr[66] 수도권 : 01, 지방 : 02

                _dtable.Columns.Add("card_barcode_new");               // dr[67] 카드사바코드 : 0 + 발송번호
                _dtable.Columns.Add("card_issue_type_new");            // dr[68] 발급구분코드_new
                _dtable.Columns.Add("card_delivery_place_type");       // dr[69] 내부수령지구분코드

                _dtable.Columns.Add("card_zipcode_new");               // dr[70]새우편번호
                _dtable.Columns.Add("card_zipcode_kind");              // dr[71]우편번호구분
                _dtable.Columns.Add("card_zipcode2_new");              // 새우편번호2
                _dtable.Columns.Add("card_zipcode2_kind");             // 우편번호
                _dtable.Columns.Add("card_zipcode3_new");              // 새우편번호3
                _dtable.Columns.Add("card_zipcode3_kind");             // dr[75]우편번호

                _dtable.Columns.Add("customer_memo");                  // dr[76]안내문구
                _dtable.Columns.Add("card_consented");                 // dr[77]프리미엄고객구분
                _dtable.Columns.Add("client_insert_type");             // dr[78]제휴선택코드

                _dtable.Columns.Add("card_client_code_1");             // dr[79]리볼빙동의
                _dtable.Columns.Add("card_traffic_code");              // dr[80]갱신



                //우편번호 관련 정보 DataSet에 담기
                _dsetZipcode = new DataSet();
                _dsetZipcdeArea = new DataSet();
                _dsetZipcode.ReadXml(xmlZipcodePath);
                _dsetZipcode.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcode.Tables[0].Columns["zipcode"] };
                _dsetZipcdeArea.ReadXml(xmlZipcodeAreaPath);
                _dsetZipcdeArea.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcdeArea.Tables[0].Columns["zipcode"] };

                //우편번호 관련 정보 DataSet에 담기
                _dsetZipcode_new = new DataSet();
                _dsetZipcdeArea_new = new DataSet();
                _dsetZipcode_new.ReadXml(xmlZipcodePath_new);
                _dsetZipcode_new.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcode_new.Tables[0].Columns["zipcode_new"] };
                _dsetZipcdeArea_new.ReadXml(xmlZipcodeAreaPath_new);
                _dsetZipcdeArea_new.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcdeArea_new.Tables[0].Columns["zipcode_new"] };

                //파일 읽기 Stream과 오류시 저장할 쓰기 Stream 생성
                _sr = new StreamReader(path, _encoding);
                _swError = new StreamWriter(path + ".Error", false, _encoding);

                while ((_strLine = _sr.ReadLine()) != null)
                {
                    if (_iSeq == 1)
                    {
                        strCard_type_detail = _strLine.Substring(_strLine.Length - 7, 7);
                    }

                    _dr = _dtable.NewRow();
                    _dr[0] = _iSeq;

                    //CSV  분리
                    _strAry = _strLine.Split(chCSV);
                    _strDeliveryPlaceType = _strAry[38];

                    _dr[5] = _strAry[0].Replace("\"", "");
                    _dr[6] = _strAry[0].Substring(1, 16);
                    _dr[7] = _strAry[1];
                    //_dr[59] = _strAry[1];
                    _dr[8] = _strAry[2];
                    _dr[9] = _strAry[3];
                    _dr[10] = _strAry[4];
                    _dr[11] = _strAry[5];
                    _dr[12] = _strAry[6];
                    _dr[13] = _strAry[7];
                    _dr[14] = _strAry[8];
                    _dr[15] = _strAry[9];

                    if (_strAry[10].Trim() == "")
                    {
                        _dr[16] = "xxxxxxxxxxxxx";
                    }
                    else
                    {
                        _dr[16] = _strAry[10].Replace('X', 'x');
                    }

                    _dr[17] = _strAry[11];
                    _dr[18] = _strAry[12];
                    _dr[19] = _strAry[13];
                    _dr[20] = _strAry[14];
                    _dr[21] = _strAry[15];
                    _dr[22] = _strAry[16];

                    _dr[23] = _strAry[17];
                    if (_strDeliveryPlaceType == "01")
                    {
                        _strZipcode = _strAry[18].Trim();

                        if (_strZipcode.Length == 5)
                        {
                            _dr[70] = _strZipcode;
                            _dr[71] = "1";
                        }
                        _dr[24] = _strZipcode;
                        _dr[25] = _strAry[19];
                        _dr[26] = _strAry[20];
                        _dr[27] = _strAry[21];
                    }
                    //else if (_strDeliveryPlaceType == "02" || _strDeliveryPlaceType == "17")
                    else if (_strDeliveryPlaceType == "02")
                    {
                        _strZipcode = _strAry[23].Trim();

                        if (_strZipcode.Length == 5)
                        {
                            _dr[70] = _strZipcode;
                            _dr[71] = "1";
                        }

                        _dr[24] = _strZipcode;
                        _dr[25] = _strAry[24];
                        _dr[26] = _strAry[25];
                        _dr[27] = _strAry[26];
                    }
                    else
                    {
                        _strZipcode = _strAry[28].Trim();

                        if (_strZipcode.Length == 5)
                        {
                            _dr[70] = _strZipcode;
                            _dr[71] = "1";
                        }

                        _dr[24] = _strZipcode;
                        _dr[25] = _strAry[29];
                        _dr[26] = _strAry[30];
                        _dr[27] = _strAry[31];
                    }

                    _dr[28] = _strAry[22];

                    //if (_strDeliveryPlaceType == "02" || _strDeliveryPlaceType == "17")
                    if (_strDeliveryPlaceType == "02")
                    {
                        _dr[29] = _strAry[18].Trim();

                        if (_dr[29].ToString().Length == 5)
                        {
                            _dr[72] = _dr[29].ToString();
                            _dr[73] = "1";
                        }

                        _dr[30] = _strAry[19];
                        _dr[31] = _strAry[20];
                        _dr[32] = _strAry[21];
                    }
                    else
                    {
                        _dr[29] = _strAry[23].Trim();

                        if (_dr[29].ToString().Length == 5)
                        {
                            _dr[72] = _dr[29].ToString();
                            _dr[73] = "1";
                        }

                        _dr[30] = _strAry[24];
                        _dr[31] = _strAry[25];
                        _dr[32] = _strAry[26];
                    }

                    _dr[33] = _strAry[27];

                    if (_strDeliveryPlaceType == "99")
                    {
                        _dr[34] = _strAry[18].Trim();

                        if (_dr[34].ToString().Length == 5)
                        {
                            _dr[74] = _dr[34].ToString();
                            _dr[75] = "1";
                        }

                        _dr[35] = _strAry[19].Trim() + "^" + _strAry[20].Trim();
                        _dr[36] = _strAry[21];
                    }
                    else
                    {
                        _dr[34] = _strAry[28].Trim();

                        if (_dr[34].ToString().Length == 5)
                        {
                            _dr[74] = _dr[34].ToString();
                            _dr[75] = "1";
                        }

                        _dr[35] = _strAry[29].Trim() + "^" + _strAry[30].Trim();
                        _dr[36] = _strAry[31];
                    }

                    _dr[37] = _strAry[32];
                    _dr[38] = _strAry[33];
                    _dr[39] = _strAry[34];
                    _dr[40] = _strAry[35];
                    _dr[41] = _strAry[36];
                    _dr[42] = _strAry[37];
                    _dr[43] = _strDeliveryPlaceType;
                    _dr[44] = _strAry[39];
                    _dr[45] = _strAry[40];
                    _dr[46] = _strAry[41];
                    _dr[47] = _strAry[42];
                    _dr[48] = _strAry[43];
                    _dr[49] = _strAry[44];
                    _dr[50] = _strAry[45];
                    _dr[51] = _strAry[46];
                    _dr[52] = _strAry[47];
                    _dr[53] = _strAry[48];
                    _dr[54] = _strAry[49];
                    //2013.01.28 태희철 수정 이용권유동의
                    //50번은 공필드값 14byte
                    _dr[55] = _strAry[51];
                    _dr[56] = _strAry[52];  //2013.04.29 태희철 추가 다이렉트 여부
                    _dr[57] = _strAry[53];  //2013.05.29 태희철 추가 무료보험 여부

                    _dr[58] = _strAry[56];

                    if (_strAry[57].Trim() == "")
                    {
                        _strAry[57] = "9";
                    }

                    if (_strAry[58].Trim() == "")
                    {
                        _strAry[58] = "9";
                    }

                    if (_strAry[59].Trim() == "")
                    {
                        _strAry[59] = "9";
                    }

                    if (_strAry[60].Trim() == "")
                    {
                        _strAry[60] = "9";
                    }

                    if (_strAry[61].Trim() == "")
                    {
                        _strAry[61] = "9";
                    }

                    if (_strAry[62].Trim() == "")
                    {
                        _strAry[62] = "9";
                    }

                    if (_strAry[63].Trim() == "")
                    {
                        _strAry[63] = "9";
                    }

                    if (_strAry[64].Trim() == "")
                    {
                        _strAry[64] = "9";
                    }

                    if (_strAry[65].Trim() == "")
                    {
                        _strAry[65] = "9";
                    }

                    if (_strAry[66].Trim() == "")
                    {
                        _strAry[66] = "9";
                    }
                    //개인정보의 선택적인 제공(3)_해외부정사용
                    if (_strAry[67].Trim() == "")
                    {
                        _strAry[67] = "9";
                    }
                    //개인정보의 선택적인 제공(4)_해외부정사용 고유식별정보 처리
                    if (_strAry[68].Trim() == "")
                    {
                        _strAry[68] = "9";
                    }

                    _dr[59] = _strAry[57] + _strAry[58] + _strAry[59] + _strAry[60] + _strAry[61] + _strAry[62] +
                        _strAry[63] + _strAry[64] + _strAry[65] + _strAry[66] + _strAry[67] + _strAry[68];

                    _dr[60] = _strAry[69];
                    _dr[61] = _strAry[70];
                    _dr[62] = _strAry[71];
                    _dr[63] = _strAry[72];
                    _dr[64] = _strAry[73];

                    if (strCard_type_detail == "0583102")
                    {
                        if ((_dr[11].ToString() == "36" || _dr[11].ToString() == "46") && _strAry[55].ToString().Trim() != "")
                        {
                            _dr[27] = _strAry[55];
                            _dr[28] = _strAry[55];
                            _dr[32] = _strAry[55];
                        }
                    }

                    //2014.12.01 태희철 수정 일반-본인만배송
                    //본인배송코드 : 00 또는 10 본인만배송
                    _strOwner_only = _strAry[74].Trim();

                    if (_strOwner_only.Trim() == "00" || _strOwner_only.Trim() == "10" || _strOwner_only.Trim() == "20"
                        || strCard_type_detail == "0581103" || strCard_type_detail == "0583103")
                    {
                        _strOwner_only = "1";
                    }
                    else
                    {
                        _strOwner_only = "0";
                    }

                    _dr[65] = _strOwner_only;

                    string strPart = "";

                    strPart = _strAry[75].Trim();

                    //수도권 01, 지방 02, 그외 03
                    if (strPart == "01")
                    {
                        _dr[66] = "1";
                    }
                    else if (strPart == "02")
                    {
                        _dr[66] = "2";
                    }
                    else
                    {
                        _dr[66] = "3";
                    }

                    //카드사바코드
                    if (_strZipcode.Trim().Length == 5)
                    {
                        _dr[67] = "0" + _strAry[0].Substring(0, 17) + _strZipcode.Trim() + " ";
                    }
                    else
                    {
                        _dr[67] = "0" + _strAry[0].Substring(0, 17) + _strZipcode.Trim();
                    }


                    //내부사용코드
                    _dr[68] = _dr[7];

                    if (_strDeliveryPlaceType == "01")
                    {
                        _dr[69] = "1";
                    }
                    else if (_strDeliveryPlaceType == "02")
                    {
                        _dr[69] = "2";
                    }
                    else
                    {
                        _dr[69] = "3";
                    }

                    string strCard_consented = _strAry[76];

                    if (strCard_consented.Length > 1)
                    {
                        _dr[77] = _strAry[76].Substring(1, 1);
                    }
                    else
                    {
                        _dr[77] = strCard_consented;
                    }

                    _dr[78] = _strAry[77];


                    //스마트특송 또는 프리미엄카드소지
                    if (_strOwner_only == "1")
                    {
                        _dr[76] = "본인지정배송 - 신분증확인필요";
                    }

                    if (_dr[11].ToString() == "37" || _dr[11].ToString() == "47" || _dr[77].ToString() != "0")
                    {
                        if (_strOwner_only == "1")
                        {
                            _dr[76] = "본인지정배송 / 사전연락 대상입니다. 반드시 고객에게 사전연락 하여주시기 바랍니다";
                        }
                        else
                        {
                            _dr[76] = "사전연락 대상입니다. 반드시 고객에게 사전연락 하여주시기 바랍니다";
                        }
                    }


                    _dr[79] = _strAry[78];
                    _dr[80] = _strAry[79];

                    if (_strZipcode != "")
                    {
                        //지역 분류 선택
                        if (_strZipcode.Length == 5)
                        {
                            _drs = _dsetZipcdeArea_new.Tables[0].Select("zipcode_new = " + _strZipcode);
                        }
                        else
                        {
                            _drs = _dsetZipcdeArea.Tables[0].Select("zipcode = " + _strZipcode);
                        }


                        if (_drs.Length < 1)
                        {
                            _strAreaGroup = "012";
                            _strBranch = "012";
                        }
                        else
                        {
                            _strAreaGroup = _drs[0][0].ToString();
                            _strBranch = _drs[0][1].ToString();
                        }

                        //우편번호 유효성 검사
                        if (_strZipcode.Length == 5)
                        {
                            _drs = _dsetZipcode_new.Tables[0].Select("zipcode_new = " + _strZipcode);
                        }
                        else
                        {
                            _drs = _dsetZipcode.Tables[0].Select("zipcode = " + _strZipcode);
                        }

                        if (_drs.Length > 0)
                        {
                            //시군구 구분 코드 : A,B,C,D...
                            _strAreaType = _drs[0]["area_type_code"].ToString();
                        }
                        else
                        {
                            _strAreaType = "";
                        }

                        //우편번호를 찾지 못 했다면 Error 파일에 기록
                        if (_strAreaType.Equals(""))
                        {
                            _swError.WriteLine(_strLine);
                            _iErrorCount++;
                        }

                        _dr[1] = _strAreaType;
                        _dr[2] = _strBranch;
                        _dr[3] = _strAreaGroup;
                        _dr[4] = _branches.GetCount(_strBranch);
                        _dtable.Rows.Add(_dr);
                    }
                    else
                    {
                        _swError.WriteLine(_strLine);
                        _iErrorCount++;
                    }
                    _iSeq++;
                }

                //변환에 성공했다면
                if (_iErrorCount < 1)
                {
                    _swError.Close();
                    _sr.Close();
                    try
                    {
                        //_dtable의 Row를 재정렬
                        _drs_HD = _dtable.Select("", "client_send_number");

                        _dtable_HD = new DataTable("Convert2");
                        //기본 컬럼
                        _dtable_HD.Columns.Add("degree_arrange_number");
                        _dtable_HD.Columns.Add("card_area_type");
                        _dtable_HD.Columns.Add("card_branch");
                        _dtable_HD.Columns.Add("card_area_group");
                        _dtable_HD.Columns.Add("area_arrange_number");
                        //세부 컬럼
                        _dtable_HD.Columns.Add("client_send_number");             // dr[5] 발송번호
                        _dtable_HD.Columns.Add("card_number");                    // 앞1자리 뺀 발송번호 16자리 
                        _dtable_HD.Columns.Add("card_issue_type_code");           // 카드신청구분코드 : 1 = 신규, 2 = 재발급, 3 = 갱신
                        _dtable_HD.Columns.Add("client_express_code");            // 업체코드 B002
                        _dtable_HD.Columns.Add("client_enterprise_code");         // 제작지점코드
                        _dtable_HD.Columns.Add("card_count");                     // dr[10] 카드매수
                        _dtable_HD.Columns.Add("client_register_type");           // 발송구분코드 : 10=당사교부, 11=특별관리, 20=일반등기 등등
                        _dtable_HD.Columns.Add("card_product_name");              // 카드상품코드명
                        _dtable_HD.Columns.Add("card_brand_code");                // 카드브랜드코드명
                        _dtable_HD.Columns.Add("card_level_code");                // 카드등급코드명
                        _dtable_HD.Columns.Add("customer_name");                  // dr[15] 고객명
                        _dtable_HD.Columns.Add("customer_SSN");                   // 고객주민번호
                        _dtable_HD.Columns.Add("family_name");                    // 가족명
                        _dtable_HD.Columns.Add("family_SSN");                     // 가족주민번호
                        _dtable_HD.Columns.Add("family_name2");                   // 가족명2
                        _dtable_HD.Columns.Add("family_SSN2");                    // dr[20] 가족주민번호2
                        _dtable_HD.Columns.Add("customer_type_code");             // 동봉카드 내역 : 00=본인, 01=본인+가족1, 02=가족1, 03=가족1+가족2
                        _dtable_HD.Columns.Add("card_family_code");               // 분리(가족)배송여부 : Y/N
                        _dtable_HD.Columns.Add("family_relation");                // 신청인관계코드명
                        _dtable_HD.Columns.Add("card_zipcode");                   // 우편번호
                        _dtable_HD.Columns.Add("card_address_local");             // dr[25] 자택주소 동이상
                        _dtable_HD.Columns.Add("card_address_detail");            // 자택주소 동이하
                        _dtable_HD.Columns.Add("card_tel1");                      // 자택전화번호
                        _dtable_HD.Columns.Add("card_mobile_tel");                // 핸드폰번호
                        _dtable_HD.Columns.Add("card_zipcode2");                  // 직장우편번호
                        _dtable_HD.Columns.Add("card_address2_local");            // dr[30] 직장주소 동이상
                        _dtable_HD.Columns.Add("card_address2_detail");           // 직장주소 동이하
                        _dtable_HD.Columns.Add("card_tel2");                      // 직장전화번호
                        _dtable_HD.Columns.Add("customer_office");                // 직장명
                        _dtable_HD.Columns.Add("card_zipcode3");                  // 기타우편번호
                        _dtable_HD.Columns.Add("card_address3");                  // dr[35] 기타주소
                        _dtable_HD.Columns.Add("card_tel3");                      // 기타전화번호
                        _dtable_HD.Columns.Add("customer_email_ID");              // EMAIL ID
                        _dtable_HD.Columns.Add("customer_email_domain");          // EMAIL 도매인
                        _dtable_HD.Columns.Add("family_tel1");                    // 가족자택전화번호
                        _dtable_HD.Columns.Add("family_tel2");                    // dr[40] 가족직장번호
                        _dtable_HD.Columns.Add("family_mobile_tel");              // 가족핸드폰번호
                        _dtable_HD.Columns.Add("card_bill_place_type");           // 청구지구분코드
                        _dtable_HD.Columns.Add("card_delivery_place_code");       // 수령지구분코드
                        _dtable_HD.Columns.Add("card_bank_account_name");         // 자동이체은행코드
                        _dtable_HD.Columns.Add("card_bank_account_no");           //  dr[45] 계좌번호
                        _dtable_HD.Columns.Add("card_payment_day");               // 결제일
                        _dtable_HD.Columns.Add("card_cooperation1");              // 제휴사명
                        _dtable_HD.Columns.Add("card_limit");                     // 희망한도
                        _dtable_HD.Columns.Add("card_agree_code");                // 동의서식별코드 : A=일반, B=MSC, C=웅진코웨이, D=동양증권
                        _dtable_HD.Columns.Add("card_terminal_issue");            // dr[50] 단말기여부 Y/N
                        _dtable_HD.Columns.Add("card_bill_way");                  // 청구서발송방법코드
                        _dtable_HD.Columns.Add("card_vip_code");                  // VIP
                        _dtable_HD.Columns.Add("card_pt_code");                   // 백화점회원등급코드
                        _dtable_HD.Columns.Add("card_register_type");             // 동의서 출력구분
                        //2013.01.28 태희철 추가
                        _dtable_HD.Columns.Add("choice_agree3");                  // dr[55] 이용권유동의
                        _dtable_HD.Columns.Add("save_agreement");                 // dr[56] 세이브다이렉트 여부 : "00" 일반, "01" 다이렉트
                        _dtable_HD.Columns.Add("card_urgency_code");              // dr[57] 무료보험
                        _dtable_HD.Columns.Add("card_cooperation_code");          // dr[58] web신청구분
                        _dtable_HD.Columns.Add("card_cooperation2");              // dr[59] 선택체크 구분

                        _dtable_HD.Columns.Add("card_design_code");               // dr[60] 표준동의코드
                        _dtable_HD.Columns.Add("card_product_code");              // 상품동의코드
                        _dtable_HD.Columns.Add("text1");                          // 제휴동의코드
                        _dtable_HD.Columns.Add("card_cost_code");                 // 신청경로
                        _dtable_HD.Columns.Add("text2");                          // 동의서코드
                        _dtable_HD.Columns.Add("card_is_for_owner_only");         // dr[65] 본인배송코드
                        _dtable_HD.Columns.Add("BC_Part");                        // dr[66] 수도권 : 01, 지방 : 02

                        _dtable_HD.Columns.Add("card_barcode_new");               // dr[67] 카드사바코드 : 0 + 발송번호
                        _dtable_HD.Columns.Add("card_issue_type_new");            // dr[68] 발급구분코드_new
                        _dtable_HD.Columns.Add("card_delivery_place_type");       // dr[69] 내부수령지구분코드

                        _dtable_HD.Columns.Add("card_zipcode_new");               // dr[70]새우편번호
                        _dtable_HD.Columns.Add("card_zipcode_kind");              // dr[71]우편번호구분
                        _dtable_HD.Columns.Add("card_zipcode2_new");              // 새우편번호2
                        _dtable_HD.Columns.Add("card_zipcode2_kind");             // 우편번호
                        _dtable_HD.Columns.Add("card_zipcode3_new");              // 새우편번호3
                        _dtable_HD.Columns.Add("card_zipcode3_kind");             // dr[75]우편번호

                        _dtable_HD.Columns.Add("customer_memo");                  // dr[76]안내문구
                        _dtable_HD.Columns.Add("card_consented");                 // dr[77]프리미엄소지구분
                        _dtable_HD.Columns.Add("client_insert_type");             // dr[78]제휴선택코드

                        _dtable_HD.Columns.Add("card_client_code_1");             // dr[79]리볼빙
                        _dtable_HD.Columns.Add("card_traffic_code");              // dr[80]갱신

                        _iSeq = 1;
                        //area_arrange_number (Total) 값을 재정의 하기 위하여 _branches를 초기화한다.
                        _branches.Clear();
                        //_dtable의 Row를 재정렬하여 _dtable_HD에 담는다
                        for (int i = 0; i < _drs_HD.Length; i++)
                        {
                            _dr_HD = _dtable_HD.NewRow();
                            for (int k = 1; k < _drs_HD[i].ItemArray.Length; k++)
                            {
                                _dr_HD[0] = _iSeq;
                                //k == 4 : area_arrange_number (Total) 값은 재정의를 한다
                                if (k == 4)
                                {
                                    _dr_HD[k] = _branches.GetCount(_drs_HD[i].ItemArray[2].ToString());
                                }
                                else
                                {
                                    _dr_HD[k] = _drs_HD[i].ItemArray[k].ToString();
                                }
                            }
                            _dtable_HD.Rows.Add(_dr_HD);
                            _iSeq++;
                        }
                        //재정렬(card_bank_ID)된 데이터를 보내준다
                        _dtable_HD.WriteXml(xmlPath);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    _strReturn = string.Format("{0}건의 데이터 변환 성공", _iSeq - 1);
                }
                else
                {
                    _strReturn = string.Format("{0}건 변환, 우편번호 미등록 {1}건 실패", _iSeq - 1, _iErrorCount);
                }
            }
            catch (Exception ex)
            {
                if (_swError != null)
                {
                    _strReturn = string.Format("{0}번 째 변환 중 오류", _iSeq);
                    _swError.WriteLine(_strLine);
                    _swError.WriteLine(ex.Message);
                }
            }
            finally
            {
                if (_swError != null) _swError.Close();
                if (_sr != null) _sr.Close();
            }
            return _strReturn;
        }

        //마감 자료
        public static string ConvertResult(System.Data.DataTable dtable, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            StreamWriter _sw01 = null, _sw02 = null, _sw03 = null, _sw04 = null, _sw05 = null, _sw06 = null, _sw07 = null, _sw08 = null, _sw11 = null, _sw12 = null, _sw13 = null, _sw14 = null, _sw15 = null, _sw16 = null, _sw17 = null, _sw18 = null, _sw19 = null, _sw20 = null, _sw21 = null, _sw22 = null;		//파일 쓰기 스트림
            int i = 0;
            StringBuilder _strLine = new StringBuilder("");
            string _strReturn = "", _strStatus = "", strClient_register_type = "", strCard_type_detail = "", strCard_in_date = "", strCustomer_ssn = "", strToDay = "";
            try
            {
                strToDay = DateTime.Now.ToString("yyyyMMdd");

                
                _sw01 = new StreamWriter(fileName + strToDay + "_B002045.TXT.동의.01", true, _encoding);
                _sw02 = new StreamWriter(fileName + strToDay + "_B002047.TXT.스마트동의.01", true, _encoding);
                _sw03 = new StreamWriter(fileName + strToDay + "_B002055.TXT.동의약식.01", true, _encoding);
                _sw04 = new StreamWriter(fileName + strToDay + "_B002057.TXT.스마트약식.01", true, _encoding);
                _sw05 = new StreamWriter(fileName + strToDay + "_B002037.TXT.스마트.01", true, _encoding);
                _sw06 = new StreamWriter(fileName + strToDay + "_B002036.TXT.블랙.01", true, _encoding);
                _sw07 = new StreamWriter(fileName + strToDay + "_B002030.TXT.일반.01", true, _encoding);
                _sw08 = new StreamWriter(fileName + strToDay + "_B0020.TXT.그외.01", true, _encoding);

                _sw11 = new StreamWriter(fileName + strToDay + "_B002045.TXT.동의.00", true, _encoding);
                _sw12 = new StreamWriter(fileName + strToDay + "_B002047.TXT.스마트동의.00", true, _encoding);
                _sw13 = new StreamWriter(fileName + strToDay + "_B002055.TXT.동의약식.00", true, _encoding);                
                _sw14 = new StreamWriter(fileName + strToDay + "_B002057.TXT.스마트약식.00", true, _encoding);                
                _sw15 = new StreamWriter(fileName + strToDay + "_B002037.TXT.스마트.00", true, _encoding);                
                _sw16 = new StreamWriter(fileName + strToDay + "_B002036.TXT.블랙.00", true, _encoding);
                _sw17 = new StreamWriter(fileName + strToDay + "_B002030.TXT.일반.00", true, _encoding);
                _sw18 = new StreamWriter(fileName + strToDay + "_B0020.TXT.그외.00", true, _encoding);

                _sw19 = new StreamWriter(fileName + strToDay + "_B002046.TXT.블랙동의.01", true, _encoding);
                _sw20 = new StreamWriter(fileName + strToDay + "_B002056.TXT.블랙약식.01", true, _encoding);

                _sw21 = new StreamWriter(fileName + strToDay + "_B002046.TXT.블랙동의.00", true, _encoding);
                _sw22 = new StreamWriter(fileName + strToDay + "_B002056.TXT.블랙약식.00", true, _encoding);
                
                
                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    _strStatus = dtable.Rows[i]["card_delivery_status"].ToString();
                    strClient_register_type = dtable.Rows[i]["client_register_type"].ToString();
                    strCard_type_detail = dtable.Rows[i]["card_type_detail"].ToString();
                    strCard_in_date = String.Format("{0:MMdd}",dtable.Rows[i]["card_in_date"]);
                    strCustomer_ssn = dtable.Rows[i]["customer_SSN"].ToString();

                    _strLine = new StringBuilder("");

                    //발송결과코드
                    if (_strStatus == "1")
                    {
                        _strLine.Append(GetStringAsLength("1", 1));
                    }
                    else if (_strStatus == "2" || _strStatus == "3")
                    {
                        _strLine.Append(GetStringAsLength("2", 1));
                    }
                    else if (_strStatus == "6")
                    {
                        _strLine.Append(GetStringAsLength("3", 1));
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength("", 1));
                    }
                    //발송번호
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_send_number"].ToString(), 17));
                    //주민번호
                    if (strCustomer_ssn == "xxxxxxxxxxxxx")
                    {
                        _strLine.Append(GetStringAsLength("", 13));    
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_SSN"].ToString(), 13));
                    }
                    //_strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_SSN"].ToString(), 13));
                    //카드신청구분코드
                    if (dtable.Rows[i]["card_issue_type_code"].ToString() == "")
                    {
                        _strLine.Append(GetStringAsLength("1", 1));
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_issue_type_code"].ToString(), 1));
                    }
                    //배송일+수령인주민번호+수령인관계코드
                    if (_strStatus == "1")
                    {
                        _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[i]["card_delivery_date"].ToString()), 8));
                        //2012.10.23 태희철 배송시간 추가
                        _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[i]["card_delivery_time"].ToString()), 4));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_SSN"].ToString(), 13, true, 'x'));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_code_change"].ToString().Replace("x", ""), 2));
                        _strLine.Append(GetStringAsLength("", 1));
                    }
                    else if (_strStatus == "2" || _strStatus == "3")
                    {
                        _strLine.Append(GetStringAsLength("", 8));
                        _strLine.Append(GetStringAsLength("", 4));
                        _strLine.Append(GetStringAsLength("", 13));
                        _strLine.Append(GetStringAsLength("", 2));
                        //2017.08.04 갱신동의코드가 있고 수취거절의 경우 코드값 "I" 를 전송
                        if (dtable.Rows[i]["card_traffic_code"].ToString() == "1" 
                            && dtable.Rows[i]["delivery_return_reason_last"].ToString() == "05")
                        {
                            _strLine.Append(GetStringAsLength("I", 1));
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["return_code_change"].ToString().Replace("0", ""), 1));
                        }
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength("", 8));
                        _strLine.Append(GetStringAsLength("", 4));
                        _strLine.Append(GetStringAsLength("", 13));
                        _strLine.Append(GetStringAsLength("", 2));
                        _strLine.Append(GetStringAsLength("", 1));
                    }
                    //배송업체코드
                    //_strLine.Append(GetStringAsLength("B002", 4));
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_express_code"].ToString(), 4));
                    //카드번호
                    //_strLine.Append(GetStringAsLength("", 16));
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["career_code"].ToString(), 16, true, ' '));

                    //동의여부
                    if (_strStatus == "2" || _strStatus == "3")
                    {
                        _strLine.Append(GetStringAsLength("N", 1));
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength("Y", 1));
                    }

                    DateTime CardInDate = DateTime.Parse(dtable.Rows[i]["card_in_date"].ToString());
                    DateTime dtDong_date_other = DateTime.Parse("2018-01-12");
                    DateTime dtDong_date = DateTime.Parse("2020-05-08");

                    //2012.08.29 현재 입회신청서변경이력 없음 변동있으므로 주의 요망
                    //입회신청서변경이력
                    if (strClient_register_type == "55" || strClient_register_type == "57")
                    {
                        if (CardInDate < dtDong_date_other)
                        {
                            _strLine.Append(GetStringAsLength("FYAK02", 8));
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength("FYAK04", 8));
                        }
                    }
                    else
                    {
                        if (CardInDate < dtDong_date)
                        {
                            _strLine.Append(GetStringAsLength("FFTA17", 8));
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength("FFTA18", 8));
                        }
                    }

                    //등기번호
                    if (dtable.Rows[i]["card_branch"].ToString() == "012")
                    {
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_SSN"].ToString(), 13, true, 'x'));//등기번호에 민증찍기
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength("", 13)); //등기번호
                    }

                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_name"].ToString(), 16, true, ' '));

                    if (_strStatus == "1" || _strStatus == "2" || _strStatus == "3")
                    {
                        //현대 일반동의서 + 체크동의서
                        if (strClient_register_type == "45")
                        {
                            _sw01.WriteLine(_strLine.ToString());
                        }
                        else if (strClient_register_type == "47")
                        {
                            _sw02.WriteLine(_strLine.ToString());
                        }
                        else if (strClient_register_type == "55")
                        {
                            _sw03.WriteLine(_strLine.ToString());
                        }
                        else if (strClient_register_type == "57")
                        {
                            _sw04.WriteLine(_strLine.ToString());
                        }
                        else if (strClient_register_type == "37")
                        {
                            _sw05.WriteLine(_strLine.ToString());
                        }
                        else if (strClient_register_type == "36")
                        {
                            _sw06.WriteLine(_strLine.ToString());
                        }
                        else if (strClient_register_type == "30")
                        {
                            _sw07.WriteLine(_strLine.ToString());
                        }
                        else if (strClient_register_type == "46")
                        {
                            _sw19.WriteLine(_strLine.ToString());
                        }
                        else if (strClient_register_type == "56")
                        {
                            _sw20.WriteLine(_strLine.ToString());
                        }
                        else
                        {
                            _sw08.WriteLine(_strLine.ToString());
                        }
                    }
                    else
                    {
                        //현대 일반동의서 + 체크동의서
                        if (strClient_register_type == "45")
                        {   
                            _sw11.WriteLine(_strLine.ToString());
                        }
                        else if (strClient_register_type == "47")
                        {
                            _sw12.WriteLine(_strLine.ToString());
                        }
                        else if (strClient_register_type == "55")
                        {
                            _sw13.WriteLine(_strLine.ToString());
                        }
                        else if (strClient_register_type == "57")
                        {
                            _sw14.WriteLine(_strLine.ToString());
                        }
                        else if (strClient_register_type == "37")
                        {
                            _sw15.WriteLine(_strLine.ToString());
                        }
                        else if (strClient_register_type == "36")
                        {
                            _sw16.WriteLine(_strLine.ToString());
                        }
                        else if (strClient_register_type == "30")
                        {
                            _sw17.WriteLine(_strLine.ToString());
                        }
                        else if (strClient_register_type == "46")
                        {
                            _sw21.WriteLine(_strLine.ToString());
                        }
                        else if (strClient_register_type == "56")
                        {
                            _sw22.WriteLine(_strLine.ToString());
                        }
                        else
                        {
                            _sw18.WriteLine(_strLine.ToString());
                        }
                    }
                }
                _strReturn = string.Format("{0}건의 인계데이터 다운 완료", i);
            }
            catch (Exception)
            {
                _strReturn = string.Format("{0}번째 데이터 인계 중 오류 발생", i + 1);
            }
            finally
            {   
                if (_sw01 != null) _sw01.Close();
                if (_sw02 != null) _sw02.Close();
                if (_sw03 != null) _sw03.Close();
                if (_sw04 != null) _sw04.Close();
                if (_sw05 != null) _sw05.Close();
                if (_sw06 != null) _sw06.Close();
                if (_sw07 != null) _sw07.Close();
                if (_sw08 != null) _sw08.Close();
                
                if (_sw11 != null) _sw11.Close();
                if (_sw12 != null) _sw12.Close();
                if (_sw13 != null) _sw13.Close();
                if (_sw14 != null) _sw14.Close();
                if (_sw15 != null) _sw15.Close();
                if (_sw16 != null) _sw16.Close();
                if (_sw17 != null) _sw17.Close();
                if (_sw18 != null) _sw18.Close();

                if (_sw19 != null) _sw19.Close();
                if (_sw20 != null) _sw20.Close();
                if (_sw21 != null) _sw21.Close();
                if (_sw22 != null) _sw22.Close();
            }
            return _strReturn;
        }

        //일일마감자료 
        public static string ConvertResultDay(System.Data.DataTable dtable, string fileName)
        {
            //return ConvertResult(dtable, fileName);
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            StreamWriter _sw00 = null, _sw01 = null, _sw02 = null, _sw03 = null, _sw04 = null, _sw05 = null, _sw06 = null, _sw07 = null, _sw08 = null, _sw09 = null, _sw10 = null, _sw11 = null, _sw12 = null; //파일 쓰기 스트림
            int i = 0, iCnt = 0;
            string strCSV = ",";

            StringBuilder _strLine = new StringBuilder("");
            string _strReturn = "", _strStatus = "", strClient_register_type = "", strToDay = "", strCustomer_ssn = "", strCard_type_detail = "";

            //전자동의서 이미지 변지
            string strClient_send_number = "", strCard_design_code = "", strCard_product_code = "", strText1 = "", strCard_cost_code = "",
                strText2 = "";
            string strChk_01 = "", strChk_02 = "", strChk_03 = "", strChk_04 = "", strChk_05 = "", strChk_06 = "", strChk_07 = "", strChk_08 = "",
                strChk_09 = "", strChk_10 = "", strChk_11 = "", strChk_12 = "", strChk_13 = "";
            string strSign_01 = "", strSign_02 = "", strSign_03 = "", strSign_04 = "", strSign_05 = ""
                , strSign_06 = "", strSign_07 = "", strSign_08 = "", strSign_09 = "";
            string strChkex_01 = "", strChkex_02 = "", strChkex_03 = "", strChkex_04 = "", strChkex_05 = "", strChkex_06 = "";

            try
            {
                strToDay = DateTime.Now.ToString("yyyyMMdd");

                _sw01 = new StreamWriter(fileName + strToDay + "_B002450" + ".TXT.동의서", true, _encoding);   //동의서
                _sw02 = new StreamWriter(fileName + strToDay + "_B002370" + ".TXT.스마트일반", true, _encoding);   //스마트일반
                _sw03 = new StreamWriter(fileName + strToDay + "_B002470" + ".TXT.스마트동의", true, _encoding);   //스마트동의
                _sw04 = new StreamWriter(fileName + strToDay + "_B002550" + ".TXT.약식동의", true, _encoding);   //약식동의
                _sw05 = new StreamWriter(fileName + strToDay + "_B002570" + ".TXT.스마트약식동의", true, _encoding);   //스마트약식동의
                _sw06 = new StreamWriter(fileName + strToDay + "_B002360" + ".TXT.블랙특송", true, _encoding);   //블랙특송
                _sw07 = new StreamWriter(fileName + strToDay + "_B002300" + ".TXT.일반", true, _encoding);   //일반
                _sw08 = new StreamWriter(fileName + strToDay + "_B002460" + ".TXT.블랙동의", true, _encoding);   //블랙동의
                _sw09 = new StreamWriter(fileName + strToDay + "_B002560" + ".TXT.블랙약식", true, _encoding);   //블랙약식
                _sw10 = new StreamWriter(fileName + "ae5_img_list.TXT", true, _encoding);   //동의서 체크값 리스트
                _sw11 = new StreamWriter(fileName + "List.dat", true, _encoding);           //동의서 이미지 리스트
                _sw12 = new StreamWriter(fileName + "List_CN_" + strToDay + ".TXT", true, _encoding);   //동의서 이미지 리스트

                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    _strStatus = dtable.Rows[i]["card_delivery_status"].ToString();
                    strClient_register_type = dtable.Rows[i]["client_register_type"].ToString();
                    strCustomer_ssn = dtable.Rows[i]["customer_SSN"].ToString();

                    //결과코드
                    if (_strStatus == "1")
                    {
                        _strLine = new StringBuilder(GetStringAsLength("1", 1, true, ' '));
                        //발송번호
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_send_number"].ToString(), 17, true, ' '));

                        //주민번호
                        if (strCustomer_ssn == "xxxxxxxxxxxxx")
                        {
                            _strLine.Append(GetStringAsLength("", 13, true, ' '));
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_SSN"].ToString(), 13, true, ' '));
                        }
                        //_strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_SSN"].ToString(), 13, true, ' '));
                        //발송구분코드
                        if (dtable.Rows[i]["card_issue_type_code"].ToString() == "")
                        {
                            _strLine.Append(GetStringAsLength("1", 1, true, ' '));
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_issue_type_code"].ToString(), 1, true, ' '));
                        }

                        _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[i]["card_delivery_date"].ToString()), 8, true, ' '));
                        //2012.10.23 태희철 배송시간 추가
                        _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[i]["card_delivery_time"].ToString()), 4, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_SSN"].ToString(), 13, true, 'x'));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_code_change"].ToString().Replace("x", ""), 2, true, ' '));
                        _strLine.Append(GetStringAsLength("", 1, true, ' '));

                        //배송업체코드
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_express_code"].ToString(), 4, true, ' '));
                        //카드번호
                        //_strLine.Append(GetStringAsLength("", 16, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["career_code"].ToString(), 16, true, ' '));
                        //동의여부
                        _strLine.Append(GetStringAsLength("Y", 1, true, ' '));

                        DateTime CardInDate = DateTime.Parse(dtable.Rows[i]["card_in_date"].ToString());
                        DateTime dtDong_date_other = DateTime.Parse("2018-01-12");
                        DateTime dtDong_date = DateTime.Parse("2020-05-08");

                        //2012.08.29 현재 입회신청서변경이력 없음 변동있으므로 주의 요망
                        //입회신청서변경이력
                        if (strClient_register_type == "55" || strClient_register_type == "57")
                        {
                            if (CardInDate < dtDong_date_other)
                            {
                                _strLine.Append(GetStringAsLength("FYAK02", 8));
                            }
                            else
                            {
                                _strLine.Append(GetStringAsLength("FYAK04", 8));
                            }
                        }
                        else
                        {
                            if (CardInDate < dtDong_date)
                            {
                                _strLine.Append(GetStringAsLength("FFTA17", 8));
                            }
                            else
                            {
                                _strLine.Append(GetStringAsLength("FFTA18", 8));
                            }
                        }

                        if (dtable.Rows[i]["card_branch"].ToString() == "012")
                        {
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_SSN"].ToString(), 13, true, 'x'));//등기번호에 민증찍기
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength("", 13, true, ' ')); //등기번호
                        }

                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_name"].ToString(), 16, true, ' '));

                        //동의서
                        if (strClient_register_type == "45")
                        {
                            iCnt++;
                            _sw01.WriteLine(_strLine.ToString());
                        }
                        //스마트
                        else if (strClient_register_type == "37")
                        {
                            iCnt++;
                            //스마트SSN
                            if (dtable.Rows[i]["receiver_SSN"].ToString().Trim() != dtable.Rows[i]["receiver_SSN_org"].ToString().Trim()
                            && dtable.Rows[i]["card_result_status"].ToString() == "61")
                            {
                                ;
                            }
                            else
                            {
                                _sw02.WriteLine(_strLine.ToString());
                            }

                        }
                        //스마트동의서
                        else if (strClient_register_type == "47")
                        {
                            iCnt++;
                            _sw03.WriteLine(_strLine.ToString());
                        }
                        //약식동의
                        else if (strClient_register_type == "55")
                        {
                            iCnt++;
                            _sw04.WriteLine(_strLine.ToString());
                        }
                        //스마트약식동의
                        else if (strClient_register_type == "57")
                        {
                            iCnt++;
                            _sw05.WriteLine(_strLine.ToString());
                        }
                        //블랙
                        else if (strClient_register_type == "36")
                        {
                            iCnt++;
                            _sw06.WriteLine(_strLine.ToString());
                        }
                        //일반
                        else if (strClient_register_type == "30")
                        {
                            iCnt++;
                            _sw07.WriteLine(_strLine.ToString());
                        }
                        //일반
                        else if (strClient_register_type == "46")
                        {
                            iCnt++;
                            _sw08.WriteLine(_strLine.ToString());
                        }
                        //일반
                        else if (strClient_register_type == "56")
                        {
                            iCnt++;
                            _sw09.WriteLine(_strLine.ToString());
                        }
                        else
                        {
                            _sw00 = new StreamWriter(fileName + ".기타", true, _encoding);
                            _sw00.WriteLine(_strLine.ToString());
                            _sw00.Close();
                        }

                        //전자동의서 개시일 12월 07일 인수데이터 부터 예정
                        //dtDong_date = DateTime.Parse("2020-12-07");

                        //전자동의서 이미지 리스트
                        if (strClient_register_type == "45" || strClient_register_type == "46" || strClient_register_type == "47" ||
                            strClient_register_type == "55" || strClient_register_type == "56" || strClient_register_type == "57")
                        {
                            strClient_send_number = dtable.Rows[i]["client_send_number"].ToString();
                            strCard_design_code = dtable.Rows[i]["card_design_code"].ToString();
                            strCard_product_code = dtable.Rows[i]["card_product_code"].ToString();
                            strText1 = dtable.Rows[i]["Text1"].ToString();
                            strText2 = dtable.Rows[i]["Text2"].ToString();
                            strCard_cost_code = dtable.Rows[i]["Card_cost_code"].ToString();
                            strCard_type_detail = dtable.Rows[i]["card_type_detail"].ToString();

                            //filname값이"" 이면 skip
                            if (dtable.Rows[i]["file_name"].ToString() == "")
                            {
                                continue;
                            }

                            strChk_01 = dtable.Rows[i]["Chk_01"].ToString();
                            strChk_02 = dtable.Rows[i]["Chk_02"].ToString();
                            strChk_03 = dtable.Rows[i]["Chk_03"].ToString();
                            strChk_04 = dtable.Rows[i]["Chk_04"].ToString();
                            strChk_05 = dtable.Rows[i]["Chk_05"].ToString();
                            strChk_06 = dtable.Rows[i]["Chk_06"].ToString();
                            strChk_07 = dtable.Rows[i]["Chk_07"].ToString();
                            strChk_08 = dtable.Rows[i]["Chk_08"].ToString();
                            strChk_09 = dtable.Rows[i]["Chk_09"].ToString();
                            strChk_10 = dtable.Rows[i]["Chk_10"].ToString();
                            strChk_11 = dtable.Rows[i]["Chk_11"].ToString();
                            strChk_12 = dtable.Rows[i]["Chk_12"].ToString();
                            strChk_13 = dtable.Rows[i]["Chk_13"].ToString();

                            strSign_01 = dtable.Rows[i]["Sign_01"].ToString();
                            strSign_02 = dtable.Rows[i]["Sign_02"].ToString();
                            strSign_03 = dtable.Rows[i]["Sign_03"].ToString();
                            strSign_04 = dtable.Rows[i]["Sign_04"].ToString();
                            strSign_05 = dtable.Rows[i]["Sign_05"].ToString();
                            strSign_06 = dtable.Rows[i]["Sign_06"].ToString();
                            strSign_07 = dtable.Rows[i]["Sign_07"].ToString();
                            strSign_08 = dtable.Rows[i]["Sign_08"].ToString();
                            strSign_09 = dtable.Rows[i]["Sign_09"].ToString();

                            strChkex_01 = dtable.Rows[i]["Chkex_01"].ToString();
                            strChkex_02 = dtable.Rows[i]["Chkex_02"].ToString();
                            strChkex_03 = dtable.Rows[i]["Chkex_03"].ToString();
                            strChkex_04 = dtable.Rows[i]["Chkex_04"].ToString();
                            strChkex_05 = dtable.Rows[i]["Chkex_05"].ToString();
                            strChkex_06 = dtable.Rows[i]["Chkex_06"].ToString();

                            if (strChk_01 == "0") strChk_01 = "2";
                            else if (strChk_01 == "" || strChk_01 == "9") strChk_01 = "4";
                            if (strChk_02 == "0") strChk_02 = "2";
                            else if (strChk_02 == "" || strChk_02 == "9") strChk_02 = "4";
                            if (strChk_03 == "0") strChk_03 = "2";
                            else if (strChk_03 == "" || strChk_03 == "9") strChk_03 = "4";
                            if (strChk_04 == "0") strChk_04 = "2";
                            else if (strChk_04 == "" || strChk_04 == "9") strChk_04 = "4";
                            if (strChk_05 == "0") strChk_05 = "2";
                            else if (strChk_05 == "" || strChk_05 == "9") strChk_05 = "4";
                            if (strChk_06 == "0") strChk_06 = "2";
                            else if (strChk_06 == "" || strChk_06 == "9") strChk_06 = "4";
                            if (strChk_07 == "0") strChk_07 = "2";
                            else if (strChk_07 == "" || strChk_07 == "9") strChk_07 = "4";
                            if (strChk_08 == "0") strChk_08 = "2";
                            else if (strChk_08 == "" || strChk_08 == "9") strChk_08 = "4";
                            if (strChk_09 == "0") strChk_09 = "2";
                            else if (strChk_09 == "" || strChk_09 == "9") strChk_09 = "4";
                            if (strChk_10 == "0") strChk_10 = "2";
                            else if (strChk_10 == "" || strChk_10 == "9") strChk_10 = "4";
                            if (strChk_11 == "0") strChk_11 = "2";
                            else if (strChk_11 == "" || strChk_11 == "9") strChk_11 = "4";
                            if (strChk_12 == "0") strChk_12 = "2";
                            else if (strChk_12 == "" || strChk_12 == "9") strChk_12 = "4";
                            if (strChk_13 == "0") strChk_13 = "2";
                            else if (strChk_13 == "" || strChk_13 == "9") strChk_13 = "4";

                            if (strSign_01 == "9") strSign_01 = "4";
                            if (strSign_02 == "9") strSign_02 = "4";
                            if (strSign_03 == "9") strSign_03 = "4";
                            if (strSign_04 == "9") strSign_04 = "4";
                            if (strSign_05 == "9") strSign_05 = "4";
                            if (strSign_06 == "9") strSign_06 = "4";
                            if (strSign_07 == "9") strSign_07 = "4";
                            if (strSign_08 == "9") strSign_08 = "4";
                            if (strSign_09 == "9") strSign_09 = "4";

                            if (strChkex_01 == "0" || strChkex_01 == "") strChkex_01 = "3";
                            else if (strChkex_01 == "9") strChkex_01 = "4";
                            if (strChkex_02 == "0" || strChkex_02 == "") strChkex_02 = "3";
                            else if (strChkex_02 == "9") strChkex_02 = "4";
                            if (strChkex_03 == "0" || strChkex_03 == "") strChkex_03 = "3";
                            else if (strChkex_03 == "9") strChkex_03 = "4";
                            if (strChkex_04 == "0" || strChkex_04 == "") strChkex_04 = "3";
                            else if (strChkex_04 == "9") strChkex_04 = "4";
                            if (strChkex_05 == "0" || strChkex_05 == "") strChkex_05 = "3";
                            else if (strChkex_05 == "9") strChkex_05 = "4";
                            if (strChkex_06 == "0" || strChkex_06 == "") strChkex_06 = "3";
                            else if (strChkex_06 == "9") strChkex_06 = "4";

                            _strLine = new StringBuilder(strToDay + "-");
                            _strLine.Append("B002");
                            _strLine.Append(strToDay.Substring(2, 6) + "-");
                            _strLine.Append(GetStringAsLength(strClient_send_number, 17, true, ' ') + strCSV);
                            _strLine.Append(GetStringAsLength(strCard_design_code, 4, true, ' ') + strCSV);
                            _strLine.Append(GetStringAsLength(strCard_product_code, 4, true, ' ') + strCSV);
                            _strLine.Append(GetStringAsLength(strText1, 4, true, ' ') + strCSV);
                            _strLine.Append(GetStringAsLength(strCard_cost_code.Trim() + strText2.Trim(), 10, true, ' ') + strCSV);

                            //베이직동의
                            if (strCard_type_detail == "0582106")
                            {
                                _strLine.Append(GetStringAsLength(strChk_01, 1, true, ' ') + strCSV);
                                _strLine.Append(GetStringAsLength(strSign_01, 1, true, ' ') + strCSV);
                                _strLine.Append(GetStringAsLength(strChk_02, 1, true, ' ') + strCSV);
                                _strLine.Append(GetStringAsLength(strSign_02, 1, true, ' ') + strCSV);
                                _strLine.Append(GetStringAsLength(strChk_03, 1, true, ' ') + strCSV);
                                _strLine.Append(GetStringAsLength(strSign_03, 1, true, ' ') + strCSV);
                                _strLine.Append("4" + strCSV);
                                _strLine.Append("4" + strCSV);
                                _strLine.Append("4" + strCSV);

                                _strLine.Append("4" + strCSV);
                                _strLine.Append("4" + strCSV);
                                _strLine.Append("4" + strCSV);
                                _strLine.Append("4" + strCSV);
                                _strLine.Append("4" + strCSV);
                                _strLine.Append("4" + strCSV);

                                _strLine.Append("4" + strCSV);
                                _strLine.Append("4" + strCSV);
                                _strLine.Append("4" + strCSV);
                                _strLine.Append("4" + strCSV);
                                _strLine.Append("4" + strCSV);
                                _strLine.Append("4" + strCSV);
                                _strLine.Append("4" + strCSV);
                                _strLine.Append("4" + strCSV);

                                _strLine.Append("4" + strCSV);
                                _strLine.Append("4" + strCSV);

                                _strLine.Append("4" + strCSV);
                                _strLine.Append("4" + strCSV);
                                _strLine.Append("4" + strCSV);
                                _strLine.Append("4");

                                _sw10.WriteLine(_strLine.ToString());
                            }
                            //일반동의서
                            else if (strClient_register_type == "45" || strClient_register_type == "46" || strClient_register_type == "47")
                            {
                                _strLine.Append(GetStringAsLength(strChk_01, 1, true, ' ') + strCSV);
                                _strLine.Append(GetStringAsLength(strSign_01, 1, true, ' ') + strCSV);
                                _strLine.Append(GetStringAsLength(strChk_02, 1, true, ' ') + strCSV);
                                _strLine.Append(GetStringAsLength(strSign_02, 1, true, ' ') + strCSV);
                                _strLine.Append(GetStringAsLength(strChk_03, 1, true, ' ') + strCSV);
                                _strLine.Append(GetStringAsLength(strSign_03, 1, true, ' ') + strCSV);
                                _strLine.Append(GetStringAsLength(strSign_04, 1, true, ' ') + strCSV);
                                _strLine.Append(GetStringAsLength(strChk_04, 1, true, ' ') + strCSV);
                                _strLine.Append(GetStringAsLength(strSign_05, 1, true, ' ') + strCSV);
                                _strLine.Append(GetStringAsLength(strChkex_01, 1, true, ' ') + strCSV);
                                _strLine.Append(GetStringAsLength(strChkex_02, 1, true, ' ') + strCSV);
                                _strLine.Append(GetStringAsLength(strChkex_03, 1, true, ' ') + strCSV);
                                _strLine.Append(GetStringAsLength(strChkex_04, 1, true, ' ') + strCSV);
                                _strLine.Append(GetStringAsLength(strChkex_05, 1, true, ' ') + strCSV);
                                _strLine.Append(GetStringAsLength(strChkex_06, 1, true, ' ') + strCSV);
                                _strLine.Append(GetStringAsLength(strChk_07, 1, true, ' ') + strCSV);
                                _strLine.Append(GetStringAsLength(strChk_08, 1, true, ' ') + strCSV);
                                _strLine.Append(GetStringAsLength(strChk_09, 1, true, ' ') + strCSV);
                                _strLine.Append(GetStringAsLength(strChk_10, 1, true, ' ') + strCSV);

                                //체크동의 구분
                                if (strCard_type_detail == "0582103")
                                {
                                    _strLine.Append("4" + strCSV);
                                    _strLine.Append("4" + strCSV);
                                }
                                else
                                {
                                    _strLine.Append(GetStringAsLength(strChk_11, 1, true, ' ') + strCSV);
                                    _strLine.Append(GetStringAsLength(strChk_12, 1, true, ' ') + strCSV);
                                }

                                _strLine.Append(GetStringAsLength(strSign_08, 1, true, ' ') + strCSV);
                                _strLine.Append(GetStringAsLength(strSign_07, 1, true, ' ') + strCSV);
                                _strLine.Append(GetStringAsLength(strChk_05, 1, true, ' ') + strCSV);
                                _strLine.Append(GetStringAsLength(strSign_06, 1, true, ' ') + strCSV);
                                _strLine.Append("4" + strCSV);
                                //체크동의 구분
                                if (strCard_type_detail == "0582103")
                                {
                                    _strLine.Append(GetStringAsLength(strChk_11, 1, true, ' ') + strCSV);
                                }
                                else
                                {
                                    _strLine.Append(GetStringAsLength(strChk_13, 1, true, ' ') + strCSV);
                                }
                                _strLine.Append(GetStringAsLength(strChk_06, 1, true, ' ') + strCSV);
                                _strLine.Append(GetStringAsLength(strSign_09, 1, true, ' '));

                                _sw10.WriteLine(_strLine.ToString());
                            }
                            //약식동의
                            else if (strClient_register_type == "55" || strClient_register_type == "56" || strClient_register_type == "57")
                            {
                                _strLine.Append("4" + strCSV);
                                _strLine.Append("4" + strCSV);
                                _strLine.Append("4" + strCSV);
                                _strLine.Append("4" + strCSV);
                                _strLine.Append("4" + strCSV);
                                _strLine.Append("4" + strCSV);
                                _strLine.Append("4" + strCSV);
                                _strLine.Append("4" + strCSV);
                                _strLine.Append("4" + strCSV);

                                _strLine.Append(GetStringAsLength(strChkex_01, 1, true, ' ') + strCSV);
                                _strLine.Append(GetStringAsLength(strChkex_02, 1, true, ' ') + strCSV);
                                _strLine.Append(GetStringAsLength(strChkex_03, 1, true, ' ') + strCSV);
                                _strLine.Append(GetStringAsLength(strChkex_04, 1, true, ' ') + strCSV);
                                _strLine.Append(GetStringAsLength(strChkex_05, 1, true, ' ') + strCSV);
                                _strLine.Append(GetStringAsLength(strChkex_06, 1, true, ' ') + strCSV);

                                _strLine.Append("4" + strCSV);
                                _strLine.Append("4" + strCSV);
                                _strLine.Append("4" + strCSV);
                                _strLine.Append("4" + strCSV);
                                _strLine.Append("4" + strCSV);
                                _strLine.Append("4" + strCSV);
                                _strLine.Append("4" + strCSV);
                                _strLine.Append("4" + strCSV);

                                _strLine.Append(GetStringAsLength(strChk_01, 1, true, ' ') + strCSV);
                                _strLine.Append(GetStringAsLength(strSign_01, 1, true, ' ') + strCSV);

                                _strLine.Append("4" + strCSV);
                                _strLine.Append("4" + strCSV);
                                _strLine.Append("4" + strCSV);
                                _strLine.Append("4");

                                _sw10.WriteLine(_strLine.ToString());
                            }
                            else
                            {
                                _sw11 = new StreamWriter(fileName + ".전자동의서_기타", true, _encoding);
                                _sw11.WriteLine(_strLine.ToString());
                                _sw11.Close();
                            }

                            _strLine = new StringBuilder(strClient_send_number);
                            _sw11.WriteLine(_strLine.ToString());

                            _strLine = new StringBuilder(strClient_send_number + strCSV);
                            _strLine.Append(dtable.Rows[i]["customer_name"].ToString());
                            _sw12.WriteLine(_strLine.ToString());
                        }
                    }
                }

                _strReturn = string.Format("{0}건의 인계데이터 다운 완료", iCnt);
            }
            catch (Exception)
            {
                _strReturn = string.Format("{0}번째 데이터 인계 중 오류 발생", i + 1);
            }
            finally
            {
                if (_sw00 != null) _sw00.Close();
                if (_sw01 != null) _sw01.Close();
                if (_sw02 != null) _sw02.Close();
                if (_sw03 != null) _sw03.Close();
                if (_sw04 != null) _sw04.Close();
                if (_sw05 != null) _sw05.Close();
                if (_sw06 != null) _sw06.Close();
                if (_sw07 != null) _sw07.Close();
                if (_sw08 != null) _sw08.Close();
                if (_sw09 != null) _sw09.Close();
                if (_sw10 != null) _sw10.Close();
                if (_sw11 != null) _sw11.Close();
                if (_sw12 != null) _sw12.Close();
            }
            return _strReturn;
        }

        //지역 번호 정리
        private static void ArrangeData(ref DataTable dtable)
        {
            string _strAreaGroup = "", _strPrevGroup = "";
            int _iAreaIndex = 1, _iIndex = -1;
            DataRow[] _dr = dtable.Select("", "card_area_group, card_zipcode, customer_name");
            for (int i = 0; i < _dr.Length; i++)
            {
                _iIndex = int.Parse(_dr[i][0].ToString());
                _strAreaGroup = _dr[i][1].ToString();
                if (_strPrevGroup != _strAreaGroup)
                {
                    _strPrevGroup = _strAreaGroup;
                    _iAreaIndex = 1;
                }
                dtable.Rows[_iIndex - 1][3] = _iAreaIndex;
                _iAreaIndex++;
            }
        }

        //문자중 -를 없앤다
        private static string RemoveDash(string value)
        {
            return value.Replace("-", "");
        }

        //문자에 대해 Length보다 길면 substring하고 짧으면 공백을 넣어서 지정 Length 만큼의 문자를 반환
        private static string GetStringAsLength(string Text, int Length)
        {
            string _strReturn = "";
            System.Text.Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);
            byte[] _byteAry = _encoding.GetBytes(Text);
            if (_byteAry.Length > Length)
            {
                _strReturn = _encoding.GetString(_byteAry, 0, Length);
            }
            else if (_byteAry.Length < Length)
            {
                _strReturn = Text;
                for (int i = 0; i < (Length - _byteAry.Length); i++)
                {
                    _strReturn += ' ';
                }
            }
            else
            {
                _strReturn = Text;
            }
            return _strReturn;
        }
        //문자에 대해 Length보다 길면 substring하고 짧으면 공백을 넣어서 지정 Length 만큼의 문자를 반환
        private static string GetStringAsLength(string Text, int Length, bool blankPositionAtBack, char chBlank)
        {
            string _strReturn = "";
            System.Text.Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);
            byte[] _byteAry = _encoding.GetBytes(Text);

            if (_byteAry.Length > Length)
            {
                _strReturn = _encoding.GetString(_byteAry, 0, Length);
            }
            else if (_byteAry.Length < Length)
            {
                if (blankPositionAtBack)
                {
                    _strReturn = Text;
                }
                else
                {
                    _strReturn = "";
                }
                for (int i = 0; i < (Length - _byteAry.Length); i++)
                {
                    _strReturn += chBlank;
                }
                if (!blankPositionAtBack)
                {
                    _strReturn += Text;
                }
            }
            else
            {
                _strReturn = Text;
            }
            return _strReturn;
        }
        #region 형식함수
        //우편번호 형식
        private static string ConvertZipcode(string value)
        {
            string _strReturn = value;
            if (_strReturn.Length == 6)
            {
                _strReturn = value.Substring(0, 3) + "-" + value.Substring(3, 3);
            }
            return _strReturn;
        }

        //주민등록번호 형식
        private static string ConvertSSN(string value)
        {
            string _strReturn = value;
            if (_strReturn.Length > 6)
            {
                _strReturn = value.Substring(0, 6) + "-" + value.Substring(6, value.Length - 6);
            }
            return _strReturn;
        }
        //전화번호 형식
        private static string ConvertTel(string value)
        {
            string _strReturn = "";
            if (value.Substring(0, 2) == "02")
            {
                if (value.Length == 9)
                {
                    _strReturn = value.Substring(0, 2) + "-" + value.Substring(2, 3) + "-" + value.Substring(5, value.Length - 5);
                }
                else if (value.Length == 10)
                {
                    _strReturn = value.Substring(0, 2) + "-" + value.Substring(2, 4) + "-" + value.Substring(6, value.Length - 6);
                }
                else if (value.Length < 9)
                {
                    if (value.Length < 6)
                    {
                        _strReturn = value.Substring(0, 2) + "-" + value.Substring(2, value.Length - 2) + "-";
                    }
                    else if (value.Length >= 6)
                    {
                        _strReturn = value.Substring(0, 2) + "-" + value.Substring(2, 3) + "-" + value.Substring(5, value.Length - 5);
                    }
                }
            }
            else
            {
                if (value.Length == 10)
                {
                    _strReturn = value.Substring(0, 3) + "-" + value.Substring(3, 3) + "-" + value.Substring(6, value.Length - 6);
                }
                else if (value.Length == 11)
                {
                    _strReturn = value.Substring(0, 3) + "-" + value.Substring(2, 4) + "-" + value.Substring(7, value.Length - 7);
                }
                else if (value.Length < 10)
                {
                    if (value.Length < 7)
                    {
                        _strReturn = value.Substring(0, 3) + "-" + value.Substring(3, value.Length - 3) + "-";

                    }
                    else if (value.Length >= 7)
                    {
                        _strReturn = value.Substring(0, 3) + "-" + value.Substring(3, 3) + "-" + value.Substring(6, value.Length - 6);
                    }
                }
            }
            return _strReturn;
        }
        #endregion
    }
    internal class Branches : Collection<BranchCount>
    {
        internal int GetCount(string strBranch)
        {
            bool _bFind = false;
            int _return = 1;

            if (base.Count > 0)
            {
                for (int i = 0; i < base.Count; i++)
                {
                    if (base[i].Branch == strBranch)
                    {
                        _return = base[i].Count + 1;
                        base[i].AddCount();
                        _bFind = true;
                        break;
                    }
                }
            }

            if (!_bFind)
            {
                base.Add(new BranchCount(strBranch));
            }

            return _return;
        }
    }

    internal class BranchCount
    {
        string branch = "";
        int count = 1;

        internal BranchCount(string strBranch)
        {
            branch = strBranch;
            count = 1;
        }

        internal string Branch
        {
            get { return branch; }
        }

        internal int Count
        {
            get { return count; }
        }

        internal void AddCount()
        {
            count++;
        }
    }
}
