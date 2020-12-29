using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace _086_롯데_test
{
    public class CONVERT
    {
        //기본 인코딩 설정
        private static string strEncoding = "ks_c_5601-1987";
        private static string strCardTypeID = "086";
        private static string strCardTypeName = "롯데_test";

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
        public static string ConvertRegister(string path, string xmlZipcodePath, string xmlZipcodeAreaPath, string xmlZipcodePath_new, string xmlZipcodeAreaPath_new, string xmlPath)
        {
            System.Text.Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            ////FileInfo _fi = null;
            StreamReader _sr = null;														//파일 읽기 스트림
            StreamWriter _swError = null;													//파일 쓰기 스트림
            DataSet _dsetZipcode = null, _dsetZipcdeArea = null;							//우편번호 관련 DataSet
            DataSet _dsetZipcode_new = null, _dsetZipcdeArea_new = null;							//우편번호 관련 DataSet
            DataTable _dtable = null;														//마스터 저장 테이블
            DataRow _dr = null;
            DataRow[] _drs = null;

            DataTable _dtable_LT = null;
            DataRow _dr_LT = null;
            DataRow[] _drs_LT = null;

            byte[] _byteAry = null;
            string _strReturn = "";
            string _strLine = "";
            string _strZipcode = "", _strAreaType = "", _strAreaGroup = "", _strBranch = "", _strDong = "", strCard_type_detail = "", strOwner = "";
            string _strDeliveryPlaceType = "";
            int _iSeq = 1, _iErrorCount = 0;

            string strText = "";
            string[] strAry_text = null;

            Branches _branches = new Branches();
            try
            {
                _dtable = new DataTable("Convert");
                //기본 컬럼
                _dtable.Columns.Add("degree_arrange_number");
                _dtable.Columns.Add("card_area_group");
                _dtable.Columns.Add("card_branch");
                _dtable.Columns.Add("card_area_type");
                _dtable.Columns.Add("area_arrange_number");
                //세부 컬럼
                _dtable.Columns.Add("client_send_date");                      //dr[5]
                _dtable.Columns.Add("client_send_number");
                _dtable.Columns.Add("card_number");
                _dtable.Columns.Add("customer_SSN");
                _dtable.Columns.Add("customer_name");
                _dtable.Columns.Add("family_name");                           //dr[10]
                _dtable.Columns.Add("family_name2");
                _dtable.Columns.Add("card_zipcode");
                _dtable.Columns.Add("card_address_local");
                _dtable.Columns.Add("card_address_detail");
                _dtable.Columns.Add("card_tel1");                             //dr[15]
                _dtable.Columns.Add("card_zipcode2");
                _dtable.Columns.Add("card_address2_local");
                _dtable.Columns.Add("card_address2_detail");
                _dtable.Columns.Add("card_tel2");
                _dtable.Columns.Add("customer_office");                       //dr[20]
                _dtable.Columns.Add("customer_branch");
                _dtable.Columns.Add("card_mobile_tel");
                _dtable.Columns.Add("card_agree_code");                     //일반동의구분 : Y=일반, N=동의
                _dtable.Columns.Add("card_issue_type_code");
                _dtable.Columns.Add("card_consented");                        //dr[25]
                _dtable.Columns.Add("card_delivery_place_code");
                _dtable.Columns.Add("client_register_type");
                _dtable.Columns.Add("card_brand_code");
                _dtable.Columns.Add("card_product_name");                 // 상품 약어명
                _dtable.Columns.Add("client_express_code");               //dr[30] 발송업체코드
                _dtable.Columns.Add("client_request_memo");               // 회원요구메모

                _dtable.Columns.Add("card_bank_account_owner");           // 예금주
                _dtable.Columns.Add("card_bank_account_name");            // 은행명
                _dtable.Columns.Add("card_bank_account_no");              // 계좌번호
                _dtable.Columns.Add("customer_type_code");                // dr[35]고객세분화
                _dtable.Columns.Add("card_product_code");                 // 상품종류
                _dtable.Columns.Add("card_terminal_issue");               // 제휴사마케팅 활용동의 대상

                //중요도 낮음[S]
                _dtable.Columns.Add("card_vip_code");                     // 카드유형
                _dtable.Columns.Add("card_level_code");                   // 유형 일련번호
                _dtable.Columns.Add("customer_no");                       // dr[40]오퍼제공 상품코드
                _dtable.Columns.Add("card_client_no_1");                  // 예약배송일자 + 민원배송유형
                //중요도 낮음[E]

                _dtable.Columns.Add("card_bill_place_type");              // 신용,하이브리드,후불교통 구분 (Y/N)
                _dtable.Columns.Add("card_bill_way");                     // 신용카드 구분 (Y/N)
                _dtable.Columns.Add("card_payment_day");                  // 비대면 동의서 제휴 상품 부가서비스 대상 (Y/N)
                //2012-02-06 태희철 추가[S] 사용안함
                _dtable.Columns.Add("card_issue_detail_code");            // dr[45]위버스카이 : 1 / 다이아몬드 : 2
                // 동의서 식별코드 1:일반, 2:SK스마트, 3:엔크린, 4:웅진, 5:정관장, 6:OK캐쉬백
                //2013.06.13 태희철 수정             
                _dtable.Columns.Add("card_register_type");               // 동의서 식별코드 : 01~09
                _dtable.Columns.Add("card_design_code");                 // 카드배송유형코드 : 발송업체코드 속성분리
                _dtable.Columns.Add("client_number");                    // 회원번호 + 일련번호 (조회키)
                _dtable.Columns.Add("family_customer_no");               // 바탕제휴카드명

                _dtable.Columns.Add("client_insert_type");               // dr[50]고유식별정보 처리동의
                _dtable.Columns.Add("card_urgency_code");               // 멤버스 기능 탑재 유무

                _dtable.Columns.Add("card_cooperation1");                // 본인여부 (11=본인만)
                _dtable.Columns.Add("card_bank_account_tel");                // 제휴사명2

                _dtable.Columns.Add("card_address_type1");             //
                _dtable.Columns.Add("card_address_type2");             //dr[55]
                _dtable.Columns.Add("card_address4_local");            //자택백업(변환) 주소
                _dtable.Columns.Add("card_address4_detail");           //직장백업(변환) 주소
                _dtable.Columns.Add("card_address5_local");            //자택백업(변환) 주소
                _dtable.Columns.Add("card_address5_detail");           //직장백업(변환) 주소

                _dtable.Columns.Add("card_barcode_new");              // dr[60] 카드사바코드
                _dtable.Columns.Add("card_issue_type_new");           // 발급구분코드_new
                _dtable.Columns.Add("card_delivery_place_type");      // 내부수령지코드

                // 제휴서비스 문구
                _dtable.Columns.Add("text1");
                _dtable.Columns.Add("text2");
                _dtable.Columns.Add("text3");                         //dr[65]
                _dtable.Columns.Add("text4");
                _dtable.Columns.Add("text5");
                _dtable.Columns.Add("text6");
                _dtable.Columns.Add("text7");
                _dtable.Columns.Add("text8");                         //dr[70]
                _dtable.Columns.Add("text9");
                _dtable.Columns.Add("text10");

                _dtable.Columns.Add("card_zipcode_new");                //dr[73] 신우편번호
                _dtable.Columns.Add("card_zipcode_kind");               //신우편번호 구분코드
                _dtable.Columns.Add("card_zipcode2_new");               //신우편번호2
                _dtable.Columns.Add("card_zipcode2_kind");              //dr[76] 신우편번호2 구분코드
                _dtable.Columns.Add("card_is_for_owner_only");          //dr[77] 제휴사명1 (본인만배송)
                _dtable.Columns.Add("customer_memo");                   //dr[78] 메모문구a
                _dtable.Columns.Add("change_add");                      //dr[79] 신분증정보

                _dtable.Columns.Add("card_cooperation2");               //dr[80] 묶음번호
                _dtable.Columns.Add("card_bank_ID");                    //dr[81] 묶음대표
                _dtable.Columns.Add("card_count");                      //dr[82] 묶음건수

                //2011-12-12 신주소 관련 추가[E]

                //우편번호 관련 정보 DataSet에 담기
                _dsetZipcode = new DataSet();
                _dsetZipcdeArea = new DataSet();
                _dsetZipcode.ReadXml(xmlZipcodePath);
                _dsetZipcode.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcode.Tables[0].Columns["zipcode"] };
                _dsetZipcdeArea.ReadXml(xmlZipcodeAreaPath);
                _dsetZipcdeArea.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcdeArea.Tables[0].Columns["zipcode"] };

                //신우편번호 관련 정보 담기
                _dsetZipcode_new = new DataSet();
                _dsetZipcdeArea_new = new DataSet();
                _dsetZipcode_new.ReadXml(xmlZipcodePath_new);
                _dsetZipcode_new.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcode_new.Tables[0].Columns["zipcode_new"] };
                _dsetZipcdeArea_new.ReadXml(xmlZipcodeAreaPath_new);
                _dsetZipcdeArea_new.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcdeArea_new.Tables[0].Columns["zipcode_new"] };

                //파일 읽기 Stream과 오류시 저장할 쓰기 Stream 생성
                _sr = new StreamReader(path, _encoding);
                _swError = new StreamWriter(path + ".Error", false, _encoding);

                string strCSS = "";

                while ((_strLine = _sr.ReadLine()) != null)
                {
                    if (_iSeq == 1)
                    {
                        strCard_type_detail = _strLine.Substring(_strLine.Length - 7, 7);
                    }

                    //인코딩, byte 배열로 담기
                    _byteAry = _encoding.GetBytes(_strLine);
                    _strDeliveryPlaceType = _encoding.GetString(_byteAry, 698, 3);
                    //본인만 유무 : 11 = 본인만
                    strOwner = _encoding.GetString(_byteAry, 1167, 2);


                    _dr = _dtable.NewRow();
                    _dr[0] = _iSeq;

                    _dr[5] = _encoding.GetString(_byteAry, 0, 8);
                    _dr[6] = _encoding.GetString(_byteAry, 8, 6);
                    _dr[7] = _encoding.GetString(_byteAry, 0, 14);
                    //strCSS = _encoding.GetString(_byteAry, 14, 13).Replace('*', 'x');
                    strCSS = _encoding.GetString(_byteAry, 14, 6) + _encoding.GetString(_byteAry, 915, 1).Replace(" ", "x") + "xxxxxx";
                    _dr[8] = strCSS;
                    _dr[9] = _encoding.GetString(_byteAry, 27, 40);
                    _dr[10] = _encoding.GetString(_byteAry, 67, 40);
                    _dr[11] = _encoding.GetString(_byteAry, 107, 40);

                    //110 = 직장, 120 = 자택, 131 = 기타
                    //자택이 아닌 코드는 직장으로 취급
                    //120 = 청구지 자택
                    if (_strDeliveryPlaceType == "120")
                    {
                        _strZipcode = _encoding.GetString(_byteAry, 147, 6).Trim();
                        _dr[12] = _strZipcode;
                        _dr[13] = _encoding.GetString(_byteAry, 153, 60);
                        _dr[14] = _encoding.GetString(_byteAry, 213, 150);
                        //_dr[15] = _encoding.GetString(_byteAry, 363, 15);
                        _dr[16] = _encoding.GetString(_byteAry, 378, 6);
                        _dr[17] = _encoding.GetString(_byteAry, 384, 60);
                        _dr[18] = _encoding.GetString(_byteAry, 444, 150); ;
                    }
                    else
                    {
                        _strZipcode = _encoding.GetString(_byteAry, 378, 6).Trim();
                        _dr[12] = _strZipcode;
                        _dr[13] = _encoding.GetString(_byteAry, 384, 60);
                        _dr[14] = _encoding.GetString(_byteAry, 444, 150);
                        //_dr[15] = _encoding.GetString(_byteAry, 594, 15);
                        _dr[16] = _encoding.GetString(_byteAry, 147, 6);
                        _dr[17] = _encoding.GetString(_byteAry, 153, 60);
                        _dr[18] = _encoding.GetString(_byteAry, 213, 150);
                    }

                    //신우편번호의 경우 (청구지)
                    if (_strZipcode.Trim().Length == 5)
                    {
                        _dr[73] = _strZipcode;
                        _dr[74] = "1";
                    }

                    //신우편번호의 경우 (비청구지)
                    if (_dr[16].ToString().Trim().Length == 5)
                    {
                        _dr[75] = _dr[16].ToString();
                        _dr[76] = "1";
                    }

                    _dr[15] = _encoding.GetString(_byteAry, 363, 15);        // 자택전화번호
                    _dr[19] = _encoding.GetString(_byteAry, 594, 15);        // 직장전화번호

                    _dr[20] = _encoding.GetString(_byteAry, 609, 30);
                    _dr[21] = _encoding.GetString(_byteAry, 639, 40);
                    _dr[22] = _encoding.GetString(_byteAry, 679, 15);
                    _strDong = _encoding.GetString(_byteAry, 694, 1);       //일반동의구분 
                    _dr[23] = _strDong;
                    _dr[24] = _encoding.GetString(_byteAry, 695, 2);
                    _dr[25] = _encoding.GetString(_byteAry, 697, 1);
                    _dr[26] = _strDeliveryPlaceType;
                    _dr[27] = _encoding.GetString(_byteAry, 701, 1);
                    _dr[28] = _encoding.GetString(_byteAry, 702, 2);
                    _dr[29] = _encoding.GetString(_byteAry, 704, 20).Replace("'", "");
                    _dr[30] = _encoding.GetString(_byteAry, 740, 2);
                    _dr[31] = _encoding.GetString(_byteAry, 742, 70);

                    _dr[32] = _encoding.GetString(_byteAry, 917, 40);
                    _dr[33] = _encoding.GetString(_byteAry, 957, 30);
                    _dr[34] = _encoding.GetString(_byteAry, 987, 20);
                    _dr[35] = _encoding.GetString(_byteAry, 1007, 2);
                    _dr[36] = _encoding.GetString(_byteAry, 1009, 5);
                    _dr[37] = _encoding.GetString(_byteAry, 1014, 1);
                    _dr[38] = _encoding.GetString(_byteAry, 1015, 3);
                    _dr[39] = _encoding.GetString(_byteAry, 1018, 7);
                    _dr[40] = _encoding.GetString(_byteAry, 1025, 7);
                    _dr[41] = _encoding.GetString(_byteAry, 1032, 9);    // 예약배송일자 + 민원배송유형
                    _dr[42] = _encoding.GetString(_byteAry, 1041, 1);    // 
                    _dr[43] = _encoding.GetString(_byteAry, 1042, 1);    // 
                    _dr[44] = _encoding.GetString(_byteAry, 1043, 1);    // 
                    _dr[45] = _encoding.GetString(_byteAry, 1044, 1);    // 위버스카이, 다이아몬드
                    _dr[46] = _encoding.GetString(_byteAry, 1045, 2);    // 동의서식별코드
                    _dr[47] = _encoding.GetString(_byteAry, 1048, 3);    // 카드배송유형코드 
                    _dr[48] = _encoding.GetString(_byteAry, 1051, 14);   // 신청일자+입회신청일련번호
                    _dr[49] = _encoding.GetString(_byteAry, 1065, 100).Replace("'", "");  // 바탕제휴카드명


                    _dr[50] = _encoding.GetString(_byteAry, 1165, 1);   //고유식별정보 처리코드
                    _dr[51] = _encoding.GetString(_byteAry, 1166, 1);   //멤버스 기능 탑재 유무


                    _dr[52] = _encoding.GetString(_byteAry, 1167, 70);      // 대면(본인)의 경우 11
                    //_dr[52] = _encoding.GetString(_byteAry, 1237, 28);    // 
                    //_dr[52] = _encoding.GetString(_byteAry, 1265, 6);     // 제니엘사용
                    _dr[53] = _encoding.GetString(_byteAry, 1271, 4);       // 2019.08.08 실번호 4자리
                    //_dr[53] = _encoding.GetString(_byteAry, 1275, 90);    // 공백
                    _dr[54] = _encoding.GetString(_byteAry, 1365, 1);
                    _dr[55] = _encoding.GetString(_byteAry, 1366, 1);
                    _dr[56] = "";
                    _dr[57] = _encoding.GetString(_byteAry, 1427, 150);
                    _dr[58] = _encoding.GetString(_byteAry, 1577, 60);
                    _dr[59] = _encoding.GetString(_byteAry, 1637, 150);

                    if (_strDong == "Y")
                    {
                        _strDong = "1";
                    }
                    else
                    {
                        _strDong = "2";
                    }

                    if (strOwner == "11")
                    {
                        _dr[77] = "1";
                        _dr[78] = "";
                        _dr[79] = "1";
                    }

                    //client_express_code(업체코드(2)) + card_zipcode(6) + 발송일자(8) + 업체일련번호(6) + 2 + 카드배송유형코드(3)
                    //총 : 26자리
                    //신우편번호의 경우 신우편번호 + " "
                    if (_strZipcode.Trim().Length == 5)
                    {
                        _dr[60] = _dr[30].ToString() + _dr[12].ToString() + " " + _dr[7].ToString() + _strDong + _dr[47].ToString();
                    }
                    else
                    {
                        _dr[60] = _dr[30].ToString() + _dr[12].ToString() + _dr[7].ToString() + _strDong + _dr[47].ToString();
                    }
                    // filler : 213 / total : 2000 byte

                    //2013-11-14 태희철 [S]
                    //NEW발급구분코드 1:신규, 2:재발급, 3:갱신, 4:재발송, 8:교체추가
                    switch (_dr[24].ToString())
                    {
                        case "11":
                            _dr[61] = "1"; break;
                        case "21":
                        case "22":
                        case "23":
                        case "25":
                        case "27":
                        case "29":
                            _dr[61] = "2"; break;
                        case "26":
                        case "31":
                        case "32":
                        case "33":
                            _dr[61] = "3"; break;
                        case "24":
                        case "28":
                        case "41":
                            _dr[61] = "4"; break;
                        case "12":
                        case "13":
                        case "14":
                        case "15":
                        case "16":
                        case "17":
                            _dr[61] = "8"; break;
                        default:
                            _dr[61] = "1";
                            break;
                    }

                    //직장 = 2, 자택 = 1
                    if (_strDeliveryPlaceType == "110")
                    {
                        _dr[62] = "2";
                    }
                    else if (_strDeliveryPlaceType == "120")
                    {
                        _dr[62] = "1";
                    }
                    else
                    {
                        _dr[62] = "3";
                    }

                    //제휴서비스문구
                    strText = _encoding.GetString(_byteAry, 2000, 1000);

                    //배열 여부 체크
                    if (strText.IndexOf('^') > -1)
                    {
                        //배열일 경우
                        strAry_text = strText.Split('^');

                        for (int i = 0; i < strAry_text.Length && i < 10; i++)
                        {
                            _dr[63 + i] = strAry_text[i].Replace("<f12,u>", "").Replace("</u,f>", "");
                        }
                    }
                    else
                    {
                        _dr[63] = strText.Trim();
                    }

                    //2020.11.04
                    //묶음배송관련 : blank = 일반건, 묶음배송키값 = 34자리
                    _dr[80] = _encoding.GetString(_byteAry, 1367, 34);
                    //묶음배송대표코드 : Y/N
                    _dr[81] = _encoding.GetString(_byteAry, 1401, 1);
                    _dr[82] = _encoding.GetString(_byteAry, 1372, 4);


                    if (_strZipcode != "")
                    {
                        //지역 분류 선택
                        //신우편번호의 경우
                        if (_strZipcode.Trim().Length == 5)
                        {
                            _drs = _dsetZipcdeArea_new.Tables[0].Select("zipcode_new = " + _strZipcode.Trim());
                        }
                        else
                        {
                            _drs = _dsetZipcdeArea.Tables[0].Select("zipcode = " + _strZipcode);
                        }
                        //_drs = _dsetZipcdeArea.Tables[0].Select("zipcode=" + _strZipcode);

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
                        if (_strZipcode.Trim().Length == 5)
                        {
                            _drs = _dsetZipcode_new.Tables[0].Select("zipcode_new = " + _strZipcode.Trim());
                        }
                        else
                        {
                            _drs = _dsetZipcode.Tables[0].Select("zipcode = " + _strZipcode);
                        }
                        //_drs = _dsetZipcode.Tables[0].Select("zipcode=" + _strZipcode);

                        if (_drs.Length > 0)
                        {
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

                        _dr[1] = _strAreaGroup;
                        _dr[2] = _strBranch;
                        _dr[3] = _strAreaType;
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

                if (strCard_type_detail.Substring(0, 4) == "0862")
                {
                    //변환에 성공했다면
                    if (_iErrorCount < 1)
                    {
                        _swError.Close();
                        _sr.Close();

                        try
                        {
                            //_dtable의 Row를 재정렬
                            _drs_LT = _dtable.Select("", "client_send_number");

                            _dtable_LT = new DataTable("Convert2");

                            _dtable_LT.Columns.Add("degree_arrange_number");
                            _dtable_LT.Columns.Add("card_area_group");
                            _dtable_LT.Columns.Add("card_branch");
                            _dtable_LT.Columns.Add("card_area_type");
                            _dtable_LT.Columns.Add("area_arrange_number");
                            //세부 컬럼
                            _dtable_LT.Columns.Add("client_send_date");
                            _dtable_LT.Columns.Add("client_send_number");
                            _dtable_LT.Columns.Add("card_number");                           //dr[7]
                            _dtable_LT.Columns.Add("customer_SSN");
                            _dtable_LT.Columns.Add("customer_name");
                            _dtable_LT.Columns.Add("family_name");                           //dr[10]
                            _dtable_LT.Columns.Add("family_name2");
                            _dtable_LT.Columns.Add("card_zipcode");                          //dr[12]
                            _dtable_LT.Columns.Add("card_address_local");
                            _dtable_LT.Columns.Add("card_address_detail");
                            _dtable_LT.Columns.Add("card_tel1");
                            _dtable_LT.Columns.Add("card_zipcode2");
                            _dtable_LT.Columns.Add("card_address2_local");                   //dr[17]
                            _dtable_LT.Columns.Add("card_address2_detail");
                            _dtable_LT.Columns.Add("card_tel2");
                            _dtable_LT.Columns.Add("customer_office");
                            _dtable_LT.Columns.Add("customer_branch");
                            _dtable_LT.Columns.Add("card_mobile_tel");                     //dr[22]
                            _dtable_LT.Columns.Add("card_agree_code");                     //일반동의구분 : Y=일반, N=동의
                            _dtable_LT.Columns.Add("card_issue_type_code");
                            _dtable_LT.Columns.Add("card_consented");
                            _dtable_LT.Columns.Add("card_delivery_place_code");
                            _dtable_LT.Columns.Add("client_register_type");                //dr[27]
                            _dtable_LT.Columns.Add("card_brand_code");
                            _dtable_LT.Columns.Add("card_product_name");                 // 상품 약어명
                            _dtable_LT.Columns.Add("client_express_code");               // 발송업체코드
                            _dtable_LT.Columns.Add("client_request_memo");               // 회원요구메모

                            _dtable_LT.Columns.Add("card_bank_account_owner");           //dr[32] 예금주
                            _dtable_LT.Columns.Add("card_bank_account_name");            // 은행명
                            _dtable_LT.Columns.Add("card_bank_account_no");              // 계좌번호
                            _dtable_LT.Columns.Add("customer_type_code");                //dr[35] 고객세분화
                            _dtable_LT.Columns.Add("card_product_code");                 // 상품종류
                            _dtable_LT.Columns.Add("card_terminal_issue");               //dr[37] 제휴사마케팅 활용동의 대상

                            //중요도 낮음[S]
                            _dtable_LT.Columns.Add("card_vip_code");                     // 카드유형
                            _dtable_LT.Columns.Add("card_level_code");                   // 유형 일련번호
                            _dtable_LT.Columns.Add("customer_no");                       //dr[40] 오퍼제공 상품코드
                            _dtable_LT.Columns.Add("card_client_no_1");                  // 예약배송일자 + 민원배송유형
                            //중요도 낮음[E]

                            _dtable_LT.Columns.Add("card_bill_place_type");              // 신용,하이브리드,후불교통 구분 (Y/N)
                            _dtable_LT.Columns.Add("card_bill_way");                     // 신용카드 구분 (Y/N)
                            _dtable_LT.Columns.Add("card_payment_day");                  // 비대면 동의서 제휴 상품 부가서비스 대상 (Y/N)

                            //2012-02-06 태희철 추가[S] 사용안함
                            _dtable_LT.Columns.Add("card_issue_detail_code");            //dr[45] 위버스카이 : 1 / 다이아몬드 : 2
                            // 동의서 식별코드 1:일반, 2:SK스마트, 3:엔크린, 4:웅진, 5:정관장, 6:OK캐쉬백
                            //2013.06.13 태희철 수정             
                            _dtable_LT.Columns.Add("card_register_type");               // 동의서 식별코드 : 01~09
                            _dtable_LT.Columns.Add("card_design_code");                 // 카드배송유형코드 : 발송업체코드 속성분리
                            _dtable_LT.Columns.Add("client_number");                    // 회원번호 + 일련번호 (조회키)
                            _dtable_LT.Columns.Add("family_customer_no");               // 바탕제휴카드명

                            _dtable_LT.Columns.Add("client_insert_type");               //dr[50] 고유식별정보 처리동의
                            _dtable_LT.Columns.Add("card_urgency_code");               // 멤버스 기능 탑재 유무

                            _dtable_LT.Columns.Add("card_cooperation1");                // 제휴사명1
                            _dtable_LT.Columns.Add("card_bank_account_tel");                // 제휴사명2

                            _dtable_LT.Columns.Add("card_address_type1");             //
                            _dtable_LT.Columns.Add("card_address_type2");             //dr[55]
                            _dtable_LT.Columns.Add("card_address4_local");            // 자택백업(변환) 주소
                            _dtable_LT.Columns.Add("card_address4_detail");           // 직장백업(변환) 주소
                            _dtable_LT.Columns.Add("card_address5_local");            //자택백업(변환) 주소
                            _dtable_LT.Columns.Add("card_address5_detail");           //직장백업(변환) 주소

                            _dtable_LT.Columns.Add("card_barcode_new");           // dr[60] 카드사바코드
                            _dtable_LT.Columns.Add("card_issue_type_new");          //발급구분코드_new
                            _dtable_LT.Columns.Add("card_delivery_place_type");     //내부수령지코드

                            // 제휴서비스 문구
                            _dtable_LT.Columns.Add("text1");                         //
                            _dtable_LT.Columns.Add("text2");
                            _dtable_LT.Columns.Add("text3");                         // dr[65]
                            _dtable_LT.Columns.Add("text4");
                            _dtable_LT.Columns.Add("text5");
                            _dtable_LT.Columns.Add("text6");
                            _dtable_LT.Columns.Add("text7");
                            _dtable_LT.Columns.Add("text8");                         // dr[70]
                            _dtable_LT.Columns.Add("text9");
                            _dtable_LT.Columns.Add("text10");                        // dr[72]

                            _dtable_LT.Columns.Add("card_zipcode_new");      //신우편번호
                            _dtable_LT.Columns.Add("card_zipcode_kind");     //신우편번호 구분코드
                            _dtable_LT.Columns.Add("card_zipcode2_new");      //신우편번호2
                            _dtable_LT.Columns.Add("card_zipcode2_kind");     //dr[76] 신우편번호2 구분코드
                            _dtable_LT.Columns.Add("card_is_for_owner_only");   //dr[77] 제휴사명1 (본인만배송)
                            _dtable_LT.Columns.Add("customer_memo");            //dr[78] 메모문구
                            _dtable_LT.Columns.Add("change_add");               //dr[79] 신분증정보

                            _dtable_LT.Columns.Add("card_cooperation2");        //dr[80] 묶음번호
                            _dtable_LT.Columns.Add("card_bank_ID");             //dr[81] 묶음대표

                            _iSeq = 1;

                            //area_arrange_number (Total) 값을 재정의 하기 위하여 _branches를 초기화한다.
                            _branches.Clear();
                            //_dtable의 Row를 재정렬하여 _dtable_LT에 담는다
                            for (int i = 0; i < _drs_LT.Length; i++)
                            {
                                _dr_LT = _dtable_LT.NewRow();
                                for (int k = 1; k < _drs_LT[i].ItemArray.Length; k++)
                                {
                                    _dr_LT[0] = _iSeq;
                                    //k == 4 : area_arrange_number (Total) 값은 재정의를 한다
                                    if (k == 4)
                                    {
                                        _dr_LT[k] = _branches.GetCount(_drs_LT[i].ItemArray[2].ToString());
                                    }
                                    else
                                    {
                                        _dr_LT[k] = _drs_LT[i].ItemArray[k].ToString();
                                    }
                                }
                                _dtable_LT.Rows.Add(_dr_LT);
                                _iSeq++;
                            }
                            //재정렬(card_bank_ID)된 데이터를 보내준다
                            _dtable_LT.WriteXml(xmlPath);
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
                else if (_iErrorCount < 1)
                {
                    _swError.Close();
                    _sr.Close();
                    _dtable.WriteXml(xmlPath);
                    _strReturn = string.Format("{0}건의 데이터 변환 성공", _iSeq - 1);
                }
                else
                {
                    _strReturn = string.Format("{0}건 변환, 우편번호 미등록 {1}건 실패", _iSeq - 1, _iErrorCount);
                }

                //변환에 성공했다면
                //if (_iErrorCount < 1)
                //{
                //    _swError.Close();
                //    _sr.Close();
                //    _dtable.WriteXml(xmlPath);
                //    _strReturn = string.Format("{0}건의 데이터 변환 성공", _iSeq - 1);
                //}
                //else
                //{
                //    _strReturn = string.Format("{0}건 변환, 우편번호 미등록 {1}건 실패", _iSeq - 1, _iErrorCount);
                //}

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


        //마감
        public static string ConvertResult(System.Data.DataTable dtable, string fileName)
        {
            string _strReturn = "";
            int _iReturn = 0;
            if (dtable.Rows.Count < 1)
            {
                _strReturn = "조회 자료 없음";
                return _strReturn;
            }
            FormSelectReceive _f = new FormSelectReceive();
            if (_f.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                _iReturn = _f.GetSelected;
            }
            switch (_iReturn)
            {
                case 1:
                    _strReturn = ConvertReceiveType1(dtable, fileName, _iReturn);
                    break;
                case 2:
                    _strReturn = ConvertReceiveType2(dtable, fileName);
                    break;
                case 3:
                    _strReturn = ConvertReceiveType1(dtable, fileName, _iReturn);
                    break;
                case 4:
                    _strReturn = ConvertReceiveType1(dtable, fileName, _iReturn);
                    break;
                default:
                    _strReturn = "";
                    break;
            }
            return _strReturn;
        }

        //마감
        public static string ConvertReceiveType1(System.Data.DataTable dtable, string fileName, int iclose_type)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            StreamWriter _sw00 = null, _sw01 = null, _sw02 = null;					//파일 쓰기 스트림
            int i = 0, icnt = 0;
            StringBuilder _strLine = new StringBuilder("");
            string _strReturn = "", _strStatus = "", strReturn_Code = "", strCard_type_detail = "", strCSS = "";
            string strZipcode_kind = "", strZipcode_kind2 = "";
            try
            {
                //string temp_time = DateTime.Now.ToShortDateString().Replace("-", "").Substring(2, 6);
                string temp_time = DateTime.Now.ToString("yyyyMMdd");

                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    _strStatus = dtable.Rows[i]["card_delivery_status"].ToString();
                    strReturn_Code = dtable.Rows[i]["delivery_return_reason_last"].ToString();
                    strCard_type_detail = dtable.Rows[i]["card_type_detail"].ToString();

                    strZipcode_kind = dtable.Rows[i]["card_zipcode_kind"].ToString();
                    strZipcode_kind2 = dtable.Rows[i]["card_zipcode2_kind"].ToString();

                    //2013.07.26 태희철 분실데이터를 반송과 동일한 형식으로 수정
                    //선반납의 경우 생성되지 않게 수정
                    if ((_strStatus == "2" || _strStatus == "3") && strReturn_Code == "30")
                    {
                        ;
                    }
                    else if (_strStatus == "2" || _strStatus == "3" || _strStatus == "6")
                    {
                        _strLine = new StringBuilder("02");
                        if (strZipcode_kind == "1")
                        {
                            _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[i]["card_zipcode_new"].ToString()), 6, true, ' '));
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[i]["card_zipcode"].ToString()), 6, true, ' '));
                        }
                        //발송일자+일련번호
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_number"].ToString(), 14, true, ' '));
                        //대면여부
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_agree_code"].ToString(), 1, true, ' '));

                        //카드반송사유코드
                        if (_strStatus == "6")
                        {
                            _strLine.Append("16");
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[i]["return_code_change"].ToString()), 2, false, '0'));
                        }

                        //반송수신구분코드 : 롯데카드 사용으로 무조건 01코드 입력
                        _strLine.Append(GetStringAsLength("01", 2, false, ' '));
                        //카드배송유형코드
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_design_code"].ToString(), 3, true, '0'));
                        //배송사원명
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["career"].ToString(), 40, true, ' '));
                        //지사코드
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_branch"].ToString(), 10, true, ' '));
                        //지사명
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["branch_name"].ToString(), 50, true, ' '));
                    }
                    else
                    {
                        strCSS = dtable.Rows[i]["customer_SSN"].ToString();
                        _strLine = new StringBuilder(GetStringAsLength(dtable.Rows[i]["card_number"].ToString(), 14, true, ' '));
                        _strLine.Append(GetStringAsLength(strCSS.Substring(0, 6).Replace("x", "*") + "*******", 13, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_name"].ToString(), 40, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["family_name"].ToString(), 40, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["family_name2"].ToString(), 40, true, ' '));

                        // 2012-02-07 태희철 수령지 : 110=직장, 120=자택
                        // 등록 시 수령지 주소는 수령지, 비수령지 구분하나
                        // 전화번호는 자택card_tel1, 직장card_tel2로 등록 한다.
                        if (dtable.Rows[i]["card_delivery_place_code"].ToString() == "110")
                        {
                            if (dtable.Rows[i]["card_zipcode2"].ToString() == "0")
                            {
                                _strLine.Append(GetStringAsLength("", 6, true, ' '));
                            }
                            else
                            {
                                if (strZipcode_kind2 == "1")
                                {
                                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_zipcode2_new"].ToString(), 6, true, ' '));
                                }
                                else
                                {
                                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_zipcode2"].ToString(), 6, true, ' '));
                                }
                                //_strLine.Append(GetStringAsLength(dtable.Rows[i]["card_zipcode2"].ToString(), 6, true, ' '));
                            }

                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_address2_local"].ToString() + " ", 60, true, ' '));
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_address2_detail"].ToString() + " ", 150, true, ' '));
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_tel1"].ToString(), 15, true, ' '));

                            if (dtable.Rows[i]["card_zipcode"].ToString() == "0")
                            {
                                _strLine.Append(GetStringAsLength("", 6, true, ' '));
                            }
                            else
                            {
                                if (strZipcode_kind == "1")
                                {
                                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_zipcode_new"].ToString(), 6, true, ' '));
                                }
                                else
                                {
                                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_zipcode"].ToString(), 6, true, ' '));
                                }
                            }

                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_address_local"].ToString() + " ", 60, true, ' '));
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_address_detail"].ToString() + " ", 150, true, ' '));
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_tel2"].ToString(), 15, true, ' '));
                        }
                        else
                        {
                            if (dtable.Rows[i]["card_zipcode"].ToString() == "0")
                            {
                                _strLine.Append(GetStringAsLength("", 6, true, ' '));
                            }
                            else
                            {
                                if (strZipcode_kind == "1")
                                {
                                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_zipcode_new"].ToString(), 6, true, ' '));
                                }
                                else
                                {
                                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_zipcode"].ToString(), 6, true, ' '));
                                }
                            }

                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_address_local"].ToString() + " ", 60, true, ' '));
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_address_detail"].ToString() + " ", 150, true, ' '));
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_tel1"].ToString(), 15, true, ' '));

                            if (dtable.Rows[i]["card_zipcode2"].ToString() == "0")
                            {
                                _strLine.Append(GetStringAsLength("", 6, true, ' '));
                            }
                            else
                            {
                                if (strZipcode_kind2 == "1")
                                {
                                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_zipcode2_new"].ToString(), 6, true, ' '));
                                }
                                else
                                {
                                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_zipcode2"].ToString(), 6, true, ' '));
                                }
                            }

                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_address2_local"].ToString() + " ", 60, true, ' '));
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_address2_detail"].ToString() + " ", 150, true, ' '));
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_tel2"].ToString(), 15, true, ' '));
                        }


                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_office"].ToString(), 30, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_branch"].ToString(), 40, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_mobile_tel"].ToString(), 15, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_agree_code"].ToString(), 1, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_issue_type_code"].ToString(), 2, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_consented"].ToString(), 1, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_delivery_place_code"].ToString(), 3, false, '0'));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_register_type"].ToString(), 1, false, '0'));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_brand_code"].ToString(), 2, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_product_name"].ToString(), 20, true, ' '));
                        //회원일련번호 : 차세대는 사용하지 않음
                        _strLine.Append(GetStringAsLength("", 16, true, ' '));
                        _strLine.Append("02");
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_request_memo"].ToString(), 70, true, ' '));

                        //배송업체구간
                        _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[i]["card_delivery_date"].ToString()), 8, true, ' '));
                        if (dtable.Rows[i]["receiver_SSN"].ToString().Length > 6)
                        {
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_SSN"].ToString().Substring(0, 6) + "*******", 13, true, ' '));
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_SSN"].ToString().Replace("x", "*"), 13, true, ' '));
                        }
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_name"].ToString(), 40, true, ' '));

                        if (_strStatus == "1" || _strStatus == "7")
                        {
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_code_change"].ToString().Replace("x", " "), 2, true, ' '));
                        }
                        else if (_strStatus == "2" || _strStatus == "3")
                        {
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["return_code_change"].ToString().Replace("x", " "), 2, true, ' '));
                        }
                        else if (_strStatus == "99")
                        {
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_code_change"].ToString().Replace("x", " "), 2, true, ' '));
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength("", 2, true, ' '));
                        }

                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_delivery_memo"].ToString(), 40, true, ' '));

                        //고객성별 2020.08.21
                        if (strCSS.Substring(6, 1) == "x")
                        {
                            if (dtable.Rows[i]["card_agree_code"].ToString() == "Y")
                            {
                                _strLine.Append(GetStringAsLength("N", 1, true, ' '));
                                _strLine.Append(GetStringAsLength("N", 1, true, ' '));
                            }
                            else
                            {
                                _strLine.Append(GetStringAsLength("Y", 1, true, ' '));
                                _strLine.Append(GetStringAsLength("Y", 1, true, ' '));
                            }
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength(strCSS.Substring(6, 1), 1, true, ' '));

                            if (dtable.Rows[i]["card_agree_code"].ToString() == "Y")
                            {
                                _strLine.Append(GetStringAsLength("N", 1, true, ' '));
                            }
                            else
                            {
                                _strLine.Append(GetStringAsLength("Y", 1, true, ' '));
                            }
                        }

                        //카드사구간
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_bank_account_owner"].ToString(), 40, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_bank_account_name"].ToString(), 30, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_bank_account_no"].ToString(), 20, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_type_code"].ToString(), 2, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_product_code"].ToString(), 5, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_terminal_issue"].ToString(), 1, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_vip_code"].ToString(), 3, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_level_code"].ToString(), 7, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_no"].ToString(), 7, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_client_no_1"].ToString(), 9, true, ' '));

                        // 신용,하이브리드,후불교통 구분 (Y/N)
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_bill_place_type"].ToString(), 1, true, ' '));
                        // 신용카드 구분 (Y/N)
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_bill_way"].ToString(), 1, true, ' '));
                        // 비대면 동의서 제휴 상품 부가서비스 대상 (Y/N)
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_payment_day"].ToString(), 1, true, ' '));

                        //카드사구간
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_issue_detail_code"].ToString(), 1, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_register_type"].ToString(), 2, true, ' '));

                        _strLine.Append(GetStringAsLength("", 1, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_design_code"].ToString(), 3, true, '0'));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_number"].ToString(), 14, true, '0'));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["family_customer_no"].ToString(), 100, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_insert_type"].ToString(), 1, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_urgency_code"].ToString(), 1, true, ' '));
                        _strLine.Append(GetStringAsLength("", 98, true, ' '));
                        _strLine.Append(GetStringAsLength("", 100, true, ' '));

                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_address_type1"].ToString(), 1, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_address_type2"].ToString(), 1, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_address4_local"].ToString(), 60, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_address4_detail"].ToString(), 150, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_address5_local"].ToString(), 60, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_address5_detail"].ToString(), 150, true, ' '));
                        _strLine.Append(GetStringAsLength("", 213, true, ' '));

                        // 제휴서비스 문구
                        if (dtable.Rows[i]["text1"].ToString() == "")
                        {
                            _strLine.Append(GetStringAsLength("", 1000, true, ' '));
                        }
                        else
                        {
                            string strText = "";

                            if (dtable.Rows[i]["text1"].ToString() != "" && dtable.Rows[i]["text2"].ToString() != "")
                            {
                                strText = dtable.Rows[i]["text1"].ToString().Trim() + "^";
                            }
                            else if (dtable.Rows[i]["text1"].ToString() != "")
                            {
                                strText = dtable.Rows[i]["text1"].ToString().Trim();
                            }

                            if (dtable.Rows[i]["text2"].ToString() != "" && dtable.Rows[i]["text3"].ToString() != "")
                            {
                                strText += dtable.Rows[i]["text2"].ToString().Trim() + "^";
                            }
                            else if (dtable.Rows[i]["text2"].ToString() != "")
                            {
                                strText += dtable.Rows[i]["text2"].ToString().Trim();
                            }

                            if (dtable.Rows[i]["text3"].ToString() != "" && dtable.Rows[i]["text4"].ToString() != "")
                            {
                                strText += dtable.Rows[i]["text3"].ToString().Trim() + "^";
                            }
                            else if (dtable.Rows[i]["text3"].ToString() != "")
                            {
                                strText += dtable.Rows[i]["text3"].ToString().Trim();
                            }

                            if (dtable.Rows[i]["text4"].ToString() != "" && dtable.Rows[i]["text5"].ToString() != "")
                            {
                                strText += dtable.Rows[i]["text4"].ToString().Trim() + "^";
                            }
                            else if (dtable.Rows[i]["text4"].ToString() != "")
                            {
                                strText += dtable.Rows[i]["text4"].ToString().Trim();
                            }

                            if (dtable.Rows[i]["text5"].ToString() != "" && dtable.Rows[i]["text6"].ToString() != "")
                            {
                                strText += dtable.Rows[i]["text5"].ToString().Trim() + "^";
                            }
                            else if (dtable.Rows[i]["text5"].ToString() != "")
                            {
                                strText += dtable.Rows[i]["text5"].ToString().Trim();
                            }

                            if (dtable.Rows[i]["text6"].ToString() != "" && dtable.Rows[i]["text7"].ToString() != "")
                            {
                                strText += dtable.Rows[i]["text6"].ToString().Trim() + "^";
                            }
                            else if (dtable.Rows[i]["text6"].ToString() != "")
                            {
                                strText += dtable.Rows[i]["text6"].ToString().Trim();
                            }

                            if (dtable.Rows[i]["text7"].ToString() != "" && dtable.Rows[i]["text8"].ToString() != "")
                            {
                                strText += dtable.Rows[i]["text7"].ToString().Trim() + "^";
                            }
                            else if (dtable.Rows[i]["text7"].ToString() != "")
                            {
                                strText += dtable.Rows[i]["text7"].ToString().Trim();
                            }

                            if (dtable.Rows[i]["text8"].ToString() != "" && dtable.Rows[i]["text9"].ToString() != "")
                            {
                                strText += dtable.Rows[i]["text8"].ToString().Trim() + "^";
                            }
                            else if (dtable.Rows[i]["text8"].ToString() != "")
                            {
                                strText += dtable.Rows[i]["text8"].ToString().Trim();
                            }

                            if (dtable.Rows[i]["text9"].ToString() != "" && dtable.Rows[i]["text10"].ToString() != "")
                            {
                                strText += dtable.Rows[i]["text9"].ToString().Trim() + "^";
                            }
                            else if (dtable.Rows[i]["text9"].ToString() != "")
                            {
                                strText += dtable.Rows[i]["text9"].ToString().Trim();
                            }

                            if (dtable.Rows[i]["text10"].ToString() != "")
                            {
                                strText += dtable.Rows[i]["text10"].ToString().Trim();
                            }

                            _strLine.Append(GetStringAsLength(strText, 1000, true, ' '));
                        }
                    }

                    if (_strStatus == "1")
                    {
                        if (iclose_type == 1)
                        {
                            _sw01 = new StreamWriter(fileName + "차수마감_02_" + temp_time + ".01", true, _encoding);
                            _sw01.WriteLine(_strLine);
                            _sw01.Close();
                            icnt++;
                        }
                        if (iclose_type == 3 && (strCard_type_detail == "0862501" || strCard_type_detail == "0862503"
                                                    || strCard_type_detail == "0863301" || strCard_type_detail == "0863303"))
                        {
                            _sw01 = new StreamWriter(fileName + "월마감_서울_02_" + temp_time + ".01", true, _encoding);
                            _sw01.WriteLine(_strLine);
                            _sw01.Close();
                            icnt++;
                        }
                        else if (iclose_type == 4 && (strCard_type_detail == "0862502" || strCard_type_detail == "0862504"
                                                    || strCard_type_detail == "0863302" || strCard_type_detail == "0863304"))
                        {
                            _sw01 = new StreamWriter(fileName + "월마감_서울외_02_" + temp_time + ".01", true, _encoding);
                            _sw01.WriteLine(_strLine);
                            _sw01.Close();
                            icnt++;
                        }
                    }
                    else if ((_strStatus == "2" || _strStatus == "3") && strReturn_Code == "30")
                    {
                        ;
                    }
                    else if (_strStatus == "2" || _strStatus == "3" || _strStatus == "6")
                    {
                        if (iclose_type == 1)
                        {
                            _sw02 = new StreamWriter(fileName + "반송_국제_" + temp_time, true, _encoding);
                            _sw02.WriteLine(_strLine);
                            _sw02.Close();
                            icnt++;
                        }
                    }
                    else
                    {
                        if (iclose_type == 1)
                        {
                            _sw00 = new StreamWriter(fileName + "차수마감_02_" + temp_time + ".00", true, _encoding);
                            _sw00.WriteLine(_strLine);
                            _sw00.Close();
                            icnt++;
                        }
                    }
                }
                _strReturn = string.Format("{0}건의 인계데이터 다운 완료", icnt);
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
            }
            return _strReturn;
        }


        //선반납마감자료
        private static string ConvertReceiveType2(DataTable dtable, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            StreamWriter _sw02 = null; 		//파일 쓰기 스트림
            int i = 0, icnt = 0;
            StringBuilder _strLine = new StringBuilder("");
            string _strReturn = "", _strStatus = "", strReturn_Code = "";
            string strCard_zipcode_kind = "";

            try
            {
                string temp_time = DateTime.Now.ToShortDateString().Replace("-", "").Substring(2, 6);

                _sw02 = new StreamWriter(fileName + ".선반송_국제_" + temp_time, true, _encoding);

                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    _strStatus = dtable.Rows[i]["card_delivery_status"].ToString();
                    strReturn_Code = dtable.Rows[i]["delivery_return_reason_last"].ToString();
                    strCard_zipcode_kind = dtable.Rows[i]["card_zipcode_kind"].ToString();

                    //선반납 + 불가지역의 경우 생성 되게 수정
                    if ((_strStatus == "2" || _strStatus == "3") && (strReturn_Code == "30" || strReturn_Code == "88"))
                    {
                        _strLine = new StringBuilder("02");

                        if (strCard_zipcode_kind == "1")
                        {
                            _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[i]["card_zipcode_new"].ToString()), 6, true, ' '));
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[i]["card_zipcode"].ToString()), 6, true, ' '));
                        }

                        //발송일자+일련번호
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_number"].ToString(), 14, true, ' '));
                        //대면여부
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_agree_code"].ToString(), 1, true, ' '));

                        //카드반송사유코드
                        if (_strStatus == "6")
                        {
                            _strLine.Append("16");
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[i]["return_code_change"].ToString()), 2, false, '0'));
                        }

                        //반송수신구분코드 : 롯데카드 사용으로 무조건 01코드 입력
                        _strLine.Append(GetStringAsLength("01", 2, false, ' '));
                        //카드배송유형코드
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_design_code"].ToString(), 3, true, '0'));
                        //배송사원명
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["career"].ToString(), 40, true, ' '));
                        //지사코드
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_branch"].ToString(), 10, true, ' '));
                        //지사명
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["branch_name"].ToString(), 50, true, ' '));
                    }

                    if ((_strStatus == "2" || _strStatus == "3") && (strReturn_Code == "30" || strReturn_Code == "88"))
                    {
                        icnt++;
                        _sw02.WriteLine(_strLine);
                    }
                }
                _strReturn = string.Format("{0}건의 인계데이터 다운 완료", icnt);
            }
            catch (Exception)
            {
                _strReturn = string.Format("{0}번째 데이터 인계 중 오류 발생", icnt + 1);
            }
            finally
            {
                if (_sw02 != null) _sw02.Close();
            }
            return _strReturn;
        }


        public static void ChangeAddress(DataTable dtable, DataTable dtable2, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            StreamWriter _sw03 = null;					//파일 쓰기 스트림
            int i = 0, j = 0;
            StringBuilder _strLine = new StringBuilder("");
            string strToDay = "";
            try
            {
                strToDay = DateTime.Now.ToString("yyMMdd");
                _sw03 = new StreamWriter(fileName + ".ADDR_02_" + strToDay, true, _encoding);

                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    //발송일자 + 발송업체일련번호
                    _strLine = new StringBuilder(GetStringAsLength(dtable.Rows[i]["card_number"].ToString(), 14, true, ' '));
                    //발송업체코드
                    _strLine.Append("02");
                    //카드수령자명
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_name"].ToString(), 40, true, ' '));
                    //변경요청일자(주소변경요청일자)
                    _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[i]["change_regdate"].ToString()), 8, true, ' '));
                    //변경된 수령 전체주소
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["change_address"].ToString() + " "
                        + dtable.Rows[i]["change_address_detail"].ToString(), 150, true, ' '));
                    //수령지주소구분코드
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_delivery_place_code"].ToString(), 3, true, ' '));

                    _strLine.Append(GetStringAsLength("1", 1, true, ' '));
                    _strLine.Append(GetStringAsLength("", 1072, true, ' '));
                    //카드배송유형코드
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_design_code"].ToString(), 3, true, '0'));
                    _sw03.WriteLine(_strLine);
                }

                //2012-05-17 민원소명자료
                for (j = 0; j < dtable2.Rows.Count; j++)
                {
                    //발송일자 + 발송업체일련번호
                    _strLine = new StringBuilder(GetStringAsLength(dtable2.Rows[j]["card_number"].ToString(), 14, true, ' '));
                    //발송업체코드
                    _strLine.Append("02");
                    //카드수령자명
                    _strLine.Append(GetStringAsLength(dtable2.Rows[j]["customer_name"].ToString(), 40, true, ' '));
                    //변경요청일자(주소변경요청일자)
                    _strLine.Append(GetStringAsLength("", 8, true, ' '));
                    //변경된 수령 전체주소
                    _strLine.Append(GetStringAsLength("", 150, true, ' '));
                    //수령지주소구분코드
                    _strLine.Append(GetStringAsLength("", 3, true, ' '));

                    //민원,주소변경 구분
                    _strLine.Append(GetStringAsLength("9", 1, true, ' '));
                    //민원소명내용
                    _strLine.Append(GetStringAsLength(dtable2.Rows[j]["appeal_memo"].ToString().Replace("\r", " ").Replace("\n", " "), 1000, true, ' '));
                    //발송업체지사명
                    _strLine.Append(GetStringAsLength(dtable2.Rows[j]["branch_name"].ToString(), 50, true, ' '));
                    //처리자사원명
                    _strLine.Append(GetStringAsLength(dtable2.Rows[j]["appeal_result_staff_name"].ToString(), 22, true, ' '));
                    //카드배송유형코드
                    _strLine.Append(GetStringAsLength(dtable2.Rows[j]["card_design_code"].ToString(), 3, true, '0'));

                    _sw03.WriteLine(_strLine);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ADDR파일 생성 중 오류 / " + ex.Message);
            }
            finally
            {
                if (_sw03 != null) _sw03.Close();
            }
        }

        //일일마감자료
        public static string ConvertResultDay(System.Data.DataTable dtable, string fileName)
        {
            return "카드사마감일 자료 다운만 가능합니다.";
        }
        //지역 번호 정리
        //2015.07.27 태희철 사용하지 않는 함수
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
