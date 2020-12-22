using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace _002_국민
{
    public class CONVERT
    {
        //기본 인코딩 설정
        private static string strEncoding = "ks_c_5601-1987";
        private static string strCardTypeID = "002";
        private static string strCardTypeName = "국민카드";

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
            StreamReader _sr = null;														//파일 읽기 스트림
            StreamWriter _swError = null;													//파일 쓰기 스트림
            DataSet _dsetZipcode = null, _dsetZipcdeArea = null;							//우편번호 관련 DataSet
            DataSet _dsetZipcode_new = null, _dsetZipcdeArea_new = null;					//우편번호 관련 DataSet
            DataTable _dtable = null;														//마스터 저장 테이블
            DataRow _dr = null;
            DataRow[] _drs = null;
            byte[] _byteAry = null;
            string _strReturn = "";
            string _strLine = "";
            string _strZipcode = "", _strAreaType = "", _strAreaGroup = "", _strBranch = "";
            string _strSendNumber = "", _strSendNumberPrev = "";
            string _strFamilyNames = "", _strFamilySSNs = "", _strFamilyCustomerNo = "", _strCustomerSSN = "", _strCardNumberNo = "";
            string _strFamilyCode = "", _strCustomerName = "", _strFamilyName = "", _strFamilySSN = "", _strCardNumber = "", _strFamilyCardNumber = "", _strOwner_only = "", _strFamily_one = "", _strOwner_one = "", _strCode2 = "", strAreaCode = "", strEmerDong = "";
            int _iSeq = 1, _iErrorCount = 0;
            int _iAddCount = 0;
            string _strNewAdd = null, strCard_type_detail = "";

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
                _dtable.Columns.Add("client_insert_type");              // dr[5]
                _dtable.Columns.Add("client_register_type");
                _dtable.Columns.Add("client_send_date");
                _dtable.Columns.Add("client_enterprise_code");          //업체코드 : 당사 인편(1), 동의(5), 긴급(51) 등등
                _dtable.Columns.Add("card_client_no_1");
                _dtable.Columns.Add("customer_SSN");                    // dr[10]
                _dtable.Columns.Add("card_number");
                _dtable.Columns.Add("customer_name");
                _dtable.Columns.Add("family_name");
                _dtable.Columns.Add("card_bank_ID");
                _dtable.Columns.Add("card_mobile_tel");                 // dr[15]
                _dtable.Columns.Add("choice_agree1");                     //개인신용정보 필수적 수집/이용 제공 동의
                _dtable.Columns.Add("choice_agree2");                     //상기목적의 고유식별정보 처리 동의
                //카드이용계약, 상기목적, 카드상품 안내, 부수서비스, 계열사 정보제공, 부정사용 방지, 사용권유(4자리)
                _dtable.Columns.Add("choice_agree3");
                _dtable.Columns.Add("card_zipcode");                    //
                _dtable.Columns.Add("card_address_local");              // dr[20]
                _dtable.Columns.Add("card_address_detail");             // 
                _dtable.Columns.Add("card_count");                      // 카드수량(굴비:P)
                _dtable.Columns.Add("card_issue_type_code");            // 발급구분
                _dtable.Columns.Add("card_delivery_place_type");        // 
                _dtable.Columns.Add("card_cooperation_code");           // dr[25] 제휴코드
                _dtable.Columns.Add("client_register_date");            // 
                _dtable.Columns.Add("client_send_number");
                _dtable.Columns.Add("client_express_code");
                _dtable.Columns.Add("card_consented");                  // 일반 / 기업 카드구분
                _dtable.Columns.Add("client_number");                   // dr[30]
                _dtable.Columns.Add("card_family_code");                // 가족발급, 본인동시
                _dtable.Columns.Add("family_SSN");
                _dtable.Columns.Add("family_customer_no");
                _dtable.Columns.Add("card_add_count");
                _dtable.Columns.Add("save_agreement");                  // dr[35]
                _dtable.Columns.Add("card_barcode_new");                // 
                _dtable.Columns.Add("client_request_memo");             // dr[37]


                //2011-10-05 태희철 추가 신주소코드
                _dtable.Columns.Add("npi_file_name");                   // dr[38]
                _dtable.Columns.Add("card_is_for_owner_only");          // dr[39] 제3자수령가능여부
                _dtable.Columns.Add("card_urgency_code");               // dr[40] 가족앞필수교부
                //2016.06.23 태희철 추가 결제계좌
                _dtable.Columns.Add("customer_position");               // dr[41] 지역구분
                _dtable.Columns.Add("card_design_code");                // dr[42] 카드결제계좌기관코드
                _dtable.Columns.Add("card_bank_account_no");            // dr[43] 카드결제계좌번호

                _dtable.Columns.Add("card_issue_type_new");             // dr[44] 발급구분코드_new

                _dtable.Columns.Add("card_zipcode_new");                // dr[45] 신우편번호
                _dtable.Columns.Add("card_zipcode_kind");               // dr[46] 우편번호구분값

                _dtable.Columns.Add("customer_order");             //메모코드
                _dtable.Columns.Add("customer_memo");              //dr[48] 메모문구
                _dtable.Columns.Add("card_pt_code");              //dr[49]신분증 진위 사후 확인
                _dtable.Columns.Add("change_add");                //dr[50]신분증 확인 입력창 여부
                _dtable.Columns.Add("card_bank_account_tel");     //dr[51]실번호 뒷4자리


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
                _dsetZipcode_new.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcode.Tables[0].Columns["zipcode_new"] };
                _dsetZipcdeArea_new.ReadXml(xmlZipcodeAreaPath_new);
                _dsetZipcdeArea_new.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcdeArea.Tables[0].Columns["zipcode_new"] };

                //파일 읽기 Stream과 오류시 저장할 쓰기 Stream 생성
                _sr = new StreamReader(path, _encoding);
                _swError = new StreamWriter(path + ".Error", false, _encoding);
                string strOther_chk = "";

                while ((_strLine = _sr.ReadLine()) != null)
                {
                    //인코딩, byte 배열로 담기
                    if (_iSeq == 1)
                    {
                        strCard_type_detail = _strLine.Substring(_strLine.Length - 7, 7);
                    }
                    //인코딩, byte 배열로 담기
                    _byteAry = _encoding.GetBytes(_strLine);

                    _dr = _dtable.NewRow();
                    _dr[0] = _iSeq;

                    //2011-10-05 태희철 수정 제휴코드 4->5byte
                    _strSendNumber = _encoding.GetString(_byteAry, 329, 16);
                    //가족만발급 (0:본인만1매or본인+가족, 1:가족만1매, 2:가족만2매이상, 3:본인동시교부)
                    _strFamilyCode = _encoding.GetString(_byteAry, 345, 1);
                    //제3자수령가능여부 1 = 동의, 0 = 미동의
                    _strOwner_one = _encoding.GetString(_byteAry, 398, 1);
                    //가족앞필수교부 : 1 = 동의, 0 = 미동의
                    _strFamily_one = _encoding.GetString(_byteAry, 399, 1);
                    //동의서 구분 : 동의서 = D
                    _strCode2 = _encoding.GetString(_byteAry, 1, 1);
                    //긴급동의
                    strEmerDong = _encoding.GetString(_byteAry, 326, 2);
                    //대외기관체크코드 : 01 = KB증권
                    strOther_chk = _encoding.GetString(_byteAry, 431, 2);

                    if ((_strCode2 == "E" || _strCode2 == "F" || _strCode2 == "L" || _strCode2 == "M"))
                    {
                        _strOwner_only = "1";
                    }
                    //2019.04.19 KB증권이며 제3자수령가능 미동의
                    else if (strOther_chk == "01" && _strOwner_one == "0")
                    {
                        _strOwner_only = "1";
                    }
                    else if (((_strFamilyCode == "0" || _strFamilyCode == "3") && _strOwner_one == "0") || (_strFamilyCode == "1" && _strFamily_one == "1"))
                    {
                        _strOwner_only = "1";
                    }
                    else
                    {
                        _strOwner_only = "0";
                    }

                    //2018.03.27 태희철 수정
                    //법인의 경우 무조건 고객명과 생년월일을 소유자로 한다
                    if (_strCode2 == "E" || _strCode2 == "F" || _strCode2 == "L" || _strCode2 == "M")
                    {
                        _strCustomerName = _encoding.GetString(_byteAry, 47, 20);
                        _strCustomerSSN = _encoding.GetString(_byteAry, 18, 13).Replace("*", "x");
                        _strFamilyName = _encoding.GetString(_byteAry, 67, 20);
                        _strFamilySSN = _encoding.GetString(_byteAry, 346, 13).Replace("*","x");
                        _strCardNumber = _encoding.GetString(_byteAry, 31, 16);
                        _strFamilyCardNumber = _strSendNumber;
                    }
                    //제3자수령가능여부 : 0 = 미동의, 1 = 동의
                    //본인1매 OR 본인+가족 중 제3자수령가능 미동의
                    //본인회원
                    else if ((_strFamilyCode == "0" && _strOwner_one == "0") &&
                        (_strCode2 != "A" || _strCode2 != "D" || _strCode2 != "Y" || _strCode2 != "S"))
                    {
                        _strCustomerName = _encoding.GetString(_byteAry, 47, 20);
                        _strCustomerSSN = _encoding.GetString(_byteAry, 346, 13).Replace("*", "x");
                        _strFamilyName = _encoding.GetString(_byteAry, 67, 20);
                        _strFamilySSN = _encoding.GetString(_byteAry, 18, 13).Replace("*", "x");
                        _strCardNumber = _encoding.GetString(_byteAry, 31, 16);
                        _strFamilyCardNumber = _strSendNumber;
                    }
                    //동의서가 아니고 가족앞필수교부 : 0 = 미동의, 1 = 동의
                    //소유자회원
                    else if ((_strFamilyCode == "1" && _strFamily_one == "1") &&
                        (_strCode2 != "A" || _strCode2 != "D" || _strCode2 != "Y" || _strCode2 != "S"))
                    {
                        _strCustomerName = _encoding.GetString(_byteAry, 47, 20);
                        _strCustomerSSN = _encoding.GetString(_byteAry, 18, 13).Replace("*", "x");
                        _strFamilyName = _encoding.GetString(_byteAry, 67, 20);
                        _strFamilySSN = _encoding.GetString(_byteAry, 346, 13).Replace("*", "x");
                        _strCardNumber = _encoding.GetString(_byteAry, 31, 16);
                        _strFamilyCardNumber = _strSendNumber;
                    }
                    //동의서 중 가족만 1매 이고 가족앞필수배송의 경우
                    //소유자회원
                    else if ((_strCode2 == "A" || _strCode2 == "D" || _strCode2 == "Y" || _strCode2 == "S") &&
                        (_strFamilyCode == "1" && _strFamily_one == "1"))
                    {
                        _strCustomerName = _encoding.GetString(_byteAry, 47, 20);
                        _strCustomerSSN = _encoding.GetString(_byteAry, 18, 13).Replace("*", "x");
                        _strFamilyName = _encoding.GetString(_byteAry, 67, 20);
                        _strFamilySSN = _encoding.GetString(_byteAry, 346, 13).Replace("*", "x");
                        _strCardNumber = _encoding.GetString(_byteAry, 31, 16);
                        _strFamilyCardNumber = _strSendNumber;
                    }
                    //소유자가본인을 포함한 카드이거나 가족카드 1건만 발급된 경우
                    //소유자회원
                    else if (_strFamilyCode.Trim() == "" || _strFamilyCode == "0" || _strFamilyCode == "1")
                    {
                        _strCustomerName = _encoding.GetString(_byteAry, 47, 20);
                        _strFamilyName = _encoding.GetString(_byteAry, 67, 20);
                        _strCustomerSSN = _encoding.GetString(_byteAry, 18, 13).Replace("*", "x");
                        _strFamilySSN = _encoding.GetString(_byteAry, 346, 13).Replace("*", "x");
                        _strCardNumber = _encoding.GetString(_byteAry, 31, 16);
                        _strFamilyCardNumber = _strSendNumber;
                    }
                    //본인동시교부
                    else if (_strFamilyCode == "3")
                    {
                        _strCustomerName = _encoding.GetString(_byteAry, 47, 20);
                        _strFamilyName = _encoding.GetString(_byteAry, 67, 20);
                        _strCustomerSSN = _encoding.GetString(_byteAry, 18, 13).Replace("*", "x");
                        _strFamilySSN = _encoding.GetString(_byteAry, 346, 13).Replace("*", "x");
                        _strCardNumber = _encoding.GetString(_byteAry, 31, 16);
                        _strFamilyCardNumber = _strSendNumber;
                    }
                    //그 외는 소유자와 가족위치를 변경한다.
                    //본인회원
                    else
                    {
                        _strCustomerName = _encoding.GetString(_byteAry, 67, 20);
                        _strFamilyName = _encoding.GetString(_byteAry, 47, 20);
                        _strCustomerSSN = _encoding.GetString(_byteAry, 346, 13).Replace("*", "x");
                        _strFamilySSN = _encoding.GetString(_byteAry, 18, 13).Replace("*", "x");
                        _strCardNumber = _encoding.GetString(_byteAry, 31, 16);
                        _strFamilyCardNumber = _strSendNumber;
                    }

                    if (_strSendNumber != _strSendNumberPrev)
                    {
                        _iAddCount = 0;
                        _strSendNumberPrev = _strSendNumber;

                        _dr[5] = _encoding.GetString(_byteAry, 0, 1);
                        _dr[6] = _encoding.GetString(_byteAry, 1, 1);
                        _dr[7] = _encoding.GetString(_byteAry, 2, 8);
                        _dr[8] = _encoding.GetString(_byteAry, 10, 2);
                        _dr[9] = _encoding.GetString(_byteAry, 12, 6);
                        //_strCustomerSSN = _encoding.GetString(_byteAry, 18, 13);
                        _dr[10] = _strCustomerSSN;
                        _dr[11] = _strCardNumber;//_encoding.GetString(_byteAry, 31, 16);
                        _dr[12] = _strCustomerName; //_encoding.GetString(_byteAry, 47, 20);
                        _dr[13] = _strFamilyName;//_encoding.GetString(_byteAry, 67, 20);
                        _dr[14] = _encoding.GetString(_byteAry, 87, 4);
                        _dr[15] = _encoding.GetString(_byteAry, 91, 15);
                        //2012.07.17 태희철 수정[S]
                        _dr[16] = _encoding.GetString(_byteAry, 106, 1);
                        _dr[17] = _encoding.GetString(_byteAry, 107, 1);

                        _dr[18] = GetStringAsLength(_encoding.GetString(_byteAry, 108, 10).Trim(), 10, true, '9');

                        _strZipcode = _encoding.GetString(_byteAry, 118, 6).Trim();

                        if (_strZipcode.Length == 5)
                        {
                            _dr[45] = _strZipcode;
                            _dr[46] = "1";
                        }

                        _dr[19] = _strZipcode;
                        _dr[20] = _encoding.GetString(_byteAry, 124, 60);
                        _dr[21] = _encoding.GetString(_byteAry, 184, 120);

                        //_dr[23] = _encoding.GetString(_byteAry, 337, 1);
                        //2011-08-17 태희철 수정 P : 굴비카드 이면 9로 사용[S]
                        //주의사항 대소문자 구분
                        if (_encoding.GetString(_byteAry, 304, 1).ToString() == "P")
                        {
                            _dr[22] = 9;
                        }
                        else if (_encoding.GetString(_byteAry, 304, 1).ToString() == "0")
                        {
                            _dr[22] = 1;
                        }
                        else
                        {
                            _dr[22] = _encoding.GetString(_byteAry, 304, 1);
                        }
                        //2011-08-17[E]

                        _dr[23] = _encoding.GetString(_byteAry, 305, 1);
                        _dr[24] = _encoding.GetString(_byteAry, 306, 1);
                        _dr[25] = _encoding.GetString(_byteAry, 307, 5);             //제휴코드
                        _dr[26] = _encoding.GetString(_byteAry, 312, 8);
                        _dr[27] = _encoding.GetString(_byteAry, 320, 6);
                        _dr[28] = _encoding.GetString(_byteAry, 326, 2);
                        _dr[29] = _encoding.GetString(_byteAry, 328, 1);             // 기업카드일경우 문구 추가
                        _dr[30] = _strSendNumber;
                        _dr[31] = _strFamilyCode;

                        if (_strCustomerSSN != _strFamilySSN)
                        {
                            _dr[32] = _strFamilySSN;
                        }

                        if (_strFamilyCode.Trim() == "" || _strFamilyCode == "0" || _strFamilyCode == "1")
                        {
                            _dr[33] = "";
                        }
                        else
                        {
                            _dr[33] = _strCardNumber;
                        }

                        _dr[34] = 0;

                        _dr[35] = _encoding.GetString(_byteAry, 359, 2);

                        if (_strCardNumber.Trim().Length == 15)
                        {
                            if (_strZipcode.Trim().Length == 5)
                            {
                                _dr[36] = _strCardNumber.Trim() + "0" + "0" + _strZipcode;
                            }
                            else
                            {
                                _dr[36] = _strCardNumber.Trim() + "0" + _strZipcode;
                            }
                        }
                        else
                        {
                            if (_strZipcode.Trim().Length == 5)
                            {
                                _dr[36] = _strCardNumber + "0" + _strZipcode;
                            }
                            else
                            {
                                _dr[36] = _strCardNumber + _strZipcode;
                            }

                        }

                        //2011-10-05 태희철 추가
                        _strNewAdd = _encoding.GetString(_byteAry, 362, 36);

                        _dr[38] = _strNewAdd;
                        _dr[39] = _strOwner_only;
                        _dr[40] = _encoding.GetString(_byteAry, 399, 1);

                        strAreaCode = _encoding.GetString(_byteAry, 400, 2);
                        _dr[41] = ConvertAreaCode(strAreaCode);
                        _dr[42] = _encoding.GetString(_byteAry, 402, 3);
                        _dr[43] = _encoding.GetString(_byteAry, 405, 20);

                        _dr[44] = _dr[23];

                        _dr[47] = _encoding.GetString(_byteAry, 425, 1);

                        _dr[49] = _encoding.GetString(_byteAry, 426, 1);

                        _dr[51] = _encoding.GetString(_byteAry, 427, 4);

                        //신분증입력 설정
                        //00226 : 국민긴급 동의 테블릿 적용 change_add 설정
                        if (_strCode2 == "P" && strEmerDong == "65")
                        {
                            _dr[50] = "1";
                        }
                        else if (strCard_type_detail.Substring(0, 4) != "0022" && _strOwner_only == "1")
                        {
                            _dr[50] = "1";
                        }

                        if (_encoding.GetString(_byteAry, 425, 1) == "1" && _dr[49].ToString() == "1")
                        {
                            _dr[48] = "친절배송요청 / 본인 신분증 확인 및 동의서 기재 확인";
                        }
                        else if (_encoding.GetString(_byteAry, 425, 1) == "1")
                        {
                            _dr[48] = "친절배송요청";
                        }

                        // 본인배송여부
                        if (((_strFamilyCode == "0" || _strFamilyCode == "3") && _strOwner_one == "0") || (_strFamilyCode == "1" && _strFamily_one == "1"))
                        {
                            //동의서는 제외
                            if ((_strCode2 == "A" || _strCode2 == "D" || _strCode2 == "Y" || _strCode2 == "S")
                                || (_strCode2 == "P" && strEmerDong == "65")
                                )
                            {
                                ;
                            }
                            else if (_encoding.GetString(_byteAry, 425, 1) == "1")
                            {
                                _dr[48] = _dr[48].ToString();
                            }
                            else
                            {
                                _dr[48] = "";
                            }
                        }
                        //2019.04.19 KB증권이며 제3자수령가능 미동의
                        else if (strOther_chk == "01" && _strOwner_one == "0")
                        {
                            if (_encoding.GetString(_byteAry, 425, 1) == "1")
                            {
                                _dr[48] = _dr[48].ToString();
                            }
                            else
                            {
                                _dr[48] = "";
                            }
                        }

                        //2011-10-04 태희철 수정 제휴코드 4-5byte 수정
                        switch (_encoding.GetString(_byteAry, 307, 5))
                        {
                            case "02060":
                            case "02062":
                            case "02063":
                            case "02064":
                                _dr[37] = "고운맘";
                                break;
                            default:
                                break;
                        }

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
                            if (_strZipcode.Length == 5)
                            {
                                _drs = _dsetZipcode_new.Tables[0].Select("zipcode_new = " + _strZipcode);
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
                    //카족카드가 1매 이상의 경우   
                    else
                    {
                        _iAddCount++;

                        _strFamilyNames = _dtable.Rows[_iSeq - 2]["family_name"].ToString().Trim();
                        _strFamilySSNs = _dtable.Rows[_iSeq - 2]["family_SSN"].ToString().Trim();
                        _strFamilyCustomerNo = _dtable.Rows[_iSeq - 2]["family_customer_no"].ToString().Trim();
                        _strCardNumberNo = _dtable.Rows[_iSeq - 2]["card_number"].ToString().Trim();

                        if (_strFamilyNames.Length > 0) _strFamilyNames += "|";
                        if (_strFamilySSNs.Length > 0) _strFamilySSNs += "|";
                        if (_strFamilyCustomerNo.Length > 0) _strFamilyCustomerNo += "|";
                        if (_strCardNumberNo.Length > 0) _strCardNumberNo += "|";

                        if (_strFamilyCode.Trim() == "" || _strFamilyCode == "0" || _strFamilyCode == "1")
                        {	//본인을 포함한 카드이거나 가족카드 1건만 발급된 경우
                            _dtable.Rows[_iSeq - 2]["family_name"] = _strFamilyNames + _encoding.GetString(_byteAry, 47, 20);
                            _dtable.Rows[_iSeq - 2]["family_SSN"] = _strFamilySSNs + _encoding.GetString(_byteAry, 18, 13);
                            _dtable.Rows[_iSeq - 2]["family_customer_no"] = _strFamilyCustomerNo + _encoding.GetString(_byteAry, 31, 16);
                            _dtable.Rows[_iSeq - 2]["card_add_count"] = _iAddCount;
                        }
                        else
                        {	//그 외는 고객과 가족위치를 변경한다.
                            _dtable.Rows[_iSeq - 2]["family_name"] = _strFamilyNames + _encoding.GetString(_byteAry, 47, 20);
                            _dtable.Rows[_iSeq - 2]["family_SSN"] = _strFamilySSNs + _encoding.GetString(_byteAry, 18, 13);
                            _dtable.Rows[_iSeq - 2]["family_customer_no"] = _strFamilyCustomerNo + _encoding.GetString(_byteAry, 31, 16);
                            _dtable.Rows[_iSeq - 2]["card_add_count"] = _iAddCount;
                        }
                    }
                }

                //변환에 성공했다면
                if (_iErrorCount < 1)
                {
                    _swError.Close();
                    _sr.Close();
                    _dtable.WriteXml(xmlPath);
                    //_fi = new FileInfo(path + ".Error");
                    //_fi.Delete();
                    _strReturn = string.Format("{0}건의 데이터 변환 성공", _iSeq - 1);
                }
                else
                    _strReturn = string.Format("{0}건 변환, 우편번호 미등록 {1}건 실패", _iSeq - _iErrorCount - 1, _iErrorCount);

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

        private static string ConvertAreaCode(string AreaCode)
        {
            string returnName = "";

            switch (AreaCode)
            {
                case "11": returnName = "서울"; break;
                case "21": returnName = "부산"; break;
                case "22": returnName = "대구"; break;
                case "23": returnName = "인천"; break;
                case "24": returnName = "광주"; break;
                case "25": returnName = "대전"; break;
                case "26": returnName = "울산"; break;
                case "31": returnName = "경기"; break;
                case "32": returnName = "강원"; break;
                case "33": returnName = "충북"; break;
                case "34": returnName = "충남"; break;
                case "35": returnName = "전북"; break;
                case "36": returnName = "전남"; break;
                case "37": returnName = "경북"; break;
                case "38": returnName = "경남"; break;
                case "39": returnName = "제주"; break;
                default:
                    break;
            }

            return returnName;
        }


        //등록 자료 생성 _ 반송등록
        public static string ConvertRegister_Ban(string path, string xmlZipcodePath, string xmlZipcodeAreaPath, string xmlBankPath, string xmlPath)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            //FileInfo _fi = null;
            StreamReader _sr = null;																					//파일 읽기 스트림
            StreamWriter _swError = null;																			//파일 쓰기 스트림
            DataSet _dsetZipcode = null, _dsetZipcdeArea = null, _dsetBank = null;											//우편번호 관련 DataSet
            DataTable _dtable = null;																					//마스터 저장 테이블
            DataRow _dr = null;
            DataRow[] _drs = null;
            byte[] _byteAry = null;
            string _strReturn = "";
            string _strLine = "";
            string _strBankID = "", _strZipcode = "", _strAddress = "", _strAreaType = "", _strAreaGroup = "", _strBranch = "", _strBankToBankID = "";
            int _iSeq = 1, _iErrorCount = 0;
            //int _iDiffLength = 0, _iLength = 237;
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
                _dtable.Columns.Add("client_insert_type");               // dr[5] 발급구분
                _dtable.Columns.Add("client_send_date");                 // 발급일자
                _dtable.Columns.Add("card_cooperation_code");            // 제휴코드
                _dtable.Columns.Add("card_product_code");                // 카드브랜드구분 영문자
                _dtable.Columns.Add("card_number");                      // 카드번호
                _dtable.Columns.Add("customer_name");                    // dr[10] 고객명
                _dtable.Columns.Add("card_zipcode");                     // 지점 우편번호
                _dtable.Columns.Add("card_address");                     // 지점 주소
                _dtable.Columns.Add("card_bank_ID");                     // 지점코드
                _dtable.Columns.Add("client_enterprise_code");           // 배송업체코드
                _dtable.Columns.Add("client_number");                    // dr[15]
                //_dtable.Columns.Add("card_barcode_new");                 // dr[16]


                //우편번호 관련 정보 DataSet에 담기
                _dsetBank = new DataSet();
                _dsetZipcode = new DataSet();
                _dsetZipcdeArea = new DataSet();
                _dsetZipcode.ReadXml(xmlZipcodePath);
                _dsetZipcode.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcode.Tables[0].Columns["zipcode"] };
                _dsetZipcdeArea.ReadXml(xmlZipcodeAreaPath);
                _dsetZipcdeArea.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcdeArea.Tables[0].Columns["zipcode"] };
                _dsetBank.ReadXml(xmlBankPath);

                //파일 읽기 Stream과 오류시 저장할 쓰기 Stream 생성
                _sr = new System.IO.StreamReader(path, _encoding);
                _swError = new System.IO.StreamWriter(path + ".Error", false, _encoding);

                while ((_strLine = _sr.ReadLine()) != null)
                {
                    //인코딩, byte 배열로 담기
                    _byteAry = _encoding.GetBytes(_strLine);

                    _dr = _dtable.NewRow();
                    _dr[0] = _iSeq;

                    _strBankID = _encoding.GetString(_byteAry, 251, 4);
                    _dr[5] = _encoding.GetString(_byteAry, 0, 1);
                    _dr[6] = _encoding.GetString(_byteAry, 1, 8);
                    //2012.07.13 태희철 수정
                    //_dr[7] = _encoding.GetString(_byteAry, 9, 4);
                    _dr[7] = _encoding.GetString(_byteAry, 9, 5);
                    _dr[8] = _encoding.GetString(_byteAry, 14, 1);
                    _dr[9] = _encoding.GetString(_byteAry, 15, 16);
                    _dr[10] = _encoding.GetString(_byteAry, 31, 14);
                    _dr[11] = _encoding.GetString(_byteAry, 45, 6);
                    _dr[12] = _encoding.GetString(_byteAry, 51, 200);
                    _dr[13] = _encoding.GetString(_byteAry, 251, 4);
                    _dr[14] = _encoding.GetString(_byteAry, 255, 5);
                    _dr[15] = _encoding.GetString(_byteAry, 260, 10);
                    //_dr[16] = _dr[9] + _strZipcode;


                    if (_strBankID != "")
                    {
                        //지역 분류 선택
                        _drs = _dsetBank.Tables[0].Select("bank_ID='" + _strBankID + "'");
                        if (_drs.Length < 1)
                        {
                            _strZipcode = "";
                            _strAddress = "";
                        }
                        else
                        {
                            _strZipcode = _drs[0]["bank_zipcode"].ToString();
                            _strAddress = _drs[0]["bank_address"].ToString();
                            _strBankToBankID = _drs[0]["bank_to_bank_ID"].ToString();
                        }
                        //_dr[17] = _strZipcode;
                        //_dr[18] = _strAddress;

                        if (_strZipcode != "")
                        {
                            //지역 분류 선택
                            _drs = _dsetZipcdeArea.Tables[0].Select("zipcode=" + _strZipcode);
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
                            _drs = _dsetZipcode.Tables[0].Select("zipcode=" + _strZipcode);
                            if (_drs.Length > 0)
                            {
                                _strAreaType = _drs[0]["area_type_code"].ToString();
                            }
                            else
                            {
                                _strAreaType = "";
                            }

                            //지점과 배송지점이 다르다면 본사로 등록시킨다.
                            /*
                            if (_strBankToBankID != _strBankID)
                            {
                                _strBranch = "000";
                                _strAreaGroup = "000";
                                _strAreaType = "A";
                                _dr["card_bank_ID"] = _strBankToBankID;
                                _dr["customer_real_SSN"] = _strBankID;
                            }
                            */

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
                    else
                    {
                        _swError.WriteLine(_strLine);
                        _iErrorCount++;
                    }
                }

                //변환에 성공했다면
                if (_iErrorCount < 1)
                {
                    _swError.Close();
                    _sr.Close();
                    _dtable.WriteXml(xmlPath);
                    //_fi = new FileInfo(path + ".Error");
                    //_fi.Delete();
                    _strReturn = string.Format("{0}건의 데이터 변환 성공", _iSeq - 1);
                }
                else
                {
                    _strReturn = string.Format("{0}건 변환, 영업점 조회 {1}건 실패", _iSeq - _iErrorCount - 1, _iErrorCount);

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
                    _strReturn = ConvertReceiveType1(dtable, fileName);
                    break;
                case 2:
                    _strReturn = ConvertReceiveType2(dtable, fileName);
                    break;
                case 3:
                    _strReturn = ConvertReceiveType3(dtable, fileName);
                    break;
                case 4:
                    _strReturn = ConvertReceiveType4(dtable, fileName);
                    break;
                default:
                    _strReturn = "종료";
                    break;
            }
            return _strReturn;
        }

        //일일마감자료
        public static string ConvertResultDay(System.Data.DataTable dtable, string fileName)
        {
            return "일일마감자료 다운은 사용하실 수 없습니다.";
        }

        //마감데이터_NEW
        private static string ConvertReceiveType1(DataTable dtable, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩
            StreamWriter _sw01 = null, _sw02 = null, _sw00 = null, _sw03 = null, _sw04 = null, _sw05 = null, _sw06 = null;  //파일 쓰기 스트림
            StreamWriter _sw11 = null, _sw12 = null, _sw10 = null, _sw13 = null, _sw14 = null, _sw15 = null, _sw16 = null;  //파일 쓰기 스트림
            string _strLine = "";
            string _strReturn = "", _strStatus = "";
            int i = -1;
            string _strCardNumber = "";
            string _strFamilyNo = "", _strFamilyCheck = "", strCheck_num = "";
            int _iAddCount = 0, icnt_01 = 0;
            string[] _strArFamilyNo = null, strCheck_num_array = null;
            //[재방관리]
            //2012-04-13 태희철 수정 
            string strChange_status = "", _strCode2 = "", strCard_type_detail = "";
            int i_imgCnt = 0;
            try
            {
                _sw01 = new StreamWriter(fileName + "국민통합마감.txt", true, _encoding);
                //_sw02 = new StreamWriter(fileName + ".02", true, _encoding);
                _sw00 = new StreamWriter(fileName + "국민통합마감_미배송_분실.txt", true, _encoding);
                _sw03 = new StreamWriter(fileName + "국민통합마감_특송_일반_미처리.txt", true, _encoding);
                _sw04 = new StreamWriter(fileName + "국민통합마감_특송_동의_미처리.txt", true, _encoding);
                _sw05 = new StreamWriter(fileName + "국민통합마감_특송_반송.txt", true, _encoding);
                _sw06 = new StreamWriter(fileName + "국민통합마감_이미지파일.txt", true, _encoding);

                _sw11 = new StreamWriter(fileName + "국민발급대행_통합마감.txt", true, _encoding);
                //_sw12 = new StreamWriter(fileName + ".02", true, _encoding);
                _sw10 = new StreamWriter(fileName + "국민발급대행_통합마감_미배송_분실.txt", true, _encoding);
                _sw13 = new StreamWriter(fileName + "국민발급대행_통합마감_특송_일반_미처리.txt", true, _encoding);
                _sw15 = new StreamWriter(fileName + "국민발급대행_통합마감_특송_반송.txt", true, _encoding);
                _sw16 = new StreamWriter(fileName + "국민발급대행_통합마감_이미지파일.txt", true, _encoding);

                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    if (dtable.Rows[i]["number"].ToString() != "")
                    {
                        strCheck_num_array = dtable.Rows[i]["number"].ToString().Split('(');
                    }

                    if (dtable.Rows[i]["card_type_detail"].ToString().Substring(0, 4) == "0024")
                    {
                        ;
                    }
                    else
                    {
                        string strChange_address = dtable.Rows[i]["change_address"].ToString() + dtable.Rows[i]["change_address_detail"].ToString();
                        _iAddCount = int.Parse(dtable.Rows[i]["card_add_count"].ToString());
                        _strStatus = dtable.Rows[i]["card_delivery_status"].ToString();
                        _strFamilyNo = dtable.Rows[i]["family_customer_no"].ToString();
                        strCard_type_detail = dtable.Rows[i]["card_type_detail"].ToString();

                        DateTime CardInDate = DateTime.Parse(dtable.Rows[i]["card_in_date"].ToString());
                        DateTime dtDong_date = DateTime.Parse("2019-06-30");

                        //국민 전송 데이터 구분 
                        //( 배송 = 11, 반송 = 12, 분실 =13, 배송 -> 반송 = 14, 반송 -> 배송 = 15, 
                        //  배송 -> 분실 = 16, 반송 -> 분실 = 17, 선반납 = 18, 선반납 배송외지역 = 19)
                        strChange_status = dtable.Rows[i]["change_delivery_status"].ToString();


                        //법인구분 : E,F,L,M
                        _strCode2 = dtable.Rows[i]["client_register_type"].ToString();

                        //데이터생성 시작
                        _strLine = "K";
                        _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["client_send_date"].ToString()), 8, true, ' ');
                        _strLine += GetStringAsLength(dtable.Rows[i]["client_express_code"].ToString(), 2, true, ' ');
                        _strLine += GetStringAsLength(dtable.Rows[i]["card_client_no_1"].ToString(), 6, true, ' ');
                        _strCardNumber = GetStringAsLength(dtable.Rows[i]["card_number"].ToString(), 16, true, ' ');
                        _strLine += "{0}";
                        _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["client_register_date"].ToString()), 8, true, ' ');

                        //배송건 중 기존배송전송완료 데이터는 데이터 생성 하지 않는다.
                        //2015.04.06 태희철 이미지수령증 파일만 생성
                        if (_strStatus == "1" && (strChange_status == "11" || strChange_status == "15"))
                        {
                            if (strCard_type_detail.Substring(0, 5) == "00226")
                            {
                                _sw06.WriteLine(GetStringAsLength("K" + GetSendCode(dtable.Rows[i]["client_register_type"].ToString(), dtable.Rows[i]["client_insert_type"].ToString()) + RemoveDash(dtable.Rows[i]["card_number"].ToString()), 20, true, ' ') + "    일마감");

                                if (_iAddCount > 0)
                                {
                                    _strArFamilyNo = _strFamilyNo.Split(new char[] { '|' });

                                    for (int j = 0; j < _strArFamilyNo.Length; j++)
                                    {
                                        _strFamilyCheck = _strArFamilyNo[j];

                                        //모체와 가족 카드번호가 다를 경우만 리스트 생성
                                        if (_strCardNumber != _strFamilyCheck)
                                        {
                                            _sw06.WriteLine(GetStringAsLength("K" + GetSendCode(dtable.Rows[i]["client_register_type"].ToString(), dtable.Rows[i]["client_insert_type"].ToString()) + _strFamilyCheck, 20, true, ' ') + "    일마감(가족)    " + dtable.Rows[i]["card_number"].ToString());
                                        }
                                    }
                                }
                            }
                            else if (strCard_type_detail.Substring(0, 4) == "0022" && dtDong_date < CardInDate)
                            {
                                _sw06.WriteLine(GetStringAsLength("K" + GetSendCode(dtable.Rows[i]["client_register_type"].ToString(), dtable.Rows[i]["client_insert_type"].ToString()) + RemoveDash(dtable.Rows[i]["card_number"].ToString()), 20, true, ' ') + "    일마감");

                                if (_iAddCount > 0)
                                {
                                    _strArFamilyNo = _strFamilyNo.Split(new char[] { '|' });

                                    for (int j = 0; j < _strArFamilyNo.Length; j++)
                                    {
                                        _strFamilyCheck = _strArFamilyNo[j];

                                        //모체와 가족 카드번호가 다를 경우만 리스트 생성
                                        if (_strCardNumber != _strFamilyCheck)
                                        {
                                            _sw06.WriteLine(GetStringAsLength("K" + GetSendCode(dtable.Rows[i]["client_register_type"].ToString(), dtable.Rows[i]["client_insert_type"].ToString()) + _strFamilyCheck, 20, true, ' ') + "    일마감(가족)    " + dtable.Rows[i]["card_number"].ToString());
                                        }
                                    }
                                }
                            }
                            else if (dtable.Rows[i]["card_kind"].ToString().ToLower() != "d")
                            {
                                _sw06.WriteLine(GetStringAsLength("K" + GetSendCode(dtable.Rows[i]["client_register_type"].ToString(), dtable.Rows[i]["client_insert_type"].ToString()) + RemoveDash(dtable.Rows[i]["card_number"].ToString()), 20, true, ' ') + "    일마감");

                                if (_iAddCount > 0)
                                {
                                    _strArFamilyNo = _strFamilyNo.Split(new char[] { '|' });

                                    for (int j = 0; j < _strArFamilyNo.Length; j++)
                                    {
                                        _strFamilyCheck = _strArFamilyNo[j];

                                        //모체와 가족 카드번호가 다를 경우만 리스트 생성
                                        if (_strCardNumber != _strFamilyCheck)
                                        {
                                            _sw06.WriteLine(GetStringAsLength("K" + GetSendCode(dtable.Rows[i]["client_register_type"].ToString(), dtable.Rows[i]["client_insert_type"].ToString()) + _strFamilyCheck, 20, true, ' ') + "    일마감(가족)    " + dtable.Rows[i]["card_number"].ToString());
                                        }
                                    }
                                }
                            }
                        }
                        //반송건 중 기존반송전송완료 데이터는 데이터 생성 하지 않는다.
                        else if ((_strStatus == "2" || _strStatus == "3") && (strChange_status == "12" || strChange_status == "14" || strChange_status == "18" || strChange_status == "19"))
                        {
                            ;
                        }
                        else if (_strStatus == "6" && (strChange_status == "13" || strChange_status == "16" || strChange_status == "17"))
                        {
                            ;
                        }
                        else
                        {
                            icnt_01++;
                            if (_strStatus == "1" || _strStatus == "7")
                            {
                                //징구구분
                                if (_strStatus == "7")
                                {
                                    if (dtable.Rows[i]["change_type"].ToString().Trim().Length > 0)
                                    {
                                        _strLine += GetStringAsLength("B", 1, true, ' ');
                                    }
                                    else
                                    {
                                        _strLine += GetStringAsLength("A", 1, true, ' ');
                                    }

                                    _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["card_delivery_date"].ToString()), 8, true, ' ');
                                    _strLine += GetStringAsLength(dtable.Rows[i]["card_result_status"].ToString(), 2, true, ' ');
                                    _strLine += GetStringAsLength("", 14, true, ' ');
                                    _strLine += GetStringAsLength("", 13, true, ' ');
                                    _strLine += GetStringAsLength("2", 1, true, ' ');
                                }
                                else
                                {
                                    //반송->배송
                                    if (strChange_status == "12" || strChange_status == "14" || strChange_status == "18" || strChange_status == "19")
                                    {
                                        _strLine += GetStringAsLength("5", 1, true, ' ');
                                    }
                                    else
                                    {
                                        _strLine += GetStringAsLength(_strStatus, 1, true, ' ');
                                    }

                                    _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["card_delivery_date"].ToString()), 8, true, ' ');

                                    //법인카드의 경우 코드 12 (기업담당자) 추가
                                    //2018.03.29 태희철 국민카드 확인 기업담당자 제외 기존과 동일 본인 : 01
                                    _strLine += GetStringAsLength(dtable.Rows[i]["receiver_code_change"].ToString().Replace("xx", "  "), 2, true, ' ');
                                    _strLine += GetStringAsLength(dtable.Rows[i]["receiver_name"].ToString(), 14, true, ' ');
                                    _strLine += GetStringAsLength(dtable.Rows[i]["receiver_SSN"].ToString().Replace("x", "*"), 13, true, ' ');

                                    _strLine += GetStringAsLength("1", 1, true, ' ');
                                }

                                //수령지구분값(재청구지구분)
                                //0-변경없음, 1-자택, 2-직장, 3-제3청구지
                                if (dtable.Rows[i]["change_type"].ToString().Trim().Length > 0)
                                {
                                    _strLine += GetStringAsLength(dtable.Rows[i]["change_place_type"].ToString(), 1, true, ' ');
                                    _strLine += GetStringAsLength(dtable.Rows[i]["change_zipcode"].ToString(), 6, true, ' ');
                                    if (strChange_address.Length > 50)
                                    {
                                        _strLine += GetStringAsLength(strChange_address.Substring(0, 50), 100, true, ' ');
                                    }
                                    else
                                    {
                                        _strLine += GetStringAsLength(strChange_address, 100, true, ' ');
                                    }
                                    _strLine += GetStringAsLength(dtable.Rows[i]["change_home_tel"].ToString(), 15, true, ' ');
                                    _strLine += GetStringAsLength(dtable.Rows[i]["change_mobile_tel"].ToString(), 15, true, ' ');

                                }
                                else
                                {
                                    _strLine += GetStringAsLength("0", 1, true, ' ');
                                    _strLine += GetStringAsLength("", 6, true, ' ');
                                    _strLine += GetStringAsLength("", 100, true, ' ');
                                    _strLine += GetStringAsLength("", 15, true, ' ');
                                    _strLine += GetStringAsLength("", 15, true, ' ');
                                }

                                //2019.06.21 전체 동의서 태블릿 적용일 : 2019.06.24 인수
                                if (strCard_type_detail.Substring(0, 5) == "00226")
                                {
                                    _strLine += GetStringAsLength("1", 1, true, ' ');

                                    if (_strStatus != "7")
                                    {
                                        i_imgCnt++;

                                        _strLine += GetStringAsLength("K" + GetSendCode(dtable.Rows[i]["client_register_type"].ToString(), dtable.Rows[i]["client_insert_type"].ToString()), 2, true, ' ');
                                        //_strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["card_number"].ToString()), 16, true, ' ');
                                        _strLine += "{1}";
                                        _strLine += GetStringAsLength(".tif", 4, true, ' ');

                                        //이미지파일명 저장
                                        _sw06.WriteLine(GetStringAsLength("K" + GetSendCode(dtable.Rows[i]["client_register_type"].ToString(), dtable.Rows[i]["client_insert_type"].ToString()) + RemoveDash(dtable.Rows[i]["card_number"].ToString()), 20, true, ' '));
                                    }
                                    else
                                    {
                                        _strLine += GetStringAsLength("", 22, true, ' ');
                                    }
                                }
                                else if (strCard_type_detail.Substring(0, 4) == "0022" && dtDong_date < CardInDate)
                                {
                                    _strLine += GetStringAsLength("1", 1, true, ' ');

                                    if (_strStatus != "7")
                                    {
                                        i_imgCnt++;

                                        _strLine += GetStringAsLength("K" + GetSendCode(dtable.Rows[i]["client_register_type"].ToString(), dtable.Rows[i]["client_insert_type"].ToString()), 2, true, ' ');
                                        //_strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["card_number"].ToString()), 16, true, ' ');
                                        _strLine += "{1}";
                                        _strLine += GetStringAsLength(".tif", 4, true, ' ');

                                        //이미지파일명 저장
                                        _sw06.WriteLine(GetStringAsLength("K" + GetSendCode(dtable.Rows[i]["client_register_type"].ToString(), dtable.Rows[i]["client_insert_type"].ToString()) + RemoveDash(dtable.Rows[i]["card_number"].ToString()), 20, true, ' '));
                                    }
                                    else
                                    {
                                        _strLine += GetStringAsLength("", 22, true, ' ');
                                    }
                                }
                                else if (dtable.Rows[i]["card_kind"].ToString().ToLower() != "d")
                                {
                                    _strLine += GetStringAsLength("1", 1, true, ' ');

                                    if (_strStatus != "7")
                                    {
                                        i_imgCnt++;

                                        _strLine += GetStringAsLength("K" + GetSendCode(dtable.Rows[i]["client_register_type"].ToString(), dtable.Rows[i]["client_insert_type"].ToString()), 2, true, ' ');
                                        //_strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["card_number"].ToString()), 16, true, ' ');
                                        _strLine += "{1}";
                                        _strLine += GetStringAsLength(".tif", 4, true, ' ');

                                        //이미지파일명 저장
                                        //KB증권구분
                                        if (strCard_type_detail.Substring(0, 5) == "00213" || strCard_type_detail.Substring(0, 5) == "00214" || strCard_type_detail.Substring(0, 5) == "00215" || strCard_type_detail.Substring(0, 5) == "00216")
                                        {
                                            _sw16.WriteLine(GetStringAsLength("K" + GetSendCode(dtable.Rows[i]["client_register_type"].ToString(), dtable.Rows[i]["client_insert_type"].ToString()) + RemoveDash(dtable.Rows[i]["card_number"].ToString()), 20, true, ' '));
                                        }
                                        else
                                        {
                                            _sw06.WriteLine(GetStringAsLength("K" + GetSendCode(dtable.Rows[i]["client_register_type"].ToString(), dtable.Rows[i]["client_insert_type"].ToString()) + RemoveDash(dtable.Rows[i]["card_number"].ToString()), 20, true, ' '));
                                        }

                                        //_sw06.WriteLine(GetStringAsLength("K" + GetSendCode(dtable.Rows[i]["client_register_type"].ToString(), dtable.Rows[i]["client_insert_type"].ToString()) + RemoveDash(dtable.Rows[i]["card_number"].ToString()), 20, true, ' '));
                                    }
                                    else
                                    {
                                        _strLine += GetStringAsLength("", 22, true, ' ');
                                    }

                                }
                                else
                                {
                                    _strLine += GetStringAsLength(dtable.Rows[i]["delivery_is_pda_register"].ToString(), 1, true, ' ');

                                    _strLine += GetStringAsLength("", 22, true, ' ');
                                }
                            }
                            // 반송
                            else if (_strStatus == "2" || _strStatus == "3")
                            {
                                //선반납 카드회수
                                if (dtable.Rows[i]["delivery_return_reason_last"].ToString() == "30")
                                {
                                    _strLine += GetStringAsLength("8", 1, true, ' ');
                                    _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["delivery_return_date_last"].ToString()), 8, true, ' ');
                                    _strLine += GetStringAsLength("22", 2, true, ' ');
                                }
                                //선반납 배송외지역
                                else if (dtable.Rows[i]["delivery_return_reason_last"].ToString() == "39")
                                {
                                    _strLine += GetStringAsLength("9", 1, true, ' ');
                                    _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["delivery_return_date_last"].ToString()), 8, true, ' ');
                                    _strLine += GetStringAsLength("22", 2, true, ' ');
                                }
                                else
                                {
                                    //기존 배송->반송
                                    if (strChange_status == "11" || strChange_status == "15")
                                    {
                                        _strLine += GetStringAsLength("4", 1, true, ' ');
                                    }
                                    else
                                    {
                                        _strLine += GetStringAsLength("2", 1, true, ' ');
                                    }

                                    _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["delivery_return_date_last"].ToString()), 8, true, ' ');
                                    _strLine += GetStringAsLength(dtable.Rows[i]["return_code_change"].ToString(), 2, true, ' ');
                                }

                                _strLine += GetStringAsLength("", 14, true, ' ');
                                _strLine += GetStringAsLength("", 13, true, ' ');
                                //징구구분
                                _strLine += GetStringAsLength("1", 1, true, ' ');
                                //수령지구분값(재청구지구분)
                                //0-변경없음, 1-자택, 2-직장, 3-제3청구지
                                if (dtable.Rows[i]["change_type"].ToString() == "")
                                {
                                    _strLine += GetStringAsLength("0", 1, true, ' ');
                                    _strLine += GetStringAsLength("", 6, true, ' ');
                                    _strLine += GetStringAsLength("", 100, true, ' ');
                                    _strLine += GetStringAsLength("", 15, true, ' ');
                                    _strLine += GetStringAsLength("", 15, true, ' ');
                                }
                                else
                                {
                                    _strLine += GetStringAsLength(dtable.Rows[i]["change_place_type"].ToString(), 1, true, ' ');
                                    _strLine += GetStringAsLength(dtable.Rows[i]["change_zipcode"].ToString(), 6, true, ' ');
                                    if (strChange_address.Length > 50)
                                    {
                                        _strLine += GetStringAsLength(strChange_address.Substring(0, 50), 100, true, ' ');
                                    }
                                    else
                                    {
                                        _strLine += GetStringAsLength(strChange_address, 100, true, ' ');
                                    }
                                    _strLine += GetStringAsLength(dtable.Rows[i]["change_home_tel"].ToString(), 15, true, ' ');
                                    _strLine += GetStringAsLength(dtable.Rows[i]["change_mobile_tel"].ToString(), 15, true, ' ');
                                }
                                _strLine += GetStringAsLength("", 23, true, ' ');
                            }
                            else
                            {
                                if (_strStatus == "6")
                                {
                                    // 6 : 기존 배송->분실
                                    if (strChange_status == "11" || strChange_status == "15")
                                    {
                                        _strLine += GetStringAsLength("6", 1, true, ' ');
                                    }
                                    // 7 : 기존 반송->분실
                                    else if (strChange_status == "12" || strChange_status == "14" || strChange_status == "18"
                                        || strChange_status == "19")
                                    {
                                        _strLine += GetStringAsLength("7", 1, true, ' ');
                                    }
                                    else
                                    {
                                        _strLine += GetStringAsLength("3", 1, true, ' ');
                                    }

                                    _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["delivery_result_regdate"].ToString()), 8, true, ' ');
                                    _strLine += GetStringAsLength("26", 2, true, ' ');
                                    _strLine += GetStringAsLength("", 14, true, ' ');
                                    _strLine += GetStringAsLength("", 13, true, ' ');
                                }
                                else
                                {
                                    _strLine += GetStringAsLength("", 1, true, ' ');
                                    _strLine += GetStringAsLength("", 8, true, ' ');
                                    _strLine += GetStringAsLength("", 2, true, ' ');
                                    _strLine += GetStringAsLength("", 14, true, ' ');
                                    _strLine += GetStringAsLength("", 13, true, ' ');
                                }

                                //징구구분
                                _strLine += GetStringAsLength("1", 1, true, ' ');

                                //수령지구분값(재청구지구분)
                                //0-변경없음, 1-자택, 2-직장, 3-제3청구지
                                if (dtable.Rows[i]["change_type"].ToString().Trim().Length > 0)
                                {
                                    _strLine += GetStringAsLength(dtable.Rows[i]["change_place_type"].ToString(), 1, true, ' ');
                                    _strLine += GetStringAsLength(dtable.Rows[i]["change_zipcode"].ToString(), 6, true, ' ');
                                    if (strChange_address.Length > 50)
                                    {
                                        _strLine += GetStringAsLength(strChange_address.Substring(0, 50), 100, true, ' ');
                                    }
                                    else
                                    {
                                        _strLine += GetStringAsLength(strChange_address, 100, true, ' ');
                                    }
                                    _strLine += GetStringAsLength(dtable.Rows[i]["change_home_tel"].ToString(), 15, true, ' ');
                                    _strLine += GetStringAsLength(dtable.Rows[i]["change_mobile_tel"].ToString(), 15, true, ' ');
                                }
                                else
                                {
                                    _strLine += GetStringAsLength("0", 1, true, ' ');
                                    _strLine += GetStringAsLength("", 6, true, ' ');
                                    _strLine += GetStringAsLength("", 100, true, ' ');
                                    _strLine += GetStringAsLength("", 15, true, ' ');
                                    _strLine += GetStringAsLength("", 15, true, ' ');
                                }
                                _strLine += GetStringAsLength("", 23, true, ' ');
                            }

                            //수령지구분값(재청구지구분)
                            //0-변경없음, 1-자택, 2-직장, 3-제3청구지
                            //2011-10-10 태희철 추가 새주소관련
                            _strLine += npi_file_name(dtable.Rows[i]["npi_file_name"].ToString(), dtable.Rows[i]["change_type"].ToString());

                            //신분증
                            if (_strStatus == "1")
                            {
                                if (dtable.Rows[i]["card_kind"].ToString().ToLower() == "d")
                                {
                                    switch (dtable.Rows[i]["code"].ToString())
                                    {
                                        case "01":
                                        case "03":
                                        case "04":
                                        case "06":
                                        case "10":
                                        case "11":
                                        case "12":
                                            _strLine += GetStringAsLength(dtable.Rows[i]["code"].ToString(), 2, true, ' ');
                                            _strLine += GetStringAsLength(dtable.Rows[i]["org"].ToString(), 10, true, ' ');
                                            _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["date"].ToString()), 10, true, ' ');
                                            break;
                                        case "02":
                                            _strLine += GetStringAsLength(dtable.Rows[i]["code"].ToString(), 2, true, ' ');
                                            _strLine += GetStringAsLength(RemoveDash(strCheck_num_array[1].Substring(0, 2)), 10, true, ' ');
                                            _strLine += GetStringAsLength(RemoveDash(strCheck_num_array[1].Substring(strCheck_num_array[1].IndexOf(")") + 2, strCheck_num_array[1].Length - 4)), 10, true, ' ');
                                            break;
                                        default:
                                            _strLine += GetStringAsLength("", 2, true, ' ');
                                            _strLine += GetStringAsLength("", 10, true, ' ');
                                            _strLine += GetStringAsLength("", 10, true, ' ');
                                            break;
                                    }
                                }
                                else
                                {
                                    switch (dtable.Rows[i]["code"].ToString())
                                    {
                                        case "01":
                                        case "03":
                                            _strLine += GetStringAsLength(dtable.Rows[i]["code"].ToString(), 2, true, ' ');
                                            _strLine += GetStringAsLength(dtable.Rows[i]["org"].ToString(), 10, true, ' ');
                                            _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["date"].ToString()), 10, true, ' ');
                                            break;
                                        case "02":
                                            _strLine += GetStringAsLength(dtable.Rows[i]["code"].ToString(), 2, true, ' ');
                                            _strLine += GetStringAsLength(RemoveDash(strCheck_num_array[1].Substring(0, 2)), 10, true, ' ');
                                            _strLine += GetStringAsLength(RemoveDash(strCheck_num_array[1].Substring(strCheck_num_array[1].IndexOf(")") + 2, strCheck_num_array[1].Length - 4)), 10, true, ' ');
                                            break;
                                        case "04":
                                        case "05":
                                        case "06":
                                        case "07":
                                        case "08":
                                        case "09":
                                        case "10":
                                        case "11":
                                        case "12":
                                            _strLine += GetStringAsLength(dtable.Rows[i]["code"].ToString(), 2, true, ' ');
                                            _strLine += GetStringAsLength("", 10, true, ' ');
                                            _strLine += GetStringAsLength("", 10, true, ' ');
                                            break;
                                        default:
                                            _strLine += GetStringAsLength("", 2, true, ' ');
                                            _strLine += GetStringAsLength("", 10, true, ' ');
                                            _strLine += GetStringAsLength("", 10, true, ' ');
                                            break;
                                    }
                                }
                            }
                            else
                            {
                                _strLine += GetStringAsLength("", 2, true, ' ');
                                _strLine += GetStringAsLength("", 10, true, ' ');
                                _strLine += GetStringAsLength("", 10, true, ' ');
                            }

                            //2012-04-13 태희철 수정 [재방관리]
                            //_sw01.WriteLine(Title + GetStringAsLength(string.Format(_strLine, _strCardNumber), 200, true, ' '));
                            if (_strStatus == "1" || _strStatus == "7" || _strStatus == "2" || _strStatus == "3")
                            {
                                if (strCard_type_detail.Substring(0, 5) == "00213" || strCard_type_detail.Substring(0, 5) == "00214" || strCard_type_detail.Substring(0, 5) == "00215" || strCard_type_detail.Substring(0, 5) == "00216")
                                {
                                    _sw11.WriteLine(GetStringAsLength(string.Format(_strLine, _strCardNumber, _strCardNumber), 298, true, ' '));
                                }
                                else
                                {
                                    _sw01.WriteLine(GetStringAsLength(string.Format(_strLine, _strCardNumber, _strCardNumber), 298, true, ' '));
                                }
                                //_sw01.WriteLine(GetStringAsLength(string.Format(_strLine, _strCardNumber, _strCardNumber), 298, true, ' '));
                            }
                            else
                            {
                                if (strCard_type_detail.Substring(0, 5) == "00213" || strCard_type_detail.Substring(0, 5) == "00214" || strCard_type_detail.Substring(0, 5) == "00215" || strCard_type_detail.Substring(0, 5) == "00216")
                                {
                                    _sw10.WriteLine(GetStringAsLength(string.Format(_strLine, _strCardNumber, _strCardNumber), 298, true, ' '));
                                }
                                else
                                {
                                    _sw00.WriteLine(GetStringAsLength(string.Format(_strLine, _strCardNumber, _strCardNumber), 298, true, ' '));
                                }
                                //_sw00.WriteLine(GetStringAsLength(string.Format(_strLine, _strCardNumber, _strCardNumber), 298, true, ' '));
                            }

                            if (_iAddCount > 0)
                            {
                                _strArFamilyNo = _strFamilyNo.Split(new char[] { '|' });

                                //2017.01.18 태희철 수정 가족건의 경우 가족 번호로 이미지 파일명 생성


                                for (int j = 0; j < _strArFamilyNo.Length; j++)
                                {
                                    _strFamilyCheck = _strArFamilyNo[j];
                                    string familyNo = GetStringAsLength(_strArFamilyNo[j], 16, true, ' ');

                                    if (_strCardNumber != _strFamilyCheck)
                                    {
                                        if (_strStatus == "1" || _strStatus == "7")
                                        {
                                            _sw01.WriteLine(GetStringAsLength(string.Format(_strLine, familyNo, familyNo), 298, true, ' '));
                                            //이미지파일명 저장
                                            if (_strStatus == "1")
                                            {
                                                //긴급동의서만 적용 : 2018.10.24 인수
                                                if (strCard_type_detail.Substring(0, 5) == "00226")
                                                {
                                                    //모체와 가족 카드번호가 다를 경우만 리스트 생성
                                                    if (dtable.Rows[i]["card_number"].ToString().Trim() != _strArFamilyNo[j].Trim())
                                                    {
                                                        _sw06.WriteLine(GetStringAsLength("K" + GetSendCode(dtable.Rows[i]["client_register_type"].ToString(), dtable.Rows[i]["client_insert_type"].ToString()) + _strFamilyCheck, 20, true, ' ') + "    가족    " + dtable.Rows[i]["card_number"].ToString());
                                                    }
                                                }
                                                //동의서 전체 태블릿 전환 : 2019.06.24 인수부터
                                                else if ((strCard_type_detail.Substring(0, 4) == "0022" && dtDong_date < CardInDate))
                                                {
                                                    //모체와 가족 카드번호가 다를 경우만 리스트 생성
                                                    if (dtable.Rows[i]["card_number"].ToString().Trim() != _strArFamilyNo[j].Trim())
                                                    {
                                                        _sw06.WriteLine(GetStringAsLength("K" + GetSendCode(dtable.Rows[i]["client_register_type"].ToString(), dtable.Rows[i]["client_insert_type"].ToString()) + _strFamilyCheck, 20, true, ' ') + "    가족    " + dtable.Rows[i]["card_number"].ToString());
                                                    }
                                                }
                                                else if (dtable.Rows[i]["card_kind"].ToString().ToLower() != "d")
                                                {
                                                    //모체와 가족 카드번호가 다를 경우만 리스트 생성
                                                    if (dtable.Rows[i]["card_number"].ToString().Trim() != _strArFamilyNo[j].Trim())
                                                    {
                                                        if (strCard_type_detail.Substring(0, 5) == "00213" || strCard_type_detail.Substring(0, 5) == "00214" || strCard_type_detail.Substring(0, 5) == "00215" || strCard_type_detail.Substring(0, 5) == "00216")
                                                        {
                                                            _sw16.WriteLine(GetStringAsLength("K" + GetSendCode(dtable.Rows[i]["client_register_type"].ToString(), dtable.Rows[i]["client_insert_type"].ToString()) + _strFamilyCheck, 20, true, ' ') + "    가족    " + dtable.Rows[i]["card_number"].ToString());
                                                        }
                                                        else
                                                        {
                                                            _sw06.WriteLine(GetStringAsLength("K" + GetSendCode(dtable.Rows[i]["client_register_type"].ToString(), dtable.Rows[i]["client_insert_type"].ToString()) + _strFamilyCheck, 20, true, ' ') + "    가족    " + dtable.Rows[i]["card_number"].ToString());
                                                        }
                                                        //_sw06.WriteLine(GetStringAsLength("K" + GetSendCode(dtable.Rows[i]["client_register_type"].ToString(), dtable.Rows[i]["client_insert_type"].ToString()) + _strFamilyCheck, 20, true, ' ') + "    가족    " + dtable.Rows[i]["card_number"].ToString());
                                                    }
                                                }
                                            }
                                        }
                                        else if (_strStatus == "2" || _strStatus == "3")
                                        {
                                            if (strCard_type_detail.Substring(0, 5) == "00213" || strCard_type_detail.Substring(0, 5) == "00214" || strCard_type_detail.Substring(0, 5) == "00215" || strCard_type_detail.Substring(0, 5) == "00216")
                                            {
                                                _sw11.WriteLine(GetStringAsLength(string.Format(_strLine, familyNo, familyNo), 298, true, ' '));
                                            }
                                            else
                                            {
                                                _sw01.WriteLine(GetStringAsLength(string.Format(_strLine, familyNo, familyNo), 298, true, ' '));
                                            }
                                            //_sw01.WriteLine(GetStringAsLength(string.Format(_strLine, familyNo, familyNo), 298, true, ' '));
                                        }
                                        else
                                        {
                                            if (strCard_type_detail.Substring(0, 5) == "00213" || strCard_type_detail.Substring(0, 5) == "00214" || strCard_type_detail.Substring(0, 5) == "00215" || strCard_type_detail.Substring(0, 5) == "00216")
                                            {
                                                _sw10.WriteLine(GetStringAsLength(string.Format(_strLine, familyNo, familyNo), 298, true, ' '));
                                            }
                                            else
                                            {
                                                _sw00.WriteLine(GetStringAsLength(string.Format(_strLine, familyNo, familyNo), 298, true, ' '));
                                            }
                                            //_sw00.WriteLine(GetStringAsLength(string.Format(_strLine, familyNo, familyNo), 298, true, ' '));
                                        }
                                    }
                                }
                            }

                            //미처리재방
                            if (_strStatus == "7")
                            {
                                //제휴사코드
                                _strLine = GetStringAsLength(dtable.Rows[i]["card_type_detail"].ToString(), 9, true, ' ');
                                //Total
                                _strLine += GetStringAsLength(dtable.Rows[i]["degree_arrange_number"].ToString(), 7, true, ' ');
                                //고객명
                                _strLine += GetStringAsLength(dtable.Rows[i]["customer_name"].ToString(), 22, true, ' ') + ",";
                                //카드번호
                                if (_strFamilyNo == "")
                                {
                                    _strLine += GetStringAsLength(dtable.Rows[i]["card_number"].ToString(), 53, true, ' ');
                                }
                                else
                                {
                                    _strLine += GetStringAsLength(dtable.Rows[i]["card_number"].ToString() + "(" + _strFamilyNo + ")", 53, true, ' ');
                                }
                                //우편번호
                                if (dtable.Rows[i]["card_zipcode_kind"].ToString() == "1")
                                {
                                    _strLine += GetStringAsLength(dtable.Rows[i]["card_zipcode_new"].ToString(), 8, true, ' ') + ", ";
                                }
                                else
                                {
                                    _strLine += GetStringAsLength(dtable.Rows[i]["card_zipcode"].ToString(), 8, true, ' ') + ", ";
                                }

                                //바코드
                                _strLine += GetStringAsLength(dtable.Rows[i]["card_barcode"].ToString(), 17, true, ' ');

                                //일반 OR 동의 구분
                                if (dtable.Rows[i]["card_kind"].ToString().ToLower() == "d")
                                {
                                    _sw04.WriteLine(_strLine);
                                }
                                else
                                {
                                    if (strCard_type_detail.Substring(0, 5) == "00213" || strCard_type_detail.Substring(0, 5) == "00214" || strCard_type_detail.Substring(0, 5) == "00215" || strCard_type_detail.Substring(0, 5) == "00216")
                                    {
                                        _sw13.WriteLine(_strLine);
                                    }
                                    else
                                    {
                                        _sw03.WriteLine(_strLine);
                                    }
                                    //_sw03.WriteLine(_strLine);
                                }
                            }
                            else if (_strStatus == "2" || _strStatus == "3")
                            {
                                //제휴사코드
                                _strLine = GetStringAsLength(dtable.Rows[i]["card_type_detail"].ToString(), 9, true, ' ');
                                //Total
                                _strLine += GetStringAsLength(dtable.Rows[i]["degree_arrange_number"].ToString(), 7, true, ' ');
                                //고객명
                                _strLine += GetStringAsLength(dtable.Rows[i]["customer_name"].ToString(), 22, true, ' ') + ",";
                                //카드번호
                                if (_strFamilyNo == "")
                                {
                                    _strLine += GetStringAsLength(dtable.Rows[i]["card_number"].ToString(), 53, true, ' ');
                                }
                                else
                                {
                                    _strLine += GetStringAsLength(dtable.Rows[i]["card_number"].ToString() + "(" + _strFamilyNo + ")", 53, true, ' ');
                                }
                                //우편번호
                                if (dtable.Rows[i]["card_zipcode_kind"].ToString() == "1")
                                {
                                    _strLine += GetStringAsLength(dtable.Rows[i]["card_zipcode_new"].ToString(), 8, true, ' ') + ", ";
                                }
                                else
                                {
                                    _strLine += GetStringAsLength(dtable.Rows[i]["card_zipcode"].ToString(), 8, true, ' ') + ", ";
                                }

                                //바코드
                                _strLine += GetStringAsLength(dtable.Rows[i]["card_barcode"].ToString(), 17, true, ' ');
                                //고객번호
                                _strLine += GetStringAsLength(dtable.Rows[i]["client_number"].ToString(), 17, true, ' ');

                                if (strCard_type_detail.Substring(0, 5) == "00213" || strCard_type_detail.Substring(0, 5) == "00214" || strCard_type_detail.Substring(0, 5) == "00215" || strCard_type_detail.Substring(0, 5) == "00216")
                                {
                                    _sw15.WriteLine(_strLine);
                                }
                                else
                                {
                                    _sw05.WriteLine(_strLine);
                                }
                                //_sw05.WriteLine(_strLine);
                            }
                        }
                    }
                    //for문 끝
                }
                _strReturn = string.Format("{0}건의 인계데이타 다운 완료", icnt_01);
            }
            catch (Exception)
            {
                _strReturn = string.Format("{0}번째 데이터 인계 중 오류 발생, {1}", i + 1, dtable.Rows[i]["card_barcode"].ToString());
            }
            finally
            {
                if (_sw00 != null) _sw00.Close();
                if (_sw01 != null) _sw01.Close();
                //if (_sw02 != null) _sw02.Close();
                if (_sw03 != null) _sw03.Close();
                if (_sw04 != null) _sw04.Close();
                if (_sw05 != null) _sw05.Close();
                if (_sw06 != null) _sw06.Close();

                if (_sw10 != null) _sw10.Close();
                if (_sw11 != null) _sw11.Close();
                //if (_sw02 != null) _sw02.Close();
                if (_sw13 != null) _sw13.Close();
                if (_sw15 != null) _sw15.Close();
                if (_sw16 != null) _sw16.Close();
            }
            return _strReturn;
        }

        //주민번호체크(2)
        private static string ConvertReceiveType2(DataTable dtable, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩
            StreamWriter _sw01 = null;//, _sw02 = null, _sw00 = null;																			//파일 쓰기 스트림
            string _strLine = "";
            string _strReturn = "", _strStatus = "";
            int i = -1;
            string _strCardNumber = "";
            string _strFamilyNo = "", _strFamilyCheck = "", strCheck_num = "", _strCode2 = "", strCard_type_detail = "";
            int _iAddCount = 0, icnt_01 = 0;
            string[] _strArFamilyNo = null, strCheck_num_array = null;

            try
            {
                _sw01 = new StreamWriter(fileName + ".01", true, _encoding);

                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    //법인구분 : E,F,L,M
                    _strCode2 = dtable.Rows[i]["client_register_type"].ToString();

                    if (dtable.Rows[i]["number"].ToString() != "")
                    {
                        strCheck_num_array = dtable.Rows[i]["number"].ToString().Split('(');
                    }

                    string Title = GetStringAsLength(dtable.Rows[i]["degree_code"].ToString(), 5, true, ' ') + " ";
                    string temp_delivery_result_editdate = dtable.Rows[i]["delivery_result_editdate"].ToString();
                    string strChange_address = dtable.Rows[i]["change_address"].ToString() + dtable.Rows[i]["change_address_detail"].ToString();

                    strCard_type_detail = dtable.Rows[i]["card_type_detail"].ToString();

                    DateTime CardInDate = DateTime.Parse(dtable.Rows[i]["card_in_date"].ToString());
                    DateTime dtDong_date = DateTime.Parse("2019-06-30");

                    if (temp_delivery_result_editdate != "")
                        temp_delivery_result_editdate = DateTime.Parse(temp_delivery_result_editdate).ToString("MMdd HH:mm:ss");
                    else
                        temp_delivery_result_editdate = "        ";

                    Title += temp_delivery_result_editdate + " ";

                    if (dtable.Rows[i]["jumnin_check"].ToString() == "1" && dtable.Rows[i]["card_delivery_status"].ToString() == "1")
                    {
                        _iAddCount = int.Parse(dtable.Rows[i]["card_add_count"].ToString());
                        _strStatus = dtable.Rows[i]["card_delivery_status"].ToString();
                        _strFamilyNo = dtable.Rows[i]["family_customer_no"].ToString();
                        _strLine = "K";
                        _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["client_send_date"].ToString()), 8, true, ' ');
                        _strLine += GetStringAsLength(dtable.Rows[i]["client_express_code"].ToString(), 2, true, ' ');
                        _strLine += GetStringAsLength(dtable.Rows[i]["card_client_no_1"].ToString(), 6, true, ' ');
                        _strCardNumber = GetStringAsLength(dtable.Rows[i]["card_number"].ToString(), 16, true, ' ');
                        _strLine += "{0}";
                        //_strLine += GetStringAsLength(dtable.Rows[i]["card_number"].ToString(), 16, true, ' ');
                        _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["client_register_date"].ToString()), 8, true, ' ');
                        if (_strStatus == "1" || _strStatus == "7")
                        {

                            //징구구분
                            if (_strStatus == "7")
                            {
                                if (dtable.Rows[i]["change_type"].ToString().Trim().Length > 0)
                                {
                                    _strLine += GetStringAsLength("B", 1, true, ' ');
                                }
                                else
                                {
                                    _strLine += GetStringAsLength("A", 1, true, ' ');
                                }
                                _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["card_delivery_date"].ToString()), 8, true, ' ');
                                _strLine += GetStringAsLength(dtable.Rows[i]["card_result_status"].ToString(), 2, true, ' ');
                                _strLine += GetStringAsLength("", 14, true, ' ');
                                _strLine += GetStringAsLength("", 13, true, ' ');
                                _strLine += GetStringAsLength("2", 1, true, ' ');
                            }
                            else
                            {
                                _strLine += GetStringAsLength(_strStatus, 1, true, ' ');
                                _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["card_delivery_date"].ToString()), 8, true, ' ');

                                //법인카드의 경우 코드 12 (기업담당자) 추가
                                //2018.03.29 태희철 국민카드 확인 기업담당자 제외 기존과 동일 본인 : 01
                                _strLine += GetStringAsLength(dtable.Rows[i]["receiver_code_change"].ToString().Replace("xx", "  "), 2, true, ' ');
                                _strLine += GetStringAsLength(dtable.Rows[i]["receiver_name"].ToString(), 14, true, ' ');
                                _strLine += GetStringAsLength(dtable.Rows[i]["receiver_SSN"].ToString().Replace("x", "*"), 13, true, ' ');



                                _strLine += GetStringAsLength("1", 1, true, ' ');
                            }
                            //수령지구분값(재청구지구분)
                            //0-변경없음, 1-자택, 2-직장, 3-제3청구지
                            if (dtable.Rows[i]["change_type"].ToString().Trim().Length > 0)
                            {
                                _strLine += GetStringAsLength(dtable.Rows[i]["change_place_type"].ToString(), 1, true, ' ');
                                _strLine += GetStringAsLength(dtable.Rows[i]["change_zipcode"].ToString(), 6, true, ' ');
                                if (strChange_address.Length > 50)
                                {
                                    _strLine += GetStringAsLength(strChange_address.Substring(0, 50), 100, true, ' ');
                                }
                                else
                                {
                                    _strLine += GetStringAsLength(strChange_address, 100, true, ' ');
                                }
                                _strLine += GetStringAsLength(dtable.Rows[i]["change_home_tel"].ToString(), 15, true, ' ');
                                _strLine += GetStringAsLength(dtable.Rows[i]["change_mobile_tel"].ToString(), 15, true, ' ');

                            }
                            else
                            {
                                _strLine += GetStringAsLength("0", 1, true, ' ');
                                _strLine += GetStringAsLength("", 6, true, ' ');
                                _strLine += GetStringAsLength("", 100, true, ' ');
                                _strLine += GetStringAsLength("", 15, true, ' ');
                                _strLine += GetStringAsLength("", 15, true, ' ');
                            }

                            if (strCard_type_detail.Substring(0, 5) == "00226")
                            {
                                _strLine += GetStringAsLength("1", 1, true, ' ');

                                if (_strStatus != "7")
                                {
                                    _strLine += GetStringAsLength("K" + GetSendCode(dtable.Rows[i]["client_register_type"].ToString(), dtable.Rows[i]["client_insert_type"].ToString()), 2, true, ' ');
                                    _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["card_number"].ToString()), 16, true, ' ');
                                    //_strLine += "{1}";
                                    _strLine += GetStringAsLength(".tif", 4, true, ' ');
                                }
                                else
                                {
                                    _strLine += GetStringAsLength("", 22, true, ' ');
                                }
                            }
                            else if (strCard_type_detail.Substring(0, 4) == "0022" && dtDong_date < CardInDate)
                            {
                                _strLine += GetStringAsLength("1", 1, true, ' ');

                                if (_strStatus != "7")
                                {
                                    _strLine += GetStringAsLength("K" + GetSendCode(dtable.Rows[i]["client_register_type"].ToString(), dtable.Rows[i]["client_insert_type"].ToString()), 2, true, ' ');
                                    _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["card_number"].ToString()), 16, true, ' ');
                                    //_strLine += "{1}";
                                    _strLine += GetStringAsLength(".tif", 4, true, ' ');
                                }
                                else
                                {
                                    _strLine += GetStringAsLength("", 22, true, ' ');
                                }
                            }
                            else if (dtable.Rows[i]["card_kind"].ToString().ToLower() != "d")
                            {
                                _strLine += GetStringAsLength("1", 1, true, ' ');

                                if (_strStatus != "7")
                                {
                                    _strLine += GetStringAsLength("K" + GetSendCode(dtable.Rows[i]["client_register_type"].ToString(), dtable.Rows[i]["client_insert_type"].ToString()), 2, true, ' ');
                                    _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["card_number"].ToString()), 16, true, ' ');
                                    //_strLine += "{1}";
                                    _strLine += GetStringAsLength(".tif", 4, true, ' ');
                                }
                                else
                                {
                                    _strLine += GetStringAsLength("", 22, true, ' ');
                                }
                            }
                            else
                            {
                                _strLine += GetStringAsLength(dtable.Rows[i]["delivery_is_pda_register"].ToString(), 1, true, ' ');

                                _strLine += GetStringAsLength("", 22, true, ' ');
                            }
                        }
                        else if (_strStatus == "2" || _strStatus == "3")
                        {
                            if (dtable.Rows[i]["delivery_return_reason_last"].ToString() == "30")
                            {
                                _strLine += GetStringAsLength("8", 1, true, ' ');
                                _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["delivery_return_date_last"].ToString()), 8, true, ' ');
                                _strLine += GetStringAsLength("22", 2, true, ' ');
                            }
                            else
                            {
                                _strLine += GetStringAsLength(_strStatus, 1, true, ' ');
                                _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["delivery_return_date_last"].ToString()), 8, true, ' ');
                                _strLine += GetStringAsLength(dtable.Rows[i]["return_code_change"].ToString(), 2, true, ' ');
                            }
                            _strLine += GetStringAsLength("", 14, true, ' ');
                            _strLine += GetStringAsLength("", 13, true, ' ');
                            //징구구분
                            _strLine += GetStringAsLength("1", 1, true, ' ');
                            //수령지구분값(재청구지구분)
                            //0-변경없음, 1-자택, 2-직장, 3-제3청구지
                            if (dtable.Rows[i]["change_type"].ToString() == "")
                            {
                                _strLine += GetStringAsLength("0", 1, true, ' ');
                                _strLine += GetStringAsLength("", 6, true, ' ');
                                _strLine += GetStringAsLength("", 100, true, ' ');
                                _strLine += GetStringAsLength("", 15, true, ' ');
                                _strLine += GetStringAsLength("", 15, true, ' ');
                            }
                            else
                            {
                                _strLine += GetStringAsLength(dtable.Rows[i]["change_place_type"].ToString(), 1, true, ' ');
                                _strLine += GetStringAsLength(dtable.Rows[i]["change_zipcode"].ToString(), 6, true, ' ');
                                if (strChange_address.Length > 50)
                                {
                                    _strLine += GetStringAsLength(strChange_address.Substring(0, 50), 100, true, ' ');
                                }
                                else
                                {
                                    _strLine += GetStringAsLength(strChange_address, 100, true, ' ');
                                }
                                _strLine += GetStringAsLength(dtable.Rows[i]["change_home_tel"].ToString(), 15, true, ' ');
                                _strLine += GetStringAsLength(dtable.Rows[i]["change_mobile_tel"].ToString(), 15, true, ' ');


                            }
                            _strLine += GetStringAsLength("", 23, true, ' ');
                        }
                        else
                        {
                            _strLine += GetStringAsLength("", 1, true, ' ');
                            _strLine += GetStringAsLength("", 8, true, ' ');
                            _strLine += GetStringAsLength("", 2, true, ' ');
                            _strLine += GetStringAsLength("", 14, true, ' ');
                            _strLine += GetStringAsLength("", 13, true, ' ');
                            //징구구분
                            _strLine += GetStringAsLength("", 1, true, ' ');
                            //수령지구분값(재청구지구분)
                            //0-변경없음, 1-자택, 2-직장, 3-제3청구지
                            if (dtable.Rows[i]["change_type"].ToString().Trim().Length > 0)
                            {
                                _strLine += GetStringAsLength(dtable.Rows[i]["change_place_type"].ToString(), 1, true, ' ');
                                _strLine += GetStringAsLength(dtable.Rows[i]["change_zipcode"].ToString(), 6, true, ' ');
                                if (strChange_address.Length > 50)
                                {
                                    _strLine += GetStringAsLength(strChange_address.Substring(0, 33), 100, true, ' ');
                                }
                                else
                                {
                                    _strLine += GetStringAsLength(strChange_address, 100, true, ' ');
                                }
                                _strLine += GetStringAsLength(dtable.Rows[i]["change_home_tel"].ToString(), 15, true, ' ');
                                _strLine += GetStringAsLength(dtable.Rows[i]["change_mobile_tel"].ToString(), 15, true, ' ');

                            }
                            else
                            {
                                _strLine += GetStringAsLength("0", 1, true, ' ');
                                _strLine += GetStringAsLength("", 6, true, ' ');
                                _strLine += GetStringAsLength("", 100, true, ' ');
                                _strLine += GetStringAsLength("", 15, true, ' ');
                                _strLine += GetStringAsLength("", 15, true, ' ');
                            }
                            _strLine += GetStringAsLength("", 23, true, ' ');
                        }
                        //수령지구분값(재청구지구분)
                        //0-변경없음, 1-자택, 2-직장, 3-제3청구지
                        //2011-10-10 태희철 추가 새주소관련
                        _strLine += npi_file_name(dtable.Rows[i]["npi_file_name"].ToString(), dtable.Rows[i]["change_type"].ToString());

                        //신분증
                        if (_strStatus == "1")
                        {
                            switch (dtable.Rows[i]["code"].ToString())
                            {
                                case "01":
                                case "03":
                                case "04":
                                case "05":
                                case "06":
                                case "07":
                                case "08":
                                case "09":
                                    _strLine += GetStringAsLength(dtable.Rows[i]["code"].ToString(), 2, true, ' ');
                                    _strLine += GetStringAsLength(dtable.Rows[i]["org"].ToString(), 10, true, ' ');
                                    _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["date"].ToString()), 10, true, ' ');
                                    break;
                                case "02":
                                    _strLine += GetStringAsLength(dtable.Rows[i]["code"].ToString(), 2, true, ' ');
                                    _strLine += GetStringAsLength(RemoveDash(strCheck_num_array[1].Substring(0, 2)), 10, true, ' ');
                                    _strLine += GetStringAsLength(RemoveDash(strCheck_num_array[1].Substring(strCheck_num_array[1].IndexOf(")") + 2, strCheck_num_array[1].Length - 4)), 10, true, ' ');
                                    break;
                                default:
                                    _strLine += GetStringAsLength("", 2, true, ' ');
                                    _strLine += GetStringAsLength("", 10, true, ' ');
                                    _strLine += GetStringAsLength("", 10, true, ' ');
                                    break;
                            }
                        }
                        else
                        {
                            _strLine += GetStringAsLength("", 2, true, ' ');
                            _strLine += GetStringAsLength("", 10, true, ' ');
                            _strLine += GetStringAsLength("", 10, true, ' ');
                        }

                        if (_strStatus == "1" || _strStatus == "7")
                        {
                            _sw01.WriteLine(Title + GetStringAsLength(string.Format(_strLine, _strCardNumber, _strCardNumber), 298, true, ' '));
                            if (_iAddCount > 0)
                            {
                                _strArFamilyNo = _strFamilyNo.Split(new char[] { '|' });

                                for (int j = 0; j < _strArFamilyNo.Length; j++)
                                {
                                    _strFamilyCheck = _strArFamilyNo[j];
                                    if (_strCardNumber != _strFamilyCheck)
                                    {
                                        string familyNo = GetStringAsLength(_strArFamilyNo[j], 16, true, ' ');

                                        _sw01.WriteLine(Title + GetStringAsLength(string.Format(_strLine, familyNo, familyNo), 298, true, ' '));
                                    }
                                }
                            }
                        }
                    }
                }
                _strReturn = string.Format("{0}건의 인계데이타 다운 완료", i);
            }
            catch (Exception)
            {
                _strReturn = string.Format("{0}번째 데이터 인계 중 오류 발생, {1}", i + 1, dtable.Rows[i]["card_barcode"].ToString());
            }
            finally
            {
                if (_sw01 != null) _sw01.Close();
            }
            return _strReturn;
        }

        //마감확인
        private static string ConvertReceiveType3(DataTable dtable, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩
            StreamWriter _sw01 = null, _sw02 = null, _sw00 = null;  //파일 쓰기 스트림

            string _strLine = "";
            string _strReturn = "", _strStatus = "";
            int i = -1;
            string _strCardNumber = "";
            string _strFamilyNo = "", _strFamilyCheck = "", strCheck_num = "", _strCode2 = "";
            int _iAddCount = 0, icnt_01 = 0;
            string[] _strArFamilyNo = null, strCheck_num_array = null;

            //[재방관리]
            //2012-04-13 태희철 수정 
            string strChange_status = "", strCard_type_detail = "";
            string strcard_barcode = "";
            try
            {
                _sw01 = new StreamWriter(fileName + ".01", true, _encoding);
                _sw02 = new StreamWriter(fileName + ".02", true, _encoding);
                _sw00 = new StreamWriter(fileName + ".00", true, _encoding);

                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    string Title = GetStringAsLength(dtable.Rows[i]["degree_code"].ToString(), 5, true, ' ') + " ";
                    string strChange_address = dtable.Rows[i]["change_address"].ToString() + dtable.Rows[i]["change_address_detail"].ToString();

                    strcard_barcode = dtable.Rows[i]["card_barcode"].ToString();
                    //국민 전송 데이터 구분 
                    //( 배송 = 11, 반송 = 12, 분실 =13, 배송 -> 반송 = 14, 반송 -> 배송 = 15, 
                    //  배송 -> 분실 = 16, 반송 -> 분실 = 17, 선반납 = 18)
                    strChange_status = dtable.Rows[i]["change_delivery_status"].ToString();
                    //법인구분 : E,F,L,M
                    _strCode2 = dtable.Rows[i]["client_register_type"].ToString();

                    Title += "        " + strcard_barcode + "  " + GetStringAsLength(strChange_status, 2, true, ' ') + " ";

                    _iAddCount = int.Parse(dtable.Rows[i]["card_add_count"].ToString());
                    _strStatus = dtable.Rows[i]["card_delivery_status"].ToString();
                    _strFamilyNo = dtable.Rows[i]["family_customer_no"].ToString();

                    strCard_type_detail = dtable.Rows[i]["card_type_detail"].ToString();

                    DateTime CardInDate = DateTime.Parse(dtable.Rows[i]["card_in_date"].ToString());
                    DateTime dtDong_date = DateTime.Parse("2019-06-30");

                    if (dtable.Rows[i]["number"].ToString() != "")
                    {
                        strCheck_num_array = dtable.Rows[i]["number"].ToString().Split('(');
                    }

                    //데이터생성 시작
                    _strLine = "K";
                    _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["client_send_date"].ToString()), 8, true, ' ');
                    _strLine += GetStringAsLength(dtable.Rows[i]["client_express_code"].ToString(), 2, true, ' ');
                    _strLine += GetStringAsLength(dtable.Rows[i]["card_client_no_1"].ToString(), 6, true, ' ');
                    _strCardNumber = GetStringAsLength(dtable.Rows[i]["card_number"].ToString(), 16, true, ' ');
                    _strLine += "{0}";
                    _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["client_register_date"].ToString()), 8, true, ' ');

                    //배송건 중 기존배송전송완료 데이터는 데이터 생성 하지 않는다.
                    if (_strStatus == "1" || _strStatus == "7")
                    {
                        //징구구분
                        if (_strStatus == "7")
                        {
                            if (dtable.Rows[i]["change_type"].ToString().Trim().Length > 0)
                            {
                                _strLine += GetStringAsLength("B", 1, true, ' ');
                            }
                            else
                            {
                                _strLine += GetStringAsLength("A", 1, true, ' ');
                            }

                            _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["card_delivery_date"].ToString()), 8, true, ' ');
                            _strLine += GetStringAsLength(dtable.Rows[i]["card_result_status"].ToString(), 2, true, ' ');
                            _strLine += GetStringAsLength("", 14, true, ' ');
                            _strLine += GetStringAsLength("", 13, true, ' ');
                            _strLine += GetStringAsLength("2", 1, true, ' ');
                        }
                        else
                        {
                            //반송->배송
                            if (strChange_status == "12" || strChange_status == "14" || strChange_status == "18" || strChange_status == "19")
                            {
                                _strLine += GetStringAsLength("5", 1, true, ' ');
                            }
                            else
                            {
                                _strLine += GetStringAsLength(_strStatus, 1, true, ' ');
                            }
                            //_strLine += GetStringAsLength(_strStatus, 1, true, ' ');

                            _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["card_delivery_date"].ToString()), 8, true, ' ');

                            //법인카드의 경우 코드 12 (기업담당자) 추가
                            //2018.03.29 태희철 국민카드 확인 기업담당자 제외 기존과 동일 본인 : 01
                            _strLine += GetStringAsLength(dtable.Rows[i]["receiver_code_change"].ToString().Replace("xx", "  "), 2, true, ' ');
                            _strLine += GetStringAsLength(dtable.Rows[i]["receiver_name"].ToString(), 14, true, ' ');
                            _strLine += GetStringAsLength(dtable.Rows[i]["receiver_SSN"].ToString().Replace("x", "*"), 13, true, ' ');

                            _strLine += GetStringAsLength("1", 1, true, ' ');
                        }

                        //수령지구분값(재청구지구분)
                        //0-변경없음, 1-자택, 2-직장, 3-제3청구지
                        if (dtable.Rows[i]["change_type"].ToString().Trim().Length > 0)
                        {
                            _strLine += GetStringAsLength(dtable.Rows[i]["change_place_type"].ToString(), 1, true, ' ');
                            _strLine += GetStringAsLength(dtable.Rows[i]["change_zipcode"].ToString(), 6, true, ' ');
                            if (strChange_address.Length > 50)
                            {
                                _strLine += GetStringAsLength(strChange_address.Substring(0, 50), 100, true, ' ');
                            }
                            else
                            {
                                _strLine += GetStringAsLength(strChange_address, 100, true, ' ');
                            }
                            _strLine += GetStringAsLength(dtable.Rows[i]["change_home_tel"].ToString(), 15, true, ' ');
                            _strLine += GetStringAsLength(dtable.Rows[i]["change_mobile_tel"].ToString(), 15, true, ' ');

                        }
                        else
                        {
                            _strLine += GetStringAsLength("0", 1, true, ' ');
                            _strLine += GetStringAsLength("", 6, true, ' ');
                            _strLine += GetStringAsLength("", 100, true, ' ');
                            _strLine += GetStringAsLength("", 15, true, ' ');
                            _strLine += GetStringAsLength("", 15, true, ' ');
                        }

                        if (strCard_type_detail.Substring(0, 5) == "00226")
                        {
                            _strLine += GetStringAsLength("1", 1, true, ' ');

                            if (_strStatus != "7")
                            {
                                _strLine += GetStringAsLength("K" + GetSendCode(dtable.Rows[i]["client_register_type"].ToString(), dtable.Rows[i]["client_insert_type"].ToString()), 2, true, ' ');
                                _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["card_number"].ToString()), 16, true, ' ');
                                //_strLine += "{1}";
                                _strLine += GetStringAsLength(".tif", 4, true, ' ');
                            }
                            else
                            {
                                _strLine += GetStringAsLength("", 22, true, ' ');
                            }
                        }
                        else if (strCard_type_detail.Substring(0, 4) == "0022" && dtDong_date < CardInDate)
                        {
                            _strLine += GetStringAsLength("1", 1, true, ' ');

                            if (_strStatus != "7")
                            {
                                _strLine += GetStringAsLength("K" + GetSendCode(dtable.Rows[i]["client_register_type"].ToString(), dtable.Rows[i]["client_insert_type"].ToString()), 2, true, ' ');
                                _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["card_number"].ToString()), 16, true, ' ');
                                //_strLine += "{1}";
                                _strLine += GetStringAsLength(".tif", 4, true, ' ');
                            }
                            else
                            {
                                _strLine += GetStringAsLength("", 22, true, ' ');
                            }
                        }
                        else if (dtable.Rows[i]["card_kind"].ToString().ToLower() != "d")
                        {
                            _strLine += GetStringAsLength("1", 1, true, ' ');

                            if (_strStatus != "7")
                            {
                                _strLine += GetStringAsLength("K" + GetSendCode(dtable.Rows[i]["client_register_type"].ToString(), dtable.Rows[i]["client_insert_type"].ToString()), 2, true, ' ');
                                _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["card_number"].ToString()), 16, true, ' ');
                                //_strLine += "{1}";
                                _strLine += GetStringAsLength(".tif", 4, true, ' ');
                            }
                            else
                            {
                                _strLine += GetStringAsLength("", 22, true, ' ');
                            }
                        }
                        else
                        {
                            _strLine += GetStringAsLength(dtable.Rows[i]["delivery_is_pda_register"].ToString(), 1, true, ' ');

                            _strLine += GetStringAsLength("", 22, true, ' ');
                        }
                    }
                    // 반송
                    else if (_strStatus == "2" || _strStatus == "3")
                    {
                        if (dtable.Rows[i]["delivery_return_reason_last"].ToString() == "30")
                        {
                            _strLine += GetStringAsLength("8", 1, true, ' ');
                            _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["delivery_return_date_last"].ToString()), 8, true, ' ');
                            _strLine += GetStringAsLength("22", 2, true, ' ');
                        }
                        //선반납 배송외지역
                        else if (dtable.Rows[i]["delivery_return_reason_last"].ToString() == "39")
                        {
                            _strLine += GetStringAsLength("9", 1, true, ' ');
                            _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["delivery_return_date_last"].ToString()), 8, true, ' ');
                            _strLine += GetStringAsLength("22", 2, true, ' ');
                        }
                        else
                        {
                            //기존 배송->반송
                            if (strChange_status == "11" || strChange_status == "15")
                            {
                                _strLine += GetStringAsLength("4", 1, true, ' ');
                            }
                            else
                            {
                                _strLine += GetStringAsLength("2", 1, true, ' ');
                            }
                            //_strLine += GetStringAsLength("2", 1, true, ' ');

                            _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["delivery_return_date_last"].ToString()), 8, true, ' ');
                            _strLine += GetStringAsLength(dtable.Rows[i]["return_code_change"].ToString(), 2, true, ' ');
                        }
                        _strLine += GetStringAsLength("", 14, true, ' ');
                        _strLine += GetStringAsLength("", 13, true, ' ');
                        //징구구분
                        _strLine += GetStringAsLength("1", 1, true, ' ');
                        //수령지구분값(재청구지구분)
                        //0-변경없음, 1-자택, 2-직장, 3-제3청구지
                        if (dtable.Rows[i]["change_type"].ToString() == "")
                        {
                            _strLine += GetStringAsLength("0", 1, true, ' ');
                            _strLine += GetStringAsLength("", 6, true, ' ');
                            _strLine += GetStringAsLength("", 100, true, ' ');
                            _strLine += GetStringAsLength("", 15, true, ' ');
                            _strLine += GetStringAsLength("", 15, true, ' ');
                        }
                        else
                        {
                            _strLine += GetStringAsLength(dtable.Rows[i]["change_place_type"].ToString(), 1, true, ' ');
                            _strLine += GetStringAsLength(dtable.Rows[i]["change_zipcode"].ToString(), 6, true, ' ');
                            if (strChange_address.Length > 50)
                            {
                                _strLine += GetStringAsLength(strChange_address.Substring(0, 50), 100, true, ' ');
                            }
                            else
                            {
                                _strLine += GetStringAsLength(strChange_address, 100, true, ' ');
                            }
                            _strLine += GetStringAsLength(dtable.Rows[i]["change_home_tel"].ToString(), 15, true, ' ');
                            _strLine += GetStringAsLength(dtable.Rows[i]["change_mobile_tel"].ToString(), 15, true, ' ');
                        }
                        _strLine += GetStringAsLength("", 23, true, ' ');
                    }
                    else
                    {
                        if (_strStatus == "6")
                        {
                            // 6 : 기존 배송->분실
                            if (strChange_status == "11" || strChange_status == "15")
                            {
                                _strLine += GetStringAsLength("6", 1, true, ' ');
                            }
                            // 7 : 기존 반송->분실
                            else if (strChange_status == "12" || strChange_status == "14" || strChange_status == "18" || strChange_status == "19")
                            {
                                _strLine += GetStringAsLength("7", 1, true, ' ');
                            }
                            else
                            {
                                _strLine += GetStringAsLength("3", 1, true, ' ');
                            }

                            _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["delivery_result_regdate"].ToString()), 8, true, ' ');
                            _strLine += GetStringAsLength("", 2, true, ' ');
                            _strLine += GetStringAsLength("", 14, true, ' ');
                            _strLine += GetStringAsLength("", 13, true, ' ');
                        }
                        else
                        {
                            _strLine += GetStringAsLength("", 1, true, ' ');
                            _strLine += GetStringAsLength("", 8, true, ' ');
                            _strLine += GetStringAsLength("", 2, true, ' ');
                            _strLine += GetStringAsLength("", 14, true, ' ');
                            _strLine += GetStringAsLength("", 13, true, ' ');
                        }

                        //징구구분
                        _strLine += GetStringAsLength("", 1, true, ' ');
                        //수령지구분값(재청구지구분)
                        //0-변경없음, 1-자택, 2-직장, 3-제3청구지
                        if (dtable.Rows[i]["change_type"].ToString().Trim().Length > 0)
                        {
                            _strLine += GetStringAsLength(dtable.Rows[i]["change_place_type"].ToString(), 1, true, ' ');
                            _strLine += GetStringAsLength(dtable.Rows[i]["change_zipcode"].ToString(), 6, true, ' ');
                            if (strChange_address.Length > 50)
                            {
                                _strLine += GetStringAsLength(strChange_address.Substring(0, 50), 100, true, ' ');
                            }
                            else
                            {
                                _strLine += GetStringAsLength(strChange_address, 100, true, ' ');
                            }
                            _strLine += GetStringAsLength(dtable.Rows[i]["change_home_tel"].ToString(), 15, true, ' ');
                            _strLine += GetStringAsLength(dtable.Rows[i]["change_mobile_tel"].ToString(), 15, true, ' ');
                        }
                        else
                        {
                            _strLine += GetStringAsLength("0", 1, true, ' ');
                            _strLine += GetStringAsLength("", 6, true, ' ');
                            _strLine += GetStringAsLength("", 100, true, ' ');
                            _strLine += GetStringAsLength("", 15, true, ' ');
                            _strLine += GetStringAsLength("", 15, true, ' ');
                        }
                        _strLine += GetStringAsLength("", 23, true, ' ');
                    }

                    //수령지구분값(재청구지구분)
                    //0-변경없음, 1-자택, 2-직장, 3-제3청구지
                    //2011-10-10 태희철 추가 새주소관련
                    _strLine += npi_file_name(dtable.Rows[i]["npi_file_name"].ToString(), dtable.Rows[i]["change_type"].ToString());

                    //신분증
                    if (_strStatus == "1")
                    {
                        switch (dtable.Rows[i]["code"].ToString())
                        {
                            case "01":
                            case "03":
                            case "04":
                            case "05":
                            case "06":
                            case "07":
                            case "08":
                            case "09":
                                _strLine += GetStringAsLength(dtable.Rows[i]["code"].ToString(), 2, true, ' ');
                                _strLine += GetStringAsLength(dtable.Rows[i]["org"].ToString(), 10, true, ' ');
                                _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["date"].ToString()), 10, true, ' ');
                                break;
                            case "02":
                                _strLine += GetStringAsLength(dtable.Rows[i]["code"].ToString(), 2, true, ' ');
                                _strLine += GetStringAsLength(RemoveDash(strCheck_num_array[1].Substring(0, 2)), 10, true, ' ');
                                _strLine += GetStringAsLength(RemoveDash(strCheck_num_array[1].Substring(strCheck_num_array[1].IndexOf(")") + 2, strCheck_num_array[1].Length - 4)), 10, true, ' ');
                                break;
                            default:
                                _strLine += GetStringAsLength("", 2, true, ' ');
                                _strLine += GetStringAsLength("", 10, true, ' ');
                                _strLine += GetStringAsLength("", 10, true, ' ');
                                break;
                        }
                    }
                    else
                    {
                        _strLine += GetStringAsLength("", 2, true, ' ');
                        _strLine += GetStringAsLength("", 10, true, ' ');
                        _strLine += GetStringAsLength("", 10, true, ' ');
                    }

                    //2012-04-13 태희철 수정 [재방관리]
                    //_sw01.WriteLine(Title + GetStringAsLength(string.Format(_strLine, _strCardNumber), 200, true, ' '));
                    if (_strStatus == "1" || _strStatus == "7")
                    {
                        _sw01.WriteLine(Title + GetStringAsLength(string.Format(_strLine, _strCardNumber, _strCardNumber), 298, true, ' '));
                    }
                    else if (_strStatus == "2" || _strStatus == "3")
                    {
                        _sw02.WriteLine(Title + GetStringAsLength(string.Format(_strLine, _strCardNumber, _strCardNumber), 298, true, ' '));
                    }
                    else
                    {
                        _sw00.WriteLine(Title + GetStringAsLength(string.Format(_strLine, _strCardNumber, _strCardNumber), 298, true, ' '));
                    }

                    if (_iAddCount > 0)
                    {
                        _strArFamilyNo = _strFamilyNo.Split(new char[] { '|' });

                        for (int j = 0; j < _strArFamilyNo.Length; j++)
                        {
                            _strFamilyCheck = _strArFamilyNo[j];

                            string familyNo = GetStringAsLength(_strArFamilyNo[j], 16, true, ' ');

                            if (_strCardNumber != _strFamilyCheck)
                            {
                                //2011-10-11 수령증 이미지 파일명 제거(?)
                                if (_strStatus == "1" || _strStatus == "7")
                                {
                                    _sw01.WriteLine(Title + GetStringAsLength(string.Format(_strLine, familyNo, familyNo), 298, true, ' '));
                                }
                                else if (_strStatus == "2" || _strStatus == "3")
                                {
                                    _sw02.WriteLine(Title + GetStringAsLength(string.Format(_strLine, familyNo, familyNo), 298, true, ' '));
                                }
                                else
                                {
                                    _sw00.WriteLine(Title + GetStringAsLength(string.Format(_strLine, familyNo, familyNo), 298, true, ' '));
                                }
                            }
                        }
                    }

                }
                _strReturn = string.Format("{0}건의 인계데이타 다운 완료", i);
            }
            catch (Exception)
            {
                _strReturn = string.Format("{0}번째 데이터 인계 중 오류 발생, {1}", i + 1, dtable.Rows[i]["card_barcode"].ToString());
            }
            finally
            {
                if (_sw00 != null) _sw00.Close();
                if (_sw01 != null) _sw01.Close();
                if (_sw02 != null) _sw02.Close();
            }
            return _strReturn;
        }

        //배송상황
        private static string ConvertReceiveType4(DataTable dtable, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩
            StreamWriter _sw01 = null, _sw02 = null, _sw00 = null;  //파일 쓰기 스트림

            string _strLine = "";
            string _strReturn = "", strGetdate = "";
            int cnt_t = 0, i = 0;
            //D_Code1 : 1 = 본사도착, 2 = 지사도착, 3 = 배송진행중, 6 = 재분류, 7 = 카드이동중,
            //8 = 선반납, 9 = 카드반송, 0 = 수령완료
            //D_Code2 : 1~5 = 1~5차 배송 중
            string strD_Code1 = "";
            string strD_date1 = "", strD_time1 = "";
            try
            {
                strGetdate = DateTime.Now.ToString("yyyyMMdd");
                _sw01 = new StreamWriter(fileName + "KS00_" + strGetdate + ".txt", true, _encoding);

                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    strD_Code1 = dtable.Rows[i]["d_code1"].ToString();

                    strD_date1 = dtable.Rows[i]["d_date1"].ToString();
                    strD_time1 = dtable.Rows[i]["d_time1"].ToString().Replace(":", "").Substring(0, 4);

                    //데이터생성 시작
                    _strLine = "K";
                    _strLine += GetStringAsLength(dtable.Rows[i]["client_number"].ToString(), 16, true, ' ');
                    _strLine += GetStringAsLength(dtable.Rows[i]["client_register_type"].ToString(), 1, true, ' ');
                    _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["client_register_date"].ToString()), 8, true, ' ');
                    _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["card_in_date"].ToString()), 8, true, ' ');
                    //배송진행구분
                    _strLine += strD_Code1;
                    _strLine += GetStringAsLength(strD_date1, 8, true, ' ');
                    _strLine += GetStringAsLength(strD_time1, 4, true, ' ');

                    _sw01.WriteLine(_strLine);
                }
                _strReturn = string.Format("{0}건의 인계데이타 다운 완료", i);
            }
            catch (Exception)
            {
                _strReturn = string.Format("{0}번째 데이터 인계 중 오류 발생, {1}", i + 1, dtable.Rows[i]["card_barcode"].ToString());
            }
            finally
            {
                if (_sw01 != null) _sw01.Close();
            }
            return _strReturn;
        }

        private static object CONV_Return_Code(string strR_Code, string strD_Code)
        {
            string strReturn = "";

            switch (strR_Code)
            {
                case "01": strReturn = strD_Code + "01"; break;
                case "04": strReturn = strD_Code + "02"; break;
                case "10": strReturn = strD_Code + "03"; break;
                default:
                    strReturn = strD_Code + "99";
                    break;
            }

            return strReturn;
        }

        // npi_file_name 값 정리
        private static string npi_file_name(string strNpi, string str_type)
        {
            string reStr = null;
            string[] _strNewAdd = null;
            string strNewAdd2 = null;

            if (strNpi.ToString().IndexOf("^") > 0)
            {
                _strNewAdd = strNpi.ToString().Split('^');
            }
            else
            {
                strNewAdd2 = strNpi.ToString();
            }

            if (strNpi == "")
            {
                reStr = GetStringAsLength("1", 36, true, ' ');
            }
            else if (strNpi.ToString().IndexOf("^") > 0)
            {
                // 주소변경타입
                if (str_type == "1")
                {
                    if (_strNewAdd[0].Substring(0, 1) == "1")
                        reStr = GetStringAsLength(_strNewAdd[0], 36, true, ' ');
                    else
                        reStr = GetStringAsLength("1", 36, true, ' ');
                }
                else if (str_type == "2")
                {
                    if (_strNewAdd[1].Substring(0, 1) == "2")
                        reStr = GetStringAsLength(_strNewAdd[1], 36, true, ' ');
                    else
                        reStr = GetStringAsLength("2", 36, true, ' ');
                }
                else
                    reStr = GetStringAsLength("1", 36, true, ' ');
            }
            else
            {
                // 주소변경타입
                if (str_type == "1")
                {
                    if (strNewAdd2.Substring(0, 1) == "1")
                        reStr = GetStringAsLength(strNewAdd2, 36, true, ' ');
                    else
                        reStr = GetStringAsLength("1", 36, true, ' ');
                }
                else if (str_type == "2")
                {
                    if (strNewAdd2.Substring(0, 1) == "2")
                        reStr = GetStringAsLength(strNewAdd2, 36, true, ' ');
                    else
                        reStr = GetStringAsLength("2", 36, true, ' ');
                }
                else
                    reStr = GetStringAsLength("1", 36, true, ' ');
            }
            return reStr;
        }
        //문자중 -를 없앤다
        private static string RemoveDash(string value)
        {
            return value.Replace("-", "");
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

        private static string GetSendCode(string value, string value2)
        {
            if (value.ToLower() == "b" || value.ToLower() == "a" || value.ToLower() == "g" ||
                value.ToLower() == "h" || value.ToLower() == "w" || value.ToLower() == "x" ||
                value.ToLower() == "c" || value.ToLower() == "f" || value.ToLower() == "m"
                )
                return "2";
            else if ((value.ToLower() == "i" || value.ToLower() == "d") && (value2 == "3" || value2 == "2"))
                return "3";
            else
                return "1";
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
