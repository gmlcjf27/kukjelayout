using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace _005_우체국_하나
{
    public class CONVERT
    {
        //기본 인코딩 설정
        private static string strEncoding = "ks_c_5601-1987";
        private static string strCardTypeID = "005";
        private static string strCardTypeName = "우체국_하나카드";

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
            StreamReader _sr = null;											//파일 읽기 스트림
            StreamWriter _swError = null;   									//파일 쓰기 스트림
            DataSet _dsetZipcode = null, _dsetZipcdeArea = null;				//우편번호 관련 DataSet
            DataSet _dsetZipcode_new = null, _dsetZipcdeArea_new = null;		//우편번호 관련 DataSet
            DataTable _dtable = null;							    			//마스터 저장 테이블
            DataRow _dr = null;
            DataRow[] _drs = null;
            byte[] _byteAry = null;
            string _strReturn = "";
            string _strLine = "";
            string _strZipcode = "", _strAreaType = "", _strAreaGroup = "", _strBranch = "", strCard_type_detail = "";
            string strCard_bank_ID = "";
            int _iSeq = 1, _iErrorCount = 0;
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
                _dtable.Columns.Add("card_level_code");                  // dr[5] 파일구분
                _dtable.Columns.Add("client_send_date");                 //발송일자
                _dtable.Columns.Add("client_send_number");             //발송일련번호
                _dtable.Columns.Add("customer_SSN");                    //주민번호
                _dtable.Columns.Add("customer_real_SSN");              //소지자 주민번호
                _dtable.Columns.Add("client_number");                    // dr[10] 회원번호 + 접수번호
                _dtable.Columns.Add("customer_name");                  //이름
                _dtable.Columns.Add("family_name");                       //소지사 성명
                _dtable.Columns.Add("card_pt_code");                     //카드구분
                _dtable.Columns.Add("card_cooperation_code");            //상품종류코드
                _dtable.Columns.Add("card_bank_ID");                     // dr[15] 관리지점코드
                _dtable.Columns.Add("card_bank_name");                 //관리점명
                _dtable.Columns.Add("card_mobile_tel");                  //핸드폰번호
                _dtable.Columns.Add("card_tel1");                          //자택전화
                _dtable.Columns.Add("card_tel2");                          //직장전화
                _dtable.Columns.Add("card_delivery_place_type");         // dr[20] 카드수령지

                _dtable.Columns.Add("card_zipcode2");                   //자택우편번호
                _dtable.Columns.Add("card_address2");                   //자택주소          
                _dtable.Columns.Add("card_zipcode3");                   //직장우편번호
                _dtable.Columns.Add("card_address3");                   //직장주소 

                _dtable.Columns.Add("card_zipcode");                     // dr[25] 수령지우편번호
                _dtable.Columns.Add("card_address_local");               //수령지 동이상
                _dtable.Columns.Add("card_address_detail");              //수령지 동이하

                _dtable.Columns.Add("client_register_date");            //발급일자
                _dtable.Columns.Add("card_issue_type_code");            //발급구분
                _dtable.Columns.Add("client_express_code");            // dr[30]동의서 종류
                _dtable.Columns.Add("client_enterprise_code");         // dr[31] 대면여부
                _dtable.Columns.Add("card_count");                        //카드매수
                _dtable.Columns.Add("card_bill_place_type");            //업체 코드

                _dtable.Columns.Add("card_consented");                  //dr[34] 체크카드 및 카드 등급

                _dtable.Columns.Add("client_request_memo");           //dr[35] 메모
                _dtable.Columns.Add("sms_status");                    //dr[36] 문자수신여부
                _dtable.Columns.Add("card_number");                   //카드넘버
                _dtable.Columns.Add("card_barcode_new");              //카드 뉴바코드
                // 2011-11-02 태희철 수정 새주소
                //새주소과련 필드를 늘리지 않고 기존 필드를 활용
                _dtable.Columns.Add("card_address4_local");             //새주소 60byte
                //2013.01.14 태희철 추가 별지동의서 구분
                _dtable.Columns.Add("choice_agree3");                 //dr[40] 01:클럽, 02:대한, 03:아시아, 04:하이플러스, 05:하이패스, 07:아모레퍼시픽
                _dtable.Columns.Add("card_client_no_1");              //dr[41]신용정보동의
                _dtable.Columns.Add("card_cooperation1");           //dr[42]상품별동의구분,항공사필수동의여부
                _dtable.Columns.Add("card_design_code");            //dr[43]항공사필수동의여부
                _dtable.Columns.Add("client_insert_type");          //dr[44]발송일자일련번호 로직값
                _dtable.Columns.Add("card_issue_type_new");           //dr[45]발급구분코드_new
                
                _dtable.Columns.Add("card_zipcode_new");              //신우편번호
                _dtable.Columns.Add("card_zipcode_kind");             //dr[47]
                _dtable.Columns.Add("card_zipcode2_new");             //
                _dtable.Columns.Add("card_zipcode2_kind");
                _dtable.Columns.Add("card_zipcode3_new");             //dr[50]
                _dtable.Columns.Add("card_zipcode3_kind");

                _dtable.Columns.Add("card_is_for_owner_only");        //dr[52]본인배송
                _dtable.Columns.Add("customer_memo");                 //dr[53]본인배송
                _dtable.Columns.Add("change_add");                    //dr[54]본인배송

                _dtable.Columns.Add("check_org");                   //dr[55]등기번호
                _dtable.Columns.Add("card_cooperation2");           //dr[56]카드사 바코드

                //파일 읽기 Stream과 오류시 저장할 쓰기 Stream 생성
                _sr = new StreamReader(path, _encoding);
                _swError = new StreamWriter(path + ".Error", false, _encoding);

                //우편번호 관련 정보 DataSet에 담기
                _dsetZipcode = new DataSet();
                _dsetZipcdeArea = new DataSet();
                _dsetZipcode.ReadXml(xmlZipcodePath);
                _dsetZipcode.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcode.Tables[0].Columns["zipcode"] };
                _dsetZipcdeArea.ReadXml(xmlZipcodeAreaPath);
                _dsetZipcdeArea.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcdeArea.Tables[0].Columns["zipcode"] };

                //신우편번호 관련 정보 DataSet에 담기
                _dsetZipcode_new = new DataSet();
                _dsetZipcdeArea_new = new DataSet();
                _dsetZipcode_new.ReadXml(xmlZipcodePath_new);
                _dsetZipcode_new.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcode_new.Tables[0].Columns["zipcode_new"] };
                _dsetZipcdeArea_new.ReadXml(xmlZipcodeAreaPath_new);
                _dsetZipcdeArea_new.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcdeArea_new.Tables[0].Columns["zipcode_new"] };

                while ((_strLine = _sr.ReadLine()) != null)
                {
                    //인코딩, byte 배열로 담기
                    if (_iSeq == 1)
                    {
                        strCard_type_detail = _strLine.Substring(_strLine.Length - 7, 7);
                    }
                    

                    //인코딩, byte 배열로 담기
                    _byteAry = _encoding.GetBytes(_strLine);

                    if (_byteAry.Length != 1207)
                    {
                        MessageBox.Show("줄번호 " + _iSeq.ToString() + " 번째 총byte오류. 제휴사코드 포함 총byte는 1207byte 입니다.", "오류");
                        throw new ArgumentNullException("총 byte 오류");
                    }

                    _dr = _dtable.NewRow();
                    _dr[0] = _iSeq;
                    _dr[5] = _encoding.GetString(_byteAry, 0, 2);
                    _dr[6] = _encoding.GetString(_byteAry, 2, 8);
                    _dr[7] = _encoding.GetString(_byteAry, 10, 8);
                    _dr[8] = _encoding.GetString(_byteAry, 18, 13).Replace("*", "x"); ;
                    _dr[9] = _encoding.GetString(_byteAry, 31, 13).Replace("*", "x"); ;
                    _dr[10] = _encoding.GetString(_byteAry, 44, 16);
                    _dr[11] = _encoding.GetString(_byteAry, 60, 25);
                    _dr[12] = _encoding.GetString(_byteAry, 85, 32);
                    _dr[13] = _encoding.GetString(_byteAry, 117, 1);
                    _dr[14] = _encoding.GetString(_byteAry, 118, 5);

                    strCard_bank_ID = _encoding.GetString(_byteAry, 123, 4);

                    if (strCard_bank_ID.Trim().Length == 3)
                    {
                        strCard_bank_ID = "0" + strCard_bank_ID;
                    }

                    _dr[15] = strCard_bank_ID;
                    _dr[16] = _encoding.GetString(_byteAry, 127, 30);
                    _dr[17] = _encoding.GetString(_byteAry, 157, 15);
                    _dr[18] = _encoding.GetString(_byteAry, 172, 15);
                    _dr[19] = _encoding.GetString(_byteAry, 187, 15);

                    _dr[20] = _encoding.GetString(_byteAry, 202, 3);

                    //하나카드 확인 요청하였으나 확인 불가 904 -> 003(기타) 변환하여 등록
                    if (_dr[20].ToString() == "904")
                    {
                        _dr[20] = "003";
                    }

                    if (_dr[20].ToString() == "001" || _dr[20].ToString() == "003")
                    {
                        _dr[21] = _encoding.GetString(_byteAry, 205, 6).Trim();

                        if (_dr[21].ToString().Length == 5)
                        {
                            _dr[48] = _dr[21].ToString();
                            _dr[49] = "1";
                        }

                        _dr[22] = _encoding.GetString(_byteAry, 211, 60).TrimEnd() + _encoding.GetString(_byteAry, 271, 100);

                        _dr[23] = _encoding.GetString(_byteAry, 371, 6).Trim();

                        if (_dr[23].ToString().Length == 5)
                        {
                            _dr[50] = _dr[21].ToString();
                            _dr[51] = "1";
                        }

                        _dr[24] = _encoding.GetString(_byteAry, 377, 60).TrimEnd() + _encoding.GetString(_byteAry, 437, 100);
                    }
                    else
                    {
                        _dr[23] = _encoding.GetString(_byteAry, 205, 6).Trim();

                        if (_dr[23].ToString().Length == 5)
                        {
                            _dr[50] = _dr[21].ToString();
                            _dr[51] = "1";
                        }

                        _dr[24] = _encoding.GetString(_byteAry, 211, 60).TrimEnd() + _encoding.GetString(_byteAry, 271, 100);

                        _dr[21] = _encoding.GetString(_byteAry, 371, 6).Trim();

                        if (_dr[21].ToString().Length == 5)
                        {
                            _dr[48] = _dr[21].ToString();
                            _dr[49] = "1";
                        }

                        _dr[22] = _encoding.GetString(_byteAry, 377, 60).TrimEnd() + _encoding.GetString(_byteAry, 437, 100);
                    }

                    _strZipcode = _encoding.GetString(_byteAry, 537, 6).Trim();
                    _dr[25] = _strZipcode;

                    if (_strZipcode.Length == 5)
                    {
                        _dr[46] = _strZipcode;
                        _dr[47] = "1";
                    }

                    _dr[26] = _encoding.GetString(_byteAry, 543, 60);
                    _dr[27] = _encoding.GetString(_byteAry, 603, 100);

                    _dr[28] = _encoding.GetString(_byteAry, 703, 8);
                    _dr[29] = _encoding.GetString(_byteAry, 711, 2);
                    _dr[30] = _encoding.GetString(_byteAry, 713, 1);
                    _dr[31] = _encoding.GetString(_byteAry, 714, 1);
                    _dr[32] = _encoding.GetString(_byteAry, 715, 2);
                    _dr[33] = _encoding.GetString(_byteAry, 717, 2);
                    _dr[34] = _encoding.GetString(_byteAry, 719, 1);
                    _dr[35] = _encoding.GetString(_byteAry, 720, 60);
                    _dr[36] = _encoding.GetString(_byteAry, 838, 1);

                    _dr[37] = _dr[6].ToString() + _dr[7].ToString();

                    int temp_int = 0;

                    if (_dr[33].ToString() == "")
                    {
                        temp_int = Convert.ToInt32(_dr[5]);

                        if (temp_int == 11 || temp_int == 12 || temp_int == 21 || temp_int == 22)
                            _dr[33] = "01";
                        else if (temp_int == 13 || temp_int == 14 || temp_int == 23 || temp_int == 24)
                            _dr[33] = "04";
                        else if (temp_int == 15 || temp_int == 16 || temp_int == 25 || temp_int == 26)
                            _dr[33] = "02";
                        else if (temp_int == 17 || temp_int == 18 || temp_int == 27 || temp_int == 28)
                            _dr[33] = "07";
                        else
                            throw new Exception("구분코드 범위 넘어감");
                    }

                    // 일반의 경우 : 0 // 동의서 : 1
                    if (_dr[31].ToString() == "N")
                        temp_int = 0;
                    else if (_dr[31].ToString() == "Y" || _dr[31].ToString() == "A")
                        temp_int = 1;
                    //본인지정의 경우 : 2
                    else if (_dr[31].ToString() == "I")
                        temp_int = 2;
                    else
                        throw new Exception("대면여부 범위 넘어감");
                    

                    //2011-11-02 태희철 추가 신주소 
                    _dr[39] = _encoding.GetString(_byteAry, 846, 60);
                    //2013.01.18 태희철 추가
                    _dr[40] = _encoding.GetString(_byteAry, 1006, 2);
                    //신용정보동의
                    _dr[41] = _encoding.GetString(_byteAry, 1008, 20);
                    //상품별동의구분,항공사필수동의여부
                    _dr[42] = _encoding.GetString(_byteAry, 1028, 40);
                    _dr[43] = _encoding.GetString(_byteAry, 1068, 5);

                    _dr[44] = _encoding.GetString(_byteAry, 1073, 1);

                    //NEW발급구분코드 추가
                    if (strCard_type_detail == "00541")
                    {
                        switch (_dr[29].ToString())
                        {
                            case "11":
                                _dr[45] = "1";
                                break;
                            case "12":
                            case "13":
                            case "14":
                            case "15":
                            case "16":
                            case "17":
                            case "18":
                            case "19":
                                _dr[45] = "6";
                                break;
                            case "30":
                            case "31":
                            case "32":
                                _dr[45] = "3";
                                break;
                            default:
                                _dr[45] = "2";
                                break;
                        }
                    }
                    else
                    {
                        switch (_dr[29].ToString())
                        {
                            case "01":
                                _dr[45] = "1"; break;
                            case "02":
                            case "03":
                                _dr[45] = "6"; break;
                            case "21":
                            case "22":
                                _dr[45] = "3"; break;
                            default:
                                _dr[45] = "2";
                                break;
                        }
                    }

                    if (strCard_type_detail == "0051102")
                    {
                        _dr[52] = "1";
                        _dr[53] = "";
                        _dr[54] = "1";
                    }

                    //등기번호
                    _dr[55] = _encoding.GetString(_byteAry, 781, 13);

                    //새우편번호의 경우 카드사바코드의 우편번호 5자리 + " "
                    if (_strZipcode.Length == 5)
                    {
                        _dr[38] = _dr[6].ToString() + _dr[7].ToString() + _strZipcode + " " + _dr[33].ToString() + _dr[29].ToString() + temp_int.ToString() + "99";
                    }
                    else
                    {
                        _dr[38] = _dr[6].ToString() + _dr[7].ToString() + (_strZipcode == "" ? "      " : _strZipcode) + _dr[33].ToString() + _dr[29].ToString() + temp_int.ToString() + "99";
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

                //변환에 성공했다면
                if (_iErrorCount < 1)
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

        //일일마감자료
        public static string ConvertResultDay(System.Data.DataTable dtable, string fileName)
        {
            return ConvertResult(dtable, fileName);
        }

        //마감 자료
        public static string ConvertResult(System.Data.DataTable dtable, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            //일반, 동의
            StreamWriter _sw00 = null, _sw01 = null, _sw02 = null, _sw10 = null, _sw11 = null, _sw12 = null;
            //지점일반, 지점동의
            StreamWriter _sw20 = null, _sw21 = null, _sw22 = null, _sw30 = null, _sw31 = null, _sw32 = null, _sw99 = null;
            //일반-본인지정
            StreamWriter _sw40 = null, _sw41 = null, _sw42 = null;

            StringBuilder _strLine = new StringBuilder("");
            string _strReturn = "", _strDeliveryStatus = "", strCard_type_detail = "";					//파일 쓰기 스트림
            int count = 0;
            int iCnt1 = 0, iCnt2 = 0, iCnt3 = 0, iCnt11 = 0, iCnt12 = 0, iCnt13 = 0;
            int iCnt21 = 0, iCnt22 = 0, iCnt23 = 0, iCnt31 = 0, iCnt32 = 0, iCnt33 = 0;
            int iCnt41 = 0, iCnt42 = 0, iCnt43 = 0;

            try
            {
                string temp_time = DateTime.Now.ToShortDateString().Replace("-", "").Substring(2, 6);
                _sw00 = new StreamWriter(fileName + "KUKJ103." + temp_time + ".00", true, _encoding);
                _sw01 = new StreamWriter(fileName + "KUKJ103." + temp_time + ".01", true, _encoding);
                _sw02 = new StreamWriter(fileName + "KUKJ103." + temp_time + ".02", true, _encoding);

                _sw10 = new StreamWriter(fileName + "KUKJ104." + temp_time + ".00", true, _encoding);
                _sw11 = new StreamWriter(fileName + "KUKJ104." + temp_time + ".01", true, _encoding);
                _sw12 = new StreamWriter(fileName + "KUKJ104." + temp_time + ".02", true, _encoding);

                _sw20 = new StreamWriter(fileName + "KUKJ107." + temp_time + ".00", true, _encoding);
                _sw21 = new StreamWriter(fileName + "KUKJ107." + temp_time + ".01", true, _encoding);
                _sw22 = new StreamWriter(fileName + "KUKJ107." + temp_time + ".02", true, _encoding);

                _sw30 = new StreamWriter(fileName + "KUKJ108." + temp_time + ".00", true, _encoding);
                _sw31 = new StreamWriter(fileName + "KUKJ108." + temp_time + ".01", true, _encoding);
                _sw32 = new StreamWriter(fileName + "KUKJ108." + temp_time + ".02", true, _encoding);

                _sw40 = new StreamWriter(fileName + "KUKJ115." + temp_time + ".00", true, _encoding);
                _sw41 = new StreamWriter(fileName + "KUKJ115." + temp_time + ".01", true, _encoding);
                _sw42 = new StreamWriter(fileName + "KUKJ115." + temp_time + ".02", true, _encoding);
                

                _strLine = _strLine.Append(GetStringAsLength("H", 1));
                _strLine.Append(GetStringAsLength(DateTime.Now.ToShortDateString().Replace("-", ""), 8));
                _strLine.Append(GetStringAsLength("", 293));

                _sw00.WriteLine(_strLine.ToString());
                _sw01.WriteLine(_strLine.ToString());
                _sw02.WriteLine(_strLine.ToString());

                _sw10.WriteLine(_strLine.ToString());
                _sw11.WriteLine(_strLine.ToString());
                _sw12.WriteLine(_strLine.ToString());

                _sw20.WriteLine(_strLine.ToString());
                _sw21.WriteLine(_strLine.ToString());
                _sw22.WriteLine(_strLine.ToString());

                _sw30.WriteLine(_strLine.ToString());
                _sw31.WriteLine(_strLine.ToString());
                _sw32.WriteLine(_strLine.ToString());

                _sw40.WriteLine(_strLine.ToString());
                _sw41.WriteLine(_strLine.ToString());
                _sw42.WriteLine(_strLine.ToString());

                for (int i = 0; i < dtable.Rows.Count; i++)
                {
                    _strDeliveryStatus = dtable.Rows[i]["card_delivery_status"].ToString();
                    strCard_type_detail = dtable.Rows[i]["card_type_detail"].ToString();

                    //쓸데이터 만들기
                    _strLine = new StringBuilder(GetStringAsLength("D01", 3));//업체코드

                    _strLine.Append(GetStringAsLength("1", 1));//파일구분(1:일반 2:영업점)
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_send_date"].ToString().Replace("-", ""), 8));//발급일자
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_send_number"].ToString().Replace("-", ""), 8)); //발급 번호
                    _strLine.Append(GetStringAsLength(delivery_stat(dtable.Rows[i]), 2)); //진행코드
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["branch_name"].ToString(), 20)); //배송지사명
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["career"].ToString(), 20)); //배송담당자명
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["delivery_out_date1"].ToString().Replace("-", ""), 8)); //1차출고일자
                    _strLine.Append(GetStringAsLength(return_reason(dtable.Rows[i]["delivery_return_reason1"].ToString()), 2)); //1차 반송코드
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["delivery_out_date2"].ToString().Replace("-", ""), 8)); //2차출고일자
                    _strLine.Append(GetStringAsLength(return_reason(dtable.Rows[i]["delivery_return_reason2"].ToString()), 2)); //1차 반송코드
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["delivery_out_date3"].ToString().Replace("-", ""), 8)); //3차출고일자
                    _strLine.Append(GetStringAsLength(return_reason(dtable.Rows[i]["delivery_return_reason3"].ToString()), 2)); //1차 반송코드
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_name"].ToString(), 20)); //수취인명
                    _strLine.Append(GetStringAsLength(receiver_code(dtable.Rows[i]["receiver_code"].ToString()), 2)); //수령인관계코드
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_SSN"].ToString().Replace("-", ""), 13)); //수취인민증 
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_delivery_date"].ToString().Replace("-", ""), 8)); //카드수령일자(결과등록일)
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_bank_id"].ToString().Replace("-", ""), 4)); //영업점코드
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_issue_type_code"].ToString().Replace("-", ""), 2)); //발급구분
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["receipt_number"].ToString().Replace("|", ""), 1)); //긴급구분                    
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_express_code"].ToString().Replace("-", ""), 1)); //포장구분
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_enterprise_code"].ToString().Replace("-", ""), 1)); //대면여부 
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_count"].ToString().Replace("-", ""), 4)); //카드매수
                    _strLine.Append(GetStringAsLength(" ", 1)); //카드매수
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_address"].ToString().Replace("-", ""), 100)); //배숑결과주소
                    _strLine.Append(GetStringAsLength("", 53));
                    //쓰기

                    //일반마감
                    if (strCard_type_detail == "0051101")
                    {
                        if (_strDeliveryStatus == "1")
                        {
                            iCnt1++;
                            _sw01.WriteLine(_strLine.ToString());
                        }
                        else if (_strDeliveryStatus == "2" || _strDeliveryStatus == "3")
                        {
                            iCnt2++;
                            _sw02.WriteLine(_strLine.ToString());
                        }
                        else
                        {
                            iCnt3++;
                            _sw00.WriteLine(_strLine.ToString());
                        }
                    }
                    //일반-본인지정
                    else if (strCard_type_detail == "0051102")
                    {
                        if (_strDeliveryStatus == "1")
                        {
                            iCnt41++;
                            _sw41.WriteLine(_strLine.ToString());
                        }
                        else if (_strDeliveryStatus == "2" || _strDeliveryStatus == "3")
                        {
                            iCnt42++;
                            _sw42.WriteLine(_strLine.ToString());
                        }
                        else
                        {
                            iCnt43++;
                            _sw40.WriteLine(_strLine.ToString());
                        }
                    }
                    //동의서마감
                    else if (strCard_type_detail.Substring(0,5) == "00521")
                    {
                        if (_strDeliveryStatus == "1")
                        {
                            iCnt11++;
                            _sw11.WriteLine(_strLine.ToString());
                        }
                        else if (_strDeliveryStatus == "2" || _strDeliveryStatus == "3")
                        {
                            iCnt12++;
                            _sw12.WriteLine(_strLine.ToString());
                        }
                        else
                        {
                            iCnt13++;
                            _sw10.WriteLine(_strLine.ToString());
                        }
                    }
                    //지점일반
                    else if (strCard_type_detail == "0054101")
                    {
                        if (_strDeliveryStatus == "1")
                        {
                            iCnt21++;
                            _sw21.WriteLine(_strLine.ToString());
                        }
                        else if (_strDeliveryStatus == "2" || _strDeliveryStatus == "3")
                        {
                            iCnt22++;
                            _sw22.WriteLine(_strLine.ToString());
                        }
                        else
                        {
                            iCnt23++;
                            _sw20.WriteLine(_strLine.ToString());
                        }
                    }
                    //지점동의
                    else if (strCard_type_detail == "0054102")
                    {
                        if (_strDeliveryStatus == "1")
                        {
                            iCnt31++;
                            _sw31.WriteLine(_strLine.ToString());
                        }
                        else if (_strDeliveryStatus == "2" || _strDeliveryStatus == "3")
                        {
                            iCnt32++;
                            _sw32.WriteLine(_strLine.ToString());
                        }
                        else
                        {
                            iCnt33++;
                            _sw30.WriteLine(_strLine.ToString());
                        }
                    }
                    count++;
                }

                _strLine = new StringBuilder("T", 1);
                _strLine.Append(GetStringAsLength(iCnt1.ToString(), 8));
                _strLine.Append(GetStringAsLength("", 293));
                _sw01.WriteLine(_strLine.ToString());

                _strLine = new StringBuilder("T", 1);
                _strLine.Append(GetStringAsLength(iCnt2.ToString(), 8));
                _strLine.Append(GetStringAsLength("", 293));
                _sw02.WriteLine(_strLine.ToString());

                _strLine = new StringBuilder("T", 1);
                _strLine.Append(GetStringAsLength(iCnt3.ToString(), 8));
                _strLine.Append(GetStringAsLength("", 293));
                _sw00.WriteLine(_strLine.ToString());

                _strLine = new StringBuilder("T", 1);
                _strLine.Append(GetStringAsLength(iCnt11.ToString(), 8));
                _strLine.Append(GetStringAsLength("", 293));
                _sw11.WriteLine(_strLine.ToString());

                _strLine = new StringBuilder("T", 1);
                _strLine.Append(GetStringAsLength(iCnt12.ToString(), 8));
                _strLine.Append(GetStringAsLength("", 293));
                _sw12.WriteLine(_strLine.ToString());

                _strLine = new StringBuilder("T", 1);
                _strLine.Append(GetStringAsLength(iCnt13.ToString(), 8));
                _strLine.Append(GetStringAsLength("", 293));
                _sw10.WriteLine(_strLine.ToString());

                _strLine = new StringBuilder("T", 1);
                _strLine.Append(GetStringAsLength(iCnt21.ToString(), 8));
                _strLine.Append(GetStringAsLength("", 293));
                _sw21.WriteLine(_strLine.ToString());

                _strLine = new StringBuilder("T", 1);
                _strLine.Append(GetStringAsLength(iCnt22.ToString(), 8));
                _strLine.Append(GetStringAsLength("", 293));
                _sw22.WriteLine(_strLine.ToString());

                _strLine = new StringBuilder("T", 1);
                _strLine.Append(GetStringAsLength(iCnt23.ToString(), 8));
                _strLine.Append(GetStringAsLength("", 293));
                _sw20.WriteLine(_strLine.ToString());

                _strLine = new StringBuilder("T", 1);
                _strLine.Append(GetStringAsLength(iCnt31.ToString(), 8));
                _strLine.Append(GetStringAsLength("", 293));
                _sw31.WriteLine(_strLine.ToString());

                _strLine = new StringBuilder("T", 1);
                _strLine.Append(GetStringAsLength(iCnt32.ToString(), 8));
                _strLine.Append(GetStringAsLength("", 293));
                _sw32.WriteLine(_strLine.ToString());

                _strLine = new StringBuilder("T", 1);
                _strLine.Append(GetStringAsLength(iCnt33.ToString(), 8));
                _strLine.Append(GetStringAsLength("", 293));
                _sw30.WriteLine(_strLine.ToString());

                _strLine = new StringBuilder("T", 1);
                _strLine.Append(GetStringAsLength(iCnt41.ToString(), 8));
                _strLine.Append(GetStringAsLength("", 293));
                _sw41.WriteLine(_strLine.ToString());

                _strLine = new StringBuilder("T", 1);
                _strLine.Append(GetStringAsLength(iCnt42.ToString(), 8));
                _strLine.Append(GetStringAsLength("", 293));
                _sw42.WriteLine(_strLine.ToString());

                _strLine = new StringBuilder("T", 1);
                _strLine.Append(GetStringAsLength(iCnt43.ToString(), 8));
                _strLine.Append(GetStringAsLength("", 293));
                _sw40.WriteLine(_strLine.ToString());

                _strReturn = string.Format("{0}건의 인계데이터 다운 완료", count);
            }
            catch (Exception)
            {
                _strReturn = string.Format("{0}번째 데이터 인계 중 오류 발생", count + 1);
            }
            finally
            {
                if (_sw00 != null) _sw00.Close();
                if (_sw01 != null) _sw01.Close();
                if (_sw02 != null) _sw02.Close();

                if (_sw10 != null) _sw10.Close();
                if (_sw11 != null) _sw11.Close();
                if (_sw12 != null) _sw12.Close();

                if (_sw20 != null) _sw20.Close();
                if (_sw21 != null) _sw21.Close();
                if (_sw22 != null) _sw22.Close();

                if (_sw30 != null) _sw30.Close();
                if (_sw31 != null) _sw31.Close();
                if (_sw32 != null) _sw32.Close();

                if (_sw40 != null) _sw40.Close();
                if (_sw41 != null) _sw41.Close();
                if (_sw42 != null) _sw42.Close();

                if (_sw99 != null) _sw99.Close();
            }
            return _strReturn;
        }

        /// <summary>
        /// 지역번호 정리
        /// </summary>
        /// <param name="dtable"></param>
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

        /// <summary>
        /// 진행 상황 산출 함수
        /// </summary>
        /// <param name="dr">레코드</param>
        /// <returns>진행 코드</returns>
        private static string delivery_stat(DataRow dr)
        {
            string temp = "";

            if (dr["card_delivery_step"].ToString() == "1")
            {
                temp = "07";
            }
            else if (dr["delivery_out_date1"].ToString() == "")
            {
                temp = "00";
            }
            else if (dr["delivery_in_date1"].ToString() == "")
            {
                temp = "01";
            }
            else if (dr["delivery_out_date2"].ToString() == "")
            {
                temp = "02";
            }
            else if (dr["delivery_in_date2"].ToString() == "")
            {
                temp = "03";
            }
            else if (dr["delivery_out_date3"].ToString() == "")
            {
                temp = "04";
            }
            else if (dr["delivery_in_date3"].ToString() == "")
            {
                temp = "05";
            }
            else
            {
                temp = "06";
            }

            return temp;
        }

        /// <summary>
        /// 반송 코트 변환
        /// </summary>
        /// <param name="value">반송 코드</param>
        /// <returns>변환 된 결과</returns>
        private static string return_reason(string value)
        {
            string temp = "";
            switch (value)
            {
                case "07":
                    temp = "01";
                    break;
                case "05":
                    temp = "02";
                    break;
                case "03":
                    temp = "03";
                    break;
                case "08":
                    temp = "04";
                    break;
                case "09":
                    temp = "05";
                    break;
                case "02":
                    temp = "06";
                    break;
                case "01":
                    temp = "07";
                    break;
                //case "09":
                //    temp = "08";
                //    break;
                case "32":
                    temp = "09";
                    break;
                case "31":
                    temp = "10";
                    break;
                case "88":
                    temp = "11";
                    break;
                case "0":
                    temp = "";
                    break;
                case "":
                    temp = "";
                    break;
                default:
                    temp = "99";
                    break;
            }

            return temp;
        }

        /// <summary>
        /// 수령인 관계코드 변환
        /// </summary>
        /// <param name="value">관계코드</param>
        /// <returns>변환된 코드</returns>
        private static string receiver_code(string value)
        {
            string temp = "";
            switch (value)
            {
                case "":
                    temp = "  ";
                    break;
                case "00":
                    temp = "  ";
                    break;
                case "01":
                    temp = "01";
                    break;
                case "06":
                    temp = "02";
                    break;
                case "07":
                    temp = "02";
                    break;
                case "04":
                    temp = "03";
                    break;
                case "05":
                    temp = "03";
                    break;
                case "08":
                    temp = "04";
                    break;
                case "09":
                    temp = "04";
                    break;
                case "10":
                    temp = "05";
                    break;
                case "11":
                    temp = "05";
                    break;
                case "12":
                    temp = "06";
                    break;
                case "13":
                    temp = "06";
                    break;
                case "19":
                    temp = "07";
                    break;
                case "20":
                    temp = "08";
                    break;
                case "23":
                    temp = "09";
                    break;
                case "24":
                    temp = "09";
                    break;
                default:
                    temp = "99";
                    break;
            }
            return temp;
        }

        /// <summary>
        /// 문자열 체우기(오버라이팅)
        /// </summary>
        /// <param name="Text">문자열</param>
        /// <param name="Length">전체 길이</param>
        /// <returns>공백으로 체워진 문자열</returns>
        private static string GetStringAsLength(string Text, int Length)
        {
            return GetStringAsLength(Text, Length, true, ' ');
        }

        /// <summary>
        /// 문자열 체우기(오버라이팅)
        /// </summary>
        /// <param name="Text">문자열</param>
        /// <param name="Length">전체 길이</param>
        /// <param name="blankPositionAtBack">왼쪽 오른쪽 전열</param>
        /// <param name="chBlank">공백에 넣을 문자</param>
        /// <returns></returns>
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
