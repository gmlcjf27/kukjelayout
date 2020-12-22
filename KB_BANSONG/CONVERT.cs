using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace _002_국민_지점
{
    public class CONVERT
    {
        //기본 인코딩 설정
        private static string strEncoding = "ks_c_5601-1987";
        private static string strCardTypeID = "201";
        private static string strCardTypeName = "국민(은)";

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
        //public static string ConvertRegister(string path, string xmlZipcodePath, string xmlZipcodeAreaPath, string xmlBankPath, string xmlPath)
        public static string ConvertRegister(string path, string xmlZipcodePath, string xmlZipcodeAreaPath, string xmlZipcodePath_new, string xmlZipcodeAreaPath_new,string xmlBankPath, string xmlPath)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            //FileInfo _fi = null;
            StreamReader _sr = null;														//파일 읽기 스트림
            StreamWriter _swError = null;													//파일 쓰기 스트림
            DataSet _dsetZipcode = null, _dsetZipcdeArea = null, _dsetBank = null;			//우편번호 관련 DataSet
            DataSet _dsetZipcode_new = null, _dsetZipcdeArea_new = null;			//우편번호 관련 DataSet
            DataTable _dtable = null;														//마스터 저장 테이블
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
                _dtable.Columns.Add("client_number");                    // dr[15] 송증번호
                _dtable.Columns.Add("client_send_number");               // dr[16] 발송번호 (송장출력번호)
                
                _dtable.Columns.Add("card_zipcode_new");                 // dr[17]
                _dtable.Columns.Add("card_zipcode_kind");                // dr[18]

                _dtable.Columns.Add("customer_ssn");                     // dr[19]
                _dtable.Columns.Add("card_is_for_owner_only");           // dr[20]


                //우편번호 관련 정보 DataSet에 담기
                _dsetBank = new DataSet();
                _dsetZipcode = new DataSet();
                _dsetZipcdeArea = new DataSet();
                _dsetZipcode.ReadXml(xmlZipcodePath);
                _dsetZipcode.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcode.Tables[0].Columns["zipcode"] };
                _dsetZipcdeArea.ReadXml(xmlZipcodeAreaPath);
                _dsetZipcdeArea.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcdeArea.Tables[0].Columns["zipcode"] };
                _dsetBank.ReadXml(xmlBankPath);

                //새우편번호 관련 정보 DataSet에 담기
                _dsetZipcode_new = new DataSet();
                _dsetZipcdeArea_new = new DataSet();
                _dsetZipcode_new.ReadXml(xmlZipcodePath_new);
                _dsetZipcode_new.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcode_new.Tables[0].Columns["zipcode_new"] };
                _dsetZipcdeArea_new.ReadXml(xmlZipcodeAreaPath_new);
                _dsetZipcdeArea_new.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcdeArea_new.Tables[0].Columns["zipcode_new"] };

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
                    _strZipcode = _encoding.GetString(_byteAry, 45, 6).Trim();
                    _dr[11] = _strZipcode;

                    if (_strZipcode.Length == 5)
                    {
                        _dr[17] = _strZipcode;
                        _dr[18] = "1";
                    }

                    _dr[12] = _encoding.GetString(_byteAry, 51, 200);
                    _dr[13] = _encoding.GetString(_byteAry, 251, 4);
                    _dr[14] = _encoding.GetString(_byteAry, 255, 5);
                    _dr[15] = _encoding.GetString(_byteAry, 260, 10);
                    _dr[16] = _encoding.GetString(_byteAry, 270, 16);

                    _dr[19] = "xxxxxxxxxxxxx";
                    _dr[20] = "2";
                    //_dr[16] = _dr[9] + _strZipcode;

                    if (_strBankID != "")
                    {
                        //지역 분류 선택
                        _drs = _dsetBank.Tables[0].Select("bank_ID= '" + _strBankID + "'");

                        if (_drs.Length < 1)
                        {
                            _strZipcode = "";
                            _strAddress = "";
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

        //마감 자료
        public static string ConvertResult(System.Data.DataTable dtable, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            StreamWriter _sw = null;																			//파일 쓰기 스트림
            string _strLine = "";
            string _strReturn = "";
            int i = -1;
            try
            {
                _sw = new StreamWriter(fileName, true, _encoding);
                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    _strLine = GetStringAsLength(RemoveDash(dtable.Rows[i]["client_send_date"].ToString().Trim()), 8);
                    _strLine += GetStringAsLength(dtable.Rows[i]["client_send_number"].ToString().Trim(), 5);
                    _strLine += GetStringAsLength(dtable.Rows[i]["customer_name"].ToString().Trim(), 40);
                    _strLine += GetStringAsLength(ConvertSSN(dtable.Rows[i]["customer_SSN"].ToString()).Trim(), 13);
                    _strLine += GetStringAsLength(dtable.Rows[i]["card_number"].ToString().Trim(), 16);
                    _strLine += GetStringAsLength(ConvertZipcode(dtable.Rows[i]["card_zipcode"].ToString().Trim()), 7);
                    _strLine += GetStringAsLength(dtable.Rows[i]["card_address"].ToString().Trim(), 100);
                    _strLine += GetStringAsLength(dtable.Rows[i]["card_tel1"].ToString().Trim(), 16);
                    _strLine += GetStringAsLength(dtable.Rows[i]["card_tel2"].ToString().Trim(), 16);
                    _strLine += GetStringAsLength(dtable.Rows[i]["card_mobile_tel"].ToString().Trim(), 16);
                    _strLine += GetStringAsLength(DeliveryType(dtable.Rows[i]["card_delivery_status"].ToString()).Trim(), 8);
                    _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["card_delivery_date"].ToString().Trim()), 8);
                    _strLine += GetStringAsLength(dtable.Rows[i]["receiver_name"].ToString().Trim(), 20);
                    _strLine += GetStringAsLength(ReceiverType(dtable.Rows[i]["receiver_code"].ToString()).Trim(), 20);
                    if (dtable.Rows[i]["receiver_SSN"].ToString().Length == 0)
                    {
                        _strLine += "";
                    }
                    else
                    {
                        _strLine += GetStringAsLength(ConvertSSN(dtable.Rows[i]["receiver_SSN"].ToString()).Trim(), 13);
                    }
                    _strLine += GetStringAsLength(ReturnType(dtable.Rows[i]["delivery_return_reason_last"].ToString()).Trim(), 20);

                    if (dtable.Rows[i]["card_delivery_status"].ToString() == "1")
                    {
                        if (dtable.Rows[i]["card_kind"].ToString().ToLower() == "d")
                            _strLine += GetStringAsLength(dtable.Rows[i]["delivery_is_pda_register"].ToString(), 1);
                        else
                            _strLine += GetStringAsLength("1", 1);
                        _strLine += GetStringAsLength("K" + GetSendCode(dtable.Rows[i]["client_register_type"].ToString(), dtable.Rows[i]["client_insert_type"].ToString()), 2);
                        _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["client_send_date"].ToString()), 8);
                        _strLine += GetStringAsLength(dtable.Rows[i]["client_enterprise_code"].ToString(), 2);
                        _strLine += GetStringAsLength(dtable.Rows[i]["card_client_no_1"].ToString(), 6);
                        _strLine += GetStringAsLength(".tif", 6);
                    }
                    else
                    {
                        _strLine += GetStringAsLength("", 23);
                    }

                    _sw.WriteLine(_strLine);
                }
                _strReturn = string.Format("{0}건의 인계데이터 다운 완료", i);
            }
            catch (Exception)
            {
                _strReturn = string.Format("{0}번째 데이터 인계 중 오류 발생, {1}", i + 1, dtable.Rows[i]["card_barcode"].ToString());
            }
            finally
            {
                _sw.Close();
            }
            return _strReturn;
        }

        //일일마감자료
        public static string ConvertResultDay(System.Data.DataTable dtable, string fileName)
        {
            return ConvertResult(dtable, fileName);
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

        //공백 모두 제거
        private static string RemoveBlank(string value)
        {
            return value.Replace(" ", "");
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
            if (_strReturn.Length == 6)
            {
                _strReturn = value.Substring(0, 6) + "-";
            }
            else
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
        #region 배송
        public static string DeliveryType(string value)
        {
            string _strReturn = value;

            switch (value)
            {
                case "0":
                    _strReturn = "미배송";
                    break;
                case "1":
                    _strReturn = "배송";
                    break;
                case "2":
                    _strReturn = "반송";
                    break;
            }
            return _strReturn;
        }
        #endregion

        #region 수령인관계
        public static string ReceiverType(string value)
        {
            string _strReceiver = value;
            switch (value)
            {
                case "00":
                    _strReceiver = "";
                    break;
                case "01":
                    _strReceiver = "본인";
                    break;
                case "02":
                    _strReceiver = "조부";
                    break;
                case "03":
                    _strReceiver = "조모";
                    break;
                case "04":
                    _strReceiver = "아버지";
                    break;
                case "05":
                    _strReceiver = "어머니";
                    break;
                case "06":
                    _strReceiver = "처";
                    break;
                case "07":
                    _strReceiver = "남편";
                    break;
                case "08":
                    _strReceiver = "형";
                    break;
                case "09":
                    _strReceiver = "동생";
                    break;
                case "10":
                    _strReceiver = "언니";
                    break;
                case "11":
                    _strReceiver = "누나";
                    break;
                case "12":
                    _strReceiver = "아들";
                    break;
                case "13":
                    _strReceiver = "딸";
                    break;
                case "14":
                    _strReceiver = "손자";
                    break;
                case "15":
                    _strReceiver = "손녀";
                    break;
                case "16":
                    _strReceiver = "며느리";
                    break;
                case "17":
                    _strReceiver = "삼촌";
                    break;
                case "18":
                    _strReceiver = "사촌";
                    break;
                case "19":
                    _strReceiver = "친척";
                    break;
                case "20":
                    _strReceiver = "직장동료";
                    break;
                case "21":
                    _strReceiver = "상사";
                    break;
                case "22":
                    _strReceiver = "친구";
                    break;
                case "23":
                    _strReceiver = "선배";
                    break;
                case "24":
                    _strReceiver = "후배";
                    break;
                case "25":
                    _strReceiver = "은행원";
                    break;
                case "26":
                    _strReceiver = "이웃";
                    break;
                case "27":
                    _strReceiver = "주인집";
                    break;
                case "28":
                    _strReceiver = "형수";
                    break;
                case "29":
                    _strReceiver = "경비원";
                    break;
                case "30":
                    _strReceiver = "친지";
                    break;
                case "31":
                    _strReceiver = "오빠";
                    break;
                case "32":
                    _strReceiver = "보증인";
                    break;
                case "33":
                    _strReceiver = "고객요청";
                    break;
                case "99":
                    _strReceiver = "기타";
                    break;

            }
            return _strReceiver;
        }
        #endregion
        #region 반송사유
        public static string ReturnType(string value)
        {
            string _strReturn = value;

            switch (value)
            {
                case "0":
                    _strReturn = "";
                    break;
                case "1":
                    _strReturn = "수취인불명";
                    break;
                case "2":
                    _strReturn = "이사불명";
                    break;
                case "3":
                    _strReturn = "주소불명";
                    break;
                case "4":
                    _strReturn = "장기폐문";
                    break;
                case "5":
                    _strReturn = "수취거절";
                    break;
                case "6":
                    _strReturn = "인터넷방송";
                    break;
                case "7":
                    _strReturn = "수취인부재";
                    break;
                case "8":
                    _strReturn = "직장퇴사";
                    break;
                case "9":
                    _strReturn = "근무지변경";
                    break;
                case "10":
                    _strReturn = "해외근무";
                    break;
                case "11":
                    _strReturn = "이민";
                    break;
                case "12":
                    _strReturn = "본인사망";
                    break;
                case "13":
                    _strReturn = "직장수취거절";
                    break;
                case "14":
                    _strReturn = "3회이상반송";
                    break;
                case "15":
                    _strReturn = "휴직";
                    break;
                case "16":
                    _strReturn = "군입대";
                    break;
                case "17":
                    _strReturn = "신청인부인";
                    break;
                case "18":
                    _strReturn = "부도";
                    break;
                case "19":
                    _strReturn = "성명오기";
                    break;
                case "20":
                    _strReturn = "우편번호착오";
                    break;
                case "21":
                    _strReturn = "주소미기재";
                    break;
                case "22":
                    _strReturn = "출장";
                    break;
                case "23":
                    _strReturn = "폐업";
                    break;
                case "24":
                    _strReturn = "재직사실무";
                    break;
                case "25":
                    _strReturn = "카드발급부정";
                    break;
                case "26":
                    _strReturn = "사고카드의심";
                    break;
                case "27":
                    _strReturn = "외근";
                    break;
                case "28":
                    _strReturn = "기재미비";
                    break;
                case "29":
                    _strReturn = "배달중민원";
                    break;
                case "30":
                    _strReturn = "카드회수";
                    break;
                case "31":
                    _strReturn = "신분증기재거부";
                    break;
                case "32":
                    _strReturn = "연락불가";
                    break;
                case "33":
                    _strReturn = "전화번호오류";
                    break;
                case "34":
                    _strReturn = "전환미동의";
                    break;
                case "88":
                    _strReturn = "불가지역";
                    break;
                case "99":
                    _strReturn = "기타";
                    break;

            }
            return _strReturn;
        }
        #endregion

        private static string GetSendCode(string value, string value2)
        {
            if (value.ToLower() == "b" || value.ToLower() == "a" || value.ToLower() == "g" ||
                value.ToLower() == "h" || value.ToLower() == "w" || value.ToLower() == "x" ||
                value.ToLower() == "c")
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
