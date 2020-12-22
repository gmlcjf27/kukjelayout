using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace Day_Close_Down
{
	public class CONVERT
	{
		//기본 인코딩 설정
		private static string strEncoding = "ks_c_5601-1987";
        private static string strCardTypeID = "999";
        private static string strCardTypeName = "통합일일마감";

        //제휴사코드 반환
        public static string GetCardType(string path)
        {
            string _strReturn = "";

            return _strReturn;
        }

		//현 DLL의 카드 타입 코드 반환
		public static string GetCardTypeID() 
        {
            int _iReturn = 0;
            FormSelectReceive _f = new FormSelectReceive();
            if (_f.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                _iReturn = _f.GetSelected;
            }
            //카드사 대분류 코드를 사용
            switch (_iReturn)
            {
                case 1:
                    strCardTypeID = "003";
                    break;
                case 2:
                    strCardTypeID = "004";
                    break;
                //하나SK
                case 3:
                    strCardTypeID = "005";
                    break;
                //하나
                case 4:
                    strCardTypeID = "006";
                    break;
                //롯데
                case 5:
                    strCardTypeID = "015";
                    break;
                //농협
                case 6:
                    strCardTypeID = "089";
                    break;
                //현대
                //case 7:
                //    strCardTypeID = "016";
                //    break;
                //국민
                case 8:
                    strCardTypeID = "002";
                    break;
                //현대캐피탈
                //case 9:
                //    strCardTypeID = "007";
                //    break;
                default:
                    strCardTypeID = "";
                    break;
            }
			return strCardTypeID;
		}

		//현 DLL의 카드 타입명 반환
		public static string GetCardTypeName() {
			return strCardTypeName;
		}

        //일일마감자료 간소화
        public static string ConvertResult(DataTable dtable, string fileName)
        {
            string _strReturn = null;

            //카드사 대분류 코드를 사용
            switch (GetStringAsLength(strCardTypeID, 3, true, ' '))
            {
                case "003": _strReturn = ConvertReceiveType1(dtable, fileName); break;
                case "004": _strReturn = ConvertReceiveType2(dtable, fileName); break;
                //하나SK
                case "005": _strReturn = ConvertReceiveType3(dtable, fileName); break;
                //하나
                case "006": _strReturn = ConvertReceiveType4(dtable, fileName); break;
                //롯데
                case "015": _strReturn = ConvertReceiveType5(dtable, fileName); break;
                //농협
                case "089": _strReturn = ConvertReceiveType6(dtable, fileName); break;
                //현대
                //case "016": _strReturn = ConvertReceiveType7(dtable, fileName); break;
                //국민
                case "002": _strReturn = ConvertReceiveType8(dtable, fileName); break;
                //현대캐피탈
                //case "007": _strReturn = ConvertReceiveType9(dtable, fileName); break;
                default:
                    _strReturn = "";
                    break; ;
            }
            
            return _strReturn;
        }
        //일일마감자료
        public static string ConvertResultDay(System.Data.DataTable dtable, string fileName)
        {
            return ConvertResult(dtable, fileName);
        }
        //신한 일일자료생성
        private static string ConvertReceiveType1(System.Data.DataTable dtable, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	        //기본 인코딩	
            StreamWriter _sw00 = null, _sw01 = null, _sw02 = null, _sw03 = null, _sw04 = null, _sw11 = null, _sw20 = null;				//파일 쓰기 스트림
            int i = 0, iCnt = 0;
            StringBuilder _strLine = new StringBuilder("");
            string _strReturn = "", strStatus = "", strClient_express_code = "", strReceiver_code = "", strToDay = "";
            string strCard_in_date_chk = "";
            DataRow[] _drs = null;
            try
            {
                strToDay = DateTime.Now.ToString("yyyyMMdd").Substring(4,4);

                _sw01 = new StreamWriter(fileName + "KJ" + strToDay + "일반일일.txt", true, _encoding);
                _sw02 = new StreamWriter(fileName + "KJ" + strToDay + "동의서일일.txt", true, _encoding);
                _sw03 = new StreamWriter(fileName + "KE" + strToDay + "LL", true, _encoding);
                _sw04 = new StreamWriter(fileName + "SHG" + strToDay + "LL", true, _encoding);
                _sw11 = new StreamWriter(fileName + "KJ" + strToDay + "LL(NEW)", true, _encoding);

                _drs = dtable.Select("", "delivery_result_editdate");

                //헤더 부분
                _sw11.WriteLine(GetStringAsLength("HDKJ" + DateTime.Now.ToString("yyyyMMdd"), 300, true, ' '));

                for (i = 0; i < _drs.Length; i++)
                {
                    strStatus = _drs[i]["card_delivery_status"].ToString();
                    strClient_express_code = _drs[i]["client_express_code"].ToString();
                    strReceiver_code = _drs[i]["receiver_code"].ToString();
                    strCard_in_date_chk = String.Format("{0:yyyyMMdd}", dtable.Rows[i]["card_in_date"]);
                    DateTime CardInDate = DateTime.Parse(_drs[i]["card_in_date"].ToString());
                    DateTime dt_date = DateTime.Parse("2014-10-21");

                    #region 구마감 배송
                    if (strStatus == "1" )
                    {
                        _strLine = new StringBuilder(GetStringAsLength(_drs[i]["card_number"].ToString(), 12, true, ' '));
                        _strLine.Append(GetStringAsLength(_drs[i]["card_brand_code"].ToString(), 1, true, ' '));
                        _strLine.Append(GetStringAsLength("", 3, true, ' '));

                        if (_drs[i]["receiver_code_change"].ToString() == "001" || strReceiver_code == "01")
                        {
                            _strLine.Append(GetStringAsLength("Y1", 2, true, ' '));
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength("Y2", 2, true, ' '));
                        }

                        _strLine.Append(GetStringAsLength(RemoveDash(_drs[i]["card_delivery_date"].ToString()), 8, true, ' '));

                        //2014.04.16 태희철 수정
                        //if (_drs[i]["receiver_SSN"].ToString().Length == 0)
                        //{
                        //    _strLine.Append(GetStringAsLength("", 6, true, '0'));
                        //    _strLine.Append(GetStringAsLength("", 8, true, ' '));
                        //}
                        //else
                        //{
                        //    _strLine.Append(GetStringAsLength(Convert_SH_SSN(_drs[i]["receiver_SSN"].ToString().Replace("x", "0")), 14, true, ' '));
                        //}

                        //민증번호
                        if (_drs[i]["card_result_status"].ToString() == "61")
                        {
                            //2014.10.21 태희철 수정 10월21일 인수분과 구분
                            if (CardInDate < dt_date)
                            {
                                _strLine.Append(GetStringAsLength(_drs[i]["customer_ssn"].ToString().Substring(7, 3), 14, true, ' '));
                            }
                            else
                            {
                                _strLine.Append(GetStringAsLength(_drs[i]["customer_ssn"].ToString().Substring(2, 4), 14, true, ' '));
                            }
                            //_strLine.Append(GetStringAsLength(_drs[i]["customer_ssn"].ToString().Substring(7, 3), 14, true, ' '));
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength(Convert_SH_SSN(_drs[i]["receiver_SSN"].ToString().Replace("x", "0")), 14, true, '0'));
                        }

                        _strLine.Append(GetStringAsLength(_drs[i]["receiver_tel"].ToString(), 15, true, ' '));
                        if (_drs[i]["card_issue_type"].ToString() == "5")
                        {
                            _strLine.Append(GetStringAsLength(RemoveDash(_drs[i]["client_send_date"].ToString()), 8, true, ' '));
                            _strLine.Append(GetStringAsLength("", 5, false, ' '));
                            _strLine.Append(GetStringAsLength(RemoveDash(_drs[i]["card_in_date"].ToString()), 8, true, ' '));
                            _strLine.Append(GetStringAsLength("", 6, true, ' '));
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength(RemoveDash(_drs[i]["client_register_date"].ToString()), 8, true, ' '));
                            _strLine.Append(GetStringAsLength(_drs[i]["client_number"].ToString(), 5, false, '0'));
                            _strLine.Append(GetStringAsLength(RemoveDash(_drs[i]["client_quick_work_date"].ToString()), 8, true, ' '));
                            _strLine.Append(GetStringAsLength(_drs[i]["client_send_number"].ToString(), 6, true, ' '));
                        }

                        _strLine.Append(GetStringAsLength("", 1, true, ' '));
                        _strLine.Append(GetStringAsLength(_drs[i]["receiver_name"].ToString(), 19, true, ' '));
                        _strLine.Append(GetStringAsLength(_drs[i]["receiver_code_change"].ToString(), 3, false, ' '));
                        _strLine.Append(GetStringAsLength(" ", 1, true, ' '));

                        _strLine.Append(GetStringAsLength("Y", 1, true, ' '));
                        _strLine.Append(GetStringAsLength("Y", 1, true, ' '));
                        _strLine.Append(GetStringAsLength("Y", 1, true, ' '));

                        _strLine.Append(GetStringAsLength("", 9, true, ' '));

                        //일반
                        if (strClient_express_code == "2002")
                        {
                            _sw01.WriteLine(_strLine.ToString());    
                        }
                        //동의서
                        else if (strClient_express_code == "2120")
                        {
                            _sw02.WriteLine(_strLine.ToString());    
                        }
                        //긴급
                        else if (strClient_express_code == "2005")
                        {
                            _sw03.WriteLine(_strLine.ToString());    
                        }
                        //GIFT
                        else if (strClient_express_code == "2126")
                        {
                            //등기제거
                            if (strReceiver_code != "98")
                            {
                                _sw04.WriteLine(_strLine.ToString());
                            }   
                        }
                        //기타
                        else
                        {
                            _sw00 = new StreamWriter(fileName + ".기타", true, _encoding);
                            _sw00.WriteLine(_strLine.ToString());
                            _sw00.Close();
                        }
                    }
                    #endregion
                    //2013.07.25 구일일마감 끝[E]

                    //2013.07.25 신마감 시작[S]
                    #region 신마감
                    if (strStatus == "1") //배송
                    {
                        _strLine = new StringBuilder("DT"); //시작코드
                        //카드번호
                        _strLine.Append(GetStringAsLength(_drs[i]["card_number"].ToString().Replace("-", ""), 16, true, ' '));

                        if (_drs[i]["receiver_code_change"].ToString() == "001" || _drs[i]["receiver_code"].ToString() == "01")
                        {
                            _strLine.Append(GetStringAsLength("Y1", 2, true, ' '));
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength("Y2", 2, true, ' '));
                        }
                        //전달일자
                        _strLine.Append(GetStringAsLength(_drs[i]["card_delivery_date"].ToString().Replace("-", ""), 8, true, ' '));
                        //민증번호
                        //_strLine.Append(GetStringAsLength(_drs[i]["receiver_SSN"].ToString().Replace("-", "").Replace("x", ""), 14, true, ' '));
                        //민증번호
                        if (_drs[i]["card_result_status"].ToString() == "61")
                        {
                            //2014.10.21 태희철 수정 10월21일 인수분과 구분
                            if (CardInDate < dt_date)
                            {
                                _strLine.Append(GetStringAsLength(_drs[i]["customer_ssn"].ToString().Substring(7, 3), 14, true, ' '));
                            }
                            else
                            {
                                _strLine.Append(GetStringAsLength(_drs[i]["customer_ssn"].ToString().Substring(2, 4), 14, true, ' '));
                            }
                            //_strLine.Append(GetStringAsLength(_drs[i]["customer_ssn"].ToString().Substring(7, 3), 14, true, ' '));
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength(Convert_SH_SSN(_drs[i]["receiver_SSN"].ToString().Replace("x", "0")), 14, true, '0'));
                        }
                        //전화번호
                        _strLine.Append(GetStringAsLength(_drs[i]["receiver_tel"].ToString().Replace("-", ""), 15, true, ' '));

                        //제작일자
                        if (_drs[i]["client_register_date"].ToString() == "")
                        {
                            _strLine.Append(GetStringAsLength(_drs[i]["client_send_date"].ToString().Replace("-", ""), 8, true, ' '));
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength(_drs[i]["client_register_date"].ToString().Replace("-", ""), 8, true, ' '));
                        }
                        //제작순번
                        _strLine.Append(GetStringAsLength(_drs[i]["client_number"].ToString(), 5, true, ' '));

                        //특송접수일자
                        if (_drs[i]["client_quick_work_date"].ToString() == "")
                            _strLine.Append(GetStringAsLength(_drs[i]["card_in_date"].ToString().Replace("-", ""), 8, true, ' '));
                        else
                            _strLine.Append(GetStringAsLength(_drs[i]["client_quick_work_date"].ToString().Replace("-", ""), 8, true, ' '));
                        //특송접수번호
                        _strLine.Append(GetStringAsLength(_drs[i]["client_send_number"].ToString(), 6, true, ' '));
                        //수령인명
                        _strLine.Append(GetStringAsLength(_drs[i]["receiver_name"].ToString(), 40, true, ' '));
                        //관계코드 - 은행 요청 코드
                        _strLine.Append(GetStringAsLength(_drs[i]["receiver_code_change"].ToString(), 3, true, ' '));
                        _strLine.Append(GetStringAsLength("", 1, true, ' '));                           //예비
                        _strLine.Append(GetStringAsLength(ConvertAgree(_drs[i]["card_agree1"].ToString()), 1, true, ' '));
                        _strLine.Append(GetStringAsLength(ConvertAgree(_drs[i]["card_agree2"].ToString()), 1, true, ' '));
                        _strLine.Append(GetStringAsLength(ConvertAgree(_drs[i]["card_agree3"].ToString()), 1, true, ' '));
                        //특송발송카드 BIN구분코드
                        _strLine.Append(GetStringAsLength(_drs[i]["card_client_no_1"].ToString(), 2, true, ' '));
                        //제휴사코드
                        _strLine.Append(GetStringAsLength(_drs[i]["client_express_code"].ToString(), 4, true, ' '));
                        _strLine.Append(GetStringAsLength("", 163, true, ' '));      //예비

                        //동의서 중 대리수령 발생 시 확인 처리
                        if (strClient_express_code == "2120" && strReceiver_code != "01")
                        {
                            _sw20 = new StreamWriter(fileName + ".대리수령", true, _encoding);
                            _sw20.WriteLine(_strLine.ToString());
                            _sw20.Close();
                        }
                        else
                        {                            
                            //신한기프트 등기 제거
                            if (strClient_express_code == "2126" && strReceiver_code == "98")
                            {
                                ;
                            }
                            else
                            {
                                iCnt++;
                                _sw11.WriteLine(_strLine.ToString());
                            }
                        }
                    }
                    #endregion
                }

                //2013.07.22 태희철 수정 [S] 신마감사용
                _strLine = new StringBuilder(GetStringAsLength("TR" + GetStringAsLength(iCnt.ToString(), 11, false, '0'), 300, true, ' '));
                _sw11.WriteLine(_strLine.ToString());
                //2013.07.22 태희철 수정 [E] 신마감사용
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

                if (_sw11 != null) _sw11.Close();
                if (_sw20 != null) _sw20.Close();
            }
            return _strReturn;

        }

        //신한 동의서 입고 데이터
        public static void ConvertRecipt_SH_In_data(DataTable dtable, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            StreamWriter _sw00 = null, _sw01 = null;		//파일 쓰기 스트림            
            StringBuilder _strLine = new StringBuilder("");
            string _strStatus = "";
            string tempday = DateTime.Now.ToString("yyyyMMdd");
            string strTemp = "|";
            string[] strECC_Code_Arrey = null;
            string strECC_Code = "";
            int i=0, itotcnt = 0;

            try
            {   
                _sw01 = new StreamWriter(fileName + tempday + "_BPR_KUKJE", true, _encoding);

                _strLine = new StringBuilder(GetStringAsLength("H", 1, true, ' ') + strTemp);
                _strLine.Append(GetStringAsLength(tempday,8,true,' ') + strTemp);
                _strLine.Append(GetStringAsLength("2120",4,true,' '));
                _strLine.Append(GetStringAsLength("",285,true,' '));

                _sw01.WriteLine(_strLine.ToString());

                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    itotcnt++;
                    _strLine = new StringBuilder("D|");

                    strECC_Code_Arrey = dtable.Rows[i]["card_cooperation1"].ToString().Split('<');
                    strECC_Code = strECC_Code_Arrey[0];

                    _strLine.Append(GetStringAsLength(strECC_Code, 35, true, ' '));
                    _strLine.Append("|");
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_name"].ToString(), 40, true, ' '));
                    _strLine.Append("|");
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_ssn"].ToString().Replace("x", "").Replace("*", ""), 20, true, ' '));
                    _strLine.Append(GetStringAsLength("", 201, true, ' '));

                    _sw01.WriteLine(_strLine.ToString());
                }
                _strLine = new StringBuilder(GetStringAsLength("T", 1, true, ' ') + strTemp);
                _strLine.Append(GetStringAsLength(itotcnt.ToString(), 11, true, ' '));
                _strLine.Append(GetStringAsLength("", 287, true, ' '));

                _sw01.WriteLine(_strLine.ToString());
            }
            catch (Exception)
            {
                MessageBox.Show(string.Format("{0}번째 동의서 입고 데이터 생성 중 오류", i));
            }
            finally
            {
                if (_sw01 != null) _sw01.Close();
            }
        }


        //삼성 일일자료생성
        private static string ConvertReceiveType2(System.Data.DataTable dtable, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            StreamWriter _sw01 = null;                                  		//파일 쓰기 스트림            
            int i = 0, iCnt = 0, iCnt2 = 0;
            StringBuilder _strLine = new StringBuilder("");
            string _strReturn = "", _strStatus = "", strCustomerSSN_type="", strCard_type_detail = "";
            string tempday = DateTime.Now.ToString("MMdd");
            try
            {
                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    _strStatus = dtable.Rows[i]["card_delivery_status"].ToString();
                    strCard_type_detail = dtable.Rows[i]["card_type_detail"].ToString();

                    if (_strStatus == "1")
                    {
                        //발급일
                        _strLine = new StringBuilder(GetStringAsLength(RemoveDash(dtable.Rows[i]["client_send_date"].ToString()), 8, true, ' '));
                        //배송업체코드
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_express_code"].ToString(), 2, true, ' '));
                        //일련번호
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_send_number"].ToString(), 7, true, ' '));
                        //카드번호
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_number"].ToString().Replace("x", "*"), 16, true, ' '));
                        //배송업체코드
                        _strLine.Append("04");
                        //대리수령인 코드 변경
                        _strLine.Append(GetStringAsLength(SM_ConvReceiver_code(dtable.Rows[i]["receiver_code"].ToString()), 2, true, ' '));
                        //고객주민번호
                        strCustomerSSN_type = dtable.Rows[i]["customer_SSN"].ToString().Replace("x", "*");

                        //if (strCustomerSSN_type.Substring(7, 1) == "*")
                        //{
                        //    ;
                        //}
                        //else
                        //{
                        //    strCustomerSSN_type = strCustomerSSN_type.Substring(0, 7) + "******" + strCustomerSSN_type.Substring(6, 3) + "****";
                        //}

                        strCustomerSSN_type = strCustomerSSN_type.Substring(0, 7) + "******";

                        _strLine.Append(GetStringAsLength(strCustomerSSN_type, 13, true, '*'));
                        //고객명
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_name"].ToString(), 8, true, ' '));
                        //배송일
                        _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[i]["card_delivery_date"].ToString()), 8, true, ' '));
                        //수령인명
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_name"].ToString(), 8, true, ' '));
                        //수령인주민번호
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_SSN"].ToString().Replace("x", "*"), 13, true, '*'));

                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_quick_seq"].ToString(), 6, true, ' '));

                        //삼성동의서 구분
                        if (strCard_type_detail.Substring(0, 5) == "00421")
                        {
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_client_no_1"].ToString(), 20, true, ' '));    
                        }
                        
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_no"].ToString(), 11, true, ' '));

                        //삼성동의서 구분
                        if (strCard_type_detail.Substring(0,5) == "00421")
                        {
                            iCnt2++;
                            _sw01 = new StreamWriter(fileName + "국제일일마감(동의서)" + tempday + ".dat", true, _encoding);
                        }
                        else
                        {
                            iCnt++;
                            _sw01 = new StreamWriter(fileName + "국제일일마감(일반)" + tempday + ".dat", true, _encoding);
                        }
                        _sw01.WriteLine(_strLine.ToString());
                        _sw01.Close();
                    }
                }
                _strReturn = string.Format("일반 : {0}건 / 동의서 : {1}건 / 합 : {2}건의 인계데이터 다운 완료", iCnt, iCnt2, iCnt + iCnt2);
            }
            catch (Exception)
            {
                _strReturn = string.Format("{0}번째 데이터 인계 중 오류 발생", i + 1);
            }
            finally
            {
                if (_sw01 != null) _sw01.Close();
            }
            return _strReturn;
        }

        //삼성 동의서 입고 데이터
        public static void ConvertRecipt_In_data(DataTable dtable, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            StreamWriter _sw01 = null;		//파일 쓰기 스트림            
            int i = 0;
            StringBuilder _strLine = new StringBuilder("");
            string _strReturn = "", _strStatus = "";
            string strCustomerSSN_type = null;
            string tempday = DateTime.Now.ToString("MMdd");

            try
            {
                _sw01 = new StreamWriter(fileName + "국제in" + tempday + ".dat", true, _encoding);

                for (int k = 0; k < dtable.Rows.Count; k++)
                {
                    _strStatus = dtable.Rows[k]["card_delivery_status"].ToString();
                    _strLine = new StringBuilder(GetStringAsLength(RemoveDash(dtable.Rows[k]["client_send_date"].ToString()), 8, true, ' '));
                    _strLine.Append(GetStringAsLength(dtable.Rows[k]["client_express_code"].ToString(), 2, false, ' '));
                    _strLine.Append(GetStringAsLength(dtable.Rows[k]["client_send_number"].ToString(), 7, true, ' '));

                    _strLine.Append(GetStringAsLength(dtable.Rows[k]["card_number"].ToString().Replace("x", "*"), 16, true, ' '));

                    //2012.10.30 태희철 수정
                    _strLine.Append("04");

                    if (dtable.Rows[k]["card_result_status"].ToString() == "61")
                    {
                        _strLine.Append(GetStringAsLength(SM_ConvReceiver_code(dtable.Rows[k]["receiver_code"].ToString()), 2, true, ' '));
                    }

                    //2012.12.26 태희철 수정[E] 대리수령인 코드 변경

                    //2011-12-20 태희철 수정[S]
                    strCustomerSSN_type = dtable.Rows[k]["customer_SSN"].ToString().Replace("x", "*");

                    //if (strCustomerSSN_type.Substring(7, 1) == "*")
                    //{
                    //    ;
                    //}
                    //else
                    //{
                    //    strCustomerSSN_type = strCustomerSSN_type.Substring(0, 3) + "***" + strCustomerSSN_type.Substring(6, 3) + "****";
                    //}
                    strCustomerSSN_type = strCustomerSSN_type.Substring(0, 7) + "******";

                    _strLine.Append(GetStringAsLength(strCustomerSSN_type, 13, true, '*'));

                    _strLine.Append(GetStringAsLength(dtable.Rows[k]["customer_name"].ToString(), 8, true, ' '));
                    _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[k]["card_delivery_date"].ToString()), 8, true, ' '));
                    _strLine.Append(GetStringAsLength(dtable.Rows[k]["receiver_name"].ToString(), 8, true, ' '));

                    if (_strStatus == "1")
                    {
                        //if (dtable.Rows[k]["card_result_status"].ToString() == "61")
                        //{
                        //    _strLine.Append(GetStringAsLength(strCustomerSSN_type, 13, true, '*'));
                        //}
                        //else
                        //{
                        //    _strLine.Append(GetStringAsLength(dtable.Rows[k]["receiver_SSN"].ToString().Replace("x", "*"), 13, true, '*'));
                        //}
                        _strLine.Append(GetStringAsLength(dtable.Rows[k]["receiver_SSN"].ToString().Replace("x", "*"), 13, true, '*'));
                    }

                    _strLine.Append(GetStringAsLength(dtable.Rows[k]["client_quick_seq"].ToString(), 6, true, ' '));

                    _strLine.Append(GetStringAsLength(dtable.Rows[k]["card_client_no_1"].ToString(), 20, true, ' '));
                    _strLine.Append(GetStringAsLength(dtable.Rows[k]["customer_no"].ToString(), 11, true, ' '));

                    if (_strStatus == "1")
                    {
                        _sw01.WriteLine(_strLine.ToString());
                    }

                }
            }
            catch (Exception)
            {
                MessageBox.Show("동의서 입고 데이터 생성 중 오류");
            }
            finally
            {
                if (_sw01 != null) _sw01.Close();
            }
        }

        //하나SK 일일자료생성
        private static string ConvertReceiveType3(System.Data.DataTable dtable, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            StreamWriter _sw01 = null;

            StringBuilder _strLine = new StringBuilder("");
            string _strReturn = "", _strStatus = "";					//파일 쓰기 스트림
            int i = 0, iCnt = 0;

            try
            {
                string temp_time = DateTime.Now.ToShortDateString().Replace("-", "").Substring(2, 6);

                _sw01 = new StreamWriter(fileName + "KUKJ102." + temp_time + ".01", true, _encoding); //배송

                _strLine.Append(GetStringAsLength("H", 1, true, ' '));
                _strLine.Append(GetStringAsLength(DateTime.Now.ToShortDateString().Replace("-", ""), 8, true, ' '));
                _strLine.Append(GetStringAsLength("", 293, true, ' '));
                
                _sw01.WriteLine(_strLine.ToString());

                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    _strStatus = dtable.Rows[i]["card_delivery_status"].ToString();

                    if (_strStatus == "1")
                    {
                        //업체코드
                        _strLine = new StringBuilder(GetStringAsLength("D01", 3, true, ' '));
                        //파일구분(1:일반 2:영업점)
                        _strLine.Append(GetStringAsLength("1", 1, true, ' '));
                        //발급일자
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_send_date"].ToString().Replace("-", ""), 8, true, ' '));
                        //발급 번호
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_send_number"].ToString().Replace("-", ""), 8, true, ' '));
                        //진행코드
                        _strLine.Append(GetStringAsLength(delivery_stat(dtable.Rows[i]), 2, true, ' '));
                        //배송지사명
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["branch_name"].ToString(), 20, true, ' '));
                        //배송담당자명
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["career"].ToString(), 20, true, ' '));
                        //1차출고일자
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["delivery_out_date1"].ToString().Replace("-", ""), 8, true, ' '));
                        //1차 반송코드
                        _strLine.Append(GetStringAsLength(return_reason(dtable.Rows[i]["delivery_return_reason1"].ToString()), 2, true, ' '));
                        //2차출고일자
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["delivery_out_date2"].ToString().Replace("-", ""), 8, true, ' '));
                        //1차 반송코드
                        _strLine.Append(GetStringAsLength(return_reason(dtable.Rows[i]["delivery_return_reason2"].ToString()), 2, true, ' '));
                        //3차출고일자
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["delivery_out_date3"].ToString().Replace("-", ""), 8, true, ' '));
                        //1차 반송코드
                        _strLine.Append(GetStringAsLength(return_reason(dtable.Rows[i]["delivery_return_reason3"].ToString()), 2, true, ' '));
                        //수취인명
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_name"].ToString(), 20, true, ' '));
                        //수령인관계코드
                        _strLine.Append(GetStringAsLength(receiver_code(dtable.Rows[i]["receiver_code"].ToString()), 2, true, ' '));
                        //수취인민증 
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_SSN"].ToString().Replace("-", ""), 13, true, ' '));
                        //카드수령일자(결과등록일)
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_delivery_date"].ToString().Replace("-", ""), 8, true, ' '));
                        //영업점코드
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_bank_id"].ToString().Replace("-", ""), 4, true, ' '));
                        //발급구분
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_issue_type_code"].ToString().Replace("-", ""), 2, true, ' '));
                        //긴급구분                    
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["receipt_number"].ToString().Replace("|", ""), 1, true, ' '));
                        //포장구분
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_express_code"].ToString().Replace("-", ""), 1, true, ' '));
                        //대면여부 
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_enterprise_code"].ToString().Replace("-", ""), 1, true, ' '));
                        //카드매수
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_count"].ToString().Replace("-", ""), 4, true, ' '));
                        //카드매수
                        _strLine.Append(GetStringAsLength(" ", 1, true, ' '));
                        //배숑결과주소
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_address"].ToString().Replace("-", ""), 100, true, ' ')); 
                        //filter
                        _strLine.Append(GetStringAsLength("", 53, true, ' '));

                        iCnt++;
                        _sw01.WriteLine(_strLine.ToString());
                    }
                }

                _strLine = new StringBuilder(GetStringAsLength("T", 1, true, ' '));
                _strLine.Append(GetStringAsLength(iCnt.ToString(), 8, true, ' '));
                _strLine.Append(GetStringAsLength("", 293, true, ' '));

                _sw01.WriteLine(_strLine.ToString());

                _strReturn = string.Format("{0}건의 인계데이터 다운 완료", iCnt);
            }
            catch (Exception)
            {
                _strReturn = string.Format("{0}번째 데이터 인계 중 오류 발생", i + 1);
            }
            finally
            {
                if (_sw01 != null) _sw01.Close();
            }
            return _strReturn;
        }

        //하나 일일자료생성
        private static string ConvertReceiveType4(System.Data.DataTable dtable, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            StreamWriter _sw01 = null, _swdug = null;										//파일 쓰기 스트림
            int i = 0, iCnt = 0;
            StringBuilder _strLine = new StringBuilder("");
            string _strReturn = "", _strStatus = "";
            string _strCardNumber = "", _strFamilyNo = "", strCard_pt_code = "", strCard_kind = "", strCard_issue_detail_code = "";
            string strToDay= "";
            int _iAddCount = 0;
            string[] _strArFamilyNo = null;
            try
            {
                strToDay = DateTime.Now.ToString("MMdd");
                _sw01 = new StreamWriter(fileName + "bug" + strToDay + ".txt", true, _encoding);
                _swdug = new StreamWriter(fileName + "dug" + strToDay + ".txt", true, _encoding);

                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    _strStatus = dtable.Rows[i]["card_delivery_status"].ToString();

                    if (_strStatus == "1")
                    {
                        _strCardNumber = GetStringAsLength(dtable.Rows[i]["card_number"].ToString(), 16, true, ' ');
                        _iAddCount = int.Parse(dtable.Rows[i]["card_add_count"].ToString());
                        _strFamilyNo = GetStringAsLength(dtable.Rows[i]["family_customer_no"].ToString(), 16, true, ' ');
                        strCard_pt_code = dtable.Rows[i]["card_pt_code"].ToString();
                        strCard_kind = dtable.Rows[i]["card_kind"].ToString();
                        strCard_issue_detail_code = dtable.Rows[i]["card_issue_detail_code"].ToString();

                        //동의서
                        if (strCard_kind == "D")
                        {
                            ;
                        }
                        //긴급
                        else if (strCard_pt_code == "K")
                        {
                            strCard_kind = strCard_pt_code;
                        }
                        else if (strCard_issue_detail_code == "S")
                        {
                            strCard_kind = "D";
                        }

                        //dug 시작
                        //_strLine = new StringBuilder(GetStringAsLength(strCard_kind, 2, true, ' '));
                        _strLine = new StringBuilder(GetStringAsLength("D", 1, true, ' '));
                        _strLine.Append("{0}");
                        _strLine.Append(GetStringAsLength("", 3, true, ' '));
                        _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[i]["client_send_date"].ToString()), 8, true, ' '));
                        _strLine.Append(GetStringAsLength("2", 1, true, ' '));
                        _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[i]["delivery_out_date1"].ToString()), 8, true, ' '));
                        _strLine.Append(GetStringAsLength("", 2, true, ' '));
                        _strLine.Append(GetStringAsLength("", 8, true, ' '));
                        _strLine.Append(GetStringAsLength("", 2, true, ' '));
                        _strLine.Append(GetStringAsLength("", 8, true, ' '));
                        _strLine.Append(GetStringAsLength("", 2, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_SSN"].ToString(), 13, true, ' '));
                        _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[i]["card_in_date"].ToString()), 8, true, ' '));

                        _swdug.WriteLine(string.Format(_strLine.ToString(), _strCardNumber));

                        if (_iAddCount == 1)
                        {
                            _swdug.WriteLine(string.Format(_strLine.ToString(), GetStringAsLength(_strFamilyNo, 16, true, ' ')));
                        }
                        if (_iAddCount > 1)
                        {
                            _strArFamilyNo = _strFamilyNo.Split(new char['|']);
                            for (int j = 0; j < _iAddCount; j++)
                            {
                                _swdug.WriteLine(string.Format(_strLine.ToString(), GetStringAsLength(_strArFamilyNo[0], 16, true, ' ')));
                            }
                        }
                        //dug 끝

                        //배송결과 시작
                        //_strLine = new StringBuilder(GetStringAsLength(strCard_kind, 2, true, ' '));
                        _strLine = new StringBuilder(GetStringAsLength("D", 1, true, ' '));
                        _strCardNumber = GetStringAsLength(dtable.Rows[i]["card_number"].ToString(), 16, true, ' ');
                        _strLine.Append("{0}");
                        _strLine.Append(GetStringAsLength("", 3, true, ' '));
                        _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[i]["client_send_date"].ToString()), 8, true, ' '));
                        _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[i]["card_in_date"].ToString()), 8, true, ' '));
                        _strLine.Append(GetStringAsLength("", 7, true, ' '));

                        _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[i]["card_delivery_date"].ToString()), 8, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_name"].ToString(), 20, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_code_change"].ToString(), 12, true, ' '));

                        //2012-11-26 태희철 수정 본인수령이면 고객주민번호 출력
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_SSN"].ToString(), 13, true, ' '));
                        _strLine.Append(GetStringAsLength("2", 1, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_insert_type"].ToString(), 1, true, ' '));
                        _strLine.Append(GetStringAsLength("0", 1, true, ' '));
                        _strLine.Append(GetStringAsLength(strCard_kind, 1, true, ' '));
                        _strLine.Append(dtable.Rows[i]["card_year"].ToString());
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["degree_code"].ToString(), 5, false, '0'));

                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_SSN"].ToString(), 13, true, ' '));

                        if (dtable.Rows[i]["card_agree_code"].ToString() == "D")
                        {
                            _strLine.Append(GetStringAsLength("03", 2, true, ' '));
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength("", 2, true, ' '));
                        }
                        
                        _sw01.WriteLine(string.Format(_strLine.ToString(), _strCardNumber));

                        if (_iAddCount == 1)
                        {
                            _sw01.WriteLine(string.Format(_strLine.ToString(), GetStringAsLength(_strFamilyNo, 16, true, ' ')));
                        }
                        if (_iAddCount > 1)
                        {
                            _strArFamilyNo = _strFamilyNo.Split(new char['|']);
                            for (int j = 0; j < _iAddCount; j++)
                            {
                                _sw01.WriteLine(string.Format(_strLine.ToString(), GetStringAsLength(_strArFamilyNo[0], 16, true, ' ')));
                            }
                        }
                        iCnt++;
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
                if (_sw01 != null) _sw01.Close();
                if (_swdug != null) _swdug.Close();
            }
            return _strReturn;
        }

        //롯데 일일자료생성
        private static string ConvertReceiveType5(System.Data.DataTable dtable, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            StreamWriter _sw01 = null;                      					//파일 쓰기 스트림
            int i = 0, iCnt = 0;
            StringBuilder _strLine = new StringBuilder("");
            string _strReturn = "", _strStatus = "", strToDay = "";
            try
            {
                strToDay = DateTime.Now.ToString("yyyyMMdd").Substring(2, 6);
                _sw01 = new StreamWriter(fileName + "일일_02_" + strToDay, true, _encoding);

                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    _strStatus = dtable.Rows[i]["card_delivery_status"].ToString();

                    if (_strStatus == "1")
                    {
                        _strLine = new StringBuilder(GetStringAsLength(dtable.Rows[i]["card_number"].ToString(), 14, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_SSN"].ToString().Replace("x", "*"), 13, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_name"].ToString(), 40, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["family_name"].ToString(), 40, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["family_name2"].ToString(), 40, true, ' '));

                        // 2012-02-07 태희철 수령지 : 001=직장, 002=자택
                        // 등록 시 수령지 주소는 수령지, 비수령지 구분하나
                        // 전화번호는 자택card_tel1, 직장card_tel2로 등록 한다.
                        // 전화번호는 자택card_tel1, 직장card_tel2로 등록 한다.
                        if (dtable.Rows[i]["card_delivery_place_code"].ToString() == "110")
                        {
                            if (dtable.Rows[i]["card_zipcode2"].ToString() == "0")
                            {
                                _strLine.Append(GetStringAsLength("", 6, true, ' '));
                            }
                            else
                            {
                                _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_zipcode2"].ToString(), 6, true, ' '));
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
                                _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_zipcode"].ToString(), 6, true, ' '));
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
                                _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_zipcode"].ToString(), 6, true, ' '));
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
                                _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_zipcode2"].ToString(), 6, true, ' '));
                            }

                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_address2_local"].ToString() + " ", 60, true, ' '));
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_address2_detail"].ToString() + " ", 150, true, ' '));
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_tel2"].ToString(), 15, true, ' '));
                        }
                        // 2012-02-07 태희철 수령지 : 001=직장, 002=자택[E]

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
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_SSN"].ToString().Substring(0, 6) + "*******", 13, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_name"].ToString(), 40, true, ' '));

                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_code_change"].ToString().Replace("x", " "), 2, true, ' '));

                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_delivery_memo"].ToString(), 40, true, ' '));

                        //2011-11-18 태희철 수정 [S]
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

                        //배송업체구간
                        if (dtable.Rows[i]["card_agree_code"].ToString() == "Y")
                        {
                            _strLine.Append(GetStringAsLength("N", 1, true, ' '));
                            _strLine.Append(GetStringAsLength("N", 1, true, ' '));
                            _strLine.Append(GetStringAsLength("N", 1, true, ' '));
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength("Y", 1, true, ' '));
                            _strLine.Append(GetStringAsLength("Y", 1, true, ' '));
                            _strLine.Append(GetStringAsLength("Y", 1, true, ' '));
                        }

                        //카드사구간
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_issue_detail_code"].ToString(), 1, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_register_type"].ToString(), 2, true, ' '));

                        _strLine.Append("Y");
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_design_code"].ToString(), 3, true, '0'));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_number"].ToString(), 14, true, '0'));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["family_customer_no"].ToString(), 100, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_cooperation1"].ToString(), 100, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_cooperation2"].ToString(), 100, true, ' '));

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
                            _strLine.Append(GetStringAsLength(GetStringAsLength(dtable.Rows[i]["text1"].ToString(),100, true,' ') + "^" +
                        GetStringAsLength(dtable.Rows[i]["text2"].ToString(), 100, true, ' ') + "^" +
                        GetStringAsLength(dtable.Rows[i]["text3"].ToString(), 100, true, ' ') + "^" +
                        GetStringAsLength(dtable.Rows[i]["text4"].ToString(), 100, true, ' ') + "^" +
                        GetStringAsLength(dtable.Rows[i]["text5"].ToString(), 100, true, ' ') + "^" +
                        GetStringAsLength(dtable.Rows[i]["text6"].ToString(), 100, true, ' ') + "^" +
                        GetStringAsLength(dtable.Rows[i]["text7"].ToString(), 100, true, ' ') + "^" +
                        GetStringAsLength(dtable.Rows[i]["text8"].ToString(), 100, true, ' ') + "^" +
                        GetStringAsLength(dtable.Rows[i]["text9"].ToString(), 100, true, ' ') + "^" +
                        GetStringAsLength(dtable.Rows[i]["text10"].ToString(), 100, true, ' '), 1000, true, ' '));
                        }

                        _sw01.WriteLine(_strLine);
                        iCnt++;
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
                if (_sw01 != null) _sw01.Close();
            }
            return _strReturn;
        }

        //롯데주소변경, 민원소명 / 일일마감과 오류 확인을 위해 분리
        public static void ChangeAddress(DataTable dtable, DataTable dtable2, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            StreamWriter _sw03 = null;					//파일 쓰기 스트림
            int i = 0, j = 0;
            StringBuilder _strLine = new StringBuilder("");
            string strToDay = "";

            try
            {
                strToDay = DateTime.Now.ToString("yyyyMMdd").Substring(2, 6);
                _sw03 = new StreamWriter(fileName + "ADDR_02_" + strToDay, true, _encoding);

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
            catch (Exception)
            {
                MessageBox.Show("ADDR파일 생성 중 오류");
            }
            finally
            {
                if (_sw03 != null) _sw03.Close();
            }
        }

        //농협 일일자료생성
        private static string ConvertReceiveType6(System.Data.DataTable dtable, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            StreamWriter _sw01 = null;                                          //파일 쓰기 스트림
            int i = 0, iCnt = 0;
            StringBuilder _strLine = new StringBuilder("");
            string _strReturn = "", _strStatus = "", strCard_issue_detail_code = "";
            try
            {
                // temp : 2 일일 // 3 = 마감
                string tempday = DateTime.Now.ToString("yyyyMMdd");
                
                _sw01 = new StreamWriter(fileName + "KU0.11262." + tempday + ".01", true, _encoding);

                _strLine = new StringBuilder(GetStringAsLength("FH", 2, true, ' '));
                _strLine.Append(GetStringAsLength(tempday, 8, true, ' '));
                _strLine.Append(GetStringAsLength("", 325, true, ' '));

                _sw01.WriteLine(_strLine);

                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    _strStatus = dtable.Rows[i]["card_delivery_status"].ToString();
                    strCard_issue_detail_code = dtable.Rows[i]["card_issue_detail_code"].ToString();

                    if (_strStatus == "1")
                    {
                        _strLine = new StringBuilder(GetStringAsLength("FD", 2, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_send_number"].ToString(), 17, true, ' ')); //배송일련번호
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_express_code"].ToString(), 2, true, ' ')); //배송차수
                        _strLine.Append(GetStringAsLength("02", 2, true, ' '));                                             //업체코드

                        //2012-01-17 태희철 정리
                        //temp_0, temp_1, temp_2 : 총건수 1:배송, 2:반송
                        string _strStatusTemp = "00";  //배송 결과

                        if (dtable.Rows[i]["receiver_code"].ToString() == "98")
                        {
                            _strStatusTemp = "05";
                        }
                        else
                        {
                            _strStatusTemp = "00";
                        }

                        _strLine.Append(GetStringAsLength(_strStatusTemp, 2, true, ' '));
                        _strLine.Append(GetStringAsLength("2", 1, true, ' '));      //배송결과수신방법코드   
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_client_no_1"].ToString(), 16, true, ' ')); //카드번호

                        if (strCard_issue_detail_code == "3")
                        {
                            _strLine.Append("0"); // 재배송의경우 무조건 "0"
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_consented"].ToString(), 1, true, ' ')); //동의서징구 결과
                        }
                        
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_delivery_place_type"].ToString(), 1, true, ' ')); //수령지 

                        if (_strStatusTemp != "05")
                        {
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_delivery_date"].ToString().Replace("-", ""), 8, true, ' ')); //배송완료일
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_name"].ToString(), 40, true, ' ')); //수령인 이름
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_SSN"].ToString(), 13, true, ' '));   //수령인 민증
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["delivery_result_regdate"].ToString().Replace("-", ""), 8, true, ' ')); //배송완료일
                            _strLine.Append(GetStringAsLength("", 40, true, ' ')); //수령인 이름
                            _strLine.Append(GetStringAsLength("", 13, true, ' '));
                        }

                        if (_strStatusTemp != "05")
                        {
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_code_change"].ToString(), 2, true, ' '));
                            _strLine.Append(GetStringAsLength("", 2, true, ' '));
                            _strLine.Append(GetStringAsLength("", 14, true, ' '));
                            _strLine.Append(GetStringAsLength("", 6, true, ' '));
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength("", 2, true, ' '));
                            _strLine.Append(GetStringAsLength("", 2, true, ' '));
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_SSN"].ToString(), 14, true, ' '));
                            // 농협등기요금
                            //2012.11.20 태희철 수정 등기요금 1960 -> 2200
                            DateTime CardInDate = DateTime.Parse(dtable.Rows[i]["card_in_date"].ToString());
                            DateTime dtDong_date = DateTime.Parse("2014-01-29");

                            if (CardInDate < dtDong_date)
                            {
                                _strLine.Append(GetStringAsLength("2230", 6, true, ' '));
                            }
                            else
                            {
                                _strLine.Append(GetStringAsLength("2440", 6, true, ' '));
                            }
                        }

                        //_strLine.Append(GetStringAsLength("", 2, true, ' '));

                        if (dtable.Rows[i]["card_delivery_place_type"].ToString() == "5")
                        {
                            _strLine.Append(GetStringAsLength("card_zipcode", 6, true, ' '));
                            _strLine.Append(GetStringAsLength("card_address", 200, true, ' '));
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength("", 6, true, ' '));
                            _strLine.Append(GetStringAsLength("", 200, true, ' '));
                        }

                        iCnt++;
                        _sw01.WriteLine(_strLine);
                    }
                }
                _strLine = new StringBuilder(GetStringAsLength("FT", 2, true, ' '));
                _strLine.Append(GetStringAsLength(iCnt.ToString(), 8, false, '0'));
                _strLine.Append(GetStringAsLength("", 325, true, ' '));

                _sw01.WriteLine(_strLine);

                _strReturn = string.Format("{0}건의 인계데이터 다운 완료", iCnt);
            }
            catch (Exception)
            {
                _strReturn = string.Format("{0}번째 데이터 인계 중 오류 발생", i + 1);
            }
            finally
            {
                if (_sw01 != null) _sw01.Close();
            }
            return _strReturn;
        }

        //현대 일일자료생성
        //private static string ConvertReceiveType7(System.Data.DataTable dtable, string fileName)
        //{
        //    //return ConvertResult(dtable, fileName);
        //    Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
        //    StreamWriter _sw00 = null, _sw01 = null, _sw02 = null, _sw03 = null, _sw04 = null, _sw05 = null, _sw06 = null; //파일 쓰기 스트림
        //    int i = 0, iCnt = 0;
        //    StringBuilder _strLine = new StringBuilder("");
        //    string _strReturn = "", _strStatus = "", strClient_register_type = "", strToDay = "", strCustomer_ssn = "";
            
        //    try
        //    {
        //        strToDay = DateTime.Now.ToString("yyyyMMdd");

        //        _sw01 = new StreamWriter(fileName + "동의" + strToDay + "일일.dat", true, _encoding);
        //        _sw02 = new StreamWriter(fileName + "스마트" + strToDay + "일일_일반.dat", true, _encoding);
        //        _sw03 = new StreamWriter(fileName + "스마트" + strToDay + "일일_동의.dat", true, _encoding);
        //        _sw04 = new StreamWriter(fileName + "약식동의" + strToDay + "일일.dat", true, _encoding);
        //        _sw05 = new StreamWriter(fileName + "약식스마트동의" + strToDay + "일일.dat", true, _encoding);
        //        _sw06 = new StreamWriter(fileName + "일반" + strToDay + "일일_일반.dat", true, _encoding);
                
        //        for (i = 0; i < dtable.Rows.Count; i++)
        //        {
        //            _strStatus = dtable.Rows[i]["card_delivery_status"].ToString();
        //            strClient_register_type = dtable.Rows[i]["client_register_type"].ToString();
        //            strCustomer_ssn = dtable.Rows[i]["customer_SSN"].ToString();

        //            //결과코드
        //            if (_strStatus == "1")
        //            {
        //                _strLine = new StringBuilder(GetStringAsLength("1", 1, true, ' '));
        //                //발송번호
        //                _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_send_number"].ToString(), 17, true, ' '));

        //                //주민번호
        //                if (strCustomer_ssn == "xxxxxxxxxxxxx")
        //                {
        //                    _strLine.Append(GetStringAsLength("", 13, true, ' '));
        //                }
        //                else
        //                {
        //                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_SSN"].ToString(), 13, true, ' '));
        //                }
        //                //_strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_SSN"].ToString(), 13, true, ' '));
        //                //발송구분코드
        //                if (dtable.Rows[i]["card_issue_type_code"].ToString() == "")
        //                {
        //                    _strLine.Append(GetStringAsLength("1", 1, true, ' '));
        //                }
        //                else
        //                {
        //                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_issue_type_code"].ToString(), 1, true, ' '));
        //                }

        //                _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[i]["card_delivery_date"].ToString()), 8, true, ' '));
        //                //2012.10.23 태희철 배송시간 추가
        //                _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[i]["card_delivery_time"].ToString()), 4, true, ' '));
        //                _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_SSN"].ToString(), 13, true, 'x'));
        //                _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_code_change"].ToString().Replace("x", ""), 2, true, ' '));
        //                _strLine.Append(GetStringAsLength("", 1, true, ' '));

        //                //배송업체코드
        //                _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_express_code"].ToString(), 4, true, ' '));
        //                //카드번호
        //                _strLine.Append(GetStringAsLength("", 16, true, ' '));
        //                //동의여부
        //                _strLine.Append(GetStringAsLength("Y", 1, true, ' '));

        //                DateTime CardInDate = DateTime.Parse(dtable.Rows[i]["card_in_date"].ToString());
        //                DateTime dtDong_date = DateTime.Parse("2014-10-01");
        //                DateTime dtDong_date_other = DateTime.Parse("2014-11-11");

        //                //2012.08.29 현재 입회신청서변경이력 없음 변동있으므로 주의 요망
        //                //입회신청서변경이력
        //                if (strClient_register_type == "55" || strClient_register_type == "57")
        //                {
        //                    if (CardInDate < dtDong_date_other)
        //                    {
        //                        _strLine.Append(GetStringAsLength("ACCA02", 8, true, ' '));
        //                    }
        //                    else
        //                    {
        //                        _strLine.Append(GetStringAsLength("ACCA03", 8, true, ' '));
        //                    }
        //                }
        //                else
        //                {
        //                    if (CardInDate < dtDong_date)
        //                    {
        //                        _strLine.Append(GetStringAsLength("FFTA32", 8, true, ' '));
        //                    }
        //                    else
        //                    {
        //                        _strLine.Append(GetStringAsLength("FFTA33", 8, true, ' '));
        //                    }
        //                }

        //                if (dtable.Rows[i]["card_branch"].ToString() == "012")
        //                {
        //                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_SSN"].ToString(), 13, true, 'x'));//등기번호에 민증찍기
        //                }
        //                else
        //                {
        //                    _strLine.Append(GetStringAsLength("", 13, true, ' ')); //등기번호
        //                }

        //                _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_name"].ToString(), 16, true, ' '));

        //                //동의서
        //                if (strClient_register_type == "45")
        //                {
        //                    iCnt++;
        //                    _sw01.WriteLine(_strLine.ToString());    
        //                }
        //                //스마트
        //                else if (strClient_register_type == "37")
        //                {
        //                    iCnt++;
        //                    //스마트SSN
        //                    if (dtable.Rows[i]["receiver_SSN"].ToString().Trim() != dtable.Rows[i]["receiver_SSN_org"].ToString().Trim()
        //                    && dtable.Rows[i]["card_result_status"].ToString() == "61")
        //                    {
        //                        ;
        //                    }
        //                    else
        //                    {
        //                        _sw02.WriteLine(_strLine.ToString());
        //                    }
                            
        //                }
        //                //스마트동의서
        //                else if (strClient_register_type == "47")
        //                {
        //                    iCnt++;
        //                    _sw03.WriteLine(_strLine.ToString());
        //                }
        //                //약식동의
        //                else if (strClient_register_type == "55")
        //                {
        //                    iCnt++;
        //                    _sw04.WriteLine(_strLine.ToString());
        //                }
        //                //스마트약식동의
        //                else if (strClient_register_type == "57")
        //                {
        //                    iCnt++;
        //                    _sw05.WriteLine(_strLine.ToString());
        //                }
        //                //일반
        //                else if (strClient_register_type == "30")
        //                {
        //                    iCnt++;
        //                    _sw06.WriteLine(_strLine.ToString());
        //                }
        //                else
        //                {
        //                    _sw00 = new StreamWriter(fileName + ".기타", true, _encoding);
        //                    _sw00.WriteLine(_strLine.ToString());
        //                    _sw00.Close();
        //                }
        //            }
        //        }

        //        _strReturn = string.Format("{0}건의 인계데이터 다운 완료", iCnt);
        //    }
        //    catch (Exception)
        //    {
        //        _strReturn = string.Format("{0}번째 데이터 인계 중 오류 발생", i + 1);
        //    }
        //    finally
        //    {
        //        if (_sw00 != null) _sw00.Close();
        //        if (_sw01 != null) _sw01.Close();
        //        if (_sw02 != null) _sw02.Close();
        //        if (_sw03 != null) _sw03.Close();
        //        if (_sw04 != null) _sw04.Close();
        //        if (_sw05 != null) _sw05.Close();
        //        if (_sw06 != null) _sw06.Close();
        //    }
        //    return _strReturn;
        //}

        //국민 일일마감

        #region 국민 일일마감
        private static string ConvertReceiveType8(System.Data.DataTable dtable, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩
            StreamWriter _sw00 = null, _sw01 = null, _sw02 = null, _sw03 = null, _sw04 = null, _sw05 = null;	//파일 쓰기 스트림
            string _strLine = "";
            string _strReturn = "", _strStatus = "";
            int i = -1;
            string _strCardNumber = "";
            string _strFamilyNo = "", _strFamilyCheck = "";
            int _iAddCount = 0, icnt_01 = 0, icnt_02 = 0;
            string[] _strArFamilyNo = null;
            string strChange_status = null;
            try
            {
                _sw01 = new StreamWriter(fileName + ".01", true, _encoding);
                _sw02 = new StreamWriter(fileName + ".02", true, _encoding);
                //완처리 : 제휴사코드 + 인수일 + 건수
                _sw03 = new StreamWriter(fileName + ".특송_완처리", true, _encoding);
                _sw04 = new StreamWriter(fileName + ".완처리_바코드.txt", true, _encoding);
                _sw05 = new StreamWriter(fileName + ".이미지_가족_리스트.txt", true, _encoding);

                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    _iAddCount = int.Parse(dtable.Rows[i]["card_add_count"].ToString());
                    _strStatus = dtable.Rows[i]["card_delivery_status"].ToString();
                    //국민 전송 데이터 구분 
                    //( 배송 = 11, 반송 = 12, 분실 =13, 배송 -> 반송 = 14, 반송 -> 배송 = 15, 
                    //  배송 -> 분실 = 16, 반송 -> 분실 = 17, 선반납 = 18)
                    strChange_status = dtable.Rows[i]["change_delivery_status"].ToString();
                    _strFamilyNo = dtable.Rows[i]["family_customer_no"].ToString();

                    //데이터생성 시작
                    _strLine = "K";
                    _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["client_send_date"].ToString()), 8, true, ' ');
                    _strLine += GetStringAsLength(dtable.Rows[i]["client_express_code"].ToString(), 2, true, ' ');
                    _strLine += GetStringAsLength(dtable.Rows[i]["card_client_no_1"].ToString(), 6, true, ' ');
                    _strCardNumber = GetStringAsLength(dtable.Rows[i]["card_number"].ToString(), 16, true, ' ');
                    _strLine += "{0}";
                    _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["client_register_date"].ToString()), 8, true, ' ');

                    if (strChange_status == "11" || strChange_status == "12" || strChange_status == "13" || strChange_status == "14"
                         || strChange_status == "15" || strChange_status == "16" || strChange_status == "17" || strChange_status == "18")
                    {
                        //결과데이터가 국민카드 전송 후 변경된 경우 일일마감에서 생성되지 않게 한다.
                        //단, 재방->배송, 재방->반송, 재방->분실 건은 생성되게 한다.
                        //즉, 반송->배송, 분실->배송, 배송->배송
                        ;
                    }
                    else if ((_strStatus == "2" || _strStatus == "3" || _strStatus == "6") &&
                        (strChange_status != "20" && strChange_status != "21" && strChange_status != "22" || strChange_status == ""))
                    {
                        //반송or분실 중 재방이 아닌 건은 생성되지 않게 한다.
                        ;
                    }
                    else
                    {
                        if (_strStatus == "1")
                        {
                            _strLine += GetStringAsLength(_strStatus, 1, true, ' ');
                            _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["card_delivery_date"].ToString()), 8, true, ' ');
                            _strLine += GetStringAsLength(dtable.Rows[i]["receiver_code_change"].ToString().Replace("xx", "  "), 2, true, ' ');
                            _strLine += GetStringAsLength(dtable.Rows[i]["receiver_name"].ToString(), 14, true, ' ');
                            _strLine += GetStringAsLength(dtable.Rows[i]["receiver_SSN"].ToString().Replace("x","*"), 13, true, ' ');
                            _strLine += GetStringAsLength("1", 1, true, ' ');

                            //수령지구분값(재청구지구분)
                            //0-변경없음, 1-자택, 2-직장, 3-제3청구지
                            if (dtable.Rows[i]["change_type"].ToString().Trim().Length > 0)
                            {
                                _strLine += GetStringAsLength(dtable.Rows[i]["change_place_type"].ToString(), 1, true, ' ');
                                _strLine += GetStringAsLength(dtable.Rows[i]["change_zipcode"].ToString(), 6, true, ' ');
                                if (dtable.Rows[i]["change_address_detail"].ToString().Length > 40)
                                {
                                    _strLine += GetStringAsLength(dtable.Rows[i]["change_address_detail"].ToString().Substring(0, 33), 60, true, ' ');
                                }
                                else
                                {
                                    _strLine += GetStringAsLength(dtable.Rows[i]["change_address_detail"].ToString(), 60, true, ' ');
                                }
                                _strLine += GetStringAsLength(dtable.Rows[i]["change_home_tel"].ToString(), 15, true, ' ');
                                _strLine += GetStringAsLength(dtable.Rows[i]["change_mobile_tel"].ToString(), 15, true, ' ');
                            }
                            else
                            {
                                _strLine += GetStringAsLength("0", 1, true, ' ');
                                _strLine += GetStringAsLength("", 6, true, ' ');
                                _strLine += GetStringAsLength("", 60, true, ' ');
                                _strLine += GetStringAsLength("", 15, true, ' ');
                                _strLine += GetStringAsLength("", 15, true, ' ');
                            }

                            if (dtable.Rows[i]["card_kind"].ToString().ToLower() == "d")
                            {
                                _strLine += GetStringAsLength(dtable.Rows[i]["delivery_is_pda_register"].ToString(), 1, true, ' ');

                                _strLine += GetStringAsLength("", 22, true, ' ');
                            }
                            else
                            {
                                _strLine += GetStringAsLength("1", 1, true, ' ');

                                _strLine += GetStringAsLength("K" + GetSendCode(dtable.Rows[i]["client_register_type"].ToString(), dtable.Rows[i]["client_insert_type"].ToString()), 2, true, ' ');
                                _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["card_number"].ToString()), 16, true, ' ');
                                _strLine += GetStringAsLength(".tif", 4, true, ' ');
                            }
                        }
                        else if (_strStatus == "2" || _strStatus == "3")
                        {
                            //기존 배송->재방->반송
                            if (strChange_status == "21")
                            {
                                _strLine += GetStringAsLength("4", 1, true, ' ');
                                _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["delivery_return_date_last"].ToString()), 8, true, ' ');
                                _strLine += GetStringAsLength(dtable.Rows[i]["return_code_change"].ToString(), 2, true, ' ');
                            }
                            else if (dtable.Rows[i]["delivery_return_reason_last"].ToString() == "30")
                            {
                                _strLine += GetStringAsLength("8", 1, true, ' ');
                                _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["delivery_return_date_last"].ToString()), 8, true, ' ');
                                _strLine += GetStringAsLength("22", 2, true, ' ');
                            }
                            else
                            {
                                _strLine += GetStringAsLength("2", 1, true, ' ');
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
                                _strLine += GetStringAsLength("", 60, true, ' ');
                                _strLine += GetStringAsLength("", 15, true, ' ');
                                _strLine += GetStringAsLength("", 15, true, ' ');
                            }
                            else
                            {
                                _strLine += GetStringAsLength(dtable.Rows[i]["change_place_type"].ToString(), 1, true, ' ');
                                _strLine += GetStringAsLength(dtable.Rows[i]["change_zipcode"].ToString(), 6, true, ' ');
                                if (dtable.Rows[i]["change_address_detail"].ToString().Length > 40)
                                {
                                    _strLine += GetStringAsLength(dtable.Rows[i]["change_address_detail"].ToString().Substring(0, 33), 60, true, ' ');
                                }
                                else
                                {
                                    _strLine += GetStringAsLength(dtable.Rows[i]["change_address_detail"].ToString(), 60, true, ' ');
                                }
                                _strLine += GetStringAsLength(dtable.Rows[i]["change_home_tel"].ToString(), 15, true, ' ');
                                _strLine += GetStringAsLength(dtable.Rows[i]["change_mobile_tel"].ToString(), 15, true, ' ');
                            }
                            _strLine += GetStringAsLength("", 23, true, ' ');
                        }
                        else if (_strStatus == "6")
                        {
                            // 6 : 기존 배송->재방->분실
                            if (strChange_status == "21")
                            {
                                _strLine += GetStringAsLength("6", 1, true, ' ');
                            }
                            // 7 : 기존 반송->재방->분실
                            if (strChange_status == "22")
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

                            //징구구분
                            _strLine += GetStringAsLength("1", 1, true, ' ');
                            //수령지구분값(재청구지구분)
                            //0-변경없음, 1-자택, 2-직장, 3-제3청구지
                            if (dtable.Rows[i]["change_type"].ToString().Trim().Length > 0)
                            {
                                _strLine += GetStringAsLength(dtable.Rows[i]["change_place_type"].ToString(), 1, true, ' ');
                                _strLine += GetStringAsLength(dtable.Rows[i]["change_zipcode"].ToString(), 6, true, ' ');
                                if (dtable.Rows[i]["change_address_detail"].ToString().Length > 40)
                                {
                                    _strLine += GetStringAsLength(dtable.Rows[i]["change_address_detail"].ToString().Substring(0, 33), 60, true, ' ');
                                }
                                else
                                {
                                    _strLine += GetStringAsLength(dtable.Rows[i]["change_address_detail"].ToString(), 60, true, ' ');
                                }
                                _strLine += GetStringAsLength(dtable.Rows[i]["change_home_tel"].ToString(), 15, true, ' ');
                                _strLine += GetStringAsLength(dtable.Rows[i]["change_mobile_tel"].ToString(), 15, true, ' ');
                            }
                            else
                            {
                                _strLine += GetStringAsLength("0", 1, true, ' ');
                                _strLine += GetStringAsLength("", 6, true, ' ');
                                _strLine += GetStringAsLength("", 60, true, ' ');
                                _strLine += GetStringAsLength("", 15, true, ' ');
                                _strLine += GetStringAsLength("", 15, true, ' ');
                            }
                            _strLine += GetStringAsLength("", 23, true, ' ');

                        }
                        //수령지구분값(재청구지구분)
                        //0-변경없음, 1-자택, 2-직장, 3-제3청구지
                        //2011-10-10 태희철 추가 새주소관련
                        _strLine += npi_file_name(dtable.Rows[i]["npi_file_name"].ToString(), dtable.Rows[i]["change_type"].ToString());

                        if (_strStatus == "1")
                        {
                            icnt_01++;
                            _sw01.WriteLine(GetStringAsLength(string.Format(_strLine, _strCardNumber), 236, true, ' '));
                        }
                        else if (_strStatus == "2" || _strStatus == "3" || _strStatus == "6")
                        {
                            icnt_02++;
                            _sw02.WriteLine(GetStringAsLength(string.Format(_strLine, _strCardNumber), 236, true, ' '));
                        }

                        if (_iAddCount > 0)
                        {
                            _strArFamilyNo = _strFamilyNo.Split(new char[] { '|' });
                            for (int j = 0; j < _strArFamilyNo.Length; j++)
                            {
                                _strFamilyCheck = _strArFamilyNo[j];
                                if (_strCardNumber != _strFamilyCheck)
                                {
                                    //2014.05.27 가족카드의 경우 수령증 이미지 파일명 제거
                                    if (_strLine.IndexOf(".tif") > -1)
                                    {
                                        _strLine = _strLine.Substring(0, _strLine.Length - 58);
                                        _strLine += GetStringAsLength("", 22, true, ' ');

                                        _strLine += npi_file_name(dtable.Rows[i]["npi_file_name"].ToString(), dtable.Rows[i]["change_type"].ToString());
                                    }

                                    if (_strStatus == "1")
                                    {
                                        icnt_01++;
                                        _sw01.WriteLine(GetStringAsLength(string.Format(_strLine, GetStringAsLength(_strArFamilyNo[j], 16, true, ' ')), 236, true, ' '));
                                    }
                                    else if (_strStatus == "2" || _strStatus == "3" || _strStatus == "6")
                                    {
                                        icnt_02++;
                                        _sw02.WriteLine(GetStringAsLength(string.Format(_strLine, GetStringAsLength(_strArFamilyNo[j], 16, true, ' ')), 236, true, ' '));
                                    }
                                }
                            }
                        }

                        //특송 담당자 전달 파일
                        //배송 / 반송 완처리재방
                        if (_strStatus == "1" || _strStatus == "2" || _strStatus == "3" || _strStatus == "6")
                        {
                            if (strChange_status == "20" || strChange_status == "21" || strChange_status == "22")
                            {
                                _strLine = dtable.Rows[i]["card_type_detail"].ToString() + ",";
                                _strLine += String.Format("{0:yyyyMMdd}", dtable.Rows[i]["card_in_date"]) + ",";
                                if (_strFamilyNo == "")
                                {
                                    _strLine += dtable.Rows[i]["card_number"].ToString();
                                }
                                else
                                {
                                    _strLine += dtable.Rows[i]["card_number"].ToString() + "(" + _strFamilyNo + ")";
                                }
                                _sw03.WriteLine(_strLine);
                            }
                        }

                        if (_strStatus == "1" && dtable.Rows[i]["card_kind"].ToString().ToLower() != "d")
                        {
                            if (strChange_status == "20" || strChange_status == "21" || strChange_status == "22")
                            {
                                _strLine = dtable.Rows[i]["card_barcode"].ToString();
                                _sw04.WriteLine(_strLine);

                                _strLine = dtable.Rows[i]["card_number"].ToString();
                                _sw05.WriteLine(_strLine);

                                if (_iAddCount > 0)
                                {
                                    _strArFamilyNo = _strFamilyNo.Split(new char[] { '|' });
                                    for (int j = 0; j < _strArFamilyNo.Length; j++)
                                    {
                                        _strFamilyCheck = _strArFamilyNo[j];
                                        if (dtable.Rows[i]["card_number"].ToString() != _strFamilyCheck)
                                        {
                                            _sw05.WriteLine(dtable.Rows[i]["card_number"].ToString() + " / " + _strFamilyCheck + "     가족");
                                        }
                                    }
                                }
                            }
                        }
                    }
                    _strReturn = string.Format("배송 {0}건 / 반송, 분실 {1}의 마감데이타 다운 완료", icnt_01, icnt_02);
                }
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
                if (_sw03 != null) _sw03.Close();
                if (_sw04 != null) _sw04.Close();
                if (_sw05 != null) _sw05.Close();
            }
            return _strReturn;
        }
        #endregion

        //현대캐피탈 일일마감
        //private static string ConvertReceiveType9(System.Data.DataTable dtable, string fileName)
        //{
        //    Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
        //    StreamWriter _sw00 = null, _sw01 = null, _sw02 = null, _sw03 = null;																			//파일 쓰기 스트림
        //    int i = 0;
        //    StringBuilder _strLine = new StringBuilder("");
        //    string _strReturn = "", _strDeliveryStatus = "";
        //    try
        //    {
        //        _sw00 = new StreamWriter(fileName + ".00", true, _encoding);
        //        _sw01 = new StreamWriter(fileName + ".01", true, _encoding);
        //        _sw02 = new StreamWriter(fileName + ".02", true, _encoding);
        //        _sw03 = new StreamWriter(fileName + ".SSN", true, _encoding);
        //        for (i = 0; i < dtable.Rows.Count; i++)
        //        {
        //            _strDeliveryStatus = dtable.Rows[i]["card_delivery_status"].ToString();
        //            //2012-03-30 태희철 수정[S] (사용안함)
        //            //_strLine = new StringBuilder(GetStringAsLength(dtable.Rows[i]["card_kind"].ToString(), 2));
        //            if (_strDeliveryStatus == "0")
        //            {
        //                //_strLine.Append(GetStringAsLength("", 1));
        //                _strLine = new StringBuilder(GetStringAsLength("", 1, true, ' '));

                        
        //            }
        //            else
        //            {
        //                //_strLine.Append(GetStringAsLength(_strDeliveryStatus, 1));
        //                _strLine = new StringBuilder(GetStringAsLength(_strDeliveryStatus, 1, true, ' '));
        //            }
        //            //2012-03-30 태희철 수정[E] (사용안함)

        //            _strLine.Append(GetStringAsLength("2", 1, true, ' '));
        //            _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_number"].ToString(), 16, true, ' '));
        //            _strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_SSN"].ToString(), 13, true, ' '));
        //            if (dtable.Rows[i]["card_issue_type_code"].ToString() == "")
        //            {
        //                _strLine.Append(GetStringAsLength("1", 1, true, ' '));
        //            }
        //            else
        //            {
        //                _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_issue_type_code"].ToString(), 1, true, ' '));
        //            }

        //            if (_strDeliveryStatus == "1" || _strDeliveryStatus == "7")
        //            {
        //                _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[i]["card_delivery_date"].ToString()), 8, true, ' '));
        //                //2012-04-18 태희철 수정
        //                //if (dtable.Rows[i]["card_result_status"].ToString() == "61")
        //                //{
        //                //    _strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_SSN"].ToString(), 13));

        //                //}
        //                //else
        //                //{
        //                //    _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_SSN"].ToString(), 13));

        //                //}
        //                _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_SSN"].ToString(), 13, true, 'x'));
        //                _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_code_change"].ToString(), 2, true, ' '));
        //                _strLine.Append(GetStringAsLength("", 1, true, ' '));
        //            }
        //            else if (_strDeliveryStatus == "2" || _strDeliveryStatus == "3")
        //            {
        //                _strLine.Append(GetStringAsLength("", 8, true, ' '));
        //                _strLine.Append(GetStringAsLength("", 13, true, ' '));
        //                _strLine.Append(GetStringAsLength("", 2, true, ' '));
        //                _strLine.Append(GetStringAsLength(ReturnType(dtable.Rows[i]["delivery_return_reason_last"].ToString()), 1, true, ' '));
        //            }
        //            else
        //            {
        //                _strLine.Append(GetStringAsLength("", 8, true, ' '));
        //                _strLine.Append(GetStringAsLength("", 13, true, ' '));
        //                _strLine.Append(GetStringAsLength("", 2, true, ' '));
        //                _strLine.Append(GetStringAsLength("", 1, true, ' '));
        //            }

        //            _strLine.Append(GetStringAsLength("B002", 4, true, ' '));
        //            _strLine.Append(GetStringAsLength("", 16, true, ' '));

        //            if (dtable.Rows[i]["card_agree1"].ToString() == "2" || dtable.Rows[i]["card_agree2"].ToString() == "2" || dtable.Rows[i]["card_agree3"].ToString() == "2")
        //            {
        //                _strLine.Append(GetStringAsLength("N", 1, true, ' '));
        //            }
        //            else
        //            {
        //                _strLine.Append(GetStringAsLength("Y", 1, true, ' '));
        //            }
        //            /*
        //           _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["입회신청서변경이력명"].ToString()), 8) + chCSV;
        //           _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["회원신청등록일자"].ToString()), 8) + chCSV;
        //           _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["회원신청일련번호"].ToString()), 6) + chCSV;
        //             */
        //            if (_strDeliveryStatus == "1" || _strDeliveryStatus == "7")
        //            {
        //                _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_name"].ToString(), 16, true, ' '));

        //            }
        //            else
        //            {
        //                _strLine.Append(GetStringAsLength("", 16, true, ' '));
        //            }

        //            if (dtable.Rows[i]["card_agree4"].ToString().Equals("1"))
        //            {
        //                _strLine.Append(GetStringAsLength("Y", 1, true, ' '));
        //            }
        //            else
        //            {
        //                _strLine.Append(GetStringAsLength("N", 1, true, ' '));
        //            }
        //            //_strLine.Append(GetStringAsLength("N", 1));
        //            _strLine.Append(GetStringAsLength("", 7, true, ' '));


        //            if (_strDeliveryStatus == "1")
        //            {
        //                if (dtable.Rows[i]["receiver_SSN"].ToString().Trim() != dtable.Rows[i]["receiver_SSN_org"].ToString().Trim() && dtable.Rows[i]["card_result_status"].ToString() == "61")
        //                {
        //                    _sw03.WriteLine(_strLine.ToString());
        //                }
        //                else
        //                {
        //                    _sw01.WriteLine(_strLine.ToString());
        //                }
        //            }
        //            else if (_strDeliveryStatus == "2" || _strDeliveryStatus == "3")
        //            {
        //                _sw02.WriteLine(_strLine.ToString());
        //            }
        //            else
        //            {
        //                _sw00.WriteLine(_strLine.ToString());
        //            }
        //        }
        //        _strReturn = string.Format("{0}건의 인계데이터 다운 완료", i);
        //    }
        //    catch (Exception)
        //    {
        //        _strReturn = string.Format("{0}번째 데이터 인계 중 오류 발생", i + 1);
        //    }
        //    finally
        //    {
        //        if (_sw00 != null) _sw00.Close();
        //        if (_sw01 != null) _sw01.Close();
        //        if (_sw02 != null) _sw02.Close();
        //        if (_sw03 != null) _sw03.Close();
        //    }
        //    return _strReturn;
        //}


        #region 기타함수
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

        //국민카드 수령인 주민번호
        private static string ConvertReceiverSSN(string value)
        {
            string _strReturn = "";

            if (value.Length == 13)
            {
                _strReturn = value.Substring(0, 3) + "***" + value.Substring(6, 3) + "****";
            }
            else if (value.Length > 6 && value.Length < 13)
            {
                _strReturn = value.Substring(0, 3) + "***" + value.Substring(6, value.Length - 3).Replace("", " ") + "****";
            }
            else if (value.Length == 6)
            {
                _strReturn = value.Substring(0, 3) + "***   ****";
            }
            else if (value.Length > 0 && value.Length < 6)
            {
                _strReturn = value.Substring(0, value.Length - 3).Replace("", " ") + " ***   ****";
            }
            else
            {
                _strReturn = value;
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

        #endregion

        private static string Convert_SH_SSN(string value)
        {
            string _strReturn = "";
            if (value.Length < 6)
            {
                _strReturn = value.Substring(0, value.Length - 0);
            }
            else if (value.Length == 6)
            {
                _strReturn = value + "-";
            }
            else if (value.Length == 13)
            {
                _strReturn = value.Substring(0, 6) + "-" + value.Substring(6, 7);
            }
            else
            {
                _strReturn = value.Substring(0, 6) + "-" + value.Substring(6, value.Length - 6);
            }
            return _strReturn;
        }

        private static string Convert_KB_SSN(string value)
        {
            string _strReturn = "";

            if (value.Length == 13)
            {
                _strReturn = value.Substring(0, 3) + "***" + value.Substring(6, 3) + "****";
            }
            else if (value.Length > 6 && value.Length < 13)
            {
                _strReturn = value.Substring(0, 3) + "***" + value.Substring(6, value.Length - 3).Replace("", " ") + "****";
            }
            else if (value.Length == 6)
            {
                _strReturn = value.Substring(0, 3) + "***   ****";
            }
            else if (value.Length > 0 && value.Length < 6)
            {
                _strReturn = value.Substring(0, value.Length - 3).Replace("", " ") + " ***   ****";
            }
            else
            {
                _strReturn = value;
            }


            return _strReturn;
        }

        //국민코드변환
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


        private static string ConvertAgree(string value)
        {
            string _strReturn = "Y";

            switch (value)
            {
                case "1":
                    _strReturn = "Y";
                    break;
                case "2":
                    _strReturn = "N";
                    break;

            }
            return _strReturn;
        }

        #region 삼성 수령인 코드 변환
        private static string SM_ConvReceiver_code(string strReceiver_code)
        {
            string strReturn = null;

            switch (strReceiver_code)
            {
                case "01": strReturn = "01"; break;
                case "02": strReturn = "11"; break;
                case "03": strReturn = "12"; break;
                case "04": strReturn = "13"; break;
                case "05": strReturn = "14"; break;
                case "06":
                case "07":
                    strReturn = "15"; break;
                case "08":
                case "17":
                case "18":
                case "19":
                case "22":
                case "23":
                    strReturn = "16"; break;
                case "09":
                case "24":
                    strReturn = "17"; break;
                case "10": strReturn = "18"; break;
                case "11": strReturn = "19"; break;
                case "12":
                case "13":
                    strReturn = "20"; break;
                case "14": strReturn = "21"; break;
                case "15": strReturn = "22"; break;
                case "16": strReturn = "26"; break;
                case "35": strReturn = "27"; break;
                case "20":
                case "21":
                case "25":
                    strReturn = "23"; break;
                case "31": strReturn = "24"; break;
                //2013.05.14 태희철 수정 대리수령 코드 추가
                //2013.05.22 적용예정
                case "28": strReturn = "28"; break; // 형수
                case "36": strReturn = "29"; break; // 제수
                case "37": strReturn = "30"; break; // 백부/모
                case "38": strReturn = "31"; break; // 숙부/모
                case "39": strReturn = "32"; break; // 고모/부
                case "40": strReturn = "33"; break; // 이모/부
                case "41": strReturn = "34"; break; // 조카
                default:
                    strReturn = "";
                    break;
            }
            return strReturn;
        }
        #endregion


        #region 하나카드 진행 상황 산출 함수     
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
        #endregion

        #region 하나카다 반송 코트 변환     
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
        #endregion

        #region 하나카드 수령인 관계코드 변환     
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
        #endregion

        #region 국민카드 npi_file_name 값 정리
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
        #endregion

        #region 현대캐피탈 반송사유
        public static string ReturnType(string value)
        {
            string _strReturn = value;

            switch (value)
            {
                case "0":
                    _strReturn = "6";
                    break;
                case "1":
                    _strReturn = "4";
                    break;
                case "2":
                    _strReturn = "2";
                    break;
                case "3":
                    _strReturn = "1";
                    break;
                case "4":
                    _strReturn = "4";
                    break;
                case "5":
                    _strReturn = "3";
                    break;
                case "6":
                    _strReturn = "3";
                    break;
                case "7":
                    _strReturn = "5";
                    break;
                case "8":
                    _strReturn = "4";
                    break;
                case "9":
                    _strReturn = "6";
                    break;
                case "10":
                    _strReturn = "5";
                    break;
                case "11":
                    _strReturn = "2";
                    break;
                case "12":
                    _strReturn = "4";
                    break;
                case "13":
                    _strReturn = "3";
                    break;
                case "14":
                    _strReturn = "5";
                    break;
                case "15":
                    _strReturn = "5";
                    break;
                case "16":
                    _strReturn = "5";
                    break;
                case "17":
                    _strReturn = "3";
                    break;
                case "18":
                    _strReturn = "4";
                    break;
                case "19":
                    _strReturn = "4";
                    break;
                case "20":
                    _strReturn = "1";
                    break;
                case "21":
                    _strReturn = "1";
                    break;
                case "22":
                    _strReturn = "5";
                    break;
                case "23":
                    _strReturn = "4";
                    break;
                case "24":
                    _strReturn = "1";
                    break;
                case "25":
                    _strReturn = "3";
                    break;
                case "26":
                    _strReturn = "6";
                    break;
                case "27":
                    _strReturn = "5";
                    break;
                case "28":
                    _strReturn = "3";
                    break;
                case "29":
                    _strReturn = "6";
                    break;
                case "30":
                    _strReturn = "6";
                    break;
                case "31":
                    _strReturn = "3";
                    break;
                case "32":
                    _strReturn = "6";
                    break;
                case "33":
                    _strReturn = "6";
                    break;
                case "34":
                    _strReturn = "3";
                    break;
                case "88":
                    _strReturn = "6";
                    break;
                case "98":
                    _strReturn = "3";
                    break;
                case "99":
                    _strReturn = "6";
                    break;
                default:
                    _strReturn = "6";
                    break;
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
