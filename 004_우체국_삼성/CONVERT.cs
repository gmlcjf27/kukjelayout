using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace _004_우체국_삼성
{
	public class CONVERT
	{
		//기본 인코딩 설정
		private static string strEncoding = "ks_c_5601-1987";
        private static string strCardTypeID = "004";
        private static string strCardTypeName = "우체국_삼성카드";

		//현 DLL의 카드 타입 코드 반환
		public static string GetCardTypeID() {
            return strCardTypeID;
		}

		//현 DLL의 카드 타입명 반환
		public static string GetCardTypeName() {
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
                    _strLine = _strLine.Replace("", " ");
                    _strLine = _strLine.Replace("", " ");
                    _strLine = _strLine.Replace("@", " ");
                    //2011-11-24 태희철 수정
                    _strLine = _strLine.Replace("Ｆ", "  ");
                    _strLine = _strLine.Replace("|", " ");


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
            byte[] _byteAry = null;
            string _strReturn = "";
            string _strLine = "", _strLine2 = "";
            string _strZipcode = "", _strAreaType = "", _strAreaGroup = "", _strBranch = "", strDong_code = "";
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
                //21.01.19 우체국 추가
                _dtable.Columns.Add("check_org");               //dr[5]등기번호
                _dtable.Columns.Add("customer_name");
                _dtable.Columns.Add("card_zipcode");
                _dtable.Columns.Add("card_address_local");
                _dtable.Columns.Add("card_address_detail");     //  
                _dtable.Columns.Add("client_send_date");        //dr[10]발송일자
                _dtable.Columns.Add("client_express_code");     //발송업체코드
                _dtable.Columns.Add("client_send_number");      //발송번호
                _dtable.Columns.Add("card_design_code");        //제휴코드
                _dtable.Columns.Add("card_register_type");      //동의서 구분코드
                _dtable.Columns.Add("client_register_type");    //dr[15]
                _dtable.Columns.Add("customer_SSN");            //고객주민번호
                _dtable.Columns.Add("card_tel1");               //자택 전화번호
                _dtable.Columns.Add("customer_office");
                _dtable.Columns.Add("card_mobile_tel");
                _dtable.Columns.Add("card_count");              //dr[20]
                _dtable.Columns.Add("customer_no");             //고객번호
                _dtable.Columns.Add("card_client_no_1");        //MF번호
                _dtable.Columns.Add("card_cooperation1");
                _dtable.Columns.Add("card_product_code");       //제휴이미지 출력 코드(파일명)
                _dtable.Columns.Add("family_relation");         //dr[25] 가족카드 확인 코드
                _dtable.Columns.Add("card_terminal_issue");     //sms수신거부 회원 표기
                _dtable.Columns.Add("client_enterprise_code");  //본인만배송 요청건
                _dtable.Columns.Add("choice_agree1");           //필수동의
                _dtable.Columns.Add("choice_agree3");           //dr[29]마케팅동의

                _dtable.Columns.Add("card_zipcode_new");            //dr[30] 신우편번호
                _dtable.Columns.Add("card_zipcode_kind");           //dr[31] 신우편번호
                _dtable.Columns.Add("card_is_for_owner_only");      //dr[32] 본인만배송
                _dtable.Columns.Add("card_delivery_place_type");    //dr[33] 내부수령지코드
                _dtable.Columns.Add("card_barcode_new");            //dr[34]

                /*
                _dtable.Columns.Add("client_quick_seq");                
                _dtable.Columns.Add("card_issue_type_code");    //발급구분
				_dtable.Columns.Add("client_number");
                _dtable.Columns.Add("card_number");
                _dtable.Columns.Add("card_bank_ID");
                _dtable.Columns.Add("card_zipcode2");
                _dtable.Columns.Add("card_address2_local");
                _dtable.Columns.Add("card_address2_detail");
                _dtable.Columns.Add("card_tel2");                  // 25 직장 전화번호
                _dtable.Columns.Add("card_zipcode3");
                _dtable.Columns.Add("card_address3");              // 수령지 주소
                _dtable.Columns.Add("card_tel3");                  // 수령지 전화번호                
                _dtable.Columns.Add("customer_branch");             //30
                _dtable.Columns.Add("customer_position");
                _dtable.Columns.Add("card_delivery_place_code");
                _dtable.Columns.Add("client_enterprise_code");      //dr[33] 본인만배송 요청건
                _dtable.Columns.Add("card_cost_code");                
                _dtable.Columns.Add("card_client_code_1");          //40
                _dtable.Columns.Add("card_bank_account_name");
                _dtable.Columns.Add("card_bank_account_no");
                _dtable.Columns.Add("card_bank_account_owner");
                _dtable.Columns.Add("client_request_memo");         // 메모
                _dtable.Columns.Add("card_barcode_new");            // dr[51]
                _dtable.Columns.Add("card_issue_type_new");         // dr[52] 발급구분코드_new
                _dtable.Columns.Add("card_delivery_place_type");    // dr[53] 내부수령지코드
                _dtable.Columns.Add("card_zipcode_new");            // dr[54] 신우편번호
                _dtable.Columns.Add("card_zipcode_kind");           // dr[55] 신우편번호
                _dtable.Columns.Add("card_zipcode2_new");           // dr[56] 신우편번호
                _dtable.Columns.Add("card_zipcode2_kind");          // dr[57] 신우편번호
                _dtable.Columns.Add("card_zipcode3_new");           // dr[58] 신우편번호
                _dtable.Columns.Add("card_zipcode3_kind");          // dr[59] 신우편번호
                _dtable.Columns.Add("card_is_for_owner_only");      //dr[60] 본인만배송
                _dtable.Columns.Add("customer_memo");               //dr[61] 팝업메모문구
                _dtable.Columns.Add("change_add");                  //dr[62] 본인여부
                _dtable.Columns.Add("card_bank_account_tel");       //dr[63] 실번호 4자리
                */

                //우편번호 관련 정보 DataSet에 담기
                _dsetZipcode = new DataSet();
                _dsetZipcdeArea = new DataSet();

                _dsetZipcode_new = new DataSet();
                _dsetZipcdeArea_new = new DataSet();
                //우편번호
                _dsetZipcode.ReadXml(xmlZipcodePath);
                _dsetZipcode.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcode.Tables[0].Columns["zipcode"] };

                //업체별우편번호 코드구분
                _dsetZipcdeArea.ReadXml(xmlZipcodeAreaPath);
                _dsetZipcdeArea.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcdeArea.Tables[0].Columns["zipcode"] };

                //신우편번호
                _dsetZipcode_new.ReadXml(xmlZipcodePath_new);
                _dsetZipcode_new.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcode.Tables[0].Columns["zipcode_new"] };

                //신업체별우편번호 코드구분
                _dsetZipcdeArea_new.ReadXml(xmlZipcodeAreaPath_new);
                _dsetZipcdeArea_new.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcdeArea.Tables[0].Columns["zipcode_new"] };

				//파일 읽기 Stream과 오류시 저장할 쓰기 Stream 생성
				_sr = new StreamReader(path, _encoding);
				_swError = new StreamWriter(path + ".Error", false, _encoding);

                while ((_strLine = _sr.ReadLine()) != null) {
                    _strLine2 = _strLine.Replace("", " ");
                    _strLine2 = _strLine2.Replace("", " ");
                    _strLine2 = _strLine2.Replace("@", " ");
                    //2011-11-24 태희철 수정
                    _strLine2 = _strLine2.Replace("Ｆ", "  ");
                    _strLine2 = _strLine2.Replace("|", " ");

                    //인코딩, byte 배열로 담기
                    _byteAry = _encoding.GetBytes(_strLine2);
                    
                    _dr = _dtable.NewRow();
                    _dr[0] = _iSeq;

                    _dr[5] = _encoding.GetString(_byteAry, 0, 13);
                    //_dr[6] = _encoding.GetString(_byteAry, 14, 5);
                    //_dr[6] = _encoding.GetString(_byteAry, 20, 3);
                    //_dr[6] = _encoding.GetString(_byteAry, 24, 7);
                    //_dr[6] = _encoding.GetString(_byteAry, 32, 3);
                    _dr[6] = _encoding.GetString(_byteAry, 36, 40);
                    //수령지우편번호, 수령지주소
                    _strZipcode = _encoding.GetString(_byteAry, 118, 6).Trim();
                    _dr[7] = _strZipcode;
                    if (_dr[7].ToString().Length == 5)
                    {
                        _dr[30] = _dr[7].ToString();
                        _dr[31] = "1";
                    }

                    _dr[8] = _encoding.GetString(_byteAry, 124, 80);
                    _dr[9] = _encoding.GetString(_byteAry, 205, 124);

                    _dr[10] = _encoding.GetString(_byteAry, 451, 8);
                    _dr[11] = _encoding.GetString(_byteAry, 460, 2);
                    _dr[12] = _encoding.GetString(_byteAry, 463, 7);
                    _dr[13] = _encoding.GetString(_byteAry, 471, 5);

                    strDong_code = _encoding.GetString(_byteAry, 477, 1);
                    _dr[14] = strDong_code;

                    //card_barcode_new 사용
                    if (strDong_code == "0")
                    {
                        // 일반
                        strDong_code = "3";
                    }
                    else
                    {
                        // 동의서
                        strDong_code = "4";
                    }

                    _dr[15] = _encoding.GetString(_byteAry, 479, 2);
                    _dr[16] = _encoding.GetString(_byteAry, 482, 13).Replace('*', 'x');
                    _dr[17] = _encoding.GetString(_byteAry, 496, 15).Replace('*', 'x');
                    _dr[18] = _encoding.GetString(_byteAry, 512, 30);
                    _dr[19] = _encoding.GetString(_byteAry, 543, 12);
                    _dr[20] = _encoding.GetString(_byteAry, 556, 2);
                    _dr[21] = _encoding.GetString(_byteAry, 559, 11);
                    _dr[22] = _encoding.GetString(_byteAry, 571, 20);
                    _dr[23] = _encoding.GetString(_byteAry, 592, 20);
                    _dr[24] = _encoding.GetString(_byteAry, 700, 4);
                    _dr[25] = _encoding.GetString(_byteAry, 713, 2);
                    //본인배송여부
                    if (_encoding.GetString(_byteAry, 716, 1) == "Y")
                    {
                        _dr[26] = "01";

                        _dr[33] = "1";
                    }
                    else
                    {
                        _dr[26] = "00";
                    }
                    
                    string str_chk1 = "", str_chk2 = "";

                    //필수동의
                    str_chk1 = _encoding.GetString(_byteAry, 720, 1).Trim();

                    if (str_chk1.Trim() == "Y")
                    {
                        _dr[28] = "1";
                    }
                    else
                    {
                        _dr[28] = "0";
                    }
                    //마케팅동의 / 출입국동의
                    _dr[29] = _encoding.GetString(_byteAry, 722, 5).Replace(' ','9');

                    _dr[34] = "1";
                    // 케리어바코드
                    if (_strZipcode.Length == 5)
                    {
                        _dr[34] = _strZipcode + " " + strDong_code + _dr[20].ToString().Trim()
                        + _dr[10].ToString().Trim() + _dr[11].ToString().Trim()
                        + _dr[12].ToString().Trim();
                    }
                    else
                    {
                        _dr[34] = _strZipcode + strDong_code + _dr[20].ToString().Trim()
                        + _dr[10].ToString().Trim() + _dr[11].ToString().Trim()
                        + _dr[12].ToString().Trim();
                    }

                    if (_strZipcode.Trim() != "")
                    {
                        //지역 분류 선택
                        if (_strZipcode.Length == 5)
                        {
                            _drs = _dsetZipcdeArea_new.Tables[0].Select("zipcode_new = " + _strZipcode.Trim());
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
					//_fi = new FileInfo(path + ".Error");
					//_fi.Delete();
                    _strReturn = string.Format("{0}건의 데이터 변환 성공", _iSeq - 1);
                }
                else
                {
                    _strReturn = string.Format("{0}건 변환, 우편번호 미등록 {1}건 실패", _iSeq - _iErrorCount - 1, _iErrorCount);
                    
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
        public static string ConvertResult(DataTable dtable, string fileName)
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
                default:
                    _strReturn = "";
                    break;
            }
            return _strReturn;
        }

        //일일마감자료
        public static string ConvertResultDay(DataTable dtable, string fileName)
        {
            return ConvertReceiveType1(dtable, fileName);
        }

        //결과자료
        private static string ConvertReceiveType1(DataTable dtable, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            StreamWriter _sw00 = null, _sw01 = null, _sw02 = null;																			//파일 쓰기 스트림
            int i = 0;
            StringBuilder _strLine = new StringBuilder("");
            string _strReturn = "", _strStatus = "", _strReason_last = "", strCard_type_detail = "";
            string strCustomerSSN_type = null, strGetDate = "", strCard_in_date = "";

            try
            {
                strGetDate = DateTime.Now.ToString("MMdd");
                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    _strStatus = dtable.Rows[i]["card_delivery_status"].ToString();
                    _strReason_last = dtable.Rows[i]["delivery_return_reason_last_name"].ToString();

                    strCard_type_detail = dtable.Rows[i]["card_type_detail"].ToString();
                    strCard_in_date = String.Format("{0:yyyyMMdd}", dtable.Rows[i]["card_in_date"]);

                    //반송사유 자리 수 over로 인한 특수기호 생성 차단
                    if (_strReason_last.Length > 5)
                    {
                        _strReason_last = _strReason_last.Substring(0, 5);
                    }
                    
                    if (_strStatus == "2" || _strStatus == "3")
                    {
                        _strLine = new StringBuilder("");
                        _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[i]["client_send_date"].ToString()), 8, true, ' '));
                        _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[i]["client_express_code"].ToString()), 2, true, ' '));
                        _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[i]["client_send_number"].ToString()), 7, true, ' '));

                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_number"].ToString().Replace("x", "*"), 16, true, ' '));

                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_name"].ToString(), 8, true, ' '));
                        _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[i]["delivery_return_date_last"].ToString()), 8, true, ' '));
                        _strLine.Append(GetStringAsLength(_strReason_last, 10, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["return_code_change"].ToString(), 2, false, ' '));

                        if (strCard_type_detail.Substring(0,5) == "00421")
                        {
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_client_no_1"].ToString(), 20, true, ' '));
                        }
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_no"].ToString(), 11, true, ' '));
                    }
                    else
                    {   
                        _strLine = new StringBuilder("");

                        _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[i]["client_send_date"].ToString()), 8, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_express_code"].ToString(), 2, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_send_number"].ToString(), 7, true, ' '));

                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_number"].ToString().Replace("x", "*"), 16, true, ' '));

                        //2012.10.30 태희철 수정
                        //_strLine.Append(GetStringAsLength(dtable.Rows[i]["card_issue_type_code"].ToString(), 2, false, ' '));
                        _strLine.Append("04");

                        if (_strStatus == "6")
                        {
                            _strLine.Append(GetStringAsLength("99", 2, true, ' '));
                        }
                        else if (_strStatus == "1")
                        {   
                            if (dtable.Rows[i]["card_result_status"].ToString() == "61" || dtable.Rows[i]["card_result_status"].ToString() == "62")
                            {
                                _strLine.Append(GetStringAsLength(ConvReceiver_code(dtable.Rows[i]["receiver_code"].ToString()), 2, true, ' '));
                            }
                            else
                            {
                                _strLine.Append(GetStringAsLength("03", 2, true, ' '));
                            }
                            //2012.12.26 태희철 수정[E] 대리수령인 코드 변경
                        }
                        else if (_strStatus == "7")
                        {
                            _strLine.Append(GetStringAsLength("07", 2, true, ' '));
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength("", 2, true, ' '));
                        }

                        //2012-01-10 태희철 수정[S]
                        // 주민번호 9자리에서 3자리 3자리 일경우를 대비
                        strCustomerSSN_type = dtable.Rows[i]["customer_SSN"].ToString().Replace("x", "*");

                        //if (strCustomerSSN_type.Substring(7,1) == "*")
                        //{
                        //    ;
                        //}
                        //else
                        //{
                        //    strCustomerSSN_type = strCustomerSSN_type.Substring(0, 3) + "***" + strCustomerSSN_type.Substring(6, 3) + "****";
                        //}

                        strCustomerSSN_type = strCustomerSSN_type.Substring(0, 7) + "******";

                        _strLine.Append(GetStringAsLength(strCustomerSSN_type, 13, true, '*'));

                        //2012-01-10 태희철 수정[E]
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_name"].ToString(), 8, true, ' '));
                        _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[i]["card_delivery_date"].ToString()), 8, true, ' '));

                        if (_strStatus == "1")
                        {
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_name"].ToString(), 8, true, ' '));

                            //if (dtable.Rows[i]["card_result_status"].ToString() == "61")
                            //{
                            //    _strLine.Append(GetStringAsLength(strCustomerSSN_type, 13, true, '*'));
                            //}
                            //else
                            //{
                            //    _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_SSN"].ToString().Replace("x", "*"), 13, true, '*'));
                            //}

                            //2014.07.07 태희철 주민번호 변경
                            //3자리 3자리(7월7일 이전) -> 6자리 3자리(7월7일 부터 2주) -> 6자리 1자리(최종변경)
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_SSN"].ToString().Replace("x", "*"), 13, true, '*'));
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength("", 8, true, ' '));
                            _strLine.Append(GetStringAsLength("", 13, true, ' '));
                        }

                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_quick_seq"].ToString(), 6, true, ' '));

                        //동의서의 경우 추가 (일반과 동의서 총 byte 다름)
                        if (strCard_type_detail.Substring(0, 4) == "0042")
                        {
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_client_no_1"].ToString(), 20, true, ' '));
                        }

                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_no"].ToString(), 11, true, ' '));
                    }

                    //일반마감
                    if (strCard_type_detail.Substring(0,5) == "00411")
                    {
                        _sw00 = new StreamWriter(fileName + strGetDate + "삼성재방(차,건)" + strCard_in_date + ".txt", true, _encoding);
                        _sw01 = new StreamWriter(fileName + "국제일반수령(" + strCard_in_date + ").dat.01", true, _encoding);
                        _sw02 = new StreamWriter(fileName + "국제일반반송(" + strCard_in_date + ").dat.02", true, _encoding);
                    }
                    //동의서마감
                    else if (strCard_type_detail.Substring(0, 4) == "0042")
                    {
                        _sw00 = new StreamWriter(fileName + "삼성동의재방(" + strCard_in_date + ").txt", true, _encoding);

                        switch (strCard_type_detail)
                        {
                            case "0042101":
                                _sw01 = new StreamWriter(fileName + "국제동의서수령(" + strCard_in_date + ").dat", true, _encoding);
                                _sw02 = new StreamWriter(fileName + "국제동의서반송(" + strCard_in_date + ").dat", true, _encoding);
                                break;
                            case "0042102":
                                _sw01 = new StreamWriter(fileName + "국제SFC수령(" + strCard_in_date + ").dat", true, _encoding);
                                _sw02 = new StreamWriter(fileName + "국제SFC반송(" + strCard_in_date + ").dat", true, _encoding);
                                break;
                            case "0042103":
                                _sw01 = new StreamWriter(fileName + "국제LIFE수령(" + strCard_in_date + ").dat", true, _encoding);
                                _sw02 = new StreamWriter(fileName + "국제LIFE반송(" + strCard_in_date + ").dat", true, _encoding);
                                break;
                            case "0042104":
                                _sw01 = new StreamWriter(fileName + "국제자체수령(" + strCard_in_date + ").dat", true, _encoding);
                                _sw02 = new StreamWriter(fileName + "국제자체반송(" + strCard_in_date + ").dat", true, _encoding);
                                break;
                            case "0042105":
                                _sw01 = new StreamWriter(fileName + "국제체크수령(" + strCard_in_date + ").dat", true, _encoding);
                                _sw02 = new StreamWriter(fileName + "국제체크반송(" + strCard_in_date + ").dat", true, _encoding);
                                break;
                            case "0042106":
                                _sw01 = new StreamWriter(fileName + "국제주유수령(" + strCard_in_date + ").dat", true, _encoding);
                                _sw02 = new StreamWriter(fileName + "국제주유반송(" + strCard_in_date + ").dat", true, _encoding);
                                break;
                            case "0042107":
                                _sw01 = new StreamWriter(fileName + "국제SOIL수령(" + strCard_in_date + ").dat", true, _encoding);
                                _sw02 = new StreamWriter(fileName + "국제SOIL반송(" + strCard_in_date + ").dat", true, _encoding);
                                break;
                            case "0042108":
                                _sw01 = new StreamWriter(fileName + "국제화재수령(" + strCard_in_date + ").dat", true, _encoding);
                                _sw02 = new StreamWriter(fileName + "국제화재반송(" + strCard_in_date + ").dat", true, _encoding);
                                break;
                            case "0042109":
                                _sw01 = new StreamWriter(fileName + "국제CMA수령(" + strCard_in_date + ").dat", true, _encoding);
                                _sw02 = new StreamWriter(fileName + "국제CMA반송(" + strCard_in_date + ").dat", true, _encoding);
                                break;
                            case "0042110":
                                _sw01 = new StreamWriter(fileName + "국제투어수령(" + strCard_in_date + ").dat", true, _encoding);
                                _sw02 = new StreamWriter(fileName + "국제투어반송(" + strCard_in_date + ").dat", true, _encoding);
                                break;
                            case "0042111":
                                _sw01 = new StreamWriter(fileName + "국제CJONE수령(" + strCard_in_date + ").dat", true, _encoding);
                                _sw02 = new StreamWriter(fileName + "국제CJONE반송(" + strCard_in_date + ").dat", true, _encoding);
                                break;
                            case "0042112":
                                _sw01 = new StreamWriter(fileName + "국제SK에너지수령(" + strCard_in_date + ").dat", true, _encoding);
                                _sw02 = new StreamWriter(fileName + "국제SK에너지반송(" + strCard_in_date + ").dat", true, _encoding);
                                break;
                            case "0042113":
                                _sw01 = new StreamWriter(fileName + "국제홈플러스수령(" + strCard_in_date + ").dat", true, _encoding);
                                _sw02 = new StreamWriter(fileName + "국제홈플러스반송(" + strCard_in_date + ").dat", true, _encoding);
                                break;
                            case "0042114":
                                _sw01 = new StreamWriter(fileName + "국제6+수령(" + strCard_in_date + ").dat", true, _encoding);
                                _sw02 = new StreamWriter(fileName + "국제6+반송(" + strCard_in_date + ").dat", true, _encoding);
                                break;
                            case "0042115":
                                _sw01 = new StreamWriter(fileName + "국제S4수령(" + strCard_in_date + ").dat", true, _encoding);
                                _sw02 = new StreamWriter(fileName + "국제S4반송(" + strCard_in_date + ").dat", true, _encoding);
                                break;
                            case "0042116":
                                _sw01 = new StreamWriter(fileName + "국제MNO수령(" + strCard_in_date + ").dat", true, _encoding);
                                _sw02 = new StreamWriter(fileName + "국제MNO반송(" + strCard_in_date + ").dat", true, _encoding);
                                break;
                            case "0042117":
                                _sw01 = new StreamWriter(fileName + "국제뷰티수령(" + strCard_in_date + ").dat", true, _encoding);
                                _sw02 = new StreamWriter(fileName + "국제뷰티반송(" + strCard_in_date + ").dat", true, _encoding);
                                break;
                            case "0042118":
                                _sw01 = new StreamWriter(fileName + "국제전자랜드수령(" + strCard_in_date + ").dat", true, _encoding);
                                _sw02 = new StreamWriter(fileName + "국제전자랜드반송(" + strCard_in_date + ").dat", true, _encoding);
                                break;
                            case "0042119":
                                _sw01 = new StreamWriter(fileName + "국제해피랜드수령(" + strCard_in_date + ").dat", true, _encoding);
                                _sw02 = new StreamWriter(fileName + "국제해피랜드반송(" + strCard_in_date + ").dat", true, _encoding);
                                break;
                            case "0042120":
                                _sw01 = new StreamWriter(fileName + "국제S클래스수령(" + strCard_in_date + ").dat", true, _encoding);
                                _sw02 = new StreamWriter(fileName + "국제S클래스반송(" + strCard_in_date + ").dat", true, _encoding);
                                break;
                            case "0042121":
                                _sw01 = new StreamWriter(fileName + "국제손보사수령(" + strCard_in_date + ").dat", true, _encoding);
                                _sw02 = new StreamWriter(fileName + "국제손보사반송(" + strCard_in_date + ").dat", true, _encoding);
                                break;
                            case "0042122":
                                _sw01 = new StreamWriter(fileName + "국제국민행복수령(" + strCard_in_date + ").dat", true, _encoding);
                                _sw02 = new StreamWriter(fileName + "국제국민행복반송(" + strCard_in_date + ").dat", true, _encoding);
                                break;
                            case "0042123":
                                _sw01 = new StreamWriter(fileName + "국제신세계수령(" + strCard_in_date + ").dat", true, _encoding);
                                _sw02 = new StreamWriter(fileName + "국제신세계반송(" + strCard_in_date + ").dat", true, _encoding);
                                break;
                            case "0042124":
                                _sw01 = new StreamWriter(fileName + "국제GS칼텍스수령(" + strCard_in_date + ").dat", true, _encoding);
                                _sw02 = new StreamWriter(fileName + "국제GS칼텍스반송(" + strCard_in_date + ").dat", true, _encoding);
                                break;
                            case "0042125":
                                _sw01 = new StreamWriter(fileName + "국제SOIL멤버십수령(" + strCard_in_date + ").dat", true, _encoding);
                                _sw02 = new StreamWriter(fileName + "국제SOIL멤버십반송(" + strCard_in_date + ").dat", true, _encoding);
                                break;
                            case "0042126":
                                _sw01 = new StreamWriter(fileName + "국제화물복지수령(" + strCard_in_date + ").dat", true, _encoding);
                                _sw02 = new StreamWriter(fileName + "국제화물복지반송(" + strCard_in_date + ").dat", true, _encoding);
                                break;
                            case "0042127":
                                _sw01 = new StreamWriter(fileName + "국제큰수레화물복지수령(" + strCard_in_date + ").dat", true, _encoding);
                                _sw02 = new StreamWriter(fileName + "국제큰수레화물복지반송(" + strCard_in_date + ").dat", true, _encoding);
                                break;
                            default:
                                _sw01 = new StreamWriter(fileName + "기타수령(" + strCard_in_date + ").dat", true, _encoding);
                                _sw02 = new StreamWriter(fileName + "기타반송(" + strCard_in_date + ").dat", true, _encoding);
                                break;
                        }
                    }

                    if (strCard_type_detail.Substring(0,5) == "00411" || strCard_type_detail.Substring(0, 4) == "0042")
                    {
                        if (_strStatus == "1")
                        {
                            _sw01.WriteLine(_strLine.ToString());
                        }
                        else if (_strStatus == "2" || _strStatus == "3")
                        {
                            _sw02.WriteLine(_strLine.ToString());
                        }
                        else
                        {
                            _sw00.WriteLine(_strLine.ToString());
                        }
                    }
                    

                    if (_sw00 != null) _sw00.Close();
                    if (_sw01 != null) _sw01.Close();
                    if (_sw02 != null) _sw02.Close();
                }
                _strReturn = string.Format("{0}건의 인계데이터 다운 완료", i);
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

        //월마감
        private static string ConvertReceiveType2(DataTable dtable, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            StreamWriter _sw00 = null, _sw01 = null, _sw02 = null, _sw03 = null;

            int i = 0, iCnt = 0;
            StringBuilder _strLine = new StringBuilder("");
            StringBuilder _strLine1 = new StringBuilder("");

            string _strReturn = "", _strStatus = "", _strCard_issue_type_code = "";
            string strCustomerSSN_type = null, strCard_type_detail = "", strCard_zipcode_kind = "";

            try
            {
                _sw00 = new StreamWriter(fileName + ".00", true, _encoding);
                _sw01 = new StreamWriter(fileName + ".배송", true, _encoding);
                _sw02 = new StreamWriter(fileName + ".반송", true, _encoding);
                _sw03 = new StreamWriter(fileName + ".분실", true, _encoding);

                _strLine1 = new StringBuilder("인수일,");
                _strLine1.Append("카드번호,");
                _strLine1.Append("주민번호,");
                _strLine1.Append("이름,");
                _strLine1.Append("수령일,");
                _strLine1.Append("우편번호,");
                _strLine1.Append("MF번호,");
                _strLine1.Append("업체코드,");
                _strLine1.Append("일련번호");

                _sw00.WriteLine(_strLine1.ToString());
                _sw01.WriteLine(_strLine1.ToString());
                _sw02.WriteLine(_strLine1.ToString());
                _sw03.WriteLine(_strLine1.ToString());
                

                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    _strCard_issue_type_code = dtable.Rows[i]["card_issue_type_code"].ToString();
                    _strStatus = dtable.Rows[i]["card_delivery_status"].ToString();
                    strCard_type_detail = dtable.Rows[i]["card_type_detail"].ToString();

                    if (strCard_type_detail.Substring(0, 5) == "00421")
                    {
                        if (_strStatus == "2" || _strStatus == "3")
                        {
                            _strLine = new StringBuilder(GetStringAsLength(RemoveDash(dtable.Rows[i]["card_in_date"].ToString()), 8, true, ' ') + ",");

                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_number"].ToString().Replace("x", "*"), 16, true, ' ') + ",");

                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_name"].ToString(), 10, true, ' ') + ",");
                            _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[i]["delivery_return_date_last"].ToString()), 8, true, ' ') + ",");
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["delivery_return_reason_last_name"].ToString(), 12, true, ' ') + ",");

                            if (strCard_zipcode_kind == "1")
                            {
                                _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_zipcode_new"].ToString(), 7, true, ' ') + ",");
                            }
                            else
                            {
                                _strLine.Append(GetStringAsLength(ConvertZipcode(dtable.Rows[i]["card_zipcode"].ToString()), 7, true, ' ') + ",");
                            }
                        }
                        else
                        {

                            _strLine = new StringBuilder(GetStringAsLength(RemoveDash(dtable.Rows[i]["card_in_date"].ToString()), 8, true, ' ') + ",");

                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_number"].ToString().Replace("x", "*"), 16, true, ' ') + ",");

                            //2011-12-20 태희철 수정[S]
                            // 주민번호 9자리에서 3자리 3자리 일경우를 대비
                            strCustomerSSN_type = dtable.Rows[i]["customer_SSN"].ToString().Replace("x", "*");
                            
                            strCustomerSSN_type = strCustomerSSN_type.Substring(0, 7) + "******";
                            _strLine.Append(GetStringAsLength(strCustomerSSN_type, 13, true, '*') + ",");

                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_name"].ToString(), 10, true, ' ') + ",");
                            _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[i]["card_delivery_date"].ToString()), 8, true, ' ') + ",");

                            if (strCard_zipcode_kind == "1")
                            {
                                _strLine.Append(GetStringAsLength(ConvertZipcode(dtable.Rows[i]["card_zipcode_new"].ToString()), 7, true, ' ') + ",");
                            }
                            else
                            {
                                _strLine.Append(GetStringAsLength(ConvertZipcode(dtable.Rows[i]["card_zipcode"].ToString()), 7, true, ' ') + ",");
                            }
                        }

                        //_strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_no"].ToString(), 11, true, ' '));
                        //2013.05.24 태희철 수정
                        //MF번호
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_client_no_1"].ToString(), 20, true, ' ') + ",");
                        //업체코드
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_express_code"].ToString(), 2, true, ' ') + ",");
                        //일련번호
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_send_number"].ToString(), 7, true, ' '));


                        //2013.05.15 태희철 수정 순수신규, 기존신규 + 일반 구분
                        iCnt++;
                        if (_strStatus == "2" || _strStatus == "3")
                        {
                            _sw02.WriteLine(_strLine.ToString());
                        }
                        else if (_strStatus == "1")
                        {
                            _sw01.WriteLine(_strLine.ToString());
                        }
                        else if (_strStatus == "6")
                        {
                            _sw03.WriteLine(_strLine.ToString());
                        }
                        else
                        {
                            _sw00 .WriteLine(_strLine.ToString());
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
            }
            return _strReturn;
        }

        //수령인 코드 변환 2012.12.26 태희철 추가
        #region 삼성 수령인 코드 변환     
        private static string ConvReceiver_code(string strReceiver_code)
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

        //지역 번호 정리
		private static void ArrangeData(ref DataTable dtable) {
			string _strAreaGroup = "", _strPrevGroup = "";
			int _iAreaIndex = 1, _iIndex = -1;
			DataRow[] _dr = dtable.Select("", "card_area_group, card_zipcode, customer_name");
			for (int i = 0; i < _dr.Length; i++) {
				_iIndex = int.Parse(_dr[i][0].ToString());
				_strAreaGroup = _dr[i][1].ToString();
				if (_strPrevGroup != _strAreaGroup) {
					_strPrevGroup = _strAreaGroup;
					_iAreaIndex = 1;
				}
				dtable.Rows[_iIndex - 1][3] = _iAreaIndex;
				_iAreaIndex++;
			}
		}

		//문자중 -를 없앤다
		private static string RemoveDash(string value) {
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
        //2012-01-09 태희철 수정 주민번호 "x"값을 "*" 값으로 수정
        //private static string CustomerSSNType(string value)
        //{
        //    string _strReturn = "";

        //    if (value.Length > 6)
        //        _strReturn = value.Substring(0, 4) + "**" + value.Substring(6);
        //    else if (value.Length > 4)
        //        _strReturn = value.Substring(0, 4) + "**";

        //    return _strReturn.Replace("x","*");
        //}
        private static string ConvertAvg(string value)
        {
            string _strReturn = "";

            if (value.LastIndexOf(".") > -1)
            {
                _strReturn = value;
            }
            else
            {
                _strReturn = value + ".0";
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
            else if (_strReturn.Length == 5)
            {
                _strReturn = value.Substring(0, 3) + "-" + value.Substring(3, 2);
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

