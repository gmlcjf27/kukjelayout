using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Text;

namespace _129NH_Re_Ban
{
    public class CONVERT
    {
        //기본 인코딩 설정
        private static string strEncoding = "ks_c_5601-1987";
        private static string strCardTypeID = "129";
        private static string strCardTypeName = "129-NH농협_재발송";

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
            string _strReturn = "0904101";

            return _strReturn;
        }

        //등록 자료 생성
        //public static string ConvertRegister(string path, string xmlZipcodePath, string xmlZipcodeAreaPath, string xmlPath)
        public static string ConvertRegister(string path, string xmlZipcodePath, string xmlZipcodeAreaPath, string xmlZipcodePath_new, string xmlZipcodeAreaPath_new, string xmlPath)
        {

            System.Text.Encoding _encoding = System.Text.Encoding.GetEncoding("ks_c_5601-1987");	//기본 인코딩	
            StreamReader _sr = null;    								//파일 읽기 스트림
            StreamWriter _swError = null;								//파일 쓰기 스트림
            DataSet _dsetZipcode = null, _dsetZipcdeArea = null;		//우편번호 관련 DataSet
            DataSet _dsetZipcode_new = null, _dsetZipcdeArea_new = null;		//우편번호 관련 DataSet
            DataTable _dtable = null;									//마스터 저장 테이블
            DataRow _dr = null;
            DataRow[] _drs = null;
            byte[] _byteAry = null;
            string _strReturn = "";
            string _strLine = "";
            string _strZipcode = "", _strAreaType = "", _strAreaGroup = "", _strBranch = "";
            int _iSeq = 1, _iErrorCount = 0;
            Branches _branches = new Branches();
            try
            {
                #region 테이블 구조
                _dtable = new DataTable("Convert");
                //기본 컬럼
                _dtable.Columns.Add("degree_arrange_number");
                _dtable.Columns.Add("card_area_group");
                _dtable.Columns.Add("card_branch");
                _dtable.Columns.Add("card_area_type");
                _dtable.Columns.Add("area_arrange_number");
                //세부 컬럼
                //자료구분      0
                _dtable.Columns.Add("client_send_number");          // dr[5]배송일련번호  0
                _dtable.Columns.Add("client_express_code");         //배송차수      0 //업체 코드 : 국제 02
                _dtable.Columns.Add("client_number");               //캐리어일련번호0
                _dtable.Columns.Add("client_send_date");            //발송일자      0
                _dtable.Columns.Add("card_issue_detail_code");          //배송구분 : 1=일반,동의서, 7=긴급, 3=재발송
                _dtable.Columns.Add("card_pt_code");                // dr[10] 카드구분      0
                _dtable.Columns.Add("client_quick_seq");              //유형코드      0
                _dtable.Columns.Add("card_request_memo");           //유형명        0
                _dtable.Columns.Add("card_client_no_1");            //시작카드번호
                _dtable.Columns.Add("customer_no");                 // dr[15]끝 카드번호
                _dtable.Columns.Add("card_count");                  //카드매수
                _dtable.Columns.Add("card_cooperation_code");       //카드 등급     0
                _dtable.Columns.Add("card_issue_type_code");        // 발급구분 
                _dtable.Columns.Add("card_consented");              // dr[18]동의서 징구여부0
                _dtable.Columns.Add("card_product_name");           // dr[19] 제휴코드
                _dtable.Columns.Add("card_bank_ID");                //dr[20] 관리점코드    0
                _dtable.Columns.Add("customer_position");           //관리점 번호
                _dtable.Columns.Add("card_zipcode3");               //관리점 우편번호
                _dtable.Columns.Add("card_address3");               //관리졈 주소
                _dtable.Columns.Add("card_bank_name");              //관리점 이름       0
                _dtable.Columns.Add("customer_name");               //dr[25] 한글이름
                _dtable.Columns.Add("customer_SSN");                //한글이름
                _dtable.Columns.Add("card_mobile_tel");             //휴대폰 번호
                _dtable.Columns.Add("card_tel1");                   //
                _dtable.Columns.Add("card_tel2");
                _dtable.Columns.Add("card_tel3");
                _dtable.Columns.Add("card_delivery_place_type");    //수령지 구분       0
                _dtable.Columns.Add("card_zipcode");
                _dtable.Columns.Add("card_address_local");          //dr[33] 동이상
                _dtable.Columns.Add("card_address_detail");         //동이하
                _dtable.Columns.Add("card_zipcode2");
                _dtable.Columns.Add("card_address2_local");         //dr[36] 동이상
                _dtable.Columns.Add("card_address2_detail");        //동이하
                _dtable.Columns.Add("family_name");
                _dtable.Columns.Add("client_bank_request_no");
                _dtable.Columns.Add("card_terminal_issue");         //dr[40]

                //2011-09-28 태희철 수정[S]
                //신주소구분
                //건물일련번호
                //도로명주소
                //2011-09-28[E]

                _dtable.Columns.Add("card_number");             //카드넘버
                _dtable.Columns.Add("card_barcode_new");        //카드 뉴바코드
                _dtable.Columns.Add("card_issue_type_new");          // dr[43] 발급구분코드_new

                _dtable.Columns.Add("card_zipcode_new");               //dr[44]신우편번호
                _dtable.Columns.Add("card_zipcode_kind");              //
                _dtable.Columns.Add("card_zipcode2_new");              //
                _dtable.Columns.Add("card_zipcode2_kind");             //
                _dtable.Columns.Add("card_zipcode3_new");              //
                _dtable.Columns.Add("card_zipcode3_kind");             //dr[49]

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
                #endregion

                #region 테이블 저장
                //파일 읽기 Stream과 오류시 저장할 쓰기 Stream 생성
                _sr = new StreamReader(path, _encoding);
                _swError = new StreamWriter(path + ".Error", false, _encoding);

                while ((_strLine = _sr.ReadLine()) != null)
                {
                    //인코딩, byte 배열로 담기
                    _byteAry = _encoding.GetBytes(_strLine);

                    if (_encoding.GetString(_byteAry, 19, _byteAry.Length - 20).Replace(" ", "") == "")
                        continue;

                    _dr = _dtable.NewRow();
                    _dr[0] = _iSeq;
                    _dr[5] = _encoding.GetString(_byteAry, 2, 17);
                    _dr[6] = _encoding.GetString(_byteAry, 19, 2);
                    _dr[7] = _encoding.GetString(_byteAry, 23, 9);
                    _dr[8] = _encoding.GetString(_byteAry, 32, 8);
                    _dr[9] = _encoding.GetString(_byteAry, 40, 1);
                    _dr[10] = _encoding.GetString(_byteAry, 41, 1);
                    _dr[11] = _encoding.GetString(_byteAry, 42, 12);
                    _dr[12] = _encoding.GetString(_byteAry, 54, 60);
                    _dr[13] = _encoding.GetString(_byteAry, 114, 16);
                    _dr[14] = _encoding.GetString(_byteAry, 130, 16);
                    _dr[15] = _encoding.GetString(_byteAry, 146, 6);
                    _dr[16] = _encoding.GetString(_byteAry, 152, 1);
                    _dr[17] = _encoding.GetString(_byteAry, 153, 1);
                    _dr[18] = _encoding.GetString(_byteAry, 154, 1);
                    _dr[19] = _encoding.GetString(_byteAry, 155, 6);
                    _dr[20] = _encoding.GetString(_byteAry, 161, 6);
                    _dr[21] = _encoding.GetString(_byteAry, 167, 12);
                    _dr[22] = _encoding.GetString(_byteAry, 179, 6).Trim();

                    if (_dr[22].ToString().Length == 5)
                    {
                        _dr[48] = _dr[22].ToString();
                        _dr[49] = "1";
                    }

                    _dr[23] = _encoding.GetString(_byteAry, 185, 100).TrimEnd() + " " + _encoding.GetString(_byteAry, 285, 100).TrimEnd();
                    // 레이아웃상 50byte이나 DB상 30byte
                    // 문제발생 시 DB 수정필요
                    _dr[24] = _encoding.GetString(_byteAry, 385, 30);
                    _dr[25] = _encoding.GetString(_byteAry, 435, 40);
                    _dr[26] = _encoding.GetString(_byteAry, 475, 13).Replace("*","x");
                    _dr[27] = _encoding.GetString(_byteAry, 488, 12);

                    _dr[29] = _encoding.GetString(_byteAry, 500, 12); //직장
                    _dr[28] = _encoding.GetString(_byteAry, 512, 12); //자택
                    _dr[30] = _encoding.GetString(_byteAry, 524, 12); //제3청구지

                    _dr[31] = _encoding.GetString(_byteAry, 536, 1);


                    _strZipcode = _encoding.GetString(_byteAry, 537, 6).Trim(); //주소1
                    _dr[32] = _strZipcode;

                    if (_strZipcode.Length == 5)
                    {
                        _dr[44] = _strZipcode;
                        _dr[45] = "1";
                    }

                    _dr[33] = _encoding.GetString(_byteAry, 543, 100).TrimEnd(); // 동이상
                    _dr[34] = _encoding.GetString(_byteAry, 643, 100).TrimEnd(); // 동이하

                    _dr[35] = _encoding.GetString(_byteAry, 743, 6).Trim(); //주소2

                    if (_dr[35].ToString().Length == 5)
                    {
                        _dr[46] = _dr[35].ToString();
                        _dr[47] = "1";
                    }

                    _dr[36] = _encoding.GetString(_byteAry, 749, 100).TrimEnd(); //주소2
                    _dr[37] = _encoding.GetString(_byteAry, 849, 100).TrimEnd(); //주소2

                    _dr[38] = _encoding.GetString(_byteAry, 949, 50);
                    _dr[39] = _encoding.GetString(_byteAry, 999, 17);
                    _dr[40] = _encoding.GetString(_byteAry, 1016, 2);

                    //2011-09-28 태희철 수정[S]
                    //신주소구분
                    //_encoding.GetString(_byteAry, 1018, 1);신주소구분여부 사용안함
                    //건물일련번호
                    //도로명주소
                    //2011-09-28[E]

                    _dr[41] = _dr[8].ToString().Substring(1, 7) + _dr[7].ToString();
                    //_dr[40] = _dr[8].ToString() + _dr[7].ToString() + "01" + "02";
                    if (_strZipcode.Length == 5)
                    {
                        _dr[42] = _dr[8].ToString() + _dr[7].ToString() + _strZipcode + " " + "01" + "02";
                    }
                    else
                    {
                        _dr[42] = _dr[8].ToString() + _dr[7].ToString() + _strZipcode + "01" + "02";
                    }

                    _dr[43] = _dr[17];

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
                #endregion
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
            return ConvertResultGubun(dtable, fileName, 3);
        }

        //일일마감자료
        public static string ConvertResultDay(System.Data.DataTable dtable, string fileName)
        {
            return ConvertResultGubun(dtable, fileName, 2);
        }

        private static string ConvertResultGubun(System.Data.DataTable dtable, string fileName, int temp)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            StreamWriter _sw00 = null, _sw01 = null, _sw02 = null;															//파일 쓰기 스트림
            int i = 0;
            StringBuilder _strLine = new StringBuilder("");
            string _strReturn = "", _strStatus = "", strDong = null;
            string strCard_zipcode_kind = "";
            int temp_0 = 0, temp_1 = 0, temp_2 = 0;
            try
            {
                string tempday = DateTime.Now.ToString("yyyyMMdd");
                if (temp == 2)
                {
                    _sw00 = new StreamWriter(fileName + "KU0.11262." + tempday + ".00", true, _encoding);
                    _sw01 = new StreamWriter(fileName + "KU0.11262." + tempday + ".01", true, _encoding);
                    _sw02 = new StreamWriter(fileName + "KU0.11262." + tempday + ".02", true, _encoding);
                    _strLine.Append(GetStringAsLength("FH", 2, true, ' '));
                    _strLine.Append(GetStringAsLength(tempday, 8, true, ' '));
                    _strLine.Append(GetStringAsLength("", 325, true, ' '));
                }
                else
                {
                    _sw00 = new StreamWriter(fileName + "KU0.11263." + tempday + ".000", true, _encoding);
                    _sw01 = new StreamWriter(fileName + "KU0.11263." + tempday + ".001", true, _encoding);
                    _sw02 = new StreamWriter(fileName + "KU0.11263." + tempday + ".002", true, _encoding);

                    _strLine.Append(GetStringAsLength("FH", 2, true, ' '));
                    _strLine.Append(GetStringAsLength(tempday, 8, true, ' '));
                    _strLine.Append("{0}");
                    _strLine.Append(GetStringAsLength("", 323, true, ' '));
                    _strLine = new StringBuilder(string.Format(_strLine.ToString(), GetStringAsLength(dtable.Rows[0]["client_express_code"].ToString(), 2, true, ' ')));
                }

                _sw01.WriteLine(_strLine);
                _sw02.WriteLine(_strLine);
                _sw00.WriteLine(_strLine);

                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    #region 시간 뽑기
                    string temp_time = "";
                    string temp_delivery_result_editdate = dtable.Rows[i]["delivery_result_editdate"].ToString();
                    if (temp_delivery_result_editdate != "")
                        temp_time = DateTime.Parse(temp_delivery_result_editdate).ToString("MMdd HH:mm") + " ";
                    else
                        temp_time = temp_delivery_result_editdate;
                    #endregion

                    _strStatus = dtable.Rows[i]["card_delivery_status"].ToString();
                    //2012.06.28 태희철 수정 재발송 데이터 중 동의서는 취급안하는 걸로 확인
                    strDong = dtable.Rows[i]["card_consented"].ToString();

                    _strLine = new StringBuilder(GetStringAsLength(temp_time, 11, true, ' '));//날자
                    _strLine.Append(GetStringAsLength("FD", 2, true, ' '));
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_send_number"].ToString(), 17, true, ' ')); //배송일련번호
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_express_code"].ToString(), 2, true, ' ')); //배송차수
                    _strLine.Append(GetStringAsLength("02", 2, true, ' '));                                             //업체코드

                    string _strStatusTemp = "00";  //배송 결과
                    //2012-01-17 태희철 정리
                    //temp_0, temp_1, temp_2 : 총건수 1:배송, 2:반송
                    //차수마감

                    if (temp == 3)
                    {
                        if (_strStatus == "1" || _strStatus == "7")
                        {
                            if (dtable.Rows[i]["receiver_code"].ToString() == "98")
                            {
                                _strStatusTemp = "05";
                            }
                            else
                            {
                                _strStatusTemp = "00";
                            }
                        }
                        else if (_strStatus == "2" || _strStatus == "3")
                        {
                            _strStatusTemp = "07";
                        }
                        else
                        {
                            _strStatusTemp = "03";
                        }
                    }
                    else
                    {
                        if (_strStatus == "1" || _strStatus == "7")
                        {
                            if (dtable.Rows[i]["receiver_code"].ToString() == "98")
                            {
                                _strStatusTemp = "05";
                            }
                            else
                            {
                                _strStatusTemp = "00";
                            }
                        }
                        else if (_strStatus == "2" || _strStatus == "3")
                        {
                            _strStatusTemp = "03";
                        }
                    }

                    //if (dtable.Rows[i]["card_kind"].ToString() == "S")
                    //{
                    //    _strStatusTemp = "06";
                    //}
                    if (_strStatusTemp == "05" && dtable.Rows[i]["card_issue_detail_code"].ToString() == "3")
                        _strLine.Append(GetStringAsLength("06", 2, true, ' '));
                    else
                        _strLine.Append(GetStringAsLength(_strStatusTemp, 2, true, ' '));

                    _strLine.Append(GetStringAsLength(temp.ToString(), 1, true, ' '));      //배송결과수신방법코드   
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_client_no_1"].ToString(), 16, true, ' ')); //카드번호
                    //_strLine.Append(GetStringAsLength(dtable.Rows[i]["card_consented"].ToString(), 1, true, ' ')); //동의서징구 결과

                    //2012.06.28 태희철 수정
                    if (strDong == "1")
                    {
                        _strLine.Append("0"); //동의서징구 결과
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
                        _strLine.Append(GetStringAsLength("", 13, true, ' '));
                        _strLine.Append(GetStringAsLength("", 40, true, ' ')); //수령인 이름
                    }

                    if (_strStatus == "2" || _strStatus == "3")
                    {
                        _strLine.Append(GetStringAsLength("", 2, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["return_code_change"].ToString(), 2, true, ' '));
                    }
                    else
                    {
                        if (_strStatusTemp != "05")
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_code_change"].ToString(), 2, true, ' '));
                        else
                            _strLine.Append(GetStringAsLength("", 2, true, ' '));

                        _strLine.Append(GetStringAsLength("", 2, true, ' '));
                    }

                    if (_strStatusTemp == "05")
                    {
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_SSN"].ToString(), 14, true, ' '));
                        // 농협등기요금
                        //2012.11.20 태희철 수정 등기요금 1960 -> 2200
                        //2014.02.03 태희철 수정 등기요금 2230 -> 2440
                        //인수일이 01월29일부터 변경
                        //_strLine.Append(GetStringAsLength("2230", 6, true, ' '));
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
                    else
                    {
                        _strLine.Append(GetStringAsLength("", 14, true, ' '));
                        _strLine.Append(GetStringAsLength("", 6, true, ' '));
                    }

                    if (dtable.Rows[i]["card_delivery_place_type"].ToString() == "5")
                    {
                        if (strCard_zipcode_kind == "1")
                        {
                            _strLine.Append(GetStringAsLength("card_zipcode_new", 6, true, ' '));
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength("card_zipcode", 6, true, ' '));
                        }
                        
                        _strLine.Append(GetStringAsLength("card_address", 200, true, ' '));
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength("", 6, true, ' '));
                        _strLine.Append(GetStringAsLength("", 200, true, ' '));
                    }

                    // 2012-01-25 태희철 수정
                    if (_strStatus == "1")
                    {
                        temp_1++;
                        _sw01.WriteLine(_strLine);
                    }
                    else if (_strStatus == "2" || _strStatus == "3")
                    {
                        temp_2++;
                        _sw02.WriteLine(_strLine);
                    }
                    else
                    {
                        temp_0++;
                        _sw00.WriteLine(_strLine);
                    }
                }

                _strLine = new StringBuilder(GetStringAsLength("", 11, true, ' '));
                _strLine.Append(GetStringAsLength("FT", 2, true, ' '));
                _strLine.Append(GetStringAsLength(temp_1.ToString(), 8, false, '0'));
                _strLine.Append(GetStringAsLength("", 325, true, ' '));

                _sw01.WriteLine(_strLine);

                _strLine = new StringBuilder(GetStringAsLength("", 11, true, ' '));
                _strLine.Append(GetStringAsLength("FT", 2, true, ' '));
                _strLine.Append(GetStringAsLength(temp_2.ToString(), 8, false, '0'));
                _strLine.Append(GetStringAsLength("", 325, true, ' '));

                _sw02.WriteLine(_strLine);

                _strLine = new StringBuilder(GetStringAsLength("", 11, true, ' '));
                _strLine.Append(GetStringAsLength("FT", 2, true, ' '));
                _strLine.Append(GetStringAsLength(temp_0.ToString(), 8, false, '0'));
                _strLine.Append(GetStringAsLength("", 325, true, ' '));
                _sw00.WriteLine(_strLine);

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

        #region 기타 기능
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
