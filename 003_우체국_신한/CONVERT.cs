﻿using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace _003_우체국_신한
{
	public class CONVERT
	{
		//기본 인코딩 설정
		private static string strEncoding = "ks_c_5601-1987";
        private static string strCardTypeID = "003";
        private static string strCardTypeName = "우체국_신한카드";

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
            StreamReader _sr = null;												//파일 읽기 스트림
            StreamWriter _swError = null;											//파일 쓰기 스트림
            DataSet _dsetZipcode = null, _dsetZipcdeArea = null;					//우편번호 관련 DataSet
            DataSet _dsetZipcode_new = null, _dsetZipcdeArea_new = null;					//우편번호 관련 DataSet
            DataTable _dtable = null;												//마스터 저장 테이블
            DataRow _dr = null;
            DataRow[] _drs = null;
            byte[] _byteAry = null;
            string _strReturn = "";
            string _strLine = "";
            string _strZipcode = "", _strAreaType = "", _strAreaGroup = "", _strBranch = "", strCard_type_detail = "", strClient_express_code = "";
            string _strDeliveryPlaceType = "", strMemo = "", strCustomer_ssn = "";
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
                _dtable.Columns.Add("client_quick_work_date");      //특송작업일자
                _dtable.Columns.Add("card_design_code");            //발송업체코드
                _dtable.Columns.Add("client_send_number");          //특송순번
                _dtable.Columns.Add("card_count");                  //카드매수 
                _dtable.Columns.Add("client_register_date");        //제작일자
                _dtable.Columns.Add("client_number");               //dr[10]제작순번

                //_dtable.Columns.Add("client_number");               //카드봉투번호
                _dtable.Columns.Add("check_org");                   //등기번호
                //_dtable.Columns.Add("client_number");               //우체국구분
                //_dtable.Columns.Add("client_number");               //종류별구분
                //_dtable.Columns.Add("client_number");               //중량
                //_dtable.Columns.Add("client_number");               //특수취급코드



                _dtable.Columns.Add("customer_name");               //이름
                _dtable.Columns.Add("card_zipcode");                //우편번호
                _dtable.Columns.Add("card_address_local");          //주소 동이상
                _dtable.Columns.Add("card_address_detail");         //dr[15]주소 동이하


                /*
                _dtable.Columns.Add("card_address");                //dr[12] 접수금액
                _dtable.Columns.Add("card_address");                //dr[12] 묶음번호
                _dtable.Columns.Add("card_address");                //dr[12] 묶음번호순번
                _dtable.Columns.Add("card_address");                //dr[12] 연계일자
                _dtable.Columns.Add("card_address");                //dr[12] 부서코드
                _dtable.Columns.Add("card_address");                //dr[12] 과명
                _dtable.Columns.Add("card_address");                //dr[12] 외부고객접수ID
                _dtable.Columns.Add("card_address");                //dr[12] 특송여부
                _dtable.Columns.Add("card_address");                //dr[12] 우편물내역
                _dtable.Columns.Add("card_address");                //dr[12] 반송필요무
                _dtable.Columns.Add("card_address");                //dr[12] SMS수신거부여부
                _dtable.Columns.Add("card_address");                //dr[12] 카드유혈별분류
                */

                //_dtable.Columns.Add("card_register_type");          //dr[15] 긴급배송여부 코드 "U"
                _dtable.Columns.Add("customer_SSN");                //생년월일
                _dtable.Columns.Add("customer_office");             //직장명
                _dtable.Columns.Add("card_mobile_tel");             //휴대폰                
                _dtable.Columns.Add("card_cooperation_code");       //상품종류명(제휴코드) 

                //_dtable.Columns.Add("card_product_name");           //SMS수신

                _dtable.Columns.Add("family_name");                 //dr[20] 가족이름
                _dtable.Columns.Add("family_SSN");                  //가족식별번호
                _dtable.Columns.Add("card_cooperation1");           //BPRECC번호
                _dtable.Columns.Add("card_cooperation2");           //dr[23]BPRECC엔코딩값
                _dtable.Columns.Add("card_product_code");           //동의서양식코드
                _dtable.Columns.Add("choice_agree1");               //dr[25] 필수동의
                _dtable.Columns.Add("choice_agree3");               //선택동의
                _dtable.Columns.Add("card_is_for_owner_only");      //본인만배송

                _dtable.Columns.Add("client_express_code");         //일반,긴급,동의 구분코드
                _dtable.Columns.Add("card_zipcode_new");            //신우편번호
                _dtable.Columns.Add("card_zipcode_kind");           //dr[30] 신우편번호구분코드
                _dtable.Columns.Add("card_delivery_place_code");    //수령지코드 1자택, 2직장
                _dtable.Columns.Add("card_barcode_new");            //케리어바코드


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

                //우편번호 관련 정보 DataSet에 담기
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

                    _byteAry = _encoding.GetBytes(_strLine);

                    if (_encoding.GetString(_byteAry, 0, 2) == "HD" || _encoding.GetString(_byteAry, 0, 2) == "TR")
                    {
                        continue;
                    }
                    //_swError.WriteLine(_strLine);
                    _strDeliveryPlaceType = _encoding.GetString(_byteAry, 45, 1);

                    _dr = _dtable.NewRow();
                    _dr[0] = _iSeq;

                    _dr[5] = _encoding.GetString(_byteAry, 2, 8);
                    _dr[6] = _encoding.GetString(_byteAry, 10, 2);
                    _dr[7] = _encoding.GetString(_byteAry, 12, 7);
                    _dr[8] = _encoding.GetString(_byteAry, 19, 2);
                    _dr[9] = _encoding.GetString(_byteAry, 21, 8);
                    _dr[10] = _encoding.GetString(_byteAry, 29, 5);
                    //_dr[22] = _encoding.GetString(_byteAry, 34, 6);       //카드봉투번호
                    //_dr[22] = _encoding.GetString(_byteAry, 40, 13);      //등기번호
                    //_dr[22] = _encoding.GetString(_byteAry, 53, 5);       //우체국구분
                    //_dr[22] = _encoding.GetString(_byteAry, 58, 3);       //종류별구분
                    //_dr[22] = _encoding.GetString(_byteAry, 61, 7);       //중량
                    //_dr[22] = _encoding.GetString(_byteAry, 68, 3);       //특수취급코드
                    _dr[11] = _encoding.GetString(_byteAry, 40, 13);      //등기번호
                    _dr[12] = _encoding.GetString(_byteAry, 71, 40);

                    _strZipcode = _encoding.GetString(_byteAry, 151, 6).Replace(" ", "").Trim();
                    _dr[13] = _strZipcode;

                    if (_strZipcode.Trim().Length == 5)
                    {
                        _dr[29] = _strZipcode;
                        _dr[30] = "1";
                    }
                    _dr[14] = _encoding.GetString(_byteAry, 157, 100);
                    _dr[15] = _encoding.GetString(_byteAry, 257, 100);
                    //_dr[12] = _encoding.GetString(_byteAry, 357, 15);
                    //_dr[12] = _encoding.GetString(_byteAry, 372, 15);
                    //_dr[12] = _encoding.GetString(_byteAry, 387, 4);
                    //_dr[12] = _encoding.GetString(_byteAry, 391, 8);
                    //_dr[12] = _encoding.GetString(_byteAry, 399, 6);
                    //_dr[12] = _encoding.GetString(_byteAry, 405, 50);
                    //_dr[12] = _encoding.GetString(_byteAry, 455, 13);
                    //_dr[13] = _encoding.GetString(_byteAry, 468, 2);

                    //_dr[12] = _encoding.GetString(_byteAry, 470, 60);
                    //_dr[12] = _encoding.GetString(_byteAry, 530, 1);
                    //_dr[12] = _encoding.GetString(_byteAry, 531, 1);
                    //_dr[12] = _encoding.GetString(_byteAry, 532, 1);

                    strCustomer_ssn = _encoding.GetString(_byteAry, 533, 7).Replace('X', 'x').Replace('*', 'x');
                    if (strCustomer_ssn.Trim().Length == 0)
                    {
                        strCustomer_ssn = "xxxxxxxxxxxxx";
                    }
                    else
                    {
                        strCustomer_ssn = strCustomer_ssn + "xxxxxx";
                    }

                    _dr[16] = strCustomer_ssn;
                    //_dr[14] = _encoding.GetString(_byteAry, 540, 6);
                    //_dr[14] = _encoding.GetString(_byteAry, 546, 15);
                    _dr[17] = _encoding.GetString(_byteAry, 561, 30);
                    _dr[18] = _encoding.GetString(_byteAry, 591, 12);
                    _dr[19] = _encoding.GetString(_byteAry, 634, 20);
                    //_dr[16] = _encoding.GetString(_byteAry, 654, 97);
                    //_dr[16] = _encoding.GetString(_byteAry, 751, 1);
                    _dr[20] = _encoding.GetString(_byteAry, 752, 80);

                    strCustomer_ssn = _encoding.GetString(_byteAry, 832, 7).Replace('X', 'x').Replace('*', 'x');
                    if (strCustomer_ssn.Trim().Length == 0)
                    {
                        strCustomer_ssn = "xxxxxxxxxxxxx";
                    }
                    else
                    {
                        strCustomer_ssn = strCustomer_ssn + "xxxxxx";
                    }

                    _dr[21] = strCustomer_ssn; //가족생년월일
                    //_dr[18] = _encoding.GetString(_byteAry, 839, 6);
                    _dr[22] = _encoding.GetString(_byteAry, 845, 25);
                    _dr[23] = _encoding.GetString(_byteAry, 870, 70); //최대
                    _dr[24] = _encoding.GetString(_byteAry, 945, 4);

                    if (_encoding.GetString(_byteAry, 949, 1) == "1" || _encoding.GetString(_byteAry, 949, 1) == "0")
                    {
                        _dr[25] = _encoding.GetString(_byteAry, 949, 1);
                    }
                    else
                    {
                        _dr[25] = "9";
                    }

                    _dr[26] = _encoding.GetString(_byteAry, 950, 10);

                    string strowner = "";

                    if (_encoding.GetString(_byteAry, 960, 1) == "Y")
                    {
                        strowner = "1";
                    }
                    _dr[27] = strowner;


                    if (strCard_type_detail.Substring(0, 5) == "00311")
                    {
                        strClient_express_code = "2002";
                    }
                    else if (strCard_type_detail.Substring(0, 5) == "00321")
                    {
                        strClient_express_code = "2120";
                    }
                    else if (strCard_type_detail.Substring(0, 4) == "0033")
                    {
                        strClient_express_code = "2005";
                    }

                    _dr[28] = strClient_express_code;
                    _dr[31] = "3";
                    _dr[32] = _dr[11];
                    ///캐리어바코드 생성
                    ///[내용] 신한의 경우 카드번호가 12자리와 11자리 두 분류로 되어있음
                    ///11자리 일 경우 중간에 공백이 생겨 바코드 리딩이 안됨.
                    ///공백을 제거하고 업데이트 하여 케리어 바코드를 22자리로 생성
                    ///리딩 할 경우에도 공백을 제거 하므로 22자리로 인식 가능
                    //_dr[27] = "0" + strClient_express_code + _dr[6].ToString().Trim() + _strZipcode;


                    if (_strZipcode != "")
                    {
                        //지역 분류 선택
                        if (_strZipcode.Trim().Length == 5)
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
                        if (_strZipcode.Trim().Length == 5)
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
            //파일 쓰기 스트림
            StreamWriter _sw01 = null, _sw00 = null, _sw11 = null, _sw10 = null;

            int i = 0, iCnt = 0, iRe_cnt = 0;
            StringBuilder _strLine = new StringBuilder("");
            string _strReturn = "", strStatus = "", _strReturnCode = "", strGetdate = "", strCard_in_date = "", strCard_type_detail = "";
            string strCard_in_date_chk = "";
            DataRow[] _drs = null;
            string[] strCheck_num_array = null;

            try
            {
                strGetdate = DateTime.Now.ToString("MMdd");

                _sw10 = new StreamWriter(fileName + "sh(" + strGetdate + ").txt.00", true, _encoding);
                _sw11 = new StreamWriter(fileName + "sh(" + strGetdate + ").txt.01", true, _encoding);

                _drs = dtable.Select("", "delivery_result_editdate");

                //헤더 부분
                string temp_head = GetStringAsLength("HDKJ" + DateTime.Now.ToString("yyyyMMdd"), 12, false, ' ');

                _sw11.WriteLine(GetStringAsLength(temp_head, 300, true, ' '));

                for (i = 0; i < _drs.Length; i++)
                {
                    strStatus = _drs[i]["card_delivery_status"].ToString();
                    strCard_in_date = String.Format("{0:MMdd}", _drs[i]["card_in_date"]);
                    strCard_in_date_chk = String.Format("{0:yyMMdd}", _drs[i]["card_in_date"]);
                    strCard_type_detail = _drs[i]["card_type_detail"].ToString();

                    _strReturnCode = _drs[i]["delivery_return_reason_last"].ToString();

                    _strLine = new StringBuilder(GetStringAsLength(_drs[i]["card_number"].ToString(), 12, true, ' '));
                    _strLine.Append(GetStringAsLength(_drs[i]["card_brand_code"].ToString(), 1, true, ' '));
                    _strLine.Append(GetStringAsLength("", 3, true, ' '));

                    DateTime CardInDate = DateTime.Parse(_drs[i]["card_in_date"].ToString());
                    DateTime dt_date = DateTime.Parse("2019-11-01");

                    #region 배송

                    if (strStatus == "1")
                    {
                        if ((_drs[i]["receiver_code_change"].ToString() == "001") ||
                            (_drs[i]["receiver_code"].ToString() == "01"))
                        {
                            _strLine.Append(GetStringAsLength("Y1", 2, true, ' '));
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength("Y2", 2, true, ' '));
                        }

                        _strLine.Append(GetStringAsLength(RemoveDash(_drs[i]["card_delivery_date"].ToString()), 8, true, ' '));

                        //민증번호
                        if (_drs[i]["card_result_status"].ToString() == "61")
                        {
                            _strLine.Append(GetStringAsLength(_drs[i]["customer_ssn"].ToString().Substring(2, 4), 14, true, ' '));
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength(ConvertLGSSN(_drs[i]["receiver_SSN"].ToString().Replace("x", "0")), 14, true, '0'));
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
                    }
                    #endregion

                    #region 반송
                    else if (strStatus == "2" || strStatus == "3")
                    {
                        _strLine.Append(GetStringAsLength(ReturnType(_strReturnCode), 2, true, ' '));
                        _strLine.Append(GetStringAsLength(RemoveDash(_drs[i]["delivery_return_date_last"].ToString()), 8, true, ' '));
                        _strLine.Append(GetStringAsLength("", 14, true, ' '));
                        _strLine.Append(GetStringAsLength("", 15, true, ' '));
                        if (_drs[i]["card_issue_type"].ToString() == "5")
                        {
                            _strLine.Append(GetStringAsLength(RemoveDash(_drs[i]["client_send_date"].ToString()), 8, true, ' '));
                            _strLine.Append(GetStringAsLength("", 5, true, ' '));
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
                        _strLine.Append(GetStringAsLength("", 20, true, ' '));
                        _strLine.Append(GetStringAsLength("", 3, false, ' '));
                        _strLine.Append(GetStringAsLength(" ", 1, true, ' '));
                    }
                    #endregion

                    //결번은 DB에서 필터됨

                    #region 분실
                    else if (strStatus == "6")
                    {
                        _strLine.Append(GetStringAsLength("LL", 2, true, ' '));
                        _strLine.Append(GetStringAsLength(RemoveDash(_drs[i]["delivery_result_regdate"].ToString()), 8, true, ' '));
                        _strLine.Append(GetStringAsLength("", 14, true, ' '));
                        _strLine.Append(GetStringAsLength("", 15, true, ' '));
                        if (_drs[i]["card_issue_type"].ToString() == "5")
                        {
                            _strLine.Append(GetStringAsLength(RemoveDash(_drs[i]["client_send_date"].ToString()), 8, true, ' '));
                            _strLine.Append(GetStringAsLength("", 5, true, ' '));
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
                        _strLine.Append(GetStringAsLength("", 20, true, ' '));
                        _strLine.Append(GetStringAsLength("", 3, false, ' '));
                        _strLine.Append(GetStringAsLength(" ", 1, true, ' '));
                    }
                    #endregion

                    #region 재방
                    else if (strStatus == "7")
                    {
                        _strLine.Append(GetStringAsLength("JB", 2, true, ' '));
                        _strLine.Append(GetStringAsLength("", 8, true, ' '));
                        _strLine.Append(GetStringAsLength("", 14, true, ' '));
                        _strLine.Append(GetStringAsLength("", 15, true, ' '));
                        if (_drs[i]["card_issue_type"].ToString() == "5")
                        {
                            _strLine.Append(GetStringAsLength(RemoveDash(_drs[i]["client_send_date"].ToString()), 8, true, ' '));
                            _strLine.Append(GetStringAsLength("", 5, true, ' '));
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
                        _strLine.Append(GetStringAsLength("", 20, true, ' '));
                        _strLine.Append(GetStringAsLength("030", 3, false, ' '));
                        _strLine.Append(GetStringAsLength(" ", 1, true, ' '));
                    }
                    #endregion

                    else
                    {
                        _strLine.Append(GetStringAsLength("", 2, true, ' '));
                        _strLine.Append(GetStringAsLength("", 8, true, ' '));
                        _strLine.Append(GetStringAsLength("", 14, true, ' '));
                        _strLine.Append(GetStringAsLength("", 15, true, ' '));

                        if (_drs[i]["client_register_date"].ToString() == "")
                            _strLine.Append(GetStringAsLength(_drs[i]["client_send_date"].ToString().Replace("-", ""), 8, true, ' '));  //제작일자
                        else
                            _strLine.Append(GetStringAsLength(_drs[i]["client_register_date"].ToString().Replace("-", ""), 8, true, ' '));  //제작일자

                        _strLine.Append(GetStringAsLength(_drs[i]["client_number"].ToString(), 5, true, ' '));                          //제작순번

                        if (_drs[i]["client_quick_work_date"].ToString() == "")
                            _strLine.Append(GetStringAsLength(_drs[i]["card_in_date"].ToString().Replace("-", ""), 8, true, ' '));//특송접수일자
                        else
                            _strLine.Append(GetStringAsLength(_drs[i]["client_quick_work_date"].ToString().Replace("-", ""), 8, true, ' '));//특송접수일자

                        _strLine.Append(GetStringAsLength(_drs[i]["client_send_number"].ToString(), 6, true, ' '));                     //특송접수번호

                        _strLine.Append(GetStringAsLength("", 20, true, ' '));
                        _strLine.Append(GetStringAsLength("", 3, false, ' '));
                        _strLine.Append(GetStringAsLength(" ", 1, true, ' '));

                    }

                    if (strStatus == "1")
                    {
                        _strLine.Append(GetStringAsLength("Y", 1, true, ' '));
                        _strLine.Append(GetStringAsLength("Y", 1, true, ' '));
                        _strLine.Append(GetStringAsLength("Y", 1, true, ' '));
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength("", 1, true, ' '));
                        _strLine.Append(GetStringAsLength("", 1, true, ' '));
                        _strLine.Append(GetStringAsLength("", 1, true, ' '));
                    }

                    _strLine.Append(GetStringAsLength("", 2, true, ' '));
                    _strLine.Append(GetStringAsLength("", 7, true, ' '));

                    if (strStatus == "1" && _drs[i]["change_add"].ToString() == "1")
                    {
                        _strLine.Append(GetStringAsLength(_drs[i]["code"].ToString(), 1, true, ' '));

                        switch (_drs[i]["code"].ToString())
                        {
                            case "1":
                            case "4":
                            case "6":
                                _strLine.Append(GetStringAsLength(RemoveDash(_drs[i]["date"].ToString()), 15, true, ' '));       //발급일자
                                _strLine.Append(GetStringAsLength(_drs[i]["org"].ToString(), 30, true, ' '));                    //발급기관
                                break;
                            case "2":
                                if (_drs[i]["number"].ToString() != "")
                                {
                                    strCheck_num_array = _drs[i]["number"].ToString().Split('(');
                                    //_strLine.Append(GetStringAsLength(RemoveDash(strCheck_num_array[1].Substring(0, 2)), 2, true, ' '));
                                    _strLine.Append(GetStringAsLength(RemoveDash(strCheck_num_array[1].Substring(strCheck_num_array[1].IndexOf(")") + 2, strCheck_num_array[1].Length - 4)), 15, true, ' '));
                                    _strLine.Append(GetStringAsLength(_drs[i]["org"].ToString(), 30, true, ' '));                    //발급기관
                                }
                                else
                                {
                                    _strLine.Append(GetStringAsLength("", 15, true, ' '));
                                    _strLine.Append(GetStringAsLength("", 30, true, ' '));
                                }
                                break;
                            case "3":
                                _strLine.Append(GetStringAsLength(RemoveDash(_drs[i]["number"].ToString()), 15, true, ' '));     //여권번호
                                _strLine.Append(GetStringAsLength(_drs[i]["org"].ToString(), 30, true, ' '));                    //발급기관
                                break;
                            default:
                                _strLine.Append(GetStringAsLength("", 15, true, ' '));
                                _strLine.Append(GetStringAsLength("", 30, true, ' '));
                                break;
                        }
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength("", 1, true, ' '));
                        _strLine.Append(GetStringAsLength("", 15, true, ' '));
                        _strLine.Append(GetStringAsLength("", 30, true, ' '));
                    }

                    //태블릿 동의값
                    if (CardInDate > dt_date && strStatus == "1" && _drs[i]["card_type_detail"].ToString().Substring(0, 4) == "0032")
                    {
                        _strLine.Append(GetStringAsLength("Y", 1, true, ' '));                                          //정보제공고객서명여부
                        _strLine.Append(GetStringAsLength("1", 1, true, ' '));                                          //이용필수개인정보동의구분
                        if (_drs[i]["chk_02"].ToString() == "9")
                        {
                            _strLine.Append(GetStringAsLength("0", 1, true, ' '));                    //조회개인정보동의구분    
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength("1", 1, true, ' '));                    //조회개인정보동의구분
                        }
                        _strLine.Append(GetStringAsLength("1", 1, true, ' '));                                          //제공필수개인정보동의구분
                        _strLine.Append(GetStringAsLength("1", 1, true, ' '));                                          //상품부가서비스개인정보동의구분
                        _strLine.Append(GetStringAsLength("1", 1, true, ' '));                                          //상품고유식별개인정보동의구분
                        _strLine.Append(GetStringAsLength("Y", 1, true, ' '));                                          //필수동의고객명서명여부

                        _strLine.Append(GetStringAsLength(_drs[i]["chk_06"].ToString(), 1, true, ' '));                 //이용선택개인정보동의구분
                        _strLine.Append(GetStringAsLength(_drs[i]["chk_07"].ToString(), 1, true, ' '));                 //이용서비스안내개인정보동의

                        _strLine.Append(GetStringAsLength(_drs[i]["chkex_02"].ToString(), 1, true, ' '));               //이용권유전화
                        _strLine.Append(GetStringAsLength(_drs[i]["chkex_03"].ToString(), 1, true, ' '));               //이용권유SMS
                        _strLine.Append(GetStringAsLength(_drs[i]["chkex_04"].ToString(), 1, true, ' '));               //이용권유서면
                        _strLine.Append(GetStringAsLength(_drs[i]["chkex_05"].ToString(), 1, true, ' '));               //이용권유Email
                        _strLine.Append(GetStringAsLength(_drs[i]["chk_08"].ToString(), 1, true, ' '));                 //이용식별번호선택정보
                        _strLine.Append(GetStringAsLength("Y", 1, true, ' '));                                          //이용고객명서명여부

                        _strLine.Append(GetStringAsLength(_drs[i]["chk_09"].ToString(), 1, true, ' '));                 //제공신한그룹개인정보
                        _strLine.Append(GetStringAsLength(_drs[i]["chk_10"].ToString(), 1, true, ' '));                 //제공부정방지개인정보
                        _strLine.Append(GetStringAsLength(_drs[i]["chk_11"].ToString(), 1, true, ' '));                 //제공고유식별개인정보
                        _strLine.Append(GetStringAsLength("Y", 1, true, ' '));                                          //제공고객명서명여부
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength("", 19, true, ' '));
                    }



                    //일반
                    if (strCard_type_detail.Substring(0, 4) == "0031")
                    {
                        _sw00 = new StreamWriter(fileName + "KJ" + strCard_in_date + "일반마감.txt.00", true, _encoding);
                        _sw01 = new StreamWriter(fileName + "KJ" + strCard_in_date + "일반마감.txt.01", true, _encoding);
                    }
                    //동의서
                    else if (strCard_type_detail.Substring(0, 4) == "0032")
                    {
                        _sw00 = new StreamWriter(fileName + "KJ" + strCard_in_date + "동의서마감.txt.00", true, _encoding);
                        _sw01 = new StreamWriter(fileName + "KJ" + strCard_in_date + "동의서마감.txt.01", true, _encoding);
                    }
                    //긴급
                    else if (strCard_type_detail.Substring(0, 4) == "0033")
                    {
                        _sw00 = new StreamWriter(fileName + "KJ" + strCard_in_date + "긴급.txt.00", true, _encoding);
                        _sw01 = new StreamWriter(fileName + "KJ" + strCard_in_date + "긴급.txt.01", true, _encoding);
                    }
                    //기프트
                    else if (strCard_type_detail.Substring(0, 4) == "0034")
                    {
                        _sw00 = new StreamWriter(fileName + "SHG" + strGetdate + ".txt.00", true, _encoding);
                        _sw01 = new StreamWriter(fileName + "SHG" + strGetdate + ".txt.01", true, _encoding);
                    }

                    if (strStatus == "0" || strStatus == "7")
                    {
                        _sw00.WriteLine(_strLine.ToString());
                    }
                    else
                    {
                        _sw01.WriteLine(_strLine.ToString());
                    }

                    if (_sw00 != null) _sw00.Close();
                    if (_sw01 != null) _sw01.Close();
                    //2013.07.22 태희철 [E] 구마감 끝


                    //2013.07.22 태희철 수정 [S] 신마감사용
                    _strLine = new StringBuilder("DT");
                    //카드번호
                    _strLine.Append(GetStringAsLength(_drs[i]["card_number"].ToString().Replace("-", ""), 16, true, ' '));

                    //배송
                    if (strStatus == "1")
                    {
                        if ((_drs[i]["receiver_code_change"].ToString() == "001") ||
                            (_drs[i]["receiver_code"].ToString() == "01"))
                        {
                            _strLine.Append(GetStringAsLength("Y1", 2, true, ' '));
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength("Y2", 2, true, ' '));
                        }
                    }
                    //재방
                    else if (strStatus == "7")
                        _strLine.Append(GetStringAsLength("JB", 2, true, ' '));
                    //결번
                    else if (strStatus == "" || strStatus == "4" || strStatus == "6")
                        _strLine.Append(GetStringAsLength("LL", 2, true, ' '));
                    //반송
                    else if (strStatus == "2" || strStatus == "3")
                        _strLine.Append(GetStringAsLength(ReturnType(_strReturnCode), 2, true, ' '));
                    else //기타
                        _strLine.Append(GetStringAsLength("", 2, true, ' '));

                    //배송
                    if (strStatus == "1")
                        _strLine.Append(GetStringAsLength(_drs[i]["card_delivery_date"].ToString().Replace("-", ""), 8, true, ' '));    //전달일자
                    //반송
                    else if (strStatus == "2" || strStatus == "3")
                        _strLine.Append(GetStringAsLength(_drs[i]["delivery_return_date_last"].ToString().Replace("-", ""), 8, true, ' '));
                    else if (strStatus == "6")
                        _strLine.Append(GetStringAsLength(_drs[i]["delivery_result_regdate"].ToString().Replace("-", ""), 8, true, ' '));
                    else
                        _strLine.Append(GetStringAsLength("", 8, true, ' '));    //전달일자


                    //생년월일
                    if (_drs[i]["card_result_status"].ToString() == "61")
                    {
                        _strLine.Append(GetStringAsLength(_drs[i]["customer_ssn"].ToString().Substring(2, 4), 14, true, ' '));
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength(ConvertLGSSN(_drs[i]["receiver_SSN"].ToString().Replace("x", "0")), 14, true, '0'));
                    }
                    //_strLine.Append(GetStringAsLength(_drs[i]["receiver_SSN"].ToString().Replace("-", "").Replace("x", ""), 14, true, ' '));         //민증번호


                    _strLine.Append(GetStringAsLength(_drs[i]["receiver_tel"].ToString().Replace("-", ""), 15, true, ' '));         //전화번호

                    if (_drs[i]["client_register_date"].ToString() == "")
                    {
                        _strLine.Append(GetStringAsLength(_drs[i]["client_send_date"].ToString().Replace("-", ""), 8, true, ' '));  //제작일자
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength(_drs[i]["client_register_date"].ToString().Replace("-", ""), 8, true, ' '));  //제작일자
                    }
                    _strLine.Append(GetStringAsLength(_drs[i]["client_number"].ToString(), 5, true, ' '));                          //제작순번

                    if (_drs[i]["client_quick_work_date"].ToString() == "")
                        _strLine.Append(GetStringAsLength(_drs[i]["card_in_date"].ToString().Replace("-", ""), 8, true, ' '));//특송접수일자
                    else
                        _strLine.Append(GetStringAsLength(_drs[i]["client_quick_work_date"].ToString().Replace("-", ""), 8, true, ' '));//특송접수일자

                    _strLine.Append(GetStringAsLength(_drs[i]["client_send_number"].ToString(), 6, true, ' '));                     //특송접수번호

                    if (strStatus == "1" || _drs[i]["card_issue_type"].ToString() == "5")
                        _strLine.Append(GetStringAsLength(_drs[i]["receiver_name"].ToString(), 40, true, ' '));
                    else
                        _strLine.Append(GetStringAsLength("", 40, true, ' '));             //수령인성명

                    if (strStatus == "1") //배송
                        _strLine.Append(GetStringAsLength(_drs[i]["receiver_code_change"].ToString(), 3, true, ' '));                   //관계코드 - 은행 요청 코드
                    else
                        _strLine.Append(GetStringAsLength("", 3, true, ' '));

                    _strLine.Append(GetStringAsLength("", 1, true, ' '));                                                           //예비

                    if (strStatus == "1")
                    {
                        _strLine.Append(GetStringAsLength(ConvertAgree(_drs[i]["card_agree1"].ToString()), 1, true, ' '));
                        _strLine.Append(GetStringAsLength(ConvertAgree(_drs[i]["card_agree2"].ToString()), 1, true, ' '));
                        _strLine.Append(GetStringAsLength(ConvertAgree(_drs[i]["card_agree3"].ToString()), 1, true, ' '));
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength("", 1, true, ' '));
                        _strLine.Append(GetStringAsLength("", 1, true, ' '));
                        _strLine.Append(GetStringAsLength("", 1, true, ' '));
                    }

                    _strLine.Append(GetStringAsLength(_drs[i]["card_client_no_1"].ToString(), 2, true, ' '));                       //특송발송카드 BIN구분코드
                    _strLine.Append(GetStringAsLength(_drs[i]["client_express_code"].ToString(), 4, true, ' '));                    //제휴사코드

                    if (strStatus == "1" && _drs[i]["change_add"].ToString() == "1")
                    {
                        _strLine.Append(GetStringAsLength(_drs[i]["code"].ToString(), 1, true, ' '));

                        switch (_drs[i]["code"].ToString())
                        {
                            case "1":
                            case "4":
                            case "6":
                                _strLine.Append(GetStringAsLength(RemoveDash(_drs[i]["date"].ToString()), 15, true, ' '));     //여권번호
                                _strLine.Append(GetStringAsLength(_drs[i]["org"].ToString(), 30, true, ' '));                    //발급기관
                                break;
                            case "2":
                                if (_drs[i]["number"].ToString() != "")
                                {
                                    strCheck_num_array = _drs[i]["number"].ToString().Split('(');
                                    //_strLine.Append(GetStringAsLength(RemoveDash(strCheck_num_array[1].Substring(0, 2)), 2, true, ' '));
                                    _strLine.Append(GetStringAsLength(RemoveDash(strCheck_num_array[1].Substring(strCheck_num_array[1].IndexOf(")") + 2, strCheck_num_array[1].Length - 4)), 15, true, ' '));
                                    _strLine.Append(GetStringAsLength(_drs[i]["org"].ToString(), 30, true, ' '));                    //발급기관
                                }
                                else
                                {
                                    _strLine.Append(GetStringAsLength("", 15, true, ' '));
                                    _strLine.Append(GetStringAsLength("", 30, true, ' '));
                                }
                                break;
                            case "3":
                                _strLine.Append(GetStringAsLength(RemoveDash(_drs[i]["number"].ToString()), 15, true, ' '));     //여권번호
                                _strLine.Append(GetStringAsLength(_drs[i]["org"].ToString(), 30, true, ' '));                    //발급기관
                                break;
                            default:

                                _strLine.Append(GetStringAsLength("", 15, true, ' '));
                                _strLine.Append(GetStringAsLength("", 30, true, ' '));
                                break;
                        }
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength("", 1, true, ' '));
                        _strLine.Append(GetStringAsLength("", 15, true, ' '));
                        _strLine.Append(GetStringAsLength("", 30, true, ' '));
                    }

                    //태블릿 동의값
                    if (CardInDate > dt_date && _drs[i]["card_type_detail"].ToString().Substring(0, 4) == "0032")
                    {
                        _strLine.Append(GetStringAsLength("Y", 1, true, ' '));                                          //정보제공고객서명여부
                        _strLine.Append(GetStringAsLength("1", 1, true, ' '));                                          //이용필수개인정보동의구분
                        if (_drs[i]["chk_02"].ToString() == "9")
                        {
                            _strLine.Append(GetStringAsLength("0", 1, true, ' '));                    //조회개인정보동의구분    
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength("1", 1, true, ' '));                    //조회개인정보동의구분
                        }
                        _strLine.Append(GetStringAsLength("1", 1, true, ' '));                                          //제공필수개인정보동의구분
                        _strLine.Append(GetStringAsLength("1", 1, true, ' '));                                          //상품부가서비스개인정보동의구분
                        _strLine.Append(GetStringAsLength("1", 1, true, ' '));                                          //상품고유식별개인정보동의구분
                        _strLine.Append(GetStringAsLength("Y", 1, true, ' '));                                          //필수동의고객명서명여부

                        _strLine.Append(GetStringAsLength(_drs[i]["chk_06"].ToString(), 1, true, ' '));                 //이용선택개인정보동의구분
                        _strLine.Append(GetStringAsLength(_drs[i]["chk_07"].ToString(), 1, true, ' '));                 //이용서비스안내개인정보동의

                        _strLine.Append(GetStringAsLength(_drs[i]["chkex_02"].ToString(), 1, true, ' '));               //이용권유전화
                        _strLine.Append(GetStringAsLength(_drs[i]["chkex_03"].ToString(), 1, true, ' '));               //이용권유SMS
                        _strLine.Append(GetStringAsLength(_drs[i]["chkex_04"].ToString(), 1, true, ' '));               //이용권유서면
                        _strLine.Append(GetStringAsLength(_drs[i]["chkex_05"].ToString(), 1, true, ' '));               //이용권유Email
                        _strLine.Append(GetStringAsLength(_drs[i]["chk_08"].ToString(), 1, true, ' '));                 //이용식별번호선택정보
                        _strLine.Append(GetStringAsLength("Y", 1, true, ' '));                                          //이용고객명서명여부

                        _strLine.Append(GetStringAsLength(_drs[i]["chk_09"].ToString(), 1, true, ' '));                 //제공신한그룹개인정보
                        _strLine.Append(GetStringAsLength(_drs[i]["chk_10"].ToString(), 1, true, ' '));                 //제공부정방지개인정보
                        _strLine.Append(GetStringAsLength(_drs[i]["chk_11"].ToString(), 1, true, ' '));                 //제공고유식별개인정보
                        _strLine.Append(GetStringAsLength("Y", 1, true, ' '));                                          //제공고객명서명여부
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength("", 19, true, ' '));
                    }

                    _strLine.Append(GetStringAsLength("", 98, true, ' '));                                                     //예비

                    if (strStatus == "0" || strStatus == "7")
                    {
                        iRe_cnt++;
                        _sw10.WriteLine(_strLine.ToString());
                    }
                    else
                    {
                        iCnt++;
                        _sw11.WriteLine(_strLine.ToString());
                    }
                }

                //2013.07.22 태희철 수정 [S] 신마감사용
                _strLine = new StringBuilder(GetStringAsLength("TR" + GetStringAsLength(iCnt.ToString(), 11, false, '0'), 300, true, ' '));
                _sw11.WriteLine(_strLine.ToString());
                //2013.07.22 태희철 수정 [E] 신마감사용

                _strReturn = string.Format("총 {0}건 / 결과완료 {1}건 / 미배송 {2}건 다운 완료", i, iCnt, iRe_cnt);

            }
            catch (Exception)
            {
                _strReturn = string.Format("{0}번째 데이터 인계 중 오류 발생", i + 1);
            }
            finally
            {
                if (_sw00 != null) _sw00.Close();
                if (_sw01 != null) _sw01.Close();
                if (_sw10 != null) _sw10.Close();
                if (_sw11 != null) _sw11.Close();
            }
            return _strReturn;
        }

        //일일마감
        public static string ConvertResultDay(System.Data.DataTable dtable, string fileName)
        {   
            return "일일마감자료 다운은 사용하실 수 없습니다.";
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

        private static string ConvertLGSSN(string value)
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

        #region 반송사유코드
        public static string ReturnType(string value)
        {
            string _strReturn = value;

            switch (value)
            {   
                case "5":
                case "6":
                case "13":
                case "17":
                case "25":
                case "28":
                case "31":
                case "34":
                case "98":
                    _strReturn = "GG";
                    break;
                case "3":
                case "20":
                case "21":
                case "88":
                    _strReturn = "BB";
                    break;
                case "1":
                case "12":
                case "18":
                case "19":
                case "23":
                case "24":
                case "33":
                    _strReturn = "CC";
                    break;
                case "2":
                case "11":
                    _strReturn = "FF";
                    break;
                case "8":
                    _strReturn = "JJ";
                    break;
                case "9":
                case "10":
                    _strReturn = "DD";
                    break;
                case "30":
                    _strReturn = "ZZ";
                    break;
                default:
                    _strReturn = "E1";
                    break;
            }
            return _strReturn;
        }
        #endregion        


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
