using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace _001_BC_TEST
{
	public class CONVERT
	{
		//기본 인코딩 설정
		private static string strEncoding = "ks_c_5601-1987";
        private static string strCardTypeID = "001";
        private static string strCardTypeName = "001_BC카드_TEST";
        private static char chCSV = ',';
        public static int _iReturn = 0;

		//현 DLL의 카드 타입 코드 반환
		public static string GetCardTypeID() 
        {
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
                    _strReturn = _strLine.Substring(_strLine.Length - 7, 7);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }

            return _strReturn;
        }

        public static string ConvertRegister(string path, string xmlZipcodePath, string xmlZipcodeAreaPath, string xmlZipcodePath_new, string xmlZipcodeAreaPath_new, string xmlPath)
        {
            System.Text.Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            //FileInfo _fi = null;
            StreamReader _sr = null;													//파일 읽기 스트림
            StreamWriter _swError = null;						    					//파일 쓰기 스트림
            DataSet _dsetZipcode = null, _dsetZipcdeArea = null;						//우편번호 관련 DataSet
            DataSet _dsetZipcode_new = null, _dsetZipcdeArea_new = null;				//우편번호 관련 DataSet
            DataTable _dtable = null;													//마스터 저장 테이블
            DataTable _dtable_NH = null;
            DataRow _dr = null;
            DataRow _dr_NH = null;
            DataRow[] _drs = null;
            DataRow[] _drs_NH = null;

            byte[] _byteAry = null;
            string _strReturn = "";
            string _strLine = "";
            string _strZipcode = "", _strAreaType = "", _strAreaGroup = "", _strBranch = "";
            string _strGubun = "", _strBankID = "", strValue = "", strNewAddress_Value = "", strCard_type_detail = "", strOwner_olny = "", strDong_chk = "", strBank_Chk = "";
            string strClient_Branch = "";
            int _iSeq = 1, _iErrorCount = 0;
            int _iGubunError = 0;
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
                _dtable.Columns.Add("client_send_date");                      // dr[5]
                _dtable.Columns.Add("card_bank_ID");
                _dtable.Columns.Add("card_number");
                _dtable.Columns.Add("customer_name");
                _dtable.Columns.Add("customer_SSN");
                _dtable.Columns.Add("client_send_number");                    // dr[10] 발송번호(케리어바코드사용)
                _dtable.Columns.Add("client_number");
                _dtable.Columns.Add("card_count");
                _dtable.Columns.Add("card_customer_regdate");                 //카드등록일
                _dtable.Columns.Add("card_issue_type_code");                  // 카드발급구분
                _dtable.Columns.Add("card_tel1");                             // dr[15]
                _dtable.Columns.Add("card_tel2");
                _dtable.Columns.Add("card_mobile_tel");
                _dtable.Columns.Add("card_zipcode");                          // 우편번호(케리어바코드사용)
                _dtable.Columns.Add("card_delivery_place_code");
                _dtable.Columns.Add("card_address_detail");                   // dr[20]
                _dtable.Columns.Add("customer_type_code");                    // 소지자구분 : 3,4 대리만 배송가능
                _dtable.Columns.Add("client_release_register");               // 개시유무
                _dtable.Columns.Add("card_traffic_code");                     // 본인배송구분
                _dtable.Columns.Add("card_is_for_owner_only");                // 고객발송구분 (본인특송)
                _dtable.Columns.Add("card_urgency_code");                     // dr[25] 
                _dtable.Columns.Add("card_agree_code");                       // 동의서구분
                _dtable.Columns.Add("card_pt_code");
                _dtable.Columns.Add("client_bank_request_no");                // 회원사(은행)코드(케리어바코드)
                _dtable.Columns.Add("save_agreement");
                _dtable.Columns.Add("card_request_memo");                     //dr[30]
                _dtable.Columns.Add("client_register_date");                  //dr[31] 생성일(케리어바코드사용)
                _dtable.Columns.Add("card_vip_code");
                _dtable.Columns.Add("delivery_limit_day");
                _dtable.Columns.Add("client_register_type");                  // 발송코드
                _dtable.Columns.Add("card_product_code");                     // dr[35] 일반,지점구분
                _dtable.Columns.Add("card_cooperation2");                     // 비씨경남 바코드
                _dtable.Columns.Add("card_barcode_new");

                //2013-05-15 태희철 추가 신주소관련[S]
                _dtable.Columns.Add("card_address_type1");                     // dr[38] 신주소구분자
                _dtable.Columns.Add("card_address_local");                     // 신주소(동이상)
                //내부코드관련
                _dtable.Columns.Add("card_issue_type_new");          // dr[40] 발급구분코드_new
                _dtable.Columns.Add("card_delivery_place_type");     // dr[41] 배송지구분 내부코드
                //2013-05-15 [E]

                //2014.09.03 태희철 동의서변경 관련
                _dtable.Columns.Add("card_brand_code");                       // dr[42] 상품제휴코드
                _dtable.Columns.Add("card_product_name");                     // dr[43] 상품제휴코드명
                // 제휴서비스 문구 : 제휴처명^제공목적^제공항목^보유기간
                _dtable.Columns.Add("text1");
                _dtable.Columns.Add("text2");                          //dr[45]
                _dtable.Columns.Add("text3");
                _dtable.Columns.Add("text4");
                _dtable.Columns.Add("text5");
                _dtable.Columns.Add("text6");
                _dtable.Columns.Add("text7");
                _dtable.Columns.Add("text8");
                _dtable.Columns.Add("text9");
                _dtable.Columns.Add("text10");                     // dr[53]
                _dtable.Columns.Add("card_client_no_1");
                _dtable.Columns.Add("client_request_memo");        //dr[55] 메모
                _dtable.Columns.Add("card_zipcode_new");           //dr[56] 새우편번호
                _dtable.Columns.Add("card_zipcode_kind");          //dr[57] 우편번호구분

                _dtable.Columns.Add("customer_order");             //dr[58] 메모코드
                _dtable.Columns.Add("customer_memo");              //dr[59] 메모문구
                _dtable.Columns.Add("change_address");             //dr[60] 수령지변경 주소
                _dtable.Columns.Add("change_zipcode");             //dr[61] 수령지변경 우편번호
                _dtable.Columns.Add("choice_agree1");              //dr[62] 동의서필수항목사전인쇄여부
                _dtable.Columns.Add("card_design_code");           //dr[63] 배송유형코드
                _dtable.Columns.Add("card_consented");              //dr[64] 우리은행별지
                _dtable.Columns.Add("card_bank_account_tel");      //dr[65] 실번호 뒷4자리
                _dtable.Columns.Add("change_add");                 //dr[66] 신분증

                _dtable.Columns.Add("card_bank_account_owner");    //dr[67] 카드사바코드 NPI용

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
                    if (_iSeq == 1)
                    {
                        strCard_type_detail = _strLine.Substring(_strLine.Length - 7, 7);
                    }

                    //인코딩, byte 배열로 담기
                    _byteAry = _encoding.GetBytes(_strLine);

                    _strBankID = _encoding.GetString(_byteAry, 8, 2);
                    strBank_Chk = _encoding.GetString(_byteAry, 250, 3);

                    _dr = _dtable.NewRow();
                    _dr[0] = _iSeq;

                    _dr[5] = _encoding.GetString(_byteAry, 0, 8);
                    _dr[6] = _encoding.GetString(_byteAry, 8, 6);
                    _dr[7] = _encoding.GetString(_byteAry, 14, 16).Replace("*", "x");
                    _dr[8] = _encoding.GetString(_byteAry, 30, 40);
                    _dr[9] = _encoding.GetString(_byteAry, 70, 13).Replace("*", "x").Replace(" ", "x");
                    _dr[10] = _encoding.GetString(_byteAry, 83, 8);
                    _dr[11] = _dr[5].ToString() + _dr[10].ToString();
                    _dr[12] = _encoding.GetString(_byteAry, 91, 3);
                    _dr[13] = _encoding.GetString(_byteAry, 94, 8);
                    _dr[14] = _encoding.GetString(_byteAry, 102, 1);        //발급구분

                    if (RemoveBlank(_encoding.GetString(_byteAry, 2753, 12)).Length > 8)
                    {
                        _dr[15] = RemoveBlank(_encoding.GetString(_byteAry, 2753, 12));
                    }
                    else
                    {
                        _dr[15] = RemoveBlank(_encoding.GetString(_byteAry, 103, 12));
                    }

                    _dr[16] = RemoveBlank(_encoding.GetString(_byteAry, 115, 12));

                    //2017.04.26 태희철수정 전화번호(가상번호)
                    if (RemoveBlank(_encoding.GetString(_byteAry, 2741, 12)).Length > 8)
                    {
                        _dr[17] = RemoveBlank(_encoding.GetString(_byteAry, 2741, 12));
                    }
                    else
                    {
                        _dr[17] = RemoveBlank(_encoding.GetString(_byteAry, 127, 12));
                    }

                    _strZipcode = _encoding.GetString(_byteAry, 139, 6).Trim();
                    _dr[18] = _strZipcode;

                    if (_strZipcode.Trim().Length == 5)
                    {
                        _dr[56] = _strZipcode.Trim();
                        _dr[57] = "1";
                    }

                    //일반_영업점 구분
                    strClient_Branch = _encoding.GetString(_byteAry, 145, 1);
                    _dr[19] = strClient_Branch;

                    _dr[20] = _encoding.GetString(_byteAry, 146, 95);       // 주소_detail
                    _dr[21] = _encoding.GetString(_byteAry, 241, 1);        // 소지자구분
                    _dr[22] = _encoding.GetString(_byteAry, 242, 1);
                    _dr[23] = _encoding.GetString(_byteAry, 244, 1);

                    //일반 중 본인만배송건
                    strOwner_olny = _encoding.GetString(_byteAry, 245, 1);

                    //2019.12.23 태희철 수정
                    //_dr[21] = _encoding.GetString(_byteAry, 241, 1);        // 소지자구분 3, 4는 대리만배송가능
                    if (strCard_type_detail.Substring(0, 5) == "00131" || (strCard_type_detail.Substring(0, 4) != "0012" && (_dr[21].ToString() == "3" || _dr[21].ToString() == "4")))
                    {
                        _dr[24] = "2";
                    }
                    //2019.12.12 태희철 수정
                    //else if (strCard_type_detail.Substring(0, 5) != "00135" &&
                    //        strCard_type_detail.Substring(0, 5) != "00113" &&
                    //        strCard_type_detail.Substring(0, 5) != "00141" && 
                    //        _dr[21].ToString() == "4")
                    //{
                    //    _dr[24] = "2";
                    //}
                    else if (strOwner_olny == "1" || strOwner_olny == "3")
                    {
                        _dr[24] = "1";
                    }
                    else if (strClient_Branch != "2" && strClient_Branch != "3" && strClient_Branch != "4" && (strCard_type_detail == "0011109" || strCard_type_detail == "0011110" || strCard_type_detail == "0011111" || strCard_type_detail == "0011112" || strCard_type_detail == "0011113" || strCard_type_detail == "0013202" || strCard_type_detail == "0013204" || strCard_type_detail == "0013303" || strCard_type_detail == "0011402"))
                    {
                        _dr[24] = "1";
                    }
                    else
                    {
                        _dr[24] = "0";
                    }

                    _dr[25] = _encoding.GetString(_byteAry, 246, 1);
                    strDong_chk = _encoding.GetString(_byteAry, 247, 1).Trim();
                    _dr[26] = strDong_chk;
                    _dr[27] = _encoding.GetString(_byteAry, 248, 2);
                    _dr[28] = strBank_Chk;
                    _dr[29] = _encoding.GetString(_byteAry, 253, 1);
                    _dr[30] = _encoding.GetString(_byteAry, 254, 34);
                    _dr[31] = _encoding.GetString(_byteAry, 288, 8);
                    _dr[32] = _encoding.GetString(_byteAry, 296, 1);
                    _dr[33] = _encoding.GetString(_byteAry, 297, 2);
                    _dr[34] = _encoding.GetString(_byteAry, 299, 1);

                    ///
                    /// 
                    /// *주의* CONVERT 시 생기는 하드코딩 : 0A, 0B, 0Z
                    /// 
                    /// 
                    //_strGubun = _encoding.GetString(_byteAry, 1741, 2).Trim();
                    //_strGubun = _encoding.GetString(_byteAry, 1766, 2).Trim();
                    _strGubun = _encoding.GetString(_byteAry, 2766, 2).Trim();

                    _dr[35] = _strGubun;



                    _dr[36] = _encoding.GetString(_byteAry, 300, 40).Trim();

                    strNewAddress_Value = _encoding.GetString(_byteAry, 341, 1);
                    _dr[38] = strNewAddress_Value;

                    //NEW내부구분코드
                    _dr[40] = _dr[14];  //발급구분
                    _dr[41] = _dr[19];  //배송지구분

                    _dr[42] = _encoding.GetString(_byteAry, 542, 6);  //상품제휴코드
                    _dr[43] = _encoding.GetString(_byteAry, 548, 40);  //상품제휴코드명

                    //제휴처1~10
                    _dr[44] = _encoding.GetString(_byteAry, 590, 40) + "^" + _encoding.GetString(_byteAry, 630, 1) + "^"
                        + _encoding.GetString(_byteAry, 631, 150) + "^" + _encoding.GetString(_byteAry, 781, 1);

                    _dr[45] = _encoding.GetString(_byteAry, 782, 40) + "^" + _encoding.GetString(_byteAry, 822, 1) + "^"
                        + _encoding.GetString(_byteAry, 823, 150) + "^" + _encoding.GetString(_byteAry, 973, 1);

                    _dr[46] = _encoding.GetString(_byteAry, 974, 40) + "^" + _encoding.GetString(_byteAry, 1014, 1) + "^"
                        + _encoding.GetString(_byteAry, 1015, 150) + "^" + _encoding.GetString(_byteAry, 1165, 1);

                    _dr[47] = _encoding.GetString(_byteAry, 1166, 40) + "^" + _encoding.GetString(_byteAry, 1206, 1) + "^"
                        + _encoding.GetString(_byteAry, 1207, 150) + "^" + _encoding.GetString(_byteAry, 1357, 1);

                    _dr[48] = _encoding.GetString(_byteAry, 1358, 40) + "^" + _encoding.GetString(_byteAry, 1398, 1) + "^"
                        + _encoding.GetString(_byteAry, 1399, 150) + "^" + _encoding.GetString(_byteAry, 1549, 1);

                    _dr[49] = _encoding.GetString(_byteAry, 1550, 40) + "^" + _encoding.GetString(_byteAry, 1590, 1) + "^"
                        + _encoding.GetString(_byteAry, 1591, 150) + "^" + _encoding.GetString(_byteAry, 1741, 1);

                    _dr[50] = _encoding.GetString(_byteAry, 1742, 40) + "^" + _encoding.GetString(_byteAry, 1782, 1) + "^"
                        + _encoding.GetString(_byteAry, 1783, 150) + "^" + _encoding.GetString(_byteAry, 1933, 1);

                    _dr[51] = _encoding.GetString(_byteAry, 1934, 40) + "^" + _encoding.GetString(_byteAry, 1974, 1) + "^"
                        + _encoding.GetString(_byteAry, 1975, 150) + "^" + _encoding.GetString(_byteAry, 2125, 1);

                    _dr[52] = _encoding.GetString(_byteAry, 2126, 40) + "^" + _encoding.GetString(_byteAry, 2166, 1) + "^"
                        + _encoding.GetString(_byteAry, 2167, 150) + "^" + _encoding.GetString(_byteAry, 2317, 1);

                    _dr[53] = _encoding.GetString(_byteAry, 2318, 40) + "^" + _encoding.GetString(_byteAry, 2358, 1) + "^"
                        + _encoding.GetString(_byteAry, 2359, 150) + "^" + _encoding.GetString(_byteAry, 2509, 1);

                    _dr[54] = _encoding.GetString(_byteAry, 2510, 20);


                    if (strCard_type_detail.Substring(0, 5) == "00131")
                    {
                        ; //2019.07.02 영업점배송의 경우 본인지정문구 및 신분증정보확인 제외
                    }
                    else if (strOwner_olny == "1" || strOwner_olny == "3")
                    {
                        _dr[59] = "";

                        if (strDong_chk.Trim() == "")
                        {
                            _dr[66] = "1";
                        }
                        else
                        {
                            _dr[66] = "";
                        }

                    }
                    else if (strClient_Branch != "2" && strClient_Branch != "3" && strClient_Branch != "4" && (strCard_type_detail == "0011109" || strCard_type_detail == "0011110" || strCard_type_detail == "0011111" || strCard_type_detail == "0011112" || strCard_type_detail == "0011113" || strCard_type_detail == "0013202" || strCard_type_detail == "0013204" || strCard_type_detail == "0013303" || strCard_type_detail == "0011402"))
                    {
                        //_dr[59] = "본인지정배송(신분증확인)"; //2019.07.22 제거
                        _dr[59] = "";
                    }
                    else if ((strClient_Branch == "2" || strClient_Branch == "3" || strClient_Branch == "4") && (strCard_type_detail == "0011109" || strCard_type_detail == "0011110" || strCard_type_detail == "0011111" || strCard_type_detail == "0011112" || strCard_type_detail == "0011113" || strCard_type_detail == "0013202" || strCard_type_detail == "0013204" || strCard_type_detail == "0013303" || strCard_type_detail == "0011402"))
                    {
                        _dr[55] = "대리수령가능";
                    }
                    else if ((strClient_Branch == "2" || strClient_Branch == "3" || strClient_Branch == "4") && strCard_type_detail == "0012108")
                    {
                        _dr[55] = "영업점";
                    }
                    else
                    {
                        _dr[59] = "";
                        _dr[66] = "";
                    }

                    _dr[58] = _encoding.GetString(_byteAry, 2530, 2).Trim();

                    if (_dr[59].ToString() != "")
                    {
                        if (_dr[58].ToString() == "01")
                        {
                            _dr[59] = _dr[59].ToString() + " / 수령지변경";
                        }
                        else if (_dr[58].ToString() == "02")
                        {
                            _dr[59] = _dr[59].ToString() + " / 배송 전 전화요청";
                        }
                        else if (_dr[58].ToString() == "03")
                        {
                            _dr[59] = _dr[59].ToString() + " / 수령지변경 및 배송 전 전화요청";
                        }
                    }
                    else
                    {
                        if (_dr[58].ToString() == "01")
                        {
                            _dr[59] = "수령지변경";
                        }
                        else if (_dr[58].ToString() == "02")
                        {
                            _dr[59] = "배송 전 전화요청";
                        }
                        else if (_dr[58].ToString() == "03")
                        {
                            _dr[59] = "수령지변경 및 배송 전 전화요청";
                        }
                        else
                        {
                            _dr[59] = _dr[58].ToString();
                        }

                    }

                    _dr[60] = _encoding.GetString(_byteAry, 2532, 100).Trim() + " " + _encoding.GetString(_byteAry, 2632, 100).Trim();
                    _dr[61] = _encoding.GetString(_byteAry, 2732, 6);

                    //2015.12.21 태희철 적용
                    _dr[62] = _encoding.GetString(_byteAry, 2738, 1);
                    //2017.02.24 태희철 적용 배송유형코드
                    _dr[63] = _encoding.GetString(_byteAry, 2739, 2);
                    //2017.08.11 태희철 적용 배송유형코드
                    _dr[64] = _encoding.GetString(_byteAry, 2765, 1);
                    //2018.05.24 태희철 적용 실번호 뒷 4자리
                    _dr[65] = _encoding.GetString(_byteAry, 135, 4);

                    //End dr[66]

                    //신주소일 경우
                    //1=구주소+구우편, 2:신주소+구우편, 3=구주소+신우편, 4=신주소+신우편
                    if (strNewAddress_Value != "1")
                    {
                        _dr[39] = _encoding.GetString(_byteAry, 342, 100).Replace("　", " ");
                        _dr[20] = _encoding.GetString(_byteAry, 442, 100).Replace("　", " ");
                    }

                    string strBank_code = "", strIssue_code = "";
                    //2017.04.17 은행코드
                    strBank_code = _encoding.GetString(_byteAry, 250, 3);
                    //발급코드
                    strIssue_code = _encoding.GetString(_byteAry, 102, 1);

                    //2012-04-20 태희철 수정 카드사바코드[S]
                    //[23]card_traffic_code // [24]card_is_for_owner_only
                    if (strCard_type_detail.Substring(0, 5) == "00121" || strCard_type_detail.Substring(0, 5) == "00123")
                    {
                        strValue = "13";
                    }
                    else if (strCard_type_detail.Substring(0, 5) == "00131")
                    {
                        strValue = "20";
                    }
                    else if (_encoding.GetString(_byteAry, 244, 1) == "1" || _encoding.GetString(_byteAry, 245, 1).ToString() == "1")
                    {
                        strValue = "10";
                    }
                    //[33]delivery_limit_day
                    else if (_dr[33].ToString() == "01")
                    {
                        // bankCode = 03,06,11,20,21,23,25,31,32,36,39,45,50,95 = true
                        if (bCard_Bank(strBank_Chk))
                        {
                            strValue = "11";
                        }
                        else
                        {
                            strValue = "99";
                        }
                    }
                    //[33]delivery_limit_day
                    else if (_dr[33].ToString() == "02")
                    {
                        if (strBank_Chk == "020" || strBank_Chk == "084" || strBank_Chk == "023")
                        {
                            strValue = "12";
                        }
                        else
                        {
                            strValue = "99";
                        }
                    }
                    //갱신
                    else if (_dr[14].ToString() == "3")
                    {
                        //2017.04.17 광주은행 갱신건의 경우 99로 예외처리
                        if (strBank_code == "034")
                        {
                            strValue = "99";
                        }
                        else
                        {
                            strValue = "14";
                        }
                    }
                    else
                    {
                        strValue = "99";
                    }

                    if (strBank_Chk == "0AA")
                    {
                        strBank_code = "037";
                    }
                    else if (strBank_Chk == "0AB")
                    {
                        strBank_code = "0GM";
                    }

                    //2012-04-20 태희철 수정 카드사바코드[E]
                    //client_register_date, client_send_number, Zipcode, client_bank_request_no
                    if (_dr[57].ToString() == "1")
                    {
                        //_dr[37] = _dr[31].ToString() + _dr[10].ToString() + "0" + _strZipcode.Trim() + strBank_code + strValue + _strGubun;
                        _dr[37] = _dr[31].ToString() + _dr[10].ToString() + "0" + _strZipcode.Trim() + strBank_code;
                        _dr[67] = _dr[31].ToString() + _dr[10].ToString() + "0" + _strZipcode.Trim() + strBank_code + strValue + _strGubun;
                    }
                    else
                    {
                        //_dr[37] = _dr[31].ToString() + _dr[10].ToString() + _strZipcode + strBank_code + strValue + _strGubun;
                        _dr[37] = _dr[31].ToString() + _dr[10].ToString() + _strZipcode + strBank_code;
                        _dr[67] = _dr[31].ToString() + _dr[10].ToString() + _strZipcode + strBank_code + strValue + _strGubun;
                    }
                    //2012-04-20 태희철 수정 카드사바코드[E]

                    if (_strZipcode != "")
                    {
                        //지역 분류 선택
                        if (_strZipcode.Trim().Length == 5)
                        {
                            _drs = _dsetZipcdeArea_new.Tables[0].Select("zipcode_new = " + _strZipcode.Trim());
                        }
                        else
                        {
                            _drs = _dsetZipcdeArea.Tables[0].Select("zipcode = " + _strZipcode.Trim());
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
                            _drs = _dsetZipcode_new.Tables[0].Select("zipcode_new = " + _strZipcode.Trim());
                        }
                        else
                        {
                            _drs = _dsetZipcode.Tables[0].Select("zipcode = " + _strZipcode.Trim());
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

                    //카드사바코드에 적용되는 내용이므로 필수사항
                    if (_strGubun == "" || (_strGubun != "0A" && _strGubun != "0B" && _strGubun != "0Z" && _strGubun != "0I" && _strGubun != "0J"))
                    {
                        _swError.WriteLine(_strLine);
                        _iGubunError++;
                        throw new ArgumentNullException("0A 또는 0B 또는 0Z 자리 오류 : 원본 총 byte와 비교하세요.");
                    }
                    _iSeq++;
                }

                //변환에 성공했다면 비씨 동의서 재정렬 
                if ((strCard_type_detail.Substring(0, 5) == "00121" || strCard_type_detail.Substring(0, 5) == "00123") && _iErrorCount < 1 && _iGubunError < 1)
                {
                    _swError.Close();
                    _sr.Close();
                    try
                    {
                        switch (strBank_Chk)
                        {
                            //비씨 제휴 농협의 경우 정렬을 다시 함
                            case "011":
                            case "012":
                                _drs_NH = _dtable.Select("", "card_bank_ID");
                                _dtable_NH = new DataTable("Convert2");

                                //기본 컬럼
                                _dtable_NH.Columns.Add("degree_arrange_number");
                                _dtable_NH.Columns.Add("card_area_group");
                                _dtable_NH.Columns.Add("card_branch");
                                _dtable_NH.Columns.Add("card_area_type");
                                _dtable_NH.Columns.Add("area_arrange_number");
                                //세부 컬럼				
                                _dtable_NH.Columns.Add("client_send_date");                 //  dr[5]
                                _dtable_NH.Columns.Add("card_bank_ID");
                                _dtable_NH.Columns.Add("card_number");
                                _dtable_NH.Columns.Add("customer_name");
                                _dtable_NH.Columns.Add("customer_SSN");
                                _dtable_NH.Columns.Add("client_send_number");            //  dr[10]
                                _dtable_NH.Columns.Add("client_number");
                                _dtable_NH.Columns.Add("card_count");
                                _dtable_NH.Columns.Add("card_customer_regdate");
                                _dtable_NH.Columns.Add("card_issue_type_code");
                                _dtable_NH.Columns.Add("card_tel1");                           //  dr[15]
                                _dtable_NH.Columns.Add("card_tel2");
                                _dtable_NH.Columns.Add("card_mobile_tel");
                                _dtable_NH.Columns.Add("card_zipcode");
                                _dtable_NH.Columns.Add("card_delivery_place_code");
                                _dtable_NH.Columns.Add("card_address_detail");             // dr[20]        
                                _dtable_NH.Columns.Add("customer_type_code");
                                _dtable_NH.Columns.Add("client_release_register");
                                _dtable_NH.Columns.Add("card_traffic_code");
                                _dtable_NH.Columns.Add("card_is_for_owner_only");
                                _dtable_NH.Columns.Add("card_urgency_code");                 // dr[25]
                                _dtable_NH.Columns.Add("card_agree_code");
                                _dtable_NH.Columns.Add("card_pt_code");
                                _dtable_NH.Columns.Add("client_bank_request_no");
                                _dtable_NH.Columns.Add("save_agreement");                     // dr[29]
                                _dtable_NH.Columns.Add("card_request_memo");                // dr[30]
                                _dtable_NH.Columns.Add("client_register_date");             // dr[31] 생성일(케리어바코드사용)
                                _dtable_NH.Columns.Add("card_vip_code");
                                _dtable_NH.Columns.Add("delivery_limit_day");
                                _dtable_NH.Columns.Add("client_register_type");
                                _dtable_NH.Columns.Add("card_product_code");                // dr[35]
                                _dtable_NH.Columns.Add("card_cooperation2");                // 경남은행 전용 바코드
                                //2012-04-20 태희철 추가
                                _dtable_NH.Columns.Add("card_barcode_new");                 // dr[37] 카드사바코드
                                //2013.06.27 태희철 추가
                                _dtable_NH.Columns.Add("card_address_type1");               // 신구주소 구분자 
                                _dtable_NH.Columns.Add("card_address_local");               // dr[39] 주소_Local
                                _dtable_NH.Columns.Add("card_issue_type_new");              // dr[40] NEW발급구분
                                _dtable_NH.Columns.Add("card_delivery_place_type");         // dr[41] 배송지구분 내부코드

                                //2014.09.03 태희철 동의서변경 관련
                                _dtable_NH.Columns.Add("card_brand_code");                       // dr[42] 상품제휴코드
                                _dtable_NH.Columns.Add("card_product_name");                     // dr[43] 상품제휴코드명
                                // 제휴서비스 문구 : 제휴처명^제공목적^제공항목^보유기간
                                _dtable_NH.Columns.Add("text1");
                                _dtable_NH.Columns.Add("text2");                          //dr[45]
                                _dtable_NH.Columns.Add("text3");
                                _dtable_NH.Columns.Add("text4");
                                _dtable_NH.Columns.Add("text5");
                                _dtable_NH.Columns.Add("text6");
                                _dtable_NH.Columns.Add("text7");
                                _dtable_NH.Columns.Add("text8");
                                _dtable_NH.Columns.Add("text9");
                                _dtable_NH.Columns.Add("text10");                        // dr[53]
                                _dtable_NH.Columns.Add("card_client_no_1");
                                _dtable_NH.Columns.Add("client_request_memo");           //dr[55] 메모
                                _dtable_NH.Columns.Add("card_zipcode_new");              //dr[56] 메모
                                _dtable_NH.Columns.Add("card_zipcode_kind");              //dr[57] 메모

                                _dtable_NH.Columns.Add("customer_order");             //dr[58] 메모코드
                                _dtable_NH.Columns.Add("customer_memo");              //메모문구
                                _dtable_NH.Columns.Add("change_address");             //dr[60] 수령지변경 주소
                                _dtable_NH.Columns.Add("change_zipcode");             //dr[61] 수령지변경 주소
                                _dtable_NH.Columns.Add("choice_agree1");             //dr[62] 수령지변경 주소
                                _dtable_NH.Columns.Add("card_design_code");          //dr[63] 배송유형코드
                                _dtable_NH.Columns.Add("card_consented");            //dr[64] 우리은행별지
                                _dtable_NH.Columns.Add("card_bank_account_tel");     //dr[65] 실번호 뒷4자리
                                _dtable_NH.Columns.Add("change_add");                 //dr[66] 신분증

                                _dtable_NH.Columns.Add("card_bank_account_owner");    //dr[67] 카드사바코드 NPI용

                                _iSeq = 1;
                                _branches.Clear();
                                for (int i = 0; i < _drs_NH.Length; i++)
                                {
                                    _dr_NH = _dtable_NH.NewRow();
                                    for (int k = 1; k < _drs_NH[i].ItemArray.Length; k++)
                                    {
                                        _dr_NH[0] = _iSeq;

                                        if (k == 4)
                                        {
                                            _dr_NH[k] = _branches.GetCount(_drs_NH[i].ItemArray[2].ToString());
                                        }
                                        else
                                        {
                                            _dr_NH[k] = _drs_NH[i].ItemArray[k].ToString();
                                        }
                                    }
                                    _dtable_NH.Rows.Add(_dr_NH);
                                    _iSeq++;
                                }
                                _dtable_NH.WriteXml(xmlPath);
                                break;
                            //비씨농협외 제휴
                            default:
                                _dtable.WriteXml(xmlPath);
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    _strReturn = string.Format("{0}건의 데이터 변환 성공", _iSeq - 1);
                }
                else if (_iErrorCount < 1 && _iGubunError < 1)
                {
                    _swError.Close();
                    _sr.Close();
                    _dtable.WriteXml(xmlPath);
                    _strReturn = string.Format("{0}건의 데이터 변환 성공", _iSeq - 1);
                }
                else
                {
                    if (_iGubunError > 0)
                    {
                        _strReturn = string.Format("{0}건 변환, 구분값 미존재 {1}건 실패", _iSeq - 1, _iGubunError);
                    }
                    else if (_iErrorCount > 0)
                    {
                        _strReturn = string.Format("{0}건 변환, 우편번호 미등록 {1}건 실패, ", _iSeq - 1, _iErrorCount);
                    }
                    else
                    {
                        _strReturn = string.Format("{0}건 변환, 우편번호 미등록 {1}건 구분값 미존재 {2}건 실패, ", _iSeq - 1, _iErrorCount, _iGubunError);
                    }
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

        private static bool bCard_Bank(string strCard_bank)
        {
            // strCard_Bank = client_bank_request_no.ToString().Substring(1,2)
            bool bBank;

            switch (strCard_bank)
            {
                case "03":
                case "10":
                case "11":
                case "12":
                case "13":
                case "14":
                case "15":
                case "16":
                case "17":
                case "18":
                case "20":
                case "84":
                    bBank = true;
                    break;
                default:
                    bBank = false;
                    break;
            }

            return bBank;
        }

        //마감 자료 생성
        public static string ConvertResult(DataTable dtable, string fileName)
        {
            //int _iReturn = 0;
            string _strReturn = "등록 프로그램 입니다.";
            //FormSelectReceive _f = new FormSelectReceive();
            //if (_f.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            //{
            //    _iReturn = _f.GetSelected;
            //}
            //switch (_iReturn)
            //{   
            //    case 1:
            //        _strReturn = ConvertReceiveType1(dtable, fileName);
            //        break;
            //    case 2:
            //        _strReturn = ConvertReceiveType2(dtable, fileName);
            //        break;
            //    case 3:
            //        _strReturn = ConvertReceiveType3(dtable, fileName);
            //        break;
            //    case 4:
            //        _strReturn = ConvertReceiveType1(dtable, fileName);
            //        break;
            //    default:
            //        _strReturn = "";
            //        break;
            //}
            return _strReturn;
        }

        //일일마감자료
        public static string ConvertResultDay(System.Data.DataTable dtable, string fileName)
        {
            return ConvertResult(dtable, fileName);
        }


        private static string CONERT_BANK_CODE(string strBankCode)
        {
            string strReturn = null;
            //우리은행 : 20, 22, 24, 83, 84 배열값[0]
            //농협중앙회 : 10, 11, 
            //농협지점 : 12, 13, 14, 15, 16, 17, 18 배열값[1]
            //제일 : 23 배열값[2]
            //기업 : 그외
            switch (strBankCode)
            {
                case "20":
                case "22":
                case "24":
                case "83":
                case "84":
                case "90":
                    strReturn = "WOORI";
                    break;
                //농협중앙회
                case "10":
                case "11":
                    strReturn = "NH1";
                    break;
                //농협지점
                case "12":
                case "13":
                case "14":
                case "15":
                case "16":
                case "17":
                case "18":
                    strReturn = "NH2";
                    break;
                case "23":
                    strReturn = "SC";
                    break;
                default:
                    strReturn = "IBK";
                    break;
            }

            return strReturn;
        }

        //결과마감
        private static string ConvertReceiveType1(System.Data.DataTable dtable, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩
            StreamWriter _sw00 = null, _sw00_1 = null;
            StreamWriter _sw01 = null, _sw01_1 = null;
            StreamWriter _sw02 = null, _sw02_1 = null;
            StreamWriter _sw012 = null, _sw012_1 = null, _swError = null;

            StreamWriter _sw00_2 = null, _sw00_3 = null;
            StreamWriter _sw01_2 = null, _sw01_3 = null;
            StreamWriter _sw02_2 = null, _sw02_3 = null;
            StreamWriter _sw012_2 = null, _sw012_3 = null;

            //076-은행긴급 결과파일생성에 사용
            StreamWriter _sw012_0 = null, _sw012_4 = null;
            StreamWriter _sw01_4 = null, _sw02_4 = null;
            StreamWriter _sw10_0 = null, _sw10_1 = null, _sw10_2 = null, _sw10_3 = null, _sw10_4 = null;
            StreamWriter _sw11_0 = null, _sw11_1 = null, _sw11_2 = null, _sw11_3 = null, _sw11_4 = null;
            StreamWriter _sw12_0 = null, _sw12_1 = null, _sw12_2 = null, _sw12_3 = null, _sw12_4 = null;
            StreamWriter _sw112_0 = null, _sw112_1 = null, _sw112_2 = null, _sw112_3 = null, _sw112_4 = null;
            StreamWriter _sw00_0 = null, _sw01_0 = null, _sw02_0 = null, _sw00_4 = null;
            
            //파일 쓰기 스트림
            StringBuilder _strLine = new StringBuilder("");

            string _strStatus = "", _strCardBranch = "";
            string _strReturn = "", _strAreaType = "";
            //2012.06.05 태희철 수정 : 카드구분추가 card_vip_code = 블리스, 인피니트
            //delivery_limit_day 추가
            string _strVIP = null, _strLimitday = null, _strowner_only = null, _str_BC_Part = null, strResult_status = "", strCard_type_detail = "";
            string _strCard_Product_code = "", _strCard_traffic_code = "", _strBankID = "";
            string _strZipCode = "", _strBank = "";
            int i = 0;

            try
            {
                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    _strStatus = dtable.Rows[i]["card_delivery_status"].ToString();
                    _strCardBranch = dtable.Rows[i]["card_area_group"].ToString();
                    _strAreaType = dtable.Rows[i]["card_area_type"].ToString();
                    _str_BC_Part = dtable.Rows[i]["BC_Part"].ToString();
                    strResult_status = dtable.Rows[i]["card_result_status"].ToString();
                    _strZipCode = dtable.Rows[i]["card_zipcode"].ToString();

                    _strVIP = dtable.Rows[i]["card_vip_code"].ToString();
                    _strLimitday = dtable.Rows[i]["delivery_limit_day"].ToString();
                    _strowner_only = dtable.Rows[i]["card_is_for_owner_only"].ToString();
                    strCard_type_detail = dtable.Rows[i]["card_type_detail"].ToString();
                    //076-은행긴급사용
                    _strCard_Product_code = dtable.Rows[i]["card_product_code"].ToString();
                    _strCard_traffic_code = dtable.Rows[i]["card_traffic_code"].ToString();
                    _strBankID = dtable.Rows[i]["card_bank_ID"].ToString();

                    //2012-02-23 태희철 정의 card_issue_type_code : 갱신여부
                    // 갱신관련 갱신배송, 갱신반송, 그외
                    // 갱신마감선택 텝을 생성
                    // 필요한건 건수만 필요
                    // 최종목적은 배송프로그램 [인수일별인수현황] 갱신선택 확인 가능
                    //[인수일별인수현황]에서 분실, 수작업은 인식을 못함

                    // 2012-03-08 태희철 수정 [갱신여부],[수령지장소] 코드 제거 사용안함
                    //_strLine =new StringBuilder(GetStringAsLength(dtable.Rows[i]["card_issue_type_code"].ToString(), 1, true, ' '));
                    //_strLine.Append(GetStringAsLength(dtable.Rows[i]["card_delivery_place_type"].ToString(), 1, true, ' '));

                    if (dtable.Rows[i]["card_bank_ID"].ToString().Length > 1)
                    {
                        _strBank = CONERT_BANK_CODE(dtable.Rows[i]["card_bank_ID"].ToString().Substring(0, 2));
                    }
                    else
                    {
                        _strBank = "";
                    }

                    _strLine = new StringBuilder(GetStringAsLength(RemoveDash(dtable.Rows[i]["client_send_date"].ToString()), 8, true, ' '));
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_bank_ID"].ToString(), 6, true, ' '));
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_number"].ToString().Replace("x", "*"), 16, true, ' '));
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_name"].ToString(), 40, true, ' '));
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_SSN"].ToString(), 13, true, ' '));
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_send_number"].ToString(), 8, true, ' '));
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_count"].ToString(), 3, false, '0'));
                    _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[i]["card_customer_regdate"].ToString()), 8, true, ' '));

                    if (_strStatus == "2" || _strStatus == "3")
                    {
                        _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[i]["delivery_return_date_last"].ToString()), 8, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["return_code_change"].ToString(), 2, false, '0'));
                        _strLine.Append(GetStringAsLength("", 8, true, ' '));
                        _strLine.Append(GetStringAsLength("", 20, true, ' '));
                        _strLine.Append(GetStringAsLength("", 2, true, ' '));
                    }
                    else if (_strStatus == "0")
                    {
                        _strLine.Append(GetStringAsLength("", 8, true, ' '));
                        _strLine.Append(GetStringAsLength("", 2, true, ' '));
                        _strLine.Append(GetStringAsLength("", 8, true, ' '));
                        _strLine.Append(GetStringAsLength("", 20, true, ' '));
                        _strLine.Append(GetStringAsLength("", 2, true, ' '));
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength("", 8, true, ' '));
                        _strLine.Append(GetStringAsLength("", 2, true, ' '));
                        _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[i]["card_delivery_date"].ToString()), 8, true, ' '));

                        if (_strStatus == "7")
                        {
                            _strLine.Append(GetStringAsLength("재배달건", 20, true, ' '));
                            _strLine.Append(GetStringAsLength("99", 2, true, ' '));
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_name"].ToString(), 20, true, ' '));
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_code_change"].ToString().Replace("x", " "), 2, true, ' '));
                        }

                    }

                    _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[i]["client_register_date"].ToString()), 8, true, ' '));
                    if (dtable.Rows[i]["card_kind"].ToString() == "D")
                    {
                        _strLine.Append(GetStringAsLength("1", 1, true, ' '));
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength("0", 1, true, ' '));
                    }

                    //등기번호
                    if (dtable.Rows[i]["card_branch"].ToString() == "012")
                    {
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_SSN"].ToString(), 13, true, ' '));
                        _strLine.Append(GetStringAsLength("", 36, true, ' '));
                    }
                    else
                    {
                        if (_strStatus == "7")
                        {
                            _strLine.Append(GetStringAsLength("", 49, true, ' '));
                        }
                        else
                        {
                            if (dtable.Rows[i]["receiver_SSN"].ToString().Length < 6)
                            {
                                _strLine.Append(GetStringAsLength("", 49, true, ' '));
                            }
                            else if (dtable.Rows[i]["card_result_status"].ToString() == "61")
                            {
                                _strLine.Append(GetStringAsLength("", 13, true, ' '));
                                _strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_SSN"].ToString(), 6, true, ' '));
                                _strLine.Append(GetStringAsLength("", 7, true, 'x'));
                                _strLine.Append(GetStringAsLength("", 23, true, ' '));
                            }
                            else
                            {
                                _strLine.Append(GetStringAsLength("", 13, true, ' '));
                                _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_SSN"].ToString(), 6, true, ' '));
                                _strLine.Append(GetStringAsLength("", 7, true, 'x'));
                                _strLine.Append(GetStringAsLength("", 23, true, ' '));
                            }
                        }
                    }

                    //우리재발송-일반
                    if (strCard_type_detail == "0011108")
                    {
                        _sw00 = new StreamWriter(fileName + "우리반송일반미배송.dat.00", true, _encoding);
                        _sw01 = new StreamWriter(fileName + "우리반송일반배송수도권.dat.01", true, _encoding);
                        _sw02 = new StreamWriter(fileName + "우리반송일반반송수도권.dat.02", true, _encoding);

                        _sw11_1 = new StreamWriter(fileName + "우리반송일반배송지방.dat.01", true, _encoding);
                        _sw11_2 = new StreamWriter(fileName + "우리반송일반반송지방.dat.02", true, _encoding);
                    }
                    //우리재발송-동의서
                    else if (strCard_type_detail == "0012401")
                    {
                        _sw00 = new StreamWriter(fileName + "우리반송동의미배송.dat.00", true, _encoding);
                        _sw01 = new StreamWriter(fileName + "우리반송동의배송.dat.01", true, _encoding);
                        _sw02 = new StreamWriter(fileName + "우리반송동의반송.dat.02", true, _encoding);
                    }
                    //001-비씨면세유
                    else if (strCard_type_detail == "0011105")
                    {
                        _sw00 = new StreamWriter(fileName + "농협면세유미배송수도권_1그룹.dat.00", true, _encoding);
                        _sw00_1 = new StreamWriter(fileName + "농협면세유미배송지방_1그룹.dat.00", true, _encoding);

                        _sw01 = new StreamWriter(fileName + "농협면세유배송수도권_1그룹.dat.01", true, _encoding);
                        _sw01_1 = new StreamWriter(fileName + "농협면세유배송지방_1그룹.dat.01", true, _encoding);


                        _sw02 = new StreamWriter(fileName + "농협면세유반송수도권_1그룹.dat.02", true, _encoding);
                        _sw02_1 = new StreamWriter(fileName + "농협면세유반송지방_1그룹.dat.02", true, _encoding);

                        _sw012 = new StreamWriter(fileName + "농협면세유등기수도권_1그룹.dat.012", true, _encoding);
                        _sw012_1 = new StreamWriter(fileName + "농협면세유등기지방_1그룹.dat.012", true, _encoding);

                        _sw00_2 = new StreamWriter(fileName + "농협면세유미배송수도권_2그룹.dat.00", true, _encoding);
                        _sw00_3 = new StreamWriter(fileName + "농협면세유미배송지방_2그룹.dat.00", true, _encoding);

                        _sw01_2 = new StreamWriter(fileName + "농협면세유배송수도권_2그룹.dat.01", true, _encoding);
                        _sw01_3 = new StreamWriter(fileName + "농협면세유배송지방_2그룹.dat.01", true, _encoding);


                        _sw02_2 = new StreamWriter(fileName + "농협면세유반송수도권_2그룹.dat.02", true, _encoding);
                        _sw02_3 = new StreamWriter(fileName + "농협면세유반송지방_2그룹.dat.02", true, _encoding);

                        _sw012_2 = new StreamWriter(fileName + "농협면세유등기수도권_2그룹.dat.012", true, _encoding);
                        _sw012_3 = new StreamWriter(fileName + "농협면세유등기지방_2그룹.dat.012", true, _encoding);
                    }
                    //001-비씨일반
                    else if (strCard_type_detail.Substring(0, 5) == "00111" && strCard_type_detail != "0011107")
                    {
                        if (_strVIP == "Z")
                        {
                            _sw00 = new StreamWriter(fileName + "일반블리스미배송_1그룹.dat.00", true, _encoding);
                            _sw01 = new StreamWriter(fileName + "일반블리스배송_1그룹.dat.01", true, _encoding);
                            _sw02 = new StreamWriter(fileName + "일반블리스반송_1그룹.dat.02", true, _encoding);
                            _sw012 = new StreamWriter(fileName + "일반블리스등기_1그룹.dat.012", true, _encoding);

                            _sw00_1 = new StreamWriter(fileName + "일반블리스미배송_2그룹.dat.00", true, _encoding);
                            _sw01_1 = new StreamWriter(fileName + "일반블리스배송_2그룹.dat.01", true, _encoding);
                            _sw02_1 = new StreamWriter(fileName + "일반블리스반송_2그룹.dat.02", true, _encoding);
                            _sw012_1 = new StreamWriter(fileName + "일반블리스등기_2그룹.dat.012", true, _encoding);
                        }
                        //2013.03.04 태희철 수정[S] BC_Part 구분 마감데이터 다운
                        else if (_strVIP == "6")
                        {
                            _sw00 = new StreamWriter(fileName + "다이아_시그니쳐미배송_1그룹.dat.00", true, _encoding);
                            _sw01 = new StreamWriter(fileName + "다이아_시그니쳐배송_1그룹.dat.01", true, _encoding);
                            _sw02 = new StreamWriter(fileName + "다이아_시그니쳐반송_1그룹.dat.02", true, _encoding);
                            _sw012 = new StreamWriter(fileName + "다이아_시그니쳐등기_1그룹.dat.012", true, _encoding);

                            _sw00_1 = new StreamWriter(fileName + "다이아_시그니쳐미배송_2그룹.dat.00", true, _encoding);
                            _sw01_1 = new StreamWriter(fileName + "다이아_시그니쳐배송_2그룹.dat.01", true, _encoding);
                            _sw02_1 = new StreamWriter(fileName + "다이아_시그니쳐반송_2그룹.dat.02", true, _encoding);
                            _sw012_1 = new StreamWriter(fileName + "다이아_시그니쳐등기_2그룹.dat.012", true, _encoding);
                        }
                        else if (_strVIP == "8")
                        {
                            _sw00 = new StreamWriter(fileName + "인피니트미배송_1그룹.dat.00", true, _encoding);
                            _sw01 = new StreamWriter(fileName + "인피니트배송_1그룹.dat.01", true, _encoding);
                            _sw02 = new StreamWriter(fileName + "인피니트반송_1그룹.dat.02", true, _encoding);
                            _sw012 = new StreamWriter(fileName + "인피니트등기_1그룹.dat.012", true, _encoding);

                            _sw00_1 = new StreamWriter(fileName + "인피니트미배송_2그룹dat.00", true, _encoding);
                            _sw01_1 = new StreamWriter(fileName + "인피니트배송_2그룹.dat.01", true, _encoding);
                            _sw02_1 = new StreamWriter(fileName + "인피니트반송_2그룹.dat.02", true, _encoding);
                            _sw012_1 = new StreamWriter(fileName + "인피니트등기_2그룹.dat.012", true, _encoding);
                        }
                        //075-비씨-YE, 001-비씨 500차와 같은 코드를 사용 주의
                        else if (_strowner_only == "1")
                        {
                            _sw00 = new StreamWriter(fileName + "기업세이브미배송_1그룹.dat.00", true, _encoding);
                            _sw01 = new StreamWriter(fileName + "기업세이브배송_1그룹.dat.01", true, _encoding);
                            _sw02 = new StreamWriter(fileName + "기업세이브반송_1그룹.dat.02", true, _encoding);
                            _sw012 = new StreamWriter(fileName + "기업세이브등기_1그룹.dat.012", true, _encoding);

                            _sw00_1 = new StreamWriter(fileName + "기업세이브미배송_2그룹.dat.00", true, _encoding);
                            _sw01_1 = new StreamWriter(fileName + "기업세이브배송_2그룹.dat.01", true, _encoding);
                            _sw02_1 = new StreamWriter(fileName + "기업세이브반송_2그룹.dat.02", true, _encoding);
                            _sw012_1 = new StreamWriter(fileName + "기업세이브등기_2그룹.dat.012", true, _encoding);
                        }
                        else
                        {
                            _sw00 = new StreamWriter(fileName + "일반미배송_1그룹.dat.00", true, _encoding);
                            _sw01 = new StreamWriter(fileName + "일반배송_1그룹.dat.01", true, _encoding);
                            _sw02 = new StreamWriter(fileName + "일반반송_1그룹.dat.02", true, _encoding);
                            _sw012 = new StreamWriter(fileName + "일반등기_1그룹.dat.012", true, _encoding);

                            _sw00_1 = new StreamWriter(fileName + "일반미배송_2그룹.dat.00", true, _encoding);
                            _sw01_1 = new StreamWriter(fileName + "일반배송_2그룹.dat.01", true, _encoding);
                            _sw02_1 = new StreamWriter(fileName + "일반반송_2그룹.dat.02", true, _encoding);
                            _sw012_1 = new StreamWriter(fileName + "일반등기_2그룹.dat.012", true, _encoding);
                        }
                    }
                    //036-비씨긴급은행2
                    else if (strCard_type_detail.Substring(0, 5) == "00133")
                    {
                        if (_strVIP == "Z")
                        {
                            _sw00 = new StreamWriter(fileName + "제일긴급2블리스미배송_1그룹.dat.00", true, _encoding);
                            _sw01 = new StreamWriter(fileName + "제일긴급2블리스배송_1그룹.dat.01", true, _encoding);
                            _sw02 = new StreamWriter(fileName + "제일긴급2블리스반송_1그룹.dat.02", true, _encoding);
                            _sw012 = new StreamWriter(fileName + "제일긴급2블리스등기_1그룹.dat.012", true, _encoding);

                            _sw00_1 = new StreamWriter(fileName + "제일긴급2블리스미배송_2그룹.dat.00", true, _encoding);
                            _sw01_1 = new StreamWriter(fileName + "제일긴급2블리스배송_2그룹.dat.01", true, _encoding);
                            _sw02_1 = new StreamWriter(fileName + "제일긴급2블리스반송_2그룹.dat.02", true, _encoding);
                            _sw012_1 = new StreamWriter(fileName + "제일긴급2블리스등기_2그룹.dat.012", true, _encoding);
                        }
                        else
                        {
                            _sw00 = new StreamWriter(fileName + "제일긴급2미배송_1그룹.dat.00", true, _encoding);
                            _sw01 = new StreamWriter(fileName + "제일긴급2배송_1그룹.dat.01", true, _encoding);
                            _sw02 = new StreamWriter(fileName + "제일긴급2반송_1그룹.dat.02", true, _encoding);
                            _sw012 = new StreamWriter(fileName + "제일긴급2등기_1그룹.dat.012", true, _encoding);

                            _sw00_1 = new StreamWriter(fileName + "제일긴급2미배송_2그룹.dat.00", true, _encoding);
                            _sw01_1 = new StreamWriter(fileName + "제일긴급2배송_2그룹.dat.01", true, _encoding);
                            _sw02_1 = new StreamWriter(fileName + "제일긴급2반송_2그룹.dat.02", true, _encoding);
                            _sw012_1 = new StreamWriter(fileName + "제일긴급2등기_2그룹.dat.012", true, _encoding);
                        }
                    }
                    //076-비씨긴급은행1
                    else if (strCard_type_detail.Substring(0, 5) == "00132")
                    {
                        if (_strCard_Product_code == "0A" && _strCard_traffic_code == "1")
                        {
                            //농협중앙회
                            _sw00_1 = new StreamWriter(fileName + "농협중앙긴급1미배송본인_1그룹.txt.00", true, _encoding);
                            _sw01_1 = new StreamWriter(fileName + "농협중앙긴급1배송본인_1그룹.txt.01", true, _encoding);
                            _sw02_1 = new StreamWriter(fileName + "농협중앙긴급1반송본인_1그룹.txt.02", true, _encoding);
                            _sw012_1 = new StreamWriter(fileName + "농협중앙긴급1등기본인_1그룹.txt.012", true, _encoding);
                            //농협긴급
                            _sw00_4 = new StreamWriter(fileName + "농협지점긴급1미배송본인_1그룹.txt.00", true, _encoding);
                            _sw01_4 = new StreamWriter(fileName + "농협지점긴급1배송본인_1그룹.txt.01", true, _encoding);
                            _sw02_4 = new StreamWriter(fileName + "농협지점긴급1반송본인_1그룹.txt.02", true, _encoding);
                            _sw012_4 = new StreamWriter(fileName + "농협지점긴급1등기본인_1그룹.txt.012", true, _encoding);

                            //농협중앙회
                            _sw10_1 = new StreamWriter(fileName + "농협중앙긴급1미배송본인_2그룹.txt.00", true, _encoding);
                            _sw11_1 = new StreamWriter(fileName + "농협중앙긴급1배송본인_2그룹.txt.01", true, _encoding);
                            _sw12_1 = new StreamWriter(fileName + "농협중앙긴급1반송본인_2그룹.txt.02", true, _encoding);
                            _sw112_1 = new StreamWriter(fileName + "농협중앙긴급1등기본인_2그룹.txt.012", true, _encoding);
                            //농협긴급
                            _sw10_4 = new StreamWriter(fileName + "농협지점긴급1미배송본인_2그룹.txt.00", true, _encoding);
                            _sw11_4 = new StreamWriter(fileName + "농협지점긴급1배송본인_2그룹.txt.01", true, _encoding);
                            _sw12_4 = new StreamWriter(fileName + "농협지점긴급1반송본인_2그룹.txt.02", true, _encoding);
                            _sw112_4 = new StreamWriter(fileName + "농협지점긴급1등기본인_2그룹.txt.012", true, _encoding);
                        }
                        else
                        {
                            // 미배송
                            _sw00_0 = new StreamWriter(fileName + "우리긴급1미배송_1그룹.dat.00", true, _encoding);
                            _sw00_1 = new StreamWriter(fileName + "농협중앙긴급1미배송_1그룹.dat.00", true, _encoding);
                            _sw00_4 = new StreamWriter(fileName + "농협지점긴급1미배송_1그룹.dat.00", true, _encoding);
                            _sw00_2 = new StreamWriter(fileName + "제일긴급1미배송_1그룹.dat.00", true, _encoding);
                            _sw00_3 = new StreamWriter(fileName + "기업긴급1미배송_1그룹.dat.00", true, _encoding);


                            // 배송
                            _sw01_0 = new StreamWriter(fileName + "우리긴급1배송_1그룹.dat.01", true, _encoding);
                            _sw01_1 = new StreamWriter(fileName + "농협중앙긴급1배송_1그룹.dat.01", true, _encoding);
                            _sw01_4 = new StreamWriter(fileName + "농협지점긴급1배송_1그룹.dat.01", true, _encoding);
                            _sw01_2 = new StreamWriter(fileName + "제일긴급1배송_1그룹.dat.01", true, _encoding);
                            _sw01_3 = new StreamWriter(fileName + "기업긴급1배송_1그룹.dat.01", true, _encoding);

                            //반송
                            _sw02_0 = new StreamWriter(fileName + "우리긴급1반송_1그룹.dat.02", true, _encoding);
                            _sw02_1 = new StreamWriter(fileName + "농협중앙긴급1반송_1그룹.dat.02", true, _encoding);
                            _sw02_4 = new StreamWriter(fileName + "농협지점긴급1반송_1그룹.dat.02", true, _encoding);
                            _sw02_2 = new StreamWriter(fileName + "제일긴급1반송_1그룹.dat.02", true, _encoding);
                            _sw02_3 = new StreamWriter(fileName + "기업긴급1반송_1그룹.dat.02", true, _encoding);

                            // 등기
                            _sw012_0 = new StreamWriter(fileName + "우리긴급1등기_1그룹.dat.012", true, _encoding);
                            _sw012_1 = new StreamWriter(fileName + "농협중앙긴급1등기_1그룹.dat.012", true, _encoding);
                            _sw012_4 = new StreamWriter(fileName + "농협지점긴급1등기_1그룹.dat.012", true, _encoding);
                            _sw012_2 = new StreamWriter(fileName + "제일긴급1등기_1그룹.dat.012", true, _encoding);
                            _sw012_3 = new StreamWriter(fileName + "기업긴급1등기_1그룹.dat.012", true, _encoding);

                            // 미배송
                            _sw10_0 = new StreamWriter(fileName + "우리긴급1미배송_2그룹.dat.00", true, _encoding);
                            _sw10_1 = new StreamWriter(fileName + "농협중앙긴급1미배송_2그룹.dat.00", true, _encoding);
                            _sw10_4 = new StreamWriter(fileName + "농협지점긴급1미배송_2그룹.dat.00", true, _encoding);
                            _sw10_2 = new StreamWriter(fileName + "제일긴급1미배송_2그룹.dat.00", true, _encoding);
                            _sw10_3 = new StreamWriter(fileName + "기업긴급1미배송_2그룹.dat.00", true, _encoding);


                            // 배송
                            _sw11_0 = new StreamWriter(fileName + "우리긴급1배송_2그룹.dat.01", true, _encoding);
                            _sw11_1 = new StreamWriter(fileName + "농협중앙긴급1배송_2그룹.dat.01", true, _encoding);
                            _sw11_4 = new StreamWriter(fileName + "농협지점긴급1배송_2그룹.dat.01", true, _encoding);
                            _sw11_2 = new StreamWriter(fileName + "제일긴급1배송_2그룹.dat.01", true, _encoding);
                            _sw11_3 = new StreamWriter(fileName + "기업긴급1배송_2그룹.dat.01", true, _encoding);

                            //반송
                            _sw12_0 = new StreamWriter(fileName + "우리긴급1반송_2그룹.dat.02", true, _encoding);
                            _sw12_1 = new StreamWriter(fileName + "농협중앙긴급1반송_2그룹.dat.02", true, _encoding);
                            _sw12_4 = new StreamWriter(fileName + "농협지점긴급1반송_2그룹.dat.02", true, _encoding);
                            _sw12_2 = new StreamWriter(fileName + "제일긴급1반송_2그룹.dat.02", true, _encoding);
                            _sw12_3 = new StreamWriter(fileName + "기업긴급1반송_2그룹.dat.02", true, _encoding);

                            // 등기
                            _sw112_0 = new StreamWriter(fileName + "우리긴급1등기_2그룹.dat.012", true, _encoding);
                            _sw112_1 = new StreamWriter(fileName + "농협중앙긴급1등기_2그룹.dat.012", true, _encoding);
                            _sw112_4 = new StreamWriter(fileName + "농협지점긴급1등기_2그룹.dat.012", true, _encoding);
                            _sw112_2 = new StreamWriter(fileName + "제일긴급1등기_2그룹.dat.012", true, _encoding);
                            _sw112_3 = new StreamWriter(fileName + "기업긴급1등기_2그룹.dat.012", true, _encoding);
                        }
                    }
                    //비씨 동의서
                    else if (strCard_type_detail.Substring(0, 5) == "00121" || strCard_type_detail.Substring(0, 5) == "00123")
                    {
                        //1그룹
                        _sw00 = new StreamWriter(fileName + "동의서미배송_1그룹.dat.00", true, _encoding);
                        _sw00_1 = new StreamWriter(fileName + "동의서미배송군단위_1그룹.dat.00", true, _encoding);

                        _sw01 = new StreamWriter(fileName + "동의서배송_1그룹.dat.01", true, _encoding);
                        _sw01_1 = new StreamWriter(fileName + "동의서배송군단위_1그룹.dat.01", true, _encoding);


                        _sw02 = new StreamWriter(fileName + "동의서반송_1그룹.dat.02", true, _encoding);
                        _sw02_1 = new StreamWriter(fileName + "동의서반송군단위_1그룹.dat.02", true, _encoding);

                        _sw012 = new StreamWriter(fileName + "동의서등기_1그룹.dat.012", true, _encoding);

                        //2그룹
                        _sw00_2 = new StreamWriter(fileName + "동의서미배송_2그룹.dat.00", true, _encoding);
                        _sw00_3 = new StreamWriter(fileName + "동의서미배송군단위_2그룹.dat.00", true, _encoding);

                        _sw01_2 = new StreamWriter(fileName + "동의서배송_2그룹.dat.01", true, _encoding);
                        _sw01_3 = new StreamWriter(fileName + "동의서배송군단위_2그룹.dat.01", true, _encoding);


                        _sw02_2 = new StreamWriter(fileName + "동의서반송_2그룹.dat.02", true, _encoding);
                        _sw02_3 = new StreamWriter(fileName + "동의서반송군단위_2그룹.dat.02", true, _encoding);

                        _sw012_1 = new StreamWriter(fileName + "동의서등기_2그룹.dat.012", true, _encoding);
                    }
                    else if (strCard_type_detail.Substring(0, 5) == "00131")
                    {
                        if (_strVIP == "Z")
                        {
                            //1그룹
                            _sw00 = new StreamWriter(fileName + "영업점블리스미배송_1그룹.dat.00", true, _encoding);
                            _sw01 = new StreamWriter(fileName + "영업점블리스배송_1그룹.dat.01", true, _encoding);
                            _sw02 = new StreamWriter(fileName + "영업점블리스반송_1그룹.dat.02", true, _encoding);
                            _sw012 = new StreamWriter(fileName + "영업점블리스등기_1그룹.dat.012", true, _encoding);
                            //2그룹
                            _sw00_1 = new StreamWriter(fileName + "영업점블리스미배송_2그룹.dat.00", true, _encoding);
                            _sw01_1 = new StreamWriter(fileName + "영업점블리스배송_2그룹.dat.01", true, _encoding);
                            _sw02_1 = new StreamWriter(fileName + "영업점블리스반송_2그룹.dat.02", true, _encoding);
                            _sw012_1 = new StreamWriter(fileName + "영업점블리스등기_2그룹.dat.012", true, _encoding);
                        }
                        else
                        {
                            //1그룹
                            _sw00 = new StreamWriter(fileName + "긴급영업미배송_1그룹.dat.00", true, _encoding);
                            _sw01 = new StreamWriter(fileName + "긴급영업배송_1그룹.dat.01", true, _encoding);
                            _sw02 = new StreamWriter(fileName + "긴급영업반송_1그룹.dat.02", true, _encoding);
                            _sw012 = new StreamWriter(fileName + "긴급영업등기_1그룹.dat.012", true, _encoding);
                            //2그룹
                            _sw00_1 = new StreamWriter(fileName + "긴급영업미배송_2그룹.dat.00", true, _encoding);
                            _sw01_1 = new StreamWriter(fileName + "긴급영업배송_2그룹.dat.01", true, _encoding);
                            _sw02_1 = new StreamWriter(fileName + "긴급영업반송_2그룹.dat.02", true, _encoding);
                            _sw012_1 = new StreamWriter(fileName + "긴급영업등기_2그룹.dat.012", true, _encoding);
                        }
                    }
                    

                    //데이터 저장
                    //우리재발송-일반
                    if (strCard_type_detail == "0011108")
                    {
                        if (_strStatus == "1" || _strStatus == "7")
                        {
                            // 2012-03-12 수도권 우편번호 100단위와 400단위구분
                            if (_strZipCode.Substring(0, 1) == "1" || _strZipCode.Substring(0, 1) == "4")
                            {
                                _sw01.WriteLine(_strLine.ToString());
                            }
                            else
                            {
                                _sw11_1.WriteLine(_strLine.ToString());
                            }
                        }
                        else if (_strStatus == "2" || _strStatus == "3")
                        {
                            if (_strZipCode.Substring(0, 1) == "1" || _strZipCode.Substring(0, 1) == "4")
                            {
                                _sw02.WriteLine(_strLine.ToString());
                            }
                            else
                            {
                                _sw11_2.WriteLine(_strLine.ToString());
                            }
                        }
                        else
                        {
                            _sw00.WriteLine(_strLine.ToString());
                        }
                    }
                    //우리재발송-동의서
                    else if (strCard_type_detail == "0012401")
                    {
                        if (_strStatus == "1" || _strStatus == "7")
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
                    //농협 면세유
                    else if (strCard_type_detail == "0011105")
                    {
                        #region 농협 면세유
                        if (_str_BC_Part == "2")
                        {
                            if (_strStatus == "2" || _strStatus == "3")
                            {
                                // 2012-03-12 수도권 우편번호 100단위와 400단위구분
                                if (_strZipCode.Substring(0, 1) == "1" || _strZipCode.Substring(0, 1) == "4")
                                {
                                    _sw02_2.WriteLine(_strLine.ToString());
                                }
                                else
                                {
                                    _sw02_3.WriteLine(_strLine.ToString());
                                }
                            }
                            else if (_strStatus == "1" || _strStatus == "7")
                            {
                                if (dtable.Rows[i]["card_branch"].ToString() == "012")
                                {
                                    // 2012-03-12 수도권 구분
                                    if (_strZipCode.Substring(0, 1) == "1" || _strZipCode.Substring(0, 1) == "4")
                                    {
                                        _sw012_2.WriteLine(_strLine.ToString());
                                    }
                                    else
                                    {
                                        _sw012_3.WriteLine(_strLine.ToString());
                                    }
                                }
                                else
                                {
                                    // 2012-03-12 수도권 구분
                                    if (_strZipCode.Substring(0, 1) == "1" || _strZipCode.Substring(0, 1) == "4")
                                    {
                                        _sw01_2.WriteLine(_strLine.ToString());
                                    }
                                    else
                                    {
                                        _sw01_3.WriteLine(_strLine.ToString());
                                    }
                                }
                            }
                            else
                            {
                                // 2012-03-12 수도권 구분
                                if (_strZipCode.Substring(0, 1) == "1" || _strZipCode.Substring(0, 1) == "4")
                                {
                                    _sw00_2.WriteLine(_strLine.ToString());
                                }
                                else
                                {
                                    _sw00_3.WriteLine(_strLine.ToString());
                                }
                            }
                        }
                        else
                        {
                            if (_strStatus == "2" || _strStatus == "3")
                            {
                                // 2012-03-12 수도권 우편번호 100단위와 400단위구분
                                if (_strZipCode.Substring(0, 1) == "1" || _strZipCode.Substring(0, 1) == "4")
                                {
                                    _sw02.WriteLine(_strLine.ToString());
                                }
                                else
                                {
                                    _sw02_1.WriteLine(_strLine.ToString());
                                }
                            }
                            else if (_strStatus == "1" || _strStatus == "7")
                            {
                                if (dtable.Rows[i]["card_branch"].ToString() == "012")
                                {
                                    // 2012-03-12 수도권 구분
                                    if (_strZipCode.Substring(0, 1) == "1" || _strZipCode.Substring(0, 1) == "4")
                                    {
                                        _sw012.WriteLine(_strLine.ToString());
                                    }
                                    else
                                    {
                                        _sw012_1.WriteLine(_strLine.ToString());
                                    }
                                }
                                else
                                {
                                    // 2012-03-12 수도권 구분
                                    if (_strZipCode.Substring(0, 1) == "1" || _strZipCode.Substring(0, 1) == "4")
                                    {
                                        _sw01.WriteLine(_strLine.ToString());
                                    }
                                    else
                                    {
                                        _sw01_1.WriteLine(_strLine.ToString());
                                    }
                                }
                            }
                            else
                            {
                                // 2012-03-12 수도권 구분
                                if (_strZipCode.Substring(0, 1) == "1" || _strZipCode.Substring(0, 1) == "4")
                                {
                                    _sw00.WriteLine(_strLine.ToString());
                                }
                                else
                                {
                                    _sw00_1.WriteLine(_strLine.ToString());

                                }
                            }
                        }
                        #endregion
                    }
                    //비씨일반, 은행긴급2
                    else if (strCard_type_detail.Substring(0, 5) == "00133" || strCard_type_detail.Substring(0, 5) == "00131" || 
                        (strCard_type_detail.Substring(0, 5) == "00111" && strCard_type_detail != "0011107"))
                    {
                        #region 비씨일반, 긴급은행긴급2, 긴급영업점
                        if (_strStatus == "2" || _strStatus == "3")
                        {
                            //2그룹 _sw02_1, 그외는 1그룹
                            if (_str_BC_Part == "2")
                            {
                                _sw02_1.WriteLine(_strLine.ToString());
                            }
                            else
                            {
                                _sw02.WriteLine(_strLine.ToString());
                            }
                        }
                        else if (_strStatus == "1" || _strStatus == "7")
                        {
                            if (dtable.Rows[i]["card_branch"].ToString() == "012")
                            {
                                if (_str_BC_Part == "2")
                                {
                                    _sw012_1.WriteLine(_strLine.ToString());
                                }
                                else
                                {
                                    _sw012.WriteLine(_strLine.ToString());
                                }
                            }
                            else
                            {
                                if (_str_BC_Part == "2")
                                {
                                    _sw01_1.WriteLine(_strLine.ToString());
                                }
                                else
                                {
                                    _sw01.WriteLine(_strLine.ToString());
                                }
                            }
                        }
                        else
                        {
                            if (_str_BC_Part == "2")
                            {
                                _sw00_1.WriteLine(_strLine.ToString());
                            }
                            else
                            {
                                _sw00.WriteLine(_strLine.ToString());
                            }
                        }
                        #endregion
                    }
                    //076-비씨은행1 결과
                    else if (strCard_type_detail.Substring(0, 5) == "00132")
                    {
                        
                        //우리은행 : 20, 22, 24, 83, 84 배열값[0]
                        //농협 : 10, 11, 12, 13, 14, 15, 16, 17, 18 배열값[1]
                        //제일 : 23 배열값[2]
                        //기업 : 그외 배열값[3]

                        #region 076-비씨은행별 결과
                        // 반송
                        if (_strStatus == "2" || _strStatus == "3")
                        {
                            #region 은행 나누기
                            //우리
                            if (_strBank == "WOORI")
                            {
                                if (dtable.Rows[i]["card_branch"].ToString() == "012")
                                {
                                    if (_str_BC_Part == "2")
                                    {
                                        _sw112_0.WriteLine(_strLine.ToString());
                                    }
                                    else
                                    {
                                        _sw012_0.WriteLine(_strLine.ToString());
                                    }
                                }
                                else
                                {
                                    if (_str_BC_Part == "2")
                                    {
                                        _sw12_0.WriteLine(_strLine.ToString());
                                    }
                                    else
                                    {
                                        _sw02_0.WriteLine(_strLine.ToString());
                                    }
                                }
                            }
                            //농협
                            else if (_strBank == "NH1")
                            {
                                if (dtable.Rows[i]["card_branch"].ToString() == "012")
                                {
                                    if (_str_BC_Part == "2")
                                    {
                                        _sw112_1.WriteLine(_strLine.ToString());
                                    }
                                    else
                                    {
                                        _sw012_1.WriteLine(_strLine.ToString());
                                    }
                                }
                                else
                                {
                                    if (_str_BC_Part == "2")
                                    {
                                        _sw12_1.WriteLine(_strLine.ToString());
                                    }
                                    else
                                    {
                                        _sw02_1.WriteLine(_strLine.ToString());
                                    }
                                }
                            }
                            else if (_strBank == "NH2")
                            {
                                if (dtable.Rows[i]["card_branch"].ToString() == "012")
                                {
                                    if (_str_BC_Part == "2")
                                    {
                                        _sw112_4.WriteLine(_strLine.ToString());
                                    }
                                    else
                                    {
                                        _sw012_4.WriteLine(_strLine.ToString());
                                    }
                                }
                                else
                                {
                                    if (_str_BC_Part == "2")
                                    {
                                        _sw12_4.WriteLine(_strLine.ToString());
                                    }
                                    else
                                    {
                                        _sw02_4.WriteLine(_strLine.ToString());
                                    }
                                }
                            }
                            //제일
                            else if (_strBank == "SC")
                            {
                                if (dtable.Rows[i]["card_branch"].ToString() == "012")
                                {
                                    if (_str_BC_Part == "2")
                                    {
                                        _sw112_2.WriteLine(_strLine.ToString());
                                    }
                                    else
                                    {
                                        _sw012_2.WriteLine(_strLine.ToString());
                                    }
                                }
                                else
                                {
                                    if (_str_BC_Part == "2")
                                    {
                                        _sw12_2.WriteLine(_strLine.ToString());
                                    }
                                    else
                                    {
                                        _sw02_2.WriteLine(_strLine.ToString());
                                    }
                                }
                            }
                            //기업
                            else
                            {
                                if (dtable.Rows[i]["card_branch"].ToString() == "012")
                                {
                                    if (_str_BC_Part == "2")
                                    {
                                        _sw112_3.WriteLine(_strLine.ToString());
                                    }
                                    else
                                    {
                                        _sw012_3.WriteLine(_strLine.ToString());
                                    }
                                }
                                else
                                {
                                    if (_str_BC_Part == "2")
                                    {
                                        _sw12_3.WriteLine(_strLine.ToString());
                                    }
                                    else
                                    {
                                        _sw02_3.WriteLine(_strLine.ToString());
                                    }
                                }
                            }
                            #endregion
                        }
                        // 배송 / 재방
                        else if (_strStatus == "1" || _strStatus == "7")
                        {
                            #region 은행 나누기
                            if (_strBank == "WOORI")
                            {
                                if (dtable.Rows[i]["card_branch"].ToString() == "012")
                                {
                                    if (_str_BC_Part == "2")
                                    {
                                        _sw112_0.WriteLine(_strLine.ToString());
                                    }
                                    else
                                    {
                                        _sw012_0.WriteLine(_strLine.ToString());
                                    }
                                }
                                else
                                {
                                    if (_str_BC_Part == "2")
                                    {
                                        _sw11_0.WriteLine(_strLine.ToString());
                                    }
                                    else
                                    {
                                        _sw01_0.WriteLine(_strLine.ToString());
                                    }
                                }
                            }
                            //농협 중앙회
                            else if (_strBank == "NH1")
                            {
                                if (dtable.Rows[i]["card_branch"].ToString() == "012")
                                {
                                    if (_str_BC_Part == "2")
                                    {
                                        _sw112_1.WriteLine(_strLine.ToString());
                                    }
                                    else
                                    {
                                        _sw012_1.WriteLine(_strLine.ToString());
                                    }
                                }
                                else
                                {
                                    if (_str_BC_Part == "2")
                                    {
                                        _sw11_1.WriteLine(_strLine.ToString());
                                    }
                                    else
                                    {
                                        _sw01_1.WriteLine(_strLine.ToString());
                                    }
                                }
                            }
                            //농협지점
                            else if (_strBank == "NH2")
                            {
                                if (dtable.Rows[i]["card_branch"].ToString() == "012")
                                {
                                    if (_str_BC_Part == "2")
                                    {
                                        _sw112_4.WriteLine(_strLine.ToString());
                                    }
                                    else
                                    {
                                        _sw012_4.WriteLine(_strLine.ToString());
                                    }
                                }
                                else
                                {
                                    if (_str_BC_Part == "2")
                                    {
                                        _sw11_4.WriteLine(_strLine.ToString());
                                    }
                                    else
                                    {
                                        _sw01_4.WriteLine(_strLine.ToString());
                                    }
                                }
                            }
                            else if (_strBank == "SC")
                            {
                                if (dtable.Rows[i]["card_branch"].ToString() == "012")
                                {
                                    if (_str_BC_Part == "2")
                                    {
                                        _sw112_2.WriteLine(_strLine.ToString());
                                    }
                                    else
                                    {
                                        _sw012_2.WriteLine(_strLine.ToString());
                                    }
                                }
                                else
                                {
                                    if (_str_BC_Part == "2")
                                    {
                                        _sw11_2.WriteLine(_strLine.ToString());
                                    }
                                    else
                                    {
                                        _sw01_2.WriteLine(_strLine.ToString());
                                    }
                                }
                            }
                            else
                            {
                                if (dtable.Rows[i]["card_branch"].ToString() == "012")
                                {
                                    if (_str_BC_Part == "2")
                                    {
                                        _sw112_3.WriteLine(_strLine.ToString());
                                    }
                                    else
                                    {
                                        _sw012_3.WriteLine(_strLine.ToString());
                                    }
                                }
                                else
                                {
                                    if (_str_BC_Part == "2")
                                    {
                                        _sw11_3.WriteLine(_strLine.ToString());
                                    }
                                    else
                                    {
                                        _sw01_3.WriteLine(_strLine.ToString());
                                    }
                                }
                            }
                            #endregion
                        }
                        // 미배송
                        else
                        {
                            if (_strBank == "WOORI")
                            {
                                if (dtable.Rows[i]["card_branch"].ToString() == "012")
                                {
                                    if (_str_BC_Part == "2")
                                    {
                                        _sw112_0.WriteLine(_strLine.ToString());
                                    }
                                    else
                                    {
                                        _sw012_0.WriteLine(_strLine.ToString());
                                    }
                                }
                                else
                                {
                                    if (_str_BC_Part == "2")
                                    {
                                        _sw10_0.WriteLine(_strLine.ToString());
                                    }
                                    else
                                    {
                                        _sw00_0.WriteLine(_strLine.ToString());
                                    }
                                }
                            }
                            //농협중앙회
                            else if (_strBank == "NH1")
                            {
                                if (dtable.Rows[i]["card_branch"].ToString() == "012")
                                {
                                    if (_str_BC_Part == "2")
                                    {
                                        _sw112_1.WriteLine(_strLine.ToString());
                                    }
                                    else
                                    {
                                        _sw012_1.WriteLine(_strLine.ToString());
                                    }
                                }
                                else
                                {
                                    if (_str_BC_Part == "2")
                                    {
                                        _sw10_1.WriteLine(_strLine.ToString());
                                    }
                                    else
                                    {
                                        _sw00_1.WriteLine(_strLine.ToString());
                                    }
                                }
                            }
                            //농협지점
                            else if (_strBank == "NH2")
                            {
                                if (dtable.Rows[i]["card_branch"].ToString() == "012")
                                {
                                    if (_str_BC_Part == "2")
                                    {
                                        _sw112_4.WriteLine(_strLine.ToString());
                                    }
                                    else
                                    {
                                        _sw012_4.WriteLine(_strLine.ToString());
                                    }
                                }
                                else
                                {
                                    if (_str_BC_Part == "2")
                                    {
                                        _sw10_4.WriteLine(_strLine.ToString());
                                    }
                                    else
                                    {
                                        _sw00_4.WriteLine(_strLine.ToString());
                                    }
                                }
                            }
                            else if (_strBank == "SC")
                            {
                                if (dtable.Rows[i]["card_branch"].ToString() == "012")
                                {
                                    if (_str_BC_Part == "2")
                                    {
                                        _sw112_2.WriteLine(_strLine.ToString());
                                    }
                                    else
                                    {
                                        _sw012_2.WriteLine(_strLine.ToString());
                                    }
                                }
                                else
                                {
                                    if (_str_BC_Part == "2")
                                    {
                                        _sw10_2.WriteLine(_strLine.ToString());
                                    }
                                    else
                                    {
                                        _sw00_2.WriteLine(_strLine.ToString());
                                    }
                                }
                            }
                            else
                            {
                                if (dtable.Rows[i]["card_branch"].ToString() == "012")
                                {
                                    if (_str_BC_Part == "2")
                                    {
                                        _sw112_3.WriteLine(_strLine.ToString());
                                    }
                                    else
                                    {
                                        _sw012_3.WriteLine(_strLine.ToString());
                                    }
                                }
                                else
                                {
                                    if (_str_BC_Part == "2")
                                    {
                                        _sw10_3.WriteLine(_strLine.ToString());
                                    }
                                    else
                                    {
                                        _sw00_3.WriteLine(_strLine.ToString());
                                    }
                                }
                            }
                        }
                    #endregion
                    }
                    //052-비씨동의서
                    else if (strCard_type_detail.Substring(0, 5) == "00121" || strCard_type_detail.Substring(0, 5) == "00123")
                    {
                        if (_strStatus == "2" || _strStatus == "3")
                        {
                            // AreaType = "E" 군지역
                            if (_strAreaType == "E")
                            {
                                if (_str_BC_Part == "2")
                                {
                                    _sw02_3.WriteLine(_strLine.ToString());
                                }
                                else
                                {
                                    _sw02_1.WriteLine(_strLine.ToString());
                                }
                            }
                            else
                            {
                                if (_str_BC_Part == "2")
                                {
                                    _sw02_2.WriteLine(_strLine.ToString());
                                }
                                else
                                {
                                    _sw02.WriteLine(_strLine.ToString());
                                }
                            }
                        }
                        else if (_strStatus == "1" || _strStatus == "7")
                        {
                            if (_strAreaType == "E")
                            {
                                if (_str_BC_Part == "2")
                                {
                                    _sw01_3.WriteLine(_strLine.ToString());
                                }
                                else
                                {
                                    _sw01_1.WriteLine(_strLine.ToString());
                                }
                            }
                            else
                            {
                                if (_str_BC_Part == "2")
                                {
                                    _sw01_2.WriteLine(_strLine.ToString());
                                }
                                else
                                {
                                    _sw01.WriteLine(_strLine.ToString());
                                }
                            }
                        }
                        else
                        {
                            if (_strAreaType == "E")
                            {
                                if (_str_BC_Part == "2")
                                {
                                    _sw00_3.WriteLine(_strLine.ToString());
                                }
                                else
                                {
                                    _sw00_1.WriteLine(_strLine.ToString());
                                }
                            }
                            else
                            {
                                if (_str_BC_Part == "2")
                                {
                                    _sw00_2.WriteLine(_strLine.ToString());
                                }
                                else
                                {
                                    _sw00.WriteLine(_strLine.ToString());
                                }
                            }
                        }
                    }

                    if (_sw00 != null) _sw00.Close();
                    if (_sw01 != null) _sw01.Close();
                    if (_sw02 != null) _sw02.Close();
                    if (_sw012 != null) _sw012.Close();

                    if (_sw00_1 != null) _sw00_1.Close();
                    if (_sw01_1 != null) _sw01_1.Close();
                    if (_sw02_1 != null) _sw02_1.Close();
                    if (_sw012_1 != null) _sw012_1.Close();

                    if (_sw00_2 != null) _sw00_2.Close();
                    if (_sw01_2 != null) _sw01_2.Close();
                    if (_sw02_2 != null) _sw02_2.Close();
                    if (_sw012_2 != null) _sw012_2.Close();

                    if (_sw00_3 != null) _sw00_3.Close();
                    if (_sw01_3 != null) _sw01_3.Close();
                    if (_sw02_3 != null) _sw02_3.Close();
                    if (_sw012_3 != null) _sw012_3.Close();

                    if (_sw00_4 != null) _sw00_4.Close();
                    if (_sw01_4 != null) _sw01_4.Close();
                    if (_sw02_4 != null) _sw02_4.Close();
                    if (_sw012_0 != null) _sw012_0.Close();
                    if (_sw012_4 != null) _sw012_4.Close();
                    

                    if (_sw10_0 != null) _sw10_0.Close();
                    if (_sw10_1 != null) _sw10_1.Close();
                    if (_sw10_2 != null) _sw10_2.Close();
                    if (_sw10_3 != null) _sw10_3.Close();
                    if (_sw10_4 != null) _sw10_4.Close();

                    if (_sw11_0 != null) _sw11_0.Close();
                    if (_sw11_1 != null) _sw11_1.Close();
                    if (_sw11_2 != null) _sw11_2.Close();
                    if (_sw11_3 != null) _sw11_3.Close();
                    if (_sw11_4 != null) _sw11_4.Close();

                    if (_sw12_0 != null) _sw12_0.Close();
                    if (_sw12_1 != null) _sw12_1.Close();
                    if (_sw12_2 != null) _sw12_2.Close();
                    if (_sw12_3 != null) _sw12_3.Close();
                    if (_sw12_4 != null) _sw12_4.Close();

                    if (_sw112_0 != null) _sw112_0.Close();
                    if (_sw112_1 != null) _sw112_1.Close();
                    if (_sw112_2 != null) _sw112_2.Close();
                    if (_sw112_3 != null) _sw112_3.Close();
                    if (_sw112_4 != null) _sw112_4.Close();

                    if (_sw00_0 != null) _sw00_0.Close();
                    if (_sw01_0 != null) _sw01_0.Close();
                    if (_sw02_0 != null) _sw02_0.Close();
                }
                _strReturn = string.Format("{0}건의 인계데이타 다운 완료", i);
            }
            catch (Exception ex)
            {
                _swError = new StreamWriter(fileName + ".ERROR", true, _encoding);
                _swError.WriteLine(_strLine.ToString());
                _swError.WriteLine(ex.Message);
                

                MessageBox.Show(ex.Message);

                _strReturn = string.Format("{0}번째 데이터 인계 중 오류 발생", i + 1);
            }
            finally
            {
                if (_sw00 != null) _sw00.Close();
                if (_sw01 != null) _sw01.Close();
                if (_sw02 != null) _sw02.Close();
                if (_sw012 != null) _sw012.Close();

                if (_sw00_1 != null) _sw00_1.Close();
                if (_sw01_1 != null) _sw01_1.Close();
                if (_sw02_1 != null) _sw02_1.Close();
                if (_sw012_1 != null) _sw012_1.Close();

                if (_sw00_2 != null) _sw00_2.Close();
                if (_sw01_2 != null) _sw01_2.Close();
                if (_sw02_2 != null) _sw02_2.Close();
                if (_sw012_2 != null) _sw012_2.Close();

                if (_sw00_3 != null) _sw00_3.Close();
                if (_sw01_3 != null) _sw01_3.Close();
                if (_sw02_3 != null) _sw02_3.Close();
                if (_sw012_3 != null) _sw012_3.Close();

                if (_sw00_4 != null) _sw00_4.Close();
                if (_sw01_4 != null) _sw01_4.Close();
                if (_sw02_4 != null) _sw02_4.Close();
                if (_sw012_0 != null) _sw012_0.Close();
                if (_sw012_4 != null) _sw012_4.Close();


                if (_sw10_0 != null) _sw10_0.Close();
                if (_sw10_1 != null) _sw10_1.Close();
                if (_sw10_2 != null) _sw10_2.Close();
                if (_sw10_3 != null) _sw10_3.Close();
                if (_sw10_4 != null) _sw10_4.Close();

                if (_sw11_0 != null) _sw11_0.Close();
                if (_sw11_1 != null) _sw11_1.Close();
                if (_sw11_2 != null) _sw11_2.Close();
                if (_sw11_3 != null) _sw11_3.Close();
                if (_sw11_4 != null) _sw11_4.Close();

                if (_sw12_0 != null) _sw12_0.Close();
                if (_sw12_1 != null) _sw12_1.Close();
                if (_sw12_2 != null) _sw12_2.Close();
                if (_sw12_3 != null) _sw12_3.Close();
                if (_sw12_4 != null) _sw12_4.Close();

                if (_sw112_0 != null) _sw112_0.Close();
                if (_sw112_1 != null) _sw112_1.Close();
                if (_sw112_2 != null) _sw112_2.Close();
                if (_sw112_3 != null) _sw112_3.Close();
                if (_sw112_4 != null) _sw112_4.Close();

                if (_sw00_0 != null) _sw00_0.Close();
                if (_sw01_0 != null) _sw01_0.Close();
                if (_sw02_0 != null) _sw02_0.Close();

                if (_swError != null) _swError.Close();
            }
            return _strReturn;
        }

        //개시 마감 자료
        private static string ConvertReceiveType2(System.Data.DataTable dtable, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩
            StreamWriter _sw012 = null;
            StreamWriter _swBank = null;
            StringBuilder _strLine = new StringBuilder("");
            string _strReturn = "", _strBankID = "", strCard_type_detail = "";
            int i = 0, i_Cnt = 0;
            try
            {
                //StreamWriter 초기화
                _sw012 = new StreamWriter(fileName + ".012", true, _encoding);
                _swBank = new StreamWriter(fileName + ".bank", true, _encoding);

                //dtable.Select("card_type_detail = '0011101', client_release_register = '1'");

                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    strCard_type_detail = dtable.Rows[i]["card_type_detail"].ToString();

                    //개시유무 코드 : client_release_register
                    if (strCard_type_detail.Substring(0,5) == "00111" && dtable.Rows[i]["client_release_register"].ToString() == "1")
                    {
                        _strLine = new StringBuilder(GetStringAsLength(dtable.Rows[i]["card_bank_id"].ToString(), 2, true, ' ') + " ");
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_number"].ToString(), 16, true, ' ') + " ");
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_name"].ToString(), 40, true, ' ') + " ");
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_send_number"].ToString(), 8, true, ' ') + " ");
                        _strLine.Append(GetStringAsLength(ConvertZipcode(dtable.Rows[i]["card_zipcode"].ToString()), 7, true, ' '));

                        if (dtable.Rows[i]["card_branch"].ToString() == "012")
                        {
                            _sw012.WriteLine(_strLine.ToString());
                        }
                        _swBank.WriteLine(_strLine.ToString());
                        i_Cnt++;
                    }
                }
                _strReturn = string.Format("{0}건의 인계데이타 다운 완료", i_Cnt);
            }
            catch (Exception)
            {
                _strReturn = string.Format("{0}번째 데이터 인계 중 오류 발생", i + 1);
            }
            finally
            {
                if (_swBank != null) _swBank.Close();
                if (_sw012 != null) _sw012.Close();
            }
            return _strReturn;
        }

        //우리은행,제일은행
        private static string ConvertReceiveType3(DataTable dtable, string fileName)
        {
            ///00.txt + 01.txt 합 하여 출력
            ///출력 시 순번 추가
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩
            StreamWriter _sw00 = null, _sw01 = null, _sw02 = null;                           //파일 쓰기 스트림
            StreamWriter _sw00_1 = null, _sw01_1 = null, _sw02_1 = null;                           //파일 쓰기 스트림
            StringBuilder _strLine = new StringBuilder("");
            string _strReturn = "", _strBankID = "", _strNumber = "", _strStatus = "", _str_BC_Part = "";
            int i = 0;
            DataRow[] _drs = null;
            try
            {
                //1그룹
                _sw00 = new StreamWriter(fileName + "_1그룹00.txt", true, _encoding);
                _sw01 = new StreamWriter(fileName + "_1그룹01.txt", true, _encoding);
                _sw02 = new StreamWriter(fileName + "_1그룹02.txt", true, _encoding);
                //2그룹
                _sw00_1 = new StreamWriter(fileName + "_2그룹00.txt", true, _encoding);
                _sw01_1 = new StreamWriter(fileName + "_2그룹01.txt", true, _encoding);
                _sw02_1 = new StreamWriter(fileName + "_2그룹02.txt", true, _encoding);

                _drs = dtable.Select("", "client_send_number");

                for (i = 0; i < _drs.Length; i++)
                {
                    _strStatus = _drs[i]["card_delivery_status"].ToString();
                    _strNumber = RemoveDash(_drs[i]["client_send_date"].ToString()).Substring(3, 5) + _drs[i]["client_send_number"].ToString();
                    _strBankID = _drs[i]["card_bank_ID"].ToString().Substring(0, 2);
                    _str_BC_Part = dtable.Rows[i]["BC_Part"].ToString();

                    if (_strBankID.Equals("23") || _strBankID.Equals("20") || _strBankID.Equals("22") ||
                        _strBankID.Equals("24") || _strBankID.Equals("83") || _strBankID.Equals("84") || _strBankID.Equals("90"))
                    {
                        _strLine = new StringBuilder(GetStringAsLength(Convert.ToString(i + 1), 5, false, '0'));
                        _strLine.Append(GetStringAsLength("", 1, false, ' '));
                        _strLine.Append(GetStringAsLength(RemoveDash(_drs[i]["client_send_date"].ToString()), 8, true, ' '));
                        _strLine.Append(GetStringAsLength("", 1, false, ' '));
                        _strLine.Append(GetStringAsLength(_drs[i]["card_bank_ID"].ToString(), 6, true, ' '));
                        _strLine.Append(GetStringAsLength("", 1, false, ' '));
                        _strLine.Append(GetStringAsLength(_drs[i]["card_number"].ToString(), 16, true, ' '));
                        _strLine.Append(GetStringAsLength("", 1, false, ' '));
                        _strLine.Append(GetStringAsLength(_drs[i]["customer_name"].ToString(), 20, true, ' '));
                        _strLine.Append(GetStringAsLength("", 1, false, ' '));
                        _strLine.Append(GetStringAsLength(_drs[i]["customer_SSN"].ToString(), 13, true, ' '));
                        _strLine.Append(GetStringAsLength("", 1, false, ' '));
                        _strLine.Append(GetStringAsLength(_drs[i]["client_send_number"].ToString(), 8, true, ' '));
                        _strLine.Append(GetStringAsLength("", 1, false, ' '));
                        _strLine.Append(GetStringAsLength(_drs[i]["card_tel1"].ToString(), 14, true, ' '));
                        _strLine.Append(GetStringAsLength("", 1, false, ' '));
                        _strLine.Append(GetStringAsLength(_drs[i]["card_tel2"].ToString(), 14, true, ' '));
                        _strLine.Append(GetStringAsLength("", 1, false, ' '));
                        _strLine.Append(GetStringAsLength(_drs[i]["card_mobile_tel"].ToString(), 13, true, ' '));
                        _strLine.Append(GetStringAsLength("", 1, false, ' '));
                        _strLine.Append(GetStringAsLength(_drs[i]["card_zipcode"].ToString(), 6, true, ' '));
                        _strLine.Append(GetStringAsLength("", 1, false, ' '));
                        _strLine.Append(GetStringAsLength(_drs[i]["card_address_detail"].ToString(), 50, true, ' '));
                        _strLine.Append(GetStringAsLength("", 1, false, ' '));
                        _strLine.Append(GetStringAsLength(_drs[i]["card_agree_code"].ToString(), 1, true, ' '));
                        _strLine.Append(GetStringAsLength("", 1, false, ' '));
                        _strLine.Append(GetStringAsLength(_strNumber, 13, true, ' '));
                        _strLine.Append(GetStringAsLength("", 1, false, ' '));
                        if (_strStatus.Equals("1"))
                        {
                            _strLine.Append(GetStringAsLength(_drs[i]["receiver_code"].ToString(), 4, true, ' '));
                            _strLine.Append(GetStringAsLength("", 1, false, ' '));
                        }
                        else if (_strStatus.Equals("2") || _strStatus.Equals("3"))
                        {
                            _strLine.Append(GetStringAsLength(_drs[i]["return_code_change"].ToString(), 4, true, ' '));
                            _strLine.Append(GetStringAsLength("", 1, false, ' '));
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength("", 4, true, ' '));
                            _strLine.Append(GetStringAsLength("", 1, false, ' '));
                        }
                        _strLine.Append(GetStringAsLength(RemoveDash(_drs[i]["client_register_date"].ToString()), 8, true, ' '));
                        _strLine.Append(GetStringAsLength("", 1, false, ' '));
                        
                        //데이터저장
                        if (_strStatus.Equals("1"))
                        {
                            if (_str_BC_Part == "2")
                            {
                                _sw01_1.WriteLine(_strLine.ToString());
                            }
                            else
                            {
                                _sw01.WriteLine(_strLine.ToString());
                            }
                        }
                        else if (_strStatus.Equals("2") || _strStatus.Equals("3"))
                        {
                            if (_str_BC_Part == "2")
                            {
                                _sw02_1.WriteLine(_strLine.ToString());
                            }
                            else
                            {
                                _sw02.WriteLine(_strLine.ToString());
                            }
                        }
                        else
                        {
                            if (_str_BC_Part == "2")
                            {
                                _sw00_1.WriteLine(_strLine.ToString());
                            }
                            else
                            {
                                _sw00.WriteLine(_strLine.ToString());
                            }
                        }
                    }
                }
                _strReturn = string.Format("{0}건의 인계데이타 다운 완료", i);
            }
            catch (Exception)
            {
                _strReturn = string.Format("{0}번째 데이터 인계 중 오류 발생", i + 1);
            }
            finally
            {
                //1그룹
                if (_sw00 != null) _sw00.Close();
                if (_sw01 != null) _sw01.Close();
                if (_sw02 != null) _sw02.Close();
                //2그룹
                if (_sw00_1 != null) _sw00_1.Close();
                if (_sw01_1 != null) _sw01_1.Close();
                if (_sw02_1 != null) _sw02_1.Close();
            }
            return _strReturn;
        }


        #region 기타함수
        //반송등록에 사용
        private static string GetBCReturnDate(string value)
        {
            string _return = value;
            if (value.Length == 6)
            {
                _return += "01";
            }
            return _return;
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
