using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace HD_CPT_DONG
{
    public class CONVERT
    {
        //기본 인코딩 설정
        private static string strEncoding = "ks_c_5601-1987";
        private static char chCSV = ';';
        private static string strCardTypeID = "008";
        private static string strCardTypeName = "현.캐_동의서";

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
            string _strReturn = "0592101";
            

            return _strReturn;
        }

        //등록 자료 생성
        //public static string ConvertRegister(string path, string xmlZipcodePath, string xmlZipcodeAreaPath, string xmlPath)
        public static string ConvertRegister(string path, string xmlZipcodePath, string xmlZipcodeAreaPath, string xmlZipcodePath_new, string xmlZipcodeAreaPath_new, string xmlPath)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            //FileInfo _fi = null;
            StreamReader _sr = null;																//파일 읽기 스트림
            StreamWriter _swError = null;						    								//파일 쓰기 스트림
            DataSet _dsetZipcode = null, _dsetZipcdeArea = null;	    							//우편번호 관련 DataSet
            DataSet _dsetZipcode_new = null, _dsetZipcdeArea_new = null;	    							//우편번호 관련 DataSet
            DataTable _dtable = null;																//마스터 저장 테이블
            DataRow _dr = null;
            DataRow[] _drs = null;
            string _strReturn = "";
            string _strLine = "";
            string[] _strAry = null;
            string _strZipcode = "", _strAreaType = "", _strAreaGroup = "", _strBranch = "";
            int _iSeq = 1, _iErrorCount = 0;
            string _strValue = "", _strDeliveryPlaceType = "";
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
                _dtable.Columns.Add("client_send_number");      //5
                _dtable.Columns.Add("card_number");
                _dtable.Columns.Add("card_issue_type_code");
                _dtable.Columns.Add("client_enterprise_code");
                _dtable.Columns.Add("card_bank_ID");
                _dtable.Columns.Add("card_count");              //dr[10] 카드매수
                _dtable.Columns.Add("client_register_type");
                _dtable.Columns.Add("card_product_name");
                _dtable.Columns.Add("card_brand_code");         // 브랜드코드
                _dtable.Columns.Add("card_level_code");         // 등급코드
                _dtable.Columns.Add("customer_name");
                _dtable.Columns.Add("customer_SSN");
                _dtable.Columns.Add("family_name");
                _dtable.Columns.Add("family_SSN");
                _dtable.Columns.Add("family_relation");
                _dtable.Columns.Add("card_zipcode");        //20
                _dtable.Columns.Add("card_address_local");
                _dtable.Columns.Add("card_address_detail");
                _dtable.Columns.Add("card_tel1");
                _dtable.Columns.Add("card_mobile_tel");
                _dtable.Columns.Add("card_zipcode2");
                _dtable.Columns.Add("card_address2_local");
                _dtable.Columns.Add("card_address2_detail");
                _dtable.Columns.Add("card_tel2");
                _dtable.Columns.Add("customer_office");
                _dtable.Columns.Add("customer_email_ID");       //30
                _dtable.Columns.Add("customer_email_domain");
                _dtable.Columns.Add("card_bill_place_type");
                _dtable.Columns.Add("card_delivery_place_code");
                _dtable.Columns.Add("card_bank_account_name");
                _dtable.Columns.Add("card_bank_account_owner");
                _dtable.Columns.Add("card_payment_day");
                _dtable.Columns.Add("card_cooperation1");       //제휴사명
                _dtable.Columns.Add("card_limit");              //희망한도
                _dtable.Columns.Add("card_agree_code");         //동의서식별코드
                _dtable.Columns.Add("card_terminal_issue");     //40
                _dtable.Columns.Add("card_bill_way");
                _dtable.Columns.Add("card_vip_code");
                _dtable.Columns.Add("card_pt_code");
                _dtable.Columns.Add("customer_name_eng_family");
                _dtable.Columns.Add("customer_name_eng_personal");
                _dtable.Columns.Add("card_barcode_new");  // 46

                _dtable.Columns.Add("card_zipcode_new");        //47
                _dtable.Columns.Add("card_zipcode_kind");        //48
                _dtable.Columns.Add("card_zipcode2_new");        //49
                _dtable.Columns.Add("card_zipcode2_kind");        //50

                /*				
                _dtable.Columns.Add("customer_birthdate");
                _dtable.Columns.Add("customer_branch");
                _dtable.Columns.Add("customer_position");
                _dtable.Columns.Add("card_bank_account_no");				
                 */


                //파일 읽기 Stream과 오류시 저장할 쓰기 Stream 생성
                _sr = new System.IO.StreamReader(path, _encoding);
                _swError = new System.IO.StreamWriter(path + ".Error", false, _encoding);

                //우편번호 관련 정보 DataSet에 담기
                _dsetZipcode = new DataSet();
                _dsetZipcdeArea = new DataSet();
                _dsetZipcode.ReadXml(xmlZipcodePath);
                _dsetZipcode.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcode.Tables[0].Columns["zipcode"] };
                _dsetZipcdeArea.ReadXml(xmlZipcodeAreaPath);
                _dsetZipcdeArea.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcdeArea.Tables[0].Columns["zipcode"] };

                //새우편번호 관련 정보 DataSet에 담기
                _dsetZipcode_new = new DataSet();
                _dsetZipcdeArea_new = new DataSet();
                _dsetZipcode_new.ReadXml(xmlZipcodePath_new);
                _dsetZipcode_new.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcode_new.Tables[0].Columns["zipcode_new"] };
                _dsetZipcdeArea_new.ReadXml(xmlZipcodeAreaPath_new);
                _dsetZipcdeArea_new.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcdeArea_new.Tables[0].Columns["zipcode_new"] };

                while ((_strLine = _sr.ReadLine()) != null)
                {
                    _dr = _dtable.NewRow();
                    _dr[0] = _iSeq;
                    //CSV 분리
                    _strAry = _strLine.Split(chCSV);


                    //01 : 자택 , 02 : 직장
                    _strDeliveryPlaceType = _strAry[27];
                    _strValue = _strAry[0];
                    _dr[5] = _strValue; ;
                    _dr[6] = _strValue.Substring(1, 16);
                    _dr[7] = _strAry[1];
                    _dr[8] = _strAry[2];
                    _dr[9] = _strAry[3];
                    _dr[10] = _strAry[4];
                    _dr[11] = _strAry[5];
                    _dr[12] = _strAry[6];
                    _dr[13] = _strAry[7];
                    _dr[14] = _strAry[8];
                    _dr[15] = _strAry[9];
                    _dr[16] = _strAry[10].Replace('X', 'x');
                    _dr[17] = _strAry[11];
                    _dr[18] = _strAry[12];
                    _dr[19] = _strAry[13];
                    if (_strDeliveryPlaceType == "01")
                    {
                        _strZipcode = _strAry[14].Trim();

                        if (_strZipcode.Length == 5)
                        {
                            _dr[47] = _strZipcode;
                            _dr[48] = "1";
                        }
                        _dr[20] = _strZipcode;
                        _dr[21] = _strAry[15];
                        _dr[22] = _strAry[16];
                        _dr[23] = _strAry[17];
                    }
                    else
                    {
                        _strZipcode = _strAry[19].Trim();

                        if (_strZipcode.Length == 5)
                        {
                            _dr[47] = _strZipcode;
                            _dr[48] = "1";
                        }
                        _dr[20] = _strZipcode;
                        _dr[21] = _strAry[20];
                        _dr[22] = _strAry[21];
                        _dr[23] = _strAry[22];
                    }

                    _dr[24] = _strAry[18];
                    if (_strDeliveryPlaceType == "01")
                    {
                        _dr[25] = _strAry[19].Trim();

                        if (_dr[25].ToString().Length == 5)
                        {
                            _dr[49] = _dr[25].ToString();
                            _dr[50] = "1";
                        }
                        _dr[26] = _strAry[20];
                        _dr[27] = _strAry[21];
                        _dr[28] = _strAry[22];

                    }
                    else
                    {
                        _dr[25] = _strAry[14].Trim();

                        if (_dr[25].ToString().Length == 5)
                        {
                            _dr[49] = _dr[25].ToString();
                            _dr[50] = "1";
                        }
                        _dr[26] = _strAry[15];
                        _dr[27] = _strAry[16];
                        _dr[28] = _strAry[17];
                    }

                    _dr[29] = _strAry[23];
                    _dr[30] = _strAry[24];
                    _dr[31] = _strAry[25];
                    _dr[32] = _strAry[26];
                    _dr[33] = _strAry[27];
                    _dr[34] = _strAry[28];
                    _dr[35] = _strAry[29];
                    _dr[36] = _strAry[30];
                    _dr[37] = _strAry[31];
                    _dr[38] = _strAry[32];
                    _dr[39] = _strAry[33];
                    _dr[40] = _strAry[34];
                    _dr[41] = _strAry[35];
                    _dr[42] = _strAry[36];
                    _dr[43] = _strAry[37];

                    if (_strAry.Length > 38)
                    {
                        _dr[44] = _strAry[38];
                        _dr[45] = _strAry[39];
                    }
                    //2015.11.24 태희철 수정
                    //_dr[46] = "02" + _strAry[0].Substring(1, 16);
                    _dr[46] = "02" + _strAry[0].Substring(1, 16) + _dr[20].ToString();

                    if (_strZipcode != "")
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
            StreamWriter _sw00 = null, _sw01 = null;																			//파일 쓰기 스트림
            int i = 0;
            StringBuilder _strLine = new StringBuilder("");
            string _strReturn = "", _strDeliveryStatus = "";
            try
            {
                _sw00 = new StreamWriter(fileName + ".000", true, _encoding);
                _sw01 = new StreamWriter(fileName + ".001", true, _encoding);
                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    _strDeliveryStatus = dtable.Rows[i]["card_delivery_status"].ToString();
                    //2012-03-30 태희철 수정[S] (사용안함)
                    //_strLine = new StringBuilder(GetStringAsLength(dtable.Rows[i]["card_kind"].ToString(), 2));
                    _strLine = new StringBuilder("");
                    if (_strDeliveryStatus == "0")
                    {
                        _strLine.Append(GetStringAsLength("", 1));
                        //_strLine = new StringBuilder(GetStringAsLength("", 1));
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength(_strDeliveryStatus, 1));
                        //_strLine = new StringBuilder(GetStringAsLength(_strDeliveryStatus, 1));
                    }
                    //2012-03-30 태희철 수정[E] (사용안함)

                    _strLine.Append(GetStringAsLength("2", 1));
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_number"].ToString(), 16));
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_SSN"].ToString(), 13));
                    if (dtable.Rows[i]["card_issue_type_code"].ToString() == "")
                    {
                        _strLine.Append(GetStringAsLength("1", 1));
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_issue_type_code"].ToString(), 1));
                    }

                    if (_strDeliveryStatus == "1" || _strDeliveryStatus == "7")
                    {
                        _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[i]["card_delivery_date"].ToString()), 8));
                        //if (dtable.Rows[i]["card_result_status"].ToString() == "61")
                        //{
                        //    _strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_SSN"].ToString().Replace("x", "0"), 13,true,'0'));

                        //}
                        //else
                        //{
                        //    _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_SSN"].ToString().Replace("x","0").Replace(" ","0"), 13));

                        //}
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_SSN"].ToString(), 13, true, 'x'));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_code_change"].ToString(), 2));
                        _strLine.Append(GetStringAsLength("", 1));
                    }
                    else if (_strDeliveryStatus == "2" || _strDeliveryStatus == "3")
                    {
                        _strLine.Append(GetStringAsLength("", 8));
                        _strLine.Append(GetStringAsLength("", 13));
                        _strLine.Append(GetStringAsLength("", 2));
                        _strLine.Append(GetStringAsLength(ReturnType(dtable.Rows[i]["delivery_return_reason_last"].ToString()), 1));
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength("", 8));
                        _strLine.Append(GetStringAsLength("", 13));
                        _strLine.Append(GetStringAsLength("", 2));
                        _strLine.Append(GetStringAsLength("", 1));
                    }

                    _strLine.Append(GetStringAsLength("B002", 4));
                    _strLine.Append(GetStringAsLength("", 16));
                    if (dtable.Rows[i]["card_agree1"].ToString() == "2" || dtable.Rows[i]["card_agree2"].ToString() == "2" || dtable.Rows[i]["card_agree3"].ToString() == "2")
                    {
                        _strLine.Append(GetStringAsLength("N", 1));
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength("Y", 1));
                    }
                    /*
                   _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["입회신청서변경이력명"].ToString()), 8) + chCSV;
                   _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["회원신청등록일자"].ToString()), 8) + chCSV;
                   _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["회원신청일련번호"].ToString()), 6) + chCSV;
                     */
                    if (_strDeliveryStatus == "1" || _strDeliveryStatus == "7")
                    {
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_name"].ToString(), 16));

                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength("", 16));
                    }

                    if (dtable.Rows[i]["card_agree4"].ToString().Equals("1"))
                    {
                        _strLine.Append(GetStringAsLength("Y", 1));
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength("N", 1));
                    }
                    _strLine.Append(GetStringAsLength("", 7));

                    if (_strDeliveryStatus == "1" || _strDeliveryStatus == "2" || _strDeliveryStatus == "3")
                    {
                        _sw01.WriteLine(_strLine.ToString());
                    }
                    else
                    {
                        _sw00.WriteLine(_strLine.ToString());
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
                if (_sw00 != null) _sw00.Close();
                if (_sw01 != null) _sw01.Close();
            }
            return _strReturn;

        }
        //일일마감자료
        public static string ConvertResultDay(System.Data.DataTable dtable, string fileName)
        {
            //return ConvertResult(dtable, fileName);
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            StreamWriter _sw00 = null, _sw01 = null, _sw02 = null, _sw03 = null;																			//파일 쓰기 스트림
            int i = 0;
            StringBuilder _strLine = new StringBuilder("");
            string _strReturn = "", _strDeliveryStatus = "";
            try
            {
                _sw00 = new StreamWriter(fileName + ".00", true, _encoding);
                _sw01 = new StreamWriter(fileName + ".01", true, _encoding);
                _sw02 = new StreamWriter(fileName + ".02", true, _encoding);
                _sw03 = new StreamWriter(fileName + ".SSN", true, _encoding);
                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    _strDeliveryStatus = dtable.Rows[i]["card_delivery_status"].ToString();
                    //2012-03-30 태희철 수정[S] (사용안함)
                    //_strLine = new StringBuilder(GetStringAsLength(dtable.Rows[i]["card_kind"].ToString(), 2));
                    if (_strDeliveryStatus == "0")
                    {
                        //_strLine.Append(GetStringAsLength("", 1));
                        _strLine = new StringBuilder(GetStringAsLength("", 1));
                    }
                    else
                    {
                        //_strLine.Append(GetStringAsLength(_strDeliveryStatus, 1));
                        _strLine = new StringBuilder(GetStringAsLength(_strDeliveryStatus, 1));
                    }
                    //2012-03-30 태희철 수정[E] (사용안함)

                    _strLine.Append(GetStringAsLength("2", 1));
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_number"].ToString(), 16));
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_SSN"].ToString(), 13));
                    if (dtable.Rows[i]["card_issue_type_code"].ToString() == "")
                    {
                        _strLine.Append(GetStringAsLength("1", 1));
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_issue_type_code"].ToString(), 1));
                    }

                    if (_strDeliveryStatus == "1" || _strDeliveryStatus == "7")
                    {
                        _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[i]["card_delivery_date"].ToString()), 8));
                        //2012-04-18 태희철 수정
                        //if (dtable.Rows[i]["card_result_status"].ToString() == "61")
                        //{
                        //    _strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_SSN"].ToString(), 13));

                        //}
                        //else
                        //{
                        //    _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_SSN"].ToString(), 13));

                        //}
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_SSN"].ToString(), 13, true, 'x'));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_code_change"].ToString(), 2));
                        _strLine.Append(GetStringAsLength("", 1));
                    }
                    else if (_strDeliveryStatus == "2" || _strDeliveryStatus == "3")
                    {
                        _strLine.Append(GetStringAsLength("", 8));
                        _strLine.Append(GetStringAsLength("", 13));
                        _strLine.Append(GetStringAsLength("", 2));
                        _strLine.Append(GetStringAsLength(ReturnType(dtable.Rows[i]["delivery_return_reason_last"].ToString()), 1));
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength("", 8));
                        _strLine.Append(GetStringAsLength("", 13));
                        _strLine.Append(GetStringAsLength("", 2));
                        _strLine.Append(GetStringAsLength("", 1));
                    }

                    _strLine.Append(GetStringAsLength("B002", 4));
                    _strLine.Append(GetStringAsLength("", 16));
                    if (dtable.Rows[i]["card_agree1"].ToString() == "2" || dtable.Rows[i]["card_agree2"].ToString() == "2" || dtable.Rows[i]["card_agree3"].ToString() == "2")
                    {
                        _strLine.Append(GetStringAsLength("N", 1));
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength("Y", 1));
                    }
                    /*
                   _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["입회신청서변경이력명"].ToString()), 8) + chCSV;
                   _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["회원신청등록일자"].ToString()), 8) + chCSV;
                   _strLine += GetStringAsLength(RemoveDash(dtable.Rows[i]["회원신청일련번호"].ToString()), 6) + chCSV;
                     */
                    if (_strDeliveryStatus == "1" || _strDeliveryStatus == "7")
                    {
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_name"].ToString(), 16));

                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength("", 16));
                    }

                    if (dtable.Rows[i]["card_agree4"].ToString().Equals("1"))
                    {
                        _strLine.Append(GetStringAsLength("Y", 1));
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength("N", 1));
                    }
                    _strLine.Append(GetStringAsLength("", 7));

                    if (_strDeliveryStatus == "1")
                    {
                        //if (dtable.Rows[i]["receiver_SSN"].ToString().Trim() != dtable.Rows[i]["receiver_SSN_org"].ToString().Trim() && dtable.Rows[i]["card_result_status"].ToString() == "61")
                        //{
                        //    _sw03.WriteLine(_strLine.ToString());
                        //}
                        //else
                        //{
                        //    _sw01.WriteLine(_strLine.ToString());
                        //}
                        _sw01.WriteLine(_strLine.ToString());
                    }
                    else if (_strDeliveryStatus == "2" || _strDeliveryStatus == "3")
                    {
                        _sw02.WriteLine(_strLine.ToString());
                    }
                    else
                    {
                        _sw00.WriteLine(_strLine.ToString());
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
                if (_sw00 != null) _sw00.Close();
                if (_sw01 != null) _sw01.Close();
                if (_sw02 != null) _sw02.Close();
                if (_sw03 != null) _sw03.Close();
            }
            return _strReturn;
        }

        #region 반송사유
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

