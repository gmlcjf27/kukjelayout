using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace _105_INTERPARK
{
    public class CONVERT
    {
        //기본 인코딩 설정
        private static string strEncoding = "ks_c_5601-1987";
        private static char chCSV = ',';
        private static string strCardTypeID = "039";
        private static string strCardTypeName = "인터파크";

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
            string _strReturn = "";

            return _strReturn;
        }


        //등록 선택 자료
        //public static string ConvertRegister(string path, string xmlZipcodePath, string xmlZipcodeAreaPath, string xmlPath)
        public static string ConvertRegister(string path, string xmlZipcodePath, string xmlZipcodeAreaPath, string xmlZipcodePath_new, string xmlZipcodeAreaPath_new, string xmlPath)

        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            //FileInfo _fi = null;
            StreamReader _sr = null;														//파일 읽기 스트림
            StreamWriter _swError = null;													//파일 쓰기 스트림
            DataSet _dsetZipcode = null, _dsetZipcodeArea = null;							//우편번호 관련 DataSet
            DataSet _dsetZipcode_new = null, _dsetZipcodeArea_new = null;					//우편번호 관련 DataSet
            DataTable _dtable = null;														//마스터 저장 테이블
            DataRow _dr = null;
            DataRow[] _drs = null;
            string _strReturn = "";
            string _strLine = "";
            string[] _strAry = null;
            string _strZipcode = "", _strAreaType = "", _strAreaGroup = "", _strBranch = "", strGetDate = "", strCustomerSSN = "";
            int _iSeq = 1, _iErrorCount = 0, _iErrorQuotaCount = 0;
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
                _dtable.Columns.Add("client_send_date");    //dr[5]
                _dtable.Columns.Add("client_send_number");
                _dtable.Columns.Add("customer_name");
                _dtable.Columns.Add("customer_SSN");
                _dtable.Columns.Add("card_number");
                _dtable.Columns.Add("card_zipcode");        //dr[10]
                _dtable.Columns.Add("card_address");
                _dtable.Columns.Add("card_tel2");
                _dtable.Columns.Add("card_mobile_tel");
                _dtable.Columns.Add("client_number");
                _dtable.Columns.Add("family_name");         //dr[15]
                _dtable.Columns.Add("customer_no");
                _dtable.Columns.Add("card_address2");
                _dtable.Columns.Add("client_quick_seq");    //dr[18]

                _dtable.Columns.Add("card_zipcode_new");    //dr[19]
                _dtable.Columns.Add("card_zipcode_kind");   //dr[20]

                _dtable.Columns.Add("customer_memo");   //dr[21]
                //_dtable.Columns.Add("card_barcode_new");   //dr[22]

                //우편번호 관련 정보 DataSet에 담기
                _dsetZipcode = new DataSet();
                _dsetZipcodeArea = new DataSet();
                _dsetZipcode.ReadXml(xmlZipcodePath);
                _dsetZipcode.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcode.Tables[0].Columns["zipcode"] };
                _dsetZipcodeArea.ReadXml(xmlZipcodeAreaPath);
                _dsetZipcodeArea.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcodeArea.Tables[0].Columns["zipcode"] };

                //우편번호 관련 정보 DataSet에 담기
                _dsetZipcode_new = new DataSet();
                _dsetZipcodeArea_new = new DataSet();
                _dsetZipcode_new.ReadXml(xmlZipcodePath_new);
                _dsetZipcode_new.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcode_new.Tables[0].Columns["zipcode_new"] };
                _dsetZipcodeArea_new.ReadXml(xmlZipcodeAreaPath_new);
                _dsetZipcodeArea_new.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcodeArea_new.Tables[0].Columns["zipcode_new"] };

                //파일 읽기 Stream과 오류시 저장할 쓰기 Stream 생성
                _sr = new System.IO.StreamReader(path, _encoding);
                _swError = new System.IO.StreamWriter(path + ".Error", false, _encoding);

                strGetDate = DateTime.Now.ToString("yyyyMMdd");

                while ((_strLine = _sr.ReadLine()) != null)
                {
                    _dr = _dtable.NewRow();
                    _dr[0] = _iSeq;
                    //CSV 분리
                    if (_strLine.IndexOf('"') > -1) _iErrorQuotaCount++;
                    _strAry = _strLine.Split(chCSV);
                    _dr[5] = strGetDate;
                    //_dr[6] = GetStringAsLength(_iSeq.ToString(), 5, false, '0');
                    _dr[6] = _strAry[14].Trim();

                    if (_dr[6].ToString().Trim().Length < 13)
                    {
                        MessageBox.Show("줄번호 " + _iSeq.ToString() + " 번째 일련번호 자릿수 오류 입니다. 일련번호는 13자리 입니다.", "오류");
                        throw new ArgumentNullException("일련번호 자릿수 오류");
                    }

                    _dr[7] = _strAry[0];

                    if (RemoveDash(_strAry[1]).Length == 11)
                    {
                        strCustomerSSN = "xx" + RemoveDash(_strAry[1]).Replace('*', 'x');
                    }
                    else if (RemoveDash(_strAry[1]).Length == 10)
                    {
                        strCustomerSSN = "xxx" + RemoveDash(_strAry[1]).Replace('*', 'x');
                    }
                    else if (RemoveDash(_strAry[1]).Length == 13)
                    {
                        strCustomerSSN = RemoveDash(_strAry[1]).Replace('*', 'x');
                    }
                    else
                    {
                        strCustomerSSN = "xxxxxxxxxxxxx";
                    }

                    _dr[8] = strCustomerSSN;
                    _dr[9] = _strAry[2].Trim();
                    _strZipcode = RemoveDash(_strAry[3]).Trim();
                    _dr[10] = _strZipcode;

                    if (_strZipcode.Length == 5)
                    {
                        _dr[19] = _strZipcode;
                        _dr[20] = "1";
                    }

                    _dr[11] = _strAry[4];
                    _dr[12] = _strAry[5];
                    _dr[13] = _strAry[6];

                    if (_strAry.Length < 7)
                    {
                        _dr[14] = "";
                        _dr[15] = "";
                        _dr[16] = "";
                        _dr[17] = "";
                        _dr[18] = "";
                    }
                    else
                    {
                        _dr[14] = _strAry[7];
                        _dr[15] = _strAry[9];
                        _dr[16] = _strAry[10];
                        _dr[17] = _strAry[13];
                        //2019.12.11 태희철 공연명(상품명)추가
                        _dr[18] = _strAry[11];
                    }

                    _dr[21] = "본인수령 시 휴대폰 뒷 4자리 등록";
                    //2019.09.10 태희철 수정 일련번호 추가
                    //_dr[6] = _strAry[14];

                    if (_strZipcode != "")
                    {
                        //지역 분류 선택
                        if (_strZipcode.Length == 5)
                        {
                            _drs = _dsetZipcodeArea_new.Tables[0].Select("zipcode_new = " + _strZipcode);
                        }
                        else
                        {
                            _drs = _dsetZipcodeArea.Tables[0].Select("zipcode = " + _strZipcode);
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
                if (_iErrorCount < 1 && _iErrorQuotaCount < 1)
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
                    _strReturn = string.Format("{0}건 변환, 우편번호 미등록 {1}건 실패, 따옴표 검출 {2}건", _iSeq - _iErrorCount - _iErrorQuotaCount - 1, _iErrorCount,
                        _iErrorQuotaCount);
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
            int _iReturn = 0;
            string _strReturn = "";
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
        public static string ConvertResultDay(System.Data.DataTable dtable, string fileName)
        {
            return ConvertResult(dtable, fileName);
        }

        //마감데이터
        private static string ConvertReceiveType1(System.Data.DataTable dtable, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            StreamWriter _sw012 = null, _sw01 = null, _sw07 = null;							//파일 쓰기 스트림
            int i = 0;
            StringBuilder _strLine = new StringBuilder("");
            string _strReturn = "", _strBranch = "", _strStatus = "", strCard_type_detail = "", strCard_in_date = "";
            string strCard_zipcode_kind = "";

            try
            {
                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    _strBranch = dtable.Rows[i]["card_branch"].ToString();
                    _strStatus = dtable.Rows[i]["card_delivery_status"].ToString();
                    strCard_type_detail = dtable.Rows[i]["card_type_detail"].ToString();
                    strCard_in_date = String.Format("{0:yyyyMMdd}", dtable.Rows[i]["card_in_date"]);

                    _strLine = new StringBuilder(GetStringAsLength(RemoveDash(dtable.Rows[i]["client_send_date"].ToString()), 8) + chCSV);
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_send_number"].ToString(), 5) + chCSV);
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_name"].ToString(), 50) + chCSV);
                    _strLine.Append(GetStringAsLength("", 14) + chCSV);
                    _strLine.Append(dtable.Rows[i]["card_number"].ToString().Trim() + chCSV);

                    if (strCard_zipcode_kind == "1")
                    {
                        _strLine.Append(GetStringAsLength(ConvertZipcode(dtable.Rows[i]["card_zipcode_new"].ToString()), 7) + chCSV);
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength(ConvertZipcode(dtable.Rows[i]["card_zipcode"].ToString()), 7) + chCSV);
                    }
                    
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_address"].ToString(), 100) + chCSV);
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_tel2"].ToString(), 50) + chCSV);
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_mobile_tel"].ToString(), 50) + chCSV);
                    _strLine.Append(GetStringAsLength(DeilveryStatusType(dtable.Rows[i]["card_delivery_status"].ToString()), 5) + chCSV);

                    if (_strStatus == "1")
                    {
                        _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[i]["card_delivery_date"].ToString()), 8) + chCSV);
                    }
                    else if (_strStatus == "2")
                    {
                        _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[i]["delivery_return_date_last"].ToString()), 8) + chCSV);
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength("", 8) + chCSV);
                    }

                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_name"].ToString(), 50) + chCSV);
                    if (dtable.Rows[i]["receiver_code_change"].ToString() == "00")
                    {
                        _strLine.Append(GetStringAsLength("", 5) + chCSV);
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_code_change"].ToString().Replace("x", " "), 5) + chCSV);
                    }

                    //본인수령의 경우 제공받은 값을 리턴함
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_SSN"].ToString(), 14) + chCSV);

                    if (_strStatus == "2" || _strStatus == "3")
                    {
                        if (dtable.Rows[i]["return_code_change"].ToString() == "00")
                        {
                            _strLine.Append(GetStringAsLength("", 5) + chCSV);
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["return_code_change"].ToString().Replace("x", " "), 5) + chCSV);
                        }
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength("", 5) + chCSV);
                    }

                    if (_strBranch == "012")
                    {
                        _strLine.Append("2");
                    }
                    else
                    {
                        _strLine.Append("1");
                    }

                    if (_strBranch == "012")
                    {
                        if (strCard_type_detail == "1051101")
                        {
                            _sw012 = new StreamWriter(fileName + "INT" + strCard_in_date + "_등기", true, _encoding);
                        }
                        else if (strCard_type_detail == "1051102")
                        {
                            _sw012 = new StreamWriter(fileName + "INT" + strCard_in_date + "m_등기", true, _encoding);
                        }
                        else if (strCard_type_detail == "1051103")
                        {
                            _sw012 = new StreamWriter(fileName + "INT" + strCard_in_date + "현대_등기", true, _encoding);
                        }
                        else if (strCard_type_detail == "1051104")
                        {
                            _sw012 = new StreamWriter(fileName + "INT" + strCard_in_date + "교보_등기", true, _encoding);
                        }
                        else
                        {
                            _sw012 = new StreamWriter(fileName + "기타" + strCard_in_date, true, _encoding);
                        }
                        _sw012.WriteLine(_strLine.ToString());
                        _sw012.Close();
                    }
                    else
                    {
                        if (_strStatus == "7")
                        {
                            if (strCard_type_detail == "1051101")
                            {
                                _sw07 = new StreamWriter(fileName + "INT_재방" + strCard_in_date, true, _encoding);
                            }
                            else if (strCard_type_detail == "1051102")
                            {
                                _sw07 = new StreamWriter(fileName + "INT_재방" + strCard_in_date + "m", true, _encoding);
                            }
                            else if (strCard_type_detail == "1051103")
                            {
                                _sw07 = new StreamWriter(fileName + "INT_재방" + strCard_in_date + "현대", true, _encoding);
                            }
                            else if (strCard_type_detail == "1051104")
                            {
                                _sw07 = new StreamWriter(fileName + "INT_재방" + strCard_in_date + "교보", true, _encoding);
                            }
                            else
                            {
                                _sw07 = new StreamWriter(fileName + "기타_재방" + strCard_in_date, true, _encoding);
                            }
                            _sw07.WriteLine(_strLine.ToString());
                            _sw07.Close();
                        }
                        else
                        {
                            if (strCard_type_detail == "1051101")
                            {
                                _sw01 = new StreamWriter(fileName + "INT" + strCard_in_date, true, _encoding);
                            }
                            else if (strCard_type_detail == "1051102")
                            {
                                _sw01 = new StreamWriter(fileName + "INT" + strCard_in_date + "m", true, _encoding);
                            }
                            else if (strCard_type_detail == "1051103")
                            {
                                _sw01 = new StreamWriter(fileName + "INT" + strCard_in_date + "현대", true, _encoding);
                            }
                            else if (strCard_type_detail == "1051104")
                            {
                                _sw01 = new StreamWriter(fileName + "INT" + strCard_in_date + "교보", true, _encoding);
                            }
                            else
                            {
                                _sw01 = new StreamWriter(fileName + "기타" + strCard_in_date, true, _encoding);
                            }
                            _sw01.WriteLine(_strLine.ToString());
                            _sw01.Close();
                        }
                    }
                }
                _strReturn = string.Format("{0}건의 인계데이터 다운 완료", i);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                _strReturn = string.Format("{0}번째 데이터 인계 중 오류 발생", i + 1);
            }
            finally
            {
                if (_sw012 != null) _sw012.Close();
                if (_sw01 != null) _sw01.Close();
            }
            return _strReturn;
        }

        //결번(등기) 데이터
        private static string ConvertReceiveType2(System.Data.DataTable dtable, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            StreamWriter _sw01 = null;											//파일 쓰기 스트림
            int i = 0, iCnt = 0;
            StringBuilder _strLine = null;
            string _strReturn = "", _strStatus = "", _strBranch = "", strCard_type_detail = "", strCard_in_date = "";
            string strCard_zipcode_kind = "";
            try
            {
                _sw01 = new StreamWriter(fileName + "INT" + strCard_in_date + "_등기", true, _encoding);
                _strLine = new StringBuilder("고객명,생년월일,예매번호,우편번호,주소,전화번호,핸드폰번호, ,예매번호,예매자명,Bdate,BdateSeq,연락처");
                _sw01.WriteLine(_strLine.ToString());
                _sw01.Close();

                _sw01 = new StreamWriter(fileName + "INT" + strCard_in_date + "m_등기", true, _encoding);
                _strLine = new StringBuilder("고객명,생년월일,예매번호,우편번호,주소,전화번호,핸드폰번호, ,예매번호,예매자명,Bdate,BdateSeq,연락처");
                _sw01.WriteLine(_strLine.ToString());
                _sw01.Close();

                _sw01 = new StreamWriter(fileName + "INT" + strCard_in_date + "현대_등기", true, _encoding);
                _strLine = new StringBuilder("고객명,생년월일,예매번호,우편번호,주소,전화번호,핸드폰번호, ,예매번호,예매자명,Bdate,BdateSeq,연락처");
                _sw01.WriteLine(_strLine.ToString());
                _sw01.Close();

                _sw01 = new StreamWriter(fileName + "INT" + strCard_in_date + "교보_등기", true, _encoding);
                _strLine = new StringBuilder("고객명,생년월일,예매번호,우편번호,주소,전화번호,핸드폰번호, ,예매번호,예매자명,Bdate,BdateSeq,연락처");
                _sw01.WriteLine(_strLine.ToString());
                _sw01.Close();


                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    _strBranch = dtable.Rows[i]["card_branch"].ToString();
                    strCard_type_detail = dtable.Rows[i]["card_type_detail"].ToString();
                    strCard_in_date = String.Format("{0:yyyyMMdd}", dtable.Rows[i]["card_in_date"]);
                    strCard_zipcode_kind = dtable.Rows[i]["card_zipcode_kind"].ToString();

                    if (_strBranch == "012")
                    {
                        _strLine = new StringBuilder(GetStringAsLength(dtable.Rows[i]["customer_name"].ToString(), 40) + chCSV);
                        _strLine.Append(GetStringAsLength("", 13) + chCSV);
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_number"].ToString(), 16) + chCSV);

                        if (strCard_zipcode_kind == "1")
                        {
                            _strLine.Append(GetStringAsLength(ConvertZipcode(dtable.Rows[i]["card_zipcode_new"].ToString()), 7) + chCSV);    
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength(ConvertZipcode(dtable.Rows[i]["card_zipcode"].ToString()), 7) + chCSV);
                        }
                        
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_address"].ToString(), 200) + chCSV);
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_tel2"].ToString(), 15) + chCSV);
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_mobile_tel"].ToString(), 15) + chCSV);
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_number"].ToString(), 15) + chCSV);

                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_number"].ToString(), 16) + chCSV);
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["family_name"].ToString(), 40) + chCSV);
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_no"].ToString(), 10) + chCSV);
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_quick_seq"].ToString(), 10) + chCSV);
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_mobile_tel"].ToString(), 15));

                        if (strCard_type_detail == "1051101")
                        {
                            _sw01 = new StreamWriter(fileName + "INT" + strCard_in_date + "_등기", true, _encoding);
                        }
                        else if (strCard_type_detail == "1051102")
                        {
                            _sw01 = new StreamWriter(fileName + "INT" + strCard_in_date + "m_등기", true, _encoding);
                        }
                        else if (strCard_type_detail == "1051103")
                        {
                            _sw01 = new StreamWriter(fileName + "INT" + strCard_in_date + "현대_등기", true, _encoding);
                        }
                        else if (strCard_type_detail == "1051104")
                        {
                            _sw01 = new StreamWriter(fileName + "INT" + strCard_in_date + "교보_등기", true, _encoding);
                        }
                        else
                        {
                            _sw01 = new StreamWriter(fileName + "기타" + strCard_in_date, true, _encoding);
                        }

                        _sw01.WriteLine(_strLine.ToString());
                        _sw01.Close();

                        iCnt++;
                    }
                }
                _strReturn = string.Format("{0}건의 등기데이터 다운 완료", iCnt);
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
            if (_strReturn.Length > 6)
            {
                _strReturn = value.Substring(0, 6) + "-" + value.Substring(6, value.Length - 6);
            }
            else if (_strReturn.Length == 6)
            {
                _strReturn = value.Substring(0, 6) + "-";
            }
            else if (_strReturn.Length == 0)
            {
                _strReturn = "      -";
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


        #region 배송결과코드
        private static string DeilveryStatusType(string value)
        {
            string _strReturn = value;

            switch (value)
            {
                case "0":
                    _strReturn = "88000";
                    break;
                case "1":
                    _strReturn = "88001";
                    break;
                case "2":
                    _strReturn = "88002";
                    break;
                case "3":
                    _strReturn = "88003";
                    break;
                case "4":
                    _strReturn = "88004";
                    break;
                case "5":
                    _strReturn = "88005";
                    break;
                case "6":
                    _strReturn = "88006";
                    break;
                case "7":
                    _strReturn = "88007";
                    break;
                case "99":
                    _strReturn = "88099";
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
