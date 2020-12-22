using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Data;
using System.Text;
using System.Windows.Forms;

namespace _151_TPLUS
{
    public class CONVERT
    {
        #region 전역변수
        private static string strEncoding = "ks_c_5601-1987";
        private static string chCSV = ",";
        private static string strCardTypeID = "151";
        private static string strCardTypeName = "151_티플러스";
        #endregion

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
            //System.Text.Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            //StreamReader _sr = null;
            //byte[] _byteAry = null;
            //string _strReturn = "";
            //string _strLine = "";

            ////파일 읽기 Stream과 오류시 저장할 쓰기 Stream 생성
            //_sr = new StreamReader(path, _encoding);
            //_strLine = _sr.ReadLine();
            //try
            //{
            //    if (_strLine.Trim() != "")
            //    {
            //        _strReturn = _strLine.Substring(_strLine.Length - 7, 7);
            //    }
            //}
            //catch (Exception e)
            //{
            //    MessageBox.Show(e.Message);
            //}

            string _strReturn = "";

            return _strReturn;
        }
        
        public static string ConvertRegister(string path, string xmlZipcodePath, string xmlZipcodeAreaPath, string xmlZipcodePath_new, string xmlZipcodeAreaPath_new, string xmlPath)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            //FileInfo _fi = null;
            StreamReader _sr = null;	                            //파일 읽기 스트림
            StreamWriter _swError = null;			                //파일 쓰기 스트림
            DataSet _dsetZipcode = null, _dsetZipcdeArea = null;	//우편번호 관련 DataSet
            DataSet _dsetZipcode_new = null, _dsetZipcdeArea_new = null;	//우편번호 관련 DataSet
            DataTable _dtable = null;			//마스터 저장 테이블
            DataRow _dr = null;
            DataRow[] _drs = null;
            byte[] _byteAry = null;
            string[] _strAry = null;
            string _strReturn = "";
            string _strLine = "";
            string _strZipcode = "", _strAreaType = "", _strAreaGroup = "", _strBranch = "", strCutmSSN = "", strNum = "";
            int _iSeq = 1, _iErrorCount = 0;
            int _iDiffLength = 0;
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
                _dtable.Columns.Add("customer_office");                 //판매점
                _dtable.Columns.Add("customer_name");                   //가입자명
                _dtable.Columns.Add("customer_ssn");                    //주민번호
                _dtable.Columns.Add("card_mobile_tel");                 //주문자핸드폰
                _dtable.Columns.Add("card_client_no_1");                //신청단말기
                _dtable.Columns.Add("card_brand_code");                 //가입유형
                _dtable.Columns.Add("card_bank_name");                  //이전통신사
                //7                                                                 //7 변경전유형
                _dtable.Columns.Add("card_cooperation1");               //인증유형
                _dtable.Columns.Add("client_express_code");             //인증번호
                _dtable.Columns.Add("card_zipcode");                    //dr[14] 우편번호
                _dtable.Columns.Add("card_address");                    //배송지주소
                //12                                                                //12 가입비구분
                _dtable.Columns.Add("card_design_code");                //유심비구분
                _dtable.Columns.Add("card_product_name");               //요금상품명
                _dtable.Columns.Add("card_cooperation2");               //휴대폰소액결제
                _dtable.Columns.Add("client_send_number");              //개통일련번호  Q
                _dtable.Columns.Add("client_number");                   //유심일련번호  R
                _dtable.Columns.Add("client_register_date");            //개통일자
                //19                                                                //19 납부방법
                _dtable.Columns.Add("family_name");                     //예금주명
                //21                                                                //22 관계
                _dtable.Columns.Add("card_bank_account_name");          //dr[23] 계좌(카드)번호
                _dtable.Columns.Add("card_pt_code");                    //약정기간
                _dtable.Columns.Add("customer_email_domain");           //출고가        Y
                _dtable.Columns.Add("text1");                           //공시지원금    Z
                _dtable.Columns.Add("text2");                           //추가지원금    AA
                _dtable.Columns.Add("text3");                           //할부원금      AB
                _dtable.Columns.Add("text4");                           //월할부금      AC
                _dtable.Columns.Add("text5");                           //할부할인      AD
                _dtable.Columns.Add("text6");                           //할부이자      AE
                _dtable.Columns.Add("text7");                           //매월납부액    AF
                _dtable.Columns.Add("card_bank_ID");                    //요금제
                _dtable.Columns.Add("text8");                           //dr[35]월정액요금
                _dtable.Columns.Add("text9");                           //월요금할인액
                _dtable.Columns.Add("customer_position");               //매월납부액    AJ
                _dtable.Columns.Add("customer_branch");                 //월기본납부액
                _dtable.Columns.Add("card_tel3");                       //업체연락처
                _dtable.Columns.Add("card_level_code");                 //dr[39]할부지원금
                _dtable.Columns.Add("card_cooperation_code");           //dr[40]총할부수수료

                _dtable.Columns.Add("client_quick_seq");                //dr[41]위약금1
                _dtable.Columns.Add("customer_no");                     //dr[42]위약금2
                _dtable.Columns.Add("card_number");                     //dr[43]고객일련번호

                _dtable.Columns.Add("card_zipcode_new");                //dr[44]
                _dtable.Columns.Add("card_zipcode_kind");               //dr[45]
                _dtable.Columns.Add("client_request_memo");             //dr[46] 메모

                //우편번호 관련 정보 DataSet에 담기
                _dsetZipcode = new DataSet();
                _dsetZipcdeArea = new DataSet();
                _dsetZipcode.ReadXml(xmlZipcodePath);
                _dsetZipcode.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcode.Tables[0].Columns["zipcode"] };
                _dsetZipcdeArea.ReadXml(xmlZipcodeAreaPath);
                _dsetZipcdeArea.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcdeArea.Tables[0].Columns["zipcode"] };

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
                    _strAry = _strLine.Split(',');

                    _iDiffLength = 0;
                    _dr = _dtable.NewRow();
                    _dr[0] = _iSeq;

                    _dr[5] = _strAry[0].Trim();
                    _dr[6] = _strAry[1].Trim();

                    string strCustm_ssn = _strAry[2].Replace("-", "").Trim();
                    if (strCustm_ssn.Length == 7)
                    {
                        _dr[7] = strCustm_ssn + "xxxxxx";
                    }
                    else
                    {
                        _dr[7] = strCustm_ssn;
                    }
                    _dr[8] = _strAry[3].Trim();
                    _dr[9] = _strAry[4].Trim();
                    _dr[10] = _strAry[5].Trim();
                    _dr[11] = _strAry[6].Trim();
                    //_strAry[7] 제외
                    _dr[12] = _strAry[8].Trim();
                    _dr[13] = _strAry[9].Trim();
                    _strZipcode = _strAry[10].Replace(" ", "").Replace("-", "").Trim();
                    _dr[14] = _strZipcode;
                    _dr[15] = _strAry[11].Trim();
                    //12
                    _dr[16] = _strAry[13].Trim();
                    _dr[17] = _strAry[14].Trim();
                    _dr[18] = _strAry[15].Trim();
                    _dr[19] = _strAry[16].Trim();
                    _dr[20] = _strAry[17].Replace("[","").Replace("]","").Trim();
                    _dr[21] = _strAry[18].Trim();
                    //19
                    _dr[22] = _strAry[20].Trim();
                    //21
                    _dr[23] = _strAry[22].Trim();
                    _dr[24] = _strAry[23].Trim();

                    if (_dr[24].ToString().Length > 2)
                        _dr[24] = _dr[24].ToString().Substring(0, 2);

                    _dr[25] = _strAry[24].Trim();
                    _dr[26] = _strAry[25].Trim();
                    _dr[27] = _strAry[26].Trim();
                    _dr[28] = _strAry[27].Trim();
                    _dr[29] = _strAry[28].Trim();
                    _dr[30] = _strAry[29].Trim();
                    _dr[31] = _strAry[30].Trim();
                    _dr[32] = _strAry[31].Trim();
                    _dr[33] = _strAry[32].Trim();
                    _dr[34] = _strAry[33].Trim();
                    _dr[35] = _strAry[34].Trim();
                    _dr[36] = _strAry[35].Trim();
                    _dr[37] = _strAry[36].Trim();
                    _dr[38] = _strAry[37].Trim();
                    _dr[39] = _strAry[38].Trim();
                    _dr[40] = _strAry[39].Trim();
                    _dr[41] = _strAry[40].Trim();
                    _dr[42] = _strAry[41].Trim();

                    _dr[43] = _strAry[42].Trim();

                    _dr[46] = _strAry[20].Trim() + " / " + _strAry[21].Trim();

                    if (_strZipcode.Length == 5)
                    {
                        _dr[44] = _strZipcode;
                        _dr[45] = "1";
                    }

                    if (_strAry.LongLength != 43)
                    {
                        MessageBox.Show("줄번호 " + _iSeq.ToString() + " 번째 배열 갯수 오류 입니다. 데이터 레이아웃을 확인하세요", "오류");
                        throw new ArgumentNullException("배열 갯수 오류");
                    }

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
            StreamWriter _sw00 = null, _sw01 = null, _sw02 = null, _sw03 = null, _sw04 = null, _sw05 = null, _sw06 = null, _sw07 = null, _sw08 = null;		//파일 쓰기 스트림
            StreamWriter _sw09 = null, _sw10 = null, _sw11 = null, _sw12 = null, _sw13 = null, _sw14 = null, _sw15 = null;		//파일 쓰기 스트림
            StreamWriter _sw21 = null, _sw22 = null, _sw23 = null, _sw24 = null, _sw25 = null, _sw26 = null, _sw27 = null, _sw28 = null, _sw29 = null;

            int i = 0, result_count1 = 0, result_count2 = 0, result_count3 = 0, result_count4 = 0, result_count5 = 0, result_count6 = 0, result_count7 = 0, result_count8 = 0, result_count9 = 0, result_count10 = 0, result_count11 = 0, result_count12 = 0, result_count13 = 0, result_count14 = 0, result_count15 = 0;

            int result_count21 = 0, result_count22 = 0, result_count23 = 0, result_count24 = 0, result_count25 = 0, result_count26 = 0, result_count27 = 0
                , result_count28 = 0, result_count29 = 0;

            StringBuilder _strLine = new StringBuilder("");
            string _strReturn = "", strStatus = "", strCard_type_detail = "";
            try
            {
                //_sw00 = new StreamWriter(fileName + ".00", true, _encoding);
                //_sw01 = new StreamWriter(fileName + ".01", true, _encoding);

                //title
                _sw01 = new StreamWriter(fileName + "시연정보통신_" + DateTime.Now.ToString("yyMMdd"), true, _encoding);
                _sw02 = new StreamWriter(fileName + "엠제이티_" + DateTime.Now.ToString("yyMMdd"), true, _encoding);
                _sw03 = new StreamWriter(fileName + "투인원01_" + DateTime.Now.ToString("yyMMdd"), true, _encoding);
                _sw04 = new StreamWriter(fileName + "투인원04_" + DateTime.Now.ToString("yyMMdd"), true, _encoding);
                _sw05 = new StreamWriter(fileName + "이레_" + DateTime.Now.ToString("yyMMdd"), true, _encoding);
                _sw06 = new StreamWriter(fileName + "미소_" + DateTime.Now.ToString("yyMMdd"), true, _encoding);
                _sw07 = new StreamWriter(fileName + "플러스엠_" + DateTime.Now.ToString("yyMMdd"), true, _encoding);
                _sw08 = new StreamWriter(fileName + "라인텔_" + DateTime.Now.ToString("yyMMdd"), true, _encoding);
                _sw09 = new StreamWriter(fileName + "이룸파트너스_" + DateTime.Now.ToString("yyMMdd"), true, _encoding);
                _sw10 = new StreamWriter(fileName + "지에스텔레콤_" + DateTime.Now.ToString("yyMMdd"), true, _encoding);
                _sw11 = new StreamWriter(fileName + "마스터_" + DateTime.Now.ToString("yyMMdd"), true, _encoding);
                _sw12 = new StreamWriter(fileName + "Cello_" + DateTime.Now.ToString("yyMMdd"), true, _encoding);
                _sw13 = new StreamWriter(fileName + "Cello2_" + DateTime.Now.ToString("yyMMdd"), true, _encoding);
                _sw14 = new StreamWriter(fileName + "TOP_" + DateTime.Now.ToString("yyMMdd"), true, _encoding);
                _sw15 = new StreamWriter(fileName + "Cello3_" + DateTime.Now.ToString("yyMMdd"), true, _encoding);

                _sw21 = new StreamWriter(fileName + "KT_시연_" + DateTime.Now.ToString("yyMMdd"), true, _encoding);
                _sw22 = new StreamWriter(fileName + "KT_Cello_" + DateTime.Now.ToString("yyMMdd"), true, _encoding);
                _sw23 = new StreamWriter(fileName + "KT_Cello2_" + DateTime.Now.ToString("yyMMdd"), true, _encoding);
                _sw24 = new StreamWriter(fileName + "KT_마스터_" + DateTime.Now.ToString("yyMMdd"), true, _encoding);
                _sw25 = new StreamWriter(fileName + "KT_이레_" + DateTime.Now.ToString("yyMMdd"), true, _encoding);
                _sw26 = new StreamWriter(fileName + "KT_이룸파트너스_" + DateTime.Now.ToString("yyMMdd"), true, _encoding);
                _sw27 = new StreamWriter(fileName + "KT_GS텔레콤_" + DateTime.Now.ToString("yyMMdd"), true, _encoding);
                _sw28 = new StreamWriter(fileName + "KT_TOP_" + DateTime.Now.ToString("yyMMdd"), true, _encoding);
                _sw29 = new StreamWriter(fileName + "KT_Cello3_" + DateTime.Now.ToString("yyMMdd"), true, _encoding);
                

                _strLine.Append("순서,업체명,인수일,휴대폰번호,성명,생년월일,수령인,관계,배송여부,배송일,반송사유");
                _sw01.WriteLine(_strLine);
                _sw02.WriteLine(_strLine);
                _sw03.WriteLine(_strLine);
                _sw04.WriteLine(_strLine);
                _sw05.WriteLine(_strLine);
                _sw06.WriteLine(_strLine);
                _sw07.WriteLine(_strLine);
                _sw08.WriteLine(_strLine);
                _sw09.WriteLine(_strLine);
                _sw10.WriteLine(_strLine);
                _sw11.WriteLine(_strLine);
                _sw12.WriteLine(_strLine);
                _sw13.WriteLine(_strLine);
                _sw14.WriteLine(_strLine);
                _sw15.WriteLine(_strLine);

                //KT
                _sw21.WriteLine(_strLine);
                _sw22.WriteLine(_strLine);
                _sw23.WriteLine(_strLine);
                _sw24.WriteLine(_strLine);
                _sw25.WriteLine(_strLine);
                _sw26.WriteLine(_strLine);
                _sw27.WriteLine(_strLine);
                _sw28.WriteLine(_strLine);
                _sw29.WriteLine(_strLine);

                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    strStatus = dtable.Rows[i]["card_delivery_status"].ToString();
                    strCard_type_detail = dtable.Rows[i]["card_type_detail"].ToString();

                    if (strStatus == "1" || strStatus == "2" || strStatus == "3")
                    {
                        switch (strCard_type_detail)
                        {
                            case "1512101":
                                result_count1++;
                                //순번
                                _strLine = new StringBuilder(GetStringAsLength(result_count1.ToString(), 5) + chCSV);
                                break;
                            case "1512102":
                                result_count2++;
                                //순번
                                _strLine = new StringBuilder(GetStringAsLength(result_count2.ToString(), 5) + chCSV);
                                break;
                            case "1512103":
                                result_count3++;
                                //순번
                                _strLine = new StringBuilder(GetStringAsLength(result_count3.ToString(), 5) + chCSV);
                                break;
                            case "1512104":
                                result_count4++;
                                //순번
                                _strLine = new StringBuilder(GetStringAsLength(result_count4.ToString(), 5) + chCSV);
                                break;
                            case "1512105":
                                result_count5++;
                                //순번
                                _strLine = new StringBuilder(GetStringAsLength(result_count5.ToString(), 5) + chCSV);
                                break;
                            case "1512106":
                                result_count6++;
                                //순번
                                _strLine = new StringBuilder(GetStringAsLength(result_count6.ToString(), 5) + chCSV);
                                break;
                            case "1512107":
                                result_count7++;
                                //순번
                                _strLine = new StringBuilder(GetStringAsLength(result_count7.ToString(), 5) + chCSV);
                                break;
                            case "1512108":
                                result_count8++;
                                //순번
                                _strLine = new StringBuilder(GetStringAsLength(result_count8.ToString(), 5) + chCSV);
                                break;
                            case "1512109":
                                result_count9++;
                                //순번
                                _strLine = new StringBuilder(GetStringAsLength(result_count9.ToString(), 5) + chCSV);
                                break;
                            case "1512110":
                                result_count10++;
                                //순번
                                _strLine = new StringBuilder(GetStringAsLength(result_count10.ToString(), 5) + chCSV);
                                break;
                            case "1512111":
                                result_count11++;
                                //순번
                                _strLine = new StringBuilder(GetStringAsLength(result_count11.ToString(), 5) + chCSV);
                                break;
                            case "1512112":
                                result_count12++;
                                //순번
                                _strLine = new StringBuilder(GetStringAsLength(result_count12.ToString(), 5) + chCSV);
                                break;
                            case "1512113":
                                result_count13++;
                                //순번
                                _strLine = new StringBuilder(GetStringAsLength(result_count13.ToString(), 5) + chCSV);
                                break;
                            case "1512114":
                                result_count14++;
                                //순번
                                _strLine = new StringBuilder(GetStringAsLength(result_count14.ToString(), 5) + chCSV);
                                break;
                            case "1512115":
                                result_count15++;
                                //순번
                                _strLine = new StringBuilder(GetStringAsLength(result_count15.ToString(), 5) + chCSV);
                                break;

                            //KT
                            case "1512201":
                                result_count21++;
                                //순번
                                _strLine = new StringBuilder(GetStringAsLength(result_count21.ToString(), 5) + chCSV);
                                break;
                            case "1512202":
                                result_count22++;
                                //순번
                                _strLine = new StringBuilder(GetStringAsLength(result_count22.ToString(), 5) + chCSV);
                                break;
                            case "1512203":
                                result_count23++;
                                //순번
                                _strLine = new StringBuilder(GetStringAsLength(result_count23.ToString(), 5) + chCSV);
                                break;
                            case "1512204":
                                result_count24++;
                                //순번
                                _strLine = new StringBuilder(GetStringAsLength(result_count24.ToString(), 5) + chCSV);
                                break;
                            case "1512205":
                                result_count25++;
                                //순번
                                _strLine = new StringBuilder(GetStringAsLength(result_count25.ToString(), 5) + chCSV);
                                break;
                            case "1512206":
                                result_count26++;
                                //순번
                                _strLine = new StringBuilder(GetStringAsLength(result_count26.ToString(), 5) + chCSV);
                                break;
                            case "1512207":
                                result_count27++;
                                //순번
                                _strLine = new StringBuilder(GetStringAsLength(result_count27.ToString(), 5) + chCSV);
                                break;
                            case "1512208":
                                result_count28++;
                                //순번
                                _strLine = new StringBuilder(GetStringAsLength(result_count28.ToString(), 5) + chCSV);
                                break;
                            case "1512209":
                                result_count29++;
                                //순번
                                _strLine = new StringBuilder(GetStringAsLength(result_count29.ToString(), 5) + chCSV);
                                break;
                            default:
                                break;
                        }
                    }
                    
                    //업체명
                    _strLine.Append(dtable.Rows[i]["customer_office"].ToString() + chCSV);
                    //인수일
                    _strLine.Append(GetStringAsLength(String.Format("{0:yyyy-MM-dd}", dtable.Rows[i]["card_in_date"].ToString()), 10) + chCSV);
                    //휴대폰번호
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_mobile_tel"].ToString(), 15) + chCSV);
                    //고객명
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_name"].ToString(), 20) + chCSV);
                    //생년월일
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["customer_ssn"].ToString().Replace("x",""),13) + chCSV);
                    //수령인
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_name"].ToString(), 20) + chCSV);

                    //수령인관계, 배송여부, 인수일, 배송일, 배송직원, 반송사유
                    if (strStatus == "1")
                    {
                        //수령인관계
                        if (dtable.Rows[i]["card_result_status"].ToString() == "61")
                        {
                            _strLine.Append(GetStringAsLength("본인", 10) + chCSV);
                        }
                        else if (dtable.Rows[i]["card_result_status"].ToString() == "62")
                        {
                            _strLine.Append(GetStringAsLength("대리", 10) + chCSV);
                        }
                        //배송여부
                        _strLine.Append(GetStringAsLength("배달", 10) + chCSV);
                        //배송일
                        _strLine.Append(GetStringAsLength(String.Format("{0:yyyy-MM-dd}", dtable.Rows[i]["card_delivery_date"].ToString()), 10) + chCSV);
                        _strLine.Append(GetStringAsLength("", 18) + chCSV);
                    }
                    else if (strStatus == "2" || strStatus == "3")
                    {
                        _strLine.Append(GetStringAsLength("", 20) + chCSV);
                        _strLine.Append(GetStringAsLength("반송", 10) + chCSV);
                        _strLine.Append(GetStringAsLength("", 10) + chCSV);
                        _strLine.Append(GetStringAsLength(ReturnType(dtable.Rows[i]["delivery_return_reason_last"].ToString()),20) + chCSV);
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength("", 20) + chCSV);
                        _strLine.Append(GetStringAsLength("", 10) + chCSV);
                        _strLine.Append(GetStringAsLength("", 10) + chCSV);
                        _strLine.Append(GetStringAsLength("", 18) + chCSV);
                    }

                    //배송, 반송
                    if (strStatus == "1" || strStatus == "2" || strStatus == "3")
                    {
                        switch (strCard_type_detail)
                        {
                            case "1512101": _sw01.WriteLine(_strLine); break;
                            case "1512102": _sw02.WriteLine(_strLine); break;
                            case "1512103": _sw03.WriteLine(_strLine); break;
                            case "1512104": _sw04.WriteLine(_strLine); break;
                            case "1512105": _sw05.WriteLine(_strLine); break;
                            case "1512106": _sw06.WriteLine(_strLine); break;
                            case "1512107": _sw07.WriteLine(_strLine); break;
                            case "1512108": _sw08.WriteLine(_strLine); break;
                            case "1512109": _sw09.WriteLine(_strLine); break;
                            case "1512110": _sw10.WriteLine(_strLine); break;
                            case "1512111": _sw11.WriteLine(_strLine); break;
                            case "1512112": _sw12.WriteLine(_strLine); break;
                            case "1512113": _sw13.WriteLine(_strLine); break;
                            case "1512114": _sw14.WriteLine(_strLine); break;
                            case "1512115": _sw15.WriteLine(_strLine); break;

                            //KT
                            case "1512201": _sw21.WriteLine(_strLine); break;
                            case "1512202": _sw22.WriteLine(_strLine); break;
                            case "1512203": _sw23.WriteLine(_strLine); break;
                            case "1512204": _sw24.WriteLine(_strLine); break;
                            case "1512205": _sw25.WriteLine(_strLine); break;
                            case "1512206": _sw26.WriteLine(_strLine); break;
                            case "1512207": _sw27.WriteLine(_strLine); break;
                            case "1512208": _sw28.WriteLine(_strLine); break;
                            case "1512209": _sw29.WriteLine(_strLine); break;
                            default:
                                break;
                        }
                    }
                    //그외
                    //else
                    //{
                    //    switch (strCard_type_detail)
                    //    {
                    //        case "1512101":
                    //            _sw00 = new StreamWriter(fileName + "미배송_시연정보통신_" + DateTime.Now.ToString("yyMMdd"), true, _encoding); break;
                    //        case "1512102":
                    //            _sw00 = new StreamWriter(fileName + "미배송_엠제이티_" + DateTime.Now.ToString("yyMMdd"), true, _encoding); break;
                    //        default:
                    //            break;
                    //    }
                    //    _sw00.WriteLine(_strLine);
                    //    _sw00.Close();
                    //}
                }
                _strReturn = string.Format("{0}건의 인계데이터 다운 완료", i);
            }
            catch (Exception ex)
            {
                _strReturn = string.Format("{0}번째 데이터 인계 중 오류 발생", i + 1);
                _strLine.Append(ex.Message);
                _sw00.WriteLine(_strLine);
            }
            finally
            {
                if (_sw01 != null) _sw01.Close();
                if (_sw02 != null) _sw02.Close();
                if (_sw03 != null) _sw03.Close();
                if (_sw04 != null) _sw04.Close();
                if (_sw05 != null) _sw05.Close();
                if (_sw06 != null) _sw06.Close();
                if (_sw07 != null) _sw07.Close();
                if (_sw08 != null) _sw08.Close();
                if (_sw09 != null) _sw09.Close();
                if (_sw10 != null) _sw10.Close();
                if (_sw11 != null) _sw11.Close();
                if (_sw12 != null) _sw12.Close();
                if (_sw13 != null) _sw13.Close();
                if (_sw14 != null) _sw14.Close();
                if (_sw15 != null) _sw15.Close();

                //KT
                if (_sw21 != null) _sw21.Close();
                if (_sw22 != null) _sw22.Close();
                if (_sw23 != null) _sw23.Close();
                if (_sw24 != null) _sw24.Close();
                if (_sw25 != null) _sw25.Close();
                if (_sw26 != null) _sw26.Close();
                if (_sw27 != null) _sw27.Close();
                if (_sw28 != null) _sw28.Close();
                if (_sw29 != null) _sw29.Close();
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
        public static string DeliveryType(string strStatus)
        {
            if (strStatus == "1")
            {
                return "배달";
            }
            else if (strStatus == "2" || strStatus == "3")
            {
                return "반송";
            }
            else
            {
                return "";
            }
        }
        #endregion

        #region 수령인관계
        public static string ReceiverType(string value)
        {
            string strReturn = null;
            string _strReceiver = value;
            switch (_strReceiver)
            {
                case "00": strReturn = "미입력"; break;
                case "01": strReturn = "본인"; break;
                case "02": strReturn = "조부"; break;
                case "03": strReturn = "조모"; break;
                case "04": strReturn = "아버지"; break;
                case "05": strReturn = "어머니"; break;
                case "06": strReturn = "배우자"; break;
                case "07": strReturn = "남편"; break;
                case "08": strReturn = "형"; break;
                case "09": strReturn = "동생"; break;
                case "10": strReturn = "언니"; break;
                case "11": strReturn = "누나"; break;
                case "12": strReturn = "아들"; break;
                case "13": strReturn = "딸"; break;
                case "14": strReturn = "손자"; break;
                case "15": strReturn = "손녀"; break;
                case "16": strReturn = "며느리"; break;
                case "17": strReturn = "삼촌"; break;
                case "18": strReturn = "사촌"; break;
                case "19": strReturn = "친척"; break;
                case "20": strReturn = "직장동료"; break;
                case "21": strReturn = "상사"; break;
                case "22": strReturn = "친구"; break;
                case "23": strReturn = "선배"; break;
                case "24": strReturn = "후배"; break;
                case "25": strReturn = "은행원"; break;
                case "26": strReturn = "이웃"; break;
                case "27": strReturn = "주인집"; break;
                case "28": strReturn = "형수"; break;
                case "29": strReturn = "경비원"; break;
                case "30": strReturn = "친지"; break;
                case "31": strReturn = "오빠"; break;
                case "32": strReturn = "보증인"; break;
                case "33": strReturn = "고객요청"; break;
                case "34": strReturn = "공무원"; break;
                case "98": strReturn = "등기"; break;
                default:
                    strReturn = "기타"; break;
            }
            return strReturn;
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
