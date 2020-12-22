using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Text;
using System.IO;
using System.Data;
using System.Windows.Forms;

namespace _147BS랜탈
{
    public class CONVERT
    {
        //기본 인코딩 설정
        private static string strEncoding = "ks_c_5601-1987";
        private static char chCSV = ',';
        private static string strCardTypeID = "147";
        private static string strCardTypeName = "BS랜탈";
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
        public static string ConvertRegister(string path, string xmlZipcodePath, string xmlZipcodeAreaPath, string xmlZipcodePath_new, string xmlZipcodeAreaPath_new, string xmlPath)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            //FileInfo _fi = null;
            StreamReader _sr = null;												//파일 읽기 스트림
            StreamWriter _swError = null;											//파일 쓰기 스트림
            DataSet _dsetZipcode = null, _dsetZipcdeArea = null;					//우편번호 관련 DataSet
            DataSet _dsetZipcode_new = null, _dsetZipcdeArea_new = null;					//우편번호 관련 DataSet
            DataTable _dtable = null;							    				//마스터 저장 테이블
            DataRow _dr = null;
            DataRow[] _drs = null;
            string _strReturn = "";
            string _strLine = "";
            string[] _strAry = null;
            string _strZipcode = "", _strZipcode2 = "", _strAreaType = "", _strAreaGroup = "", _strBranch = "";
            int _iSeq = 1, _iErrorCount = 0, _iErrorQuotaCount = 0;
            Branches _branches = new Branches();
            try
            {
                _dtable = new DataTable("Convert");
                //기본 컬럼
                _dtable.Columns.Add("degree_arrange_number");
                _dtable.Columns.Add("card_area_group");
                _dtable.Columns.Add("Card_branch");
                _dtable.Columns.Add("card_area_type");
                _dtable.Columns.Add("area_arrange_number");
                //세부 컬럼
                _dtable.Columns.Add("client_send_number");         //dr[5] 계약번호
                _dtable.Columns.Add("customer_name");              //계약자명
                _dtable.Columns.Add("customer_ssn");               //생년월일 6자리
                _dtable.Columns.Add("card_register_type");         //남녀구분 (출력용) 남:1, 여:2 (외국인, 2000년 이후 출생 무관)
                _dtable.Columns.Add("card_bank_ID");                //상품명
                _dtable.Columns.Add("customer_email_id");          //dr[10]월렌탈료
                _dtable.Columns.Add("card_pt_code");               //할부개월
                _dtable.Columns.Add("customer_email_domain");      //양도가
                _dtable.Columns.Add("customer_office");            //계약처
                _dtable.Columns.Add("card_zipcode");              //계약자우편번호
                _dtable.Columns.Add("card_address");             //dr[15]계약자(수령지)주소
                _dtable.Columns.Add("card_zipcode2");            //설치(물품)우편번호
                _dtable.Columns.Add("card_address2");            //(물품)주소1
                _dtable.Columns.Add("card_tel1");                //계약자연락처
                _dtable.Columns.Add("card_mobile_tel");          //계약자휴대폰
                _dtable.Columns.Add("card_request_memo");         //dr[20]배송메모
                _dtable.Columns.Add("card_bank_account_name");   // 카드사명
                _dtable.Columns.Add("card_bank_account_no");     //결제카드번호
                _dtable.Columns.Add("card_limit");               //유효기간
                _dtable.Columns.Add("card_bank_account_owner");  //카드소유주
                _dtable.Columns.Add("card_payment_day");         //dr[25]납입일자
                _dtable.Columns.Add("card_agree_code");         //구분코드
                _dtable.Columns.Add("card_number");              //주문번호 조회용
                _dtable.Columns.Add("card_is_for_owner_only");    //dr[28]

                _dtable.Columns.Add("card_zipcode_new");          //dr[29]
                _dtable.Columns.Add("card_zipcode_kind");         //dr[30]
                _dtable.Columns.Add("card_zipcode2_new");         //dr[31]
                _dtable.Columns.Add("card_zipcode2_kind");        //dr[32]

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
                    _dr = _dtable.NewRow();
                    _dr[0] = _iSeq;
                    if (_strLine.IndexOf('"') > -1) _iErrorQuotaCount++;
                    //CSV 분리
                    _strAry = _strLine.Split(chCSV);

                    _dr[5] = RemoveDash(_strAry[0]);
                    _dr[6] = _strAry[1];
                    _dr[7] = _strAry[2].Trim() + "xxxxxxx";
                    _dr[8] = _strAry[3].ToUpper().Replace("F", "2").Replace("M", "1");
                    _dr[9] = _strAry[4];
                    _dr[10] = _strAry[5];
                    _dr[11] = _strAry[6];
                    _dr[12] = _strAry[7];
                    _dr[13] = _strAry[8];
                    _strZipcode = RemoveDash(_strAry[9]).Trim();

                    if (_strZipcode.Length == 5)
                    {
                        _dr[29] = _strZipcode;
                        _dr[30] = "1";
                    }
                    _dr[14] = _strZipcode;
                    _dr[15] = _strAry[10];
                    _strZipcode2 = RemoveDash(_strAry[11]).Trim();

                    if (_strZipcode2.Length ==5)
                    {
                        _dr[31] = _strZipcode2;
                        _dr[32] = "1";
                    }
                    _dr[16] = _strZipcode2;
                    _dr[17] = _strAry[12];
                    _dr[18] = _strAry[13];
                    _dr[19] = _strAry[14];
                    _dr[20] = _strAry[15];
                    _dr[21] = _strAry[16];
                    _dr[22] = _strAry[17];
                    _dr[23] = _strAry[18];
                    _dr[24] = _strAry[19];
                    _dr[25] = _strAry[20].Replace("일","");
                    _dr[26] = _strAry[21];
                    _dr[27] = RemoveDash(_strAry[0]);
                    _dr[28] = "1";

                    if (_strAry.LongLength != 23)
                    {
                        MessageBox.Show("줄번호 " + _iSeq.ToString() +  " 번째 배열 갯수 오류 입니다. 데이터 레이아웃을 확인하세요", "오류");
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
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            StreamWriter _sw00 = null, _sw01 = null;								//파일 쓰기 스트림
            int i = 0, k =0;
            //_strLine = CJ오쇼핑 마감
            StringBuilder _strLine = new StringBuilder("");
            string _strReturn = "", _strStatus = "";

            try
            {
                _sw00 = new StreamWriter(fileName + "비에스_미배송", true, _encoding);
                _sw01 = new StreamWriter(fileName + "비에스_배송_반송", true, _encoding);

                _sw01.WriteLine("순번,계약번호,계약자,상품명,배송결과,고객서명일,반송사유");
                

                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    _strStatus = dtable.Rows[i]["card_delivery_status"].ToString();

                    if (_strStatus == "1" || _strStatus == "2" || _strStatus == "3")
                    {
                        k++;
                    }

                    _strLine = new StringBuilder(k.ToString() + chCSV);
                    _strLine.Append(dtable.Rows[i]["client_send_number"].ToString() + chCSV); //주문번호
                    _strLine.Append(dtable.Rows[i]["customer_name"].ToString() + chCSV);
                    _strLine.Append(dtable.Rows[i]["card_bank_ID"].ToString() + chCSV);

                    //배송
                    if (_strStatus == "1")
                    {
                        _strLine.Append("배송" + chCSV);
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_delivery_date"].ToString().Trim(), 10) + chCSV);
                        _strLine.Append("");
                    }
                    //반송
                    else if (_strStatus == "2" || _strStatus == "3")
                    {   
                        _strLine.Append("반송" + chCSV);
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["delivery_return_date_last"].ToString().Trim(), 10) + chCSV);
                        _strLine.Append(ReturnType(dtable.Rows[i]["delivery_return_reason_last"].ToString()));
                    }
                    else
                    {   
                        _strLine.Append("미배송" + chCSV);
                        _strLine.Append("" + chCSV);
                        _strLine.Append("");
                    }

                    if (_strStatus == "1" || _strStatus == "2" || _strStatus == "3")
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
                if (_sw01 != null) _sw01.Close();
                if (_sw00 != null) _sw00.Close();
            }
            return _strReturn;
        }

        //일일마감자료
        public static string ConvertResultDay(System.Data.DataTable dtable, string fileName)
        {
            return ConvertResult(dtable, fileName);
        }


        // 수령인관계
        #region 수령인관계
        private static string Receiver_Code_Name(string _strReceiver_code)
        {
            string strReturn = null;
            switch (_strReceiver_code)
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
