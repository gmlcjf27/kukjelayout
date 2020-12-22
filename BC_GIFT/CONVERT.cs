using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace BC_GIFT
{
	public class CONVERT
	{
		//기본 인코딩 설정
		private static string strEncoding = "ks_c_5601-1987";
        private static string strCardTyepID = "042";
        private static string strCardTypeName = "BC기프트";

		//현 DLL의 카드 타입 코드 반환
		public static string GetCardTypeID() {
            return strCardTyepID;
		}

		//현 DLL의 카드 타입명 반환
		public static string GetCardTypeName() {
            return strCardTypeName;
		}

        //제휴사코드 반환
        public static string GetCardType(string path)
        {
            string _strReturn = "0013501";

            return _strReturn;
        }

		//등록 자료 생성
		//public static string ConvertRegister(string path, string xmlZipcodePath, string xmlZipcodeAreaPath, string xmlPath)
        public static string ConvertRegister(string path, string xmlZipcodePath, string xmlZipcodeAreaPath, string xmlZipcodePath_new, string xmlZipcodeAreaPath_new, string xmlPath)
		{
			System.Text.Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
			FileInfo _fi = null;
			StreamReader _sr = null;														//파일 읽기 스트림
			StreamWriter _swError = null;													//파일 쓰기 스트림
			DataSet _dsetZipcode = null, _dsetZipcdeArea = null;							//우편번호 관련 DataSet
            DataSet _dsetZipcode_new = null, _dsetZipcdeArea_new = null;					//우편번호 관련 DataSet
			DataTable _dtable = null;														//마스터 저장 테이블
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
				_dtable = new DataTable("Convert");
				//기본 컬럼
				_dtable.Columns.Add("degree_arrange_number");
				_dtable.Columns.Add("card_area_group");
				_dtable.Columns.Add("card_branch");
				_dtable.Columns.Add("card_area_type");
				_dtable.Columns.Add("area_arrange_number");

				//세부 컬럼
				_dtable.Columns.Add("client_send_date");             // dr[5]
				_dtable.Columns.Add("client_send_number");
				_dtable.Columns.Add("card_denomination");
				_dtable.Columns.Add("card_count");
				_dtable.Columns.Add("card_bank_name");
                _dtable.Columns.Add("card_bank_ID");                 // dr[10]
				_dtable.Columns.Add("card_tel1");
				_dtable.Columns.Add("card_zipcode");
                _dtable.Columns.Add("card_address");
                _dtable.Columns.Add("customer_name");
                _dtable.Columns.Add("card_number");                  // dr[15]

                _dtable.Columns.Add("card_zipcode_new");             // dr[17]
                _dtable.Columns.Add("card_zipcode_kind");            // dr[18]

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
                _dsetZipcdeArea_new.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcdeArea.Tables[0].Columns["zipcode_new"] };


				while ((_strLine = _sr.ReadLine()) != null)
				{
					//인코딩, byte 배열로 담기
					_byteAry = _encoding.GetBytes(_strLine);

					_dr = _dtable.NewRow();
					_dr[0] = _iSeq;
					_dr[5] = _encoding.GetString(_byteAry, 0, 8);
					_dr[6] = _encoding.GetString(_byteAry, 8, 8);
					_dr[7] = _encoding.GetString(_byteAry, 33, 6);
					_dr[8] = _encoding.GetString(_byteAry, 39, 6);
					_dr[9] = _encoding.GetString(_byteAry, 46, 20);
					_dr[10] = _encoding.GetString(_byteAry, 67, 6);
					_dr[11] = _encoding.GetString(_byteAry, 74, 13);
					_strZipcode = _encoding.GetString(_byteAry, 87, 7).Trim();
					_dr[12] = _strZipcode;

                    if (_strZipcode.Length == 5)
                    {
                        _dr[16] = _strZipcode;
                        _dr[17] = "1";
                    }

                    _dr[13] = _encoding.GetString(_byteAry, 94, 50);
                    _dr[14] = _encoding.GetString(_byteAry, 46, 20);

                    if (_encoding.GetString(_byteAry, 144, 1) != "1")
                    {
                        _dr[13] = _encoding.GetString(_byteAry, 145, 100).Trim() + _encoding.GetString(_byteAry, 245, 100).Trim();
                    }

                    //카드번호
                    _dr[15] = _encoding.GetString(_byteAry, 0, 16); 

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

                        //_drs = _dsetZipcdeArea.Tables[0].Select("zipcode=" + _strZipcode);

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

                        //_drs = _dsetZipcode.Tables[0].Select("zipcode=" + _strZipcode);

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
                    _fi = new FileInfo(path + ".Error");
                    _fi.Delete();
                    _strReturn = string.Format("{0}건의 데이터 변환 성공", _iSeq-1);
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
			finally {
				if (_swError != null) _swError.Close();
				if (_sr != null) _sr.Close();
			}
			return _strReturn;
		}

		//기본 마감 자료
        public static string ConvertResult(DataTable dtable, string fileName) 
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩
			StreamWriter _sw = null;																			//파일 쓰기 스트림
            StringBuilder _strLine = new StringBuilder("");
            string _strReturn = "";
            // 배송상태, 수령인관계명, 반송사유
            string _strStatus = null, strReceiver_Code_Name = null, _strReturnCode_Name = null;
            int i = 0;

            try 
            {
                _sw = new StreamWriter(fileName, true, _encoding);

				for (i = 0; i < dtable.Rows.Count; i++) 
                {
                    _strStatus = dtable.Rows[i]["card_delivery_status"].ToString();
                    strReceiver_Code_Name = Receiver_Code_Name(dtable.Rows[i]["receiver_code"].ToString());
                    _strReturnCode_Name = ReturnType(dtable.Rows[i]["delivery_return_reason_last"].ToString());

                    _strLine = new StringBuilder(GetStringAsLength(RemoveDash(dtable.Rows[i]["client_send_date"].ToString()), 8));
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_send_number"].ToString(), 8));
                    _strLine.Append(GetStringAsLength("", 17)); // 구분값(공백)
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_denomination"].ToString(), 6));
                    _strLine.Append(GetStringAsLength("", 3));   // 구분값(공백)
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_count"].ToString(), 6));
                    _strLine.Append(GetStringAsLength("", 1));   // 구분값(공백)
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_bank_name"].ToString(), 20));
                    _strLine.Append(GetStringAsLength("", 1));   // 구분값(공백)
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_bank_ID"].ToString(), 6));
                    _strLine.Append(GetStringAsLength("", 1));   // 구분값(공백)
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_tel1"].ToString(), 13));
                    _strLine.Append(GetStringAsLength("", 1));   // 구분값(공백)
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_zipcode"].ToString(), 6));
                    _strLine.Append(GetStringAsLength("", 1)); // 구분값(공백)
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_address"].ToString(), 200));

                    // 배송이면 수령일
                    if (_strStatus == "1")
                    {
                        _strLine.Append(GetStringAsLength(RemoveDash(dtable.Rows[i]["card_delivery_date"].ToString()), 8));
                        _strLine.Append(GetStringAsLength("", 1)); // 구분값(공백)
                        if (dtable.Rows[i]["receiver_code"].ToString() == "01")
                        {
                            _strLine.Append(GetStringAsLength("본인수령", 8));
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength("대리수령", 8));
                        }
                        _strLine.Append(GetStringAsLength("", 1)); // 구분값(공백)
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_name"].ToString(), 30));
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength("", 8));
                        _strLine.Append(GetStringAsLength("", 1)); // 구분값(공백)
                        _strLine.Append(GetStringAsLength("", 8));
                        _strLine.Append(GetStringAsLength("", 1)); // 구분값(공백)
                        _strLine.Append(GetStringAsLength("", 30));
                    }
                    _strLine.Append(GetStringAsLength("", 1));   // 구분값(공백)
                    // 수령인 관계
                    _strLine.Append(GetStringAsLength(strReceiver_Code_Name, 20));

                    if (_strStatus == "2" || _strStatus == "3")
                    {
                        // 반송사유
                        _strLine.Append(GetStringAsLength(_strReturnCode_Name, 20));
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength("", 20));
                    }
                    
					_sw.WriteLine(_strLine);
				}
                _strReturn = string.Format("{0}건의 인계데이터 다운 완료", i + 1);
			}
			catch (Exception)
			{
                _strReturn = string.Format("{0}번째 데이터 인계 중 오류 발생", i + 1);
			}
			finally
			{
				_sw.Close();
			}
			return _strReturn;
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
                case "99": strReturn = "기타"; break;
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
