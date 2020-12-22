using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Text;

namespace _003_신한_인수_data
{
	public class CONVERT
	{
		//기본 인코딩 설정
		private static string strEncoding = "ks_c_5601-1987";
        private static string strCardTypeID = "003";
        //주의 : 신한_인수 문구 변경 시 배송프로그램 영향 받음
        private static string strCardTypeName = "신한_인수";

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
            string _strReturn = "";

            return _strReturn;
        }

        //등록 자료 생성
        public static string ConvertRegister(string path, string xmlZipcodePath, string xmlZipcodeAreaPath, string xmlPath)
        {   
            string _strReturn = "";

            _strReturn = "발송 전용 레이아웃 입니다.";
            return _strReturn;
        }

        //마감 자료
        public static string ConvertResult(System.Data.DataTable dtable, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            StreamWriter _sw = null;											//파일 쓰기 스트림
            StringBuilder _strLine = new StringBuilder("");
            string _strReturn = "", strCareer = "";
            DataRow[] _drs = null;
            int i = 0;

            try
            {
                string temp_Indate = "";

                if (dtable.Rows.Count > 0)
                {
                    temp_Indate = dtable.Rows[0]["card_in_date"].ToString().Replace("-", "").Substring(4, 4);
                }
                _sw = new StreamWriter(fileName+ "sh" + temp_Indate + "_JK02.txt", true, _encoding);

                //업체코드 정렬
                _drs = dtable.Select("", "client_express_code");

                //헤더 부분
                _sw.WriteLine(GetStringAsLength("HDKJ" + DateTime.Now.ToString("yyyyMMdd"), 300, true, ' ')); //11은 공백

                if (_drs.Length > 0)
                {
                    for (i = 0; i < _drs.Length; i++)
                    {
                        //배송사원명을 5자로 지정
                        strCareer = _drs[i]["career"].ToString();

                        if (strCareer.Length > 6)
                        {
                            strCareer = strCareer.Substring(0, 5);
                        }

                        //시작코드
                        _strLine = new StringBuilder(GetStringAsLength("DT",2,true,' '));
                        //카드번호
                        _strLine.Append(GetStringAsLength(_drs[i]["card_number"].ToString(), 12, true, ' '));
                        //업체코드
                        _strLine.Append(GetStringAsLength(_drs[i]["client_express_code"].ToString(), 4, true, ' '));
                        //공백
                        _strLine.Append(GetStringAsLength("", 2, true, ' '));
                        //지사명
                        _strLine.Append(GetStringAsLength(_drs[i]["branch_name"].ToString(), 15, true, ' '));
                        //공백
                        _strLine.Append(GetStringAsLength("", 2, true, ' '));
                        //배송사원명
                        _strLine.Append(GetStringAsLength(strCareer, 10, true, ' '));
                        //발송일자
                        _strLine.Append(GetStringAsLength(_drs[i]["client_register_date"].ToString().Replace("-",""), 8, true, ' '));
                        //지사이관일자
                        _strLine.Append(GetStringAsLength(_drs[i]["branch_in_date"].ToString(), 8, true, ' '));
                        //카드처리상태
                        _strLine.Append(GetStringAsLength("", 3, true, ' '));
                        //기타
                        _strLine.Append(GetStringAsLength("", 4, true, ' '));
                        //filter
                        _strLine.Append(GetStringAsLength("", 230, true, ' '));

                        _sw.WriteLine(_strLine);
                    }
                    _strReturn = string.Format("{0}건의 인계데이터 다운 완료", i);
                }
                _sw.WriteLine(GetStringAsLength("TR" + GetStringAsLength(i.ToString(), 11, false, '0'), 300, true, ' '));
            }
            catch (Exception)
            {
                _strReturn = string.Format("{0}번째 데이터 인계 중 오류 발생", i);
            }
            finally
            {

                if (_sw != null) _sw.Close();
            }
            return _strReturn;
        }

        //일일마감자료
        public static string ConvertResultDay(System.Data.DataTable dtable, string fileName)
        {
            return ConvertResult(dtable, fileName);
           
        }

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
                _strReturn = value.Substring(0, 6) + "-" + value.Substring(6, value.Length - 6);
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
        private static string ConvertReturnStatus(string value)
        {
            string _strReturn = "";

            switch (value.Substring(0,2))
            {
                case "01":
                    _strReturn = "CC";
                    break;
                case "02":
                    _strReturn = "FF";
                    break;
                case "03":
                    _strReturn = "BB";
                    break;
                case "04":
                    _strReturn = "E1";
                    break;
                case "05":
                    _strReturn = "GG";
                    break;
                case "06":
                    _strReturn = "JJ";
                    break;
                case "07":
                    _strReturn = "DD";
                    break;
                case "08":
                    _strReturn = "B2";
                    break;
                default:
                    _strReturn = "ZZ";
                    break;

               
            }
            return _strReturn;
            
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
