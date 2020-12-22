using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Text;

namespace _003_신한_이미지
{
	public class CONVERT
	{
		//기본 인코딩 설정
		private static string strEncoding = "ks_c_5601-1987";
        private static string strCardTypeID = "003";
        //주의 : 신한_인수 문구 변경 시 배송프로그램 영향 받음
        private static string strCardTypeName = "신한_이미지";

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
            StreamWriter _sw00 = null, _sw01 = null;		//파일 쓰기 스트림            
            StringBuilder _strLine = new StringBuilder("");
            string _strStatus = "";
            string tempday = DateTime.Now.ToString("yyyyMMdd");
            string strTemp = "|";
            string _strReturn = "", strECC_Code = "", strDelivery_date = "", strCardTypeDetail = "", strStatus = "";
            int i = 0, itotcnt = 0;

            try
            {
                _sw01 = new StreamWriter(fileName + tempday + "_IMG_BPR_KUKJE", true, _encoding);

                _strLine = new StringBuilder(GetStringAsLength("HD", 2, true, ' '));
                _strLine.Append(GetStringAsLength(tempday, 8, true, ' '));
                _strLine.Append(GetStringAsLength("2002", 4, true, ' '));
                _strLine.Append(GetStringAsLength("", 284, true, ' '));

                _sw01.WriteLine(_strLine.ToString());

                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    strStatus = dtable.Rows[i]["card_delivery_status"].ToString();

                    if (strStatus == "1")
                    {
                        itotcnt++;

                        _strLine = new StringBuilder("DT");

                        strECC_Code = dtable.Rows[i]["card_cooperation1"].ToString();
                        strDelivery_date = RemoveDash(dtable.Rows[i]["card_delivery_date"].ToString());
                        strCardTypeDetail = dtable.Rows[i]["card_type_detail"].ToString();

                        _strLine.Append(GetStringAsLength(strDelivery_date, 8, true, ' '));
                        if (strCardTypeDetail.Substring(0, 4) == "0032")
                        {
                            _strLine.Append("48");
                        }
                        else
                        {
                            _strLine.Append("73");
                        }
                        _strLine.Append(GetStringAsLength(strECC_Code, 35, true, ' '));
                        _strLine.Append(GetStringAsLength("", 253, true, ' '));
                        _sw01.WriteLine(_strLine.ToString());
                    }
                }

                _strLine = new StringBuilder(GetStringAsLength("TR", 2, true, ' '));
                _strLine.Append(GetStringAsLength(itotcnt.ToString(), 10, true, ' '));
                _strLine.Append(GetStringAsLength("", 287, true, ' '));

                _sw01.WriteLine(_strLine.ToString());

                _strReturn = string.Format("{0}건의 이미지데이터 다운 완료", itotcnt);
            }
            catch (Exception)
            {
                _strReturn = string.Format("{0}번째 이미지데이터 생성 중 오류", itotcnt);
            }
            finally
            {
                if (_sw01 != null) _sw01.Close();
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
