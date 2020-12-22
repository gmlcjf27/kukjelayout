using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace BC_GIFT_NEW
{
	public class CONVERT
	{
		//기본 인코딩 설정
		private static string strEncoding = "ks_c_5601-1987";
		private static char chCSV = ',';
        private static string strCardTypeID = "043";
        private static string strCardTypeName = "BC별도";

		//현 DLL의 카드 타입 코드 반환
		public static string GetCardTypeID() {
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

		//등록 자료 생성
		//public static string ConvertRegister(string path, string xmlZipcodePath, string xmlZipcodeAreaPath, string xmlPath)
        public static string ConvertRegister(string path, string xmlZipcodePath, string xmlZipcodeAreaPath, string xmlZipcodePath_new, string xmlZipcodeAreaPath_new, string xmlPath)
		{
			System.Text.Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
			//FileInfo _fi = null;
			StreamReader _sr = null;													//파일 읽기 스트림
			StreamWriter _swError = null;												//파일 쓰기 스트림
			DataSet _dsetZipcode = null, _dsetZipcdeArea = null;						//우편번호 관련 DataSet
            DataSet _dsetZipcode_new = null, _dsetZipcdeArea_new = null;				//우편번호 관련 DataSet
			DataTable _dtable = null;													//마스터 저장 테이블
			DataRow _dr = null;
			DataRow[] _drs = null;
            string _strReturn = "";
			string _strLine = "";
			string[] _strAry = null;
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
				_dtable.Columns.Add("client_number");               // dr[5]
				_dtable.Columns.Add("card_product_name");
                _dtable.Columns.Add("customer_SSN");
				_dtable.Columns.Add("customer_name");
				_dtable.Columns.Add("card_mobile_tel");
				_dtable.Columns.Add("card_number");                 // dr[10]
                _dtable.Columns.Add("card_zipcode");                // dr[11]
				_dtable.Columns.Add("card_address");                

                _dtable.Columns.Add("card_zipcode_new");             // dr[13]
                _dtable.Columns.Add("card_zipcode_kind");            // dr[14]


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
                _sr = new StreamReader(path, _encoding);
                _swError = new StreamWriter(path + ".Error", false, _encoding);

                while ((_strLine = _sr.ReadLine()) != null)
                {
                    _dr = _dtable.NewRow();
                    _dr[0] = _iSeq;
                    //CSV 분리
                    _strAry = _strLine.Split(chCSV);
                    _dr[5] = _strAry[0];
                    _dr[6] = _strAry[1];
                    _dr[7] = _strAry[2] + _strAry[3];
                    //_dr[8] = _strAry[5] + " " + _strAry[6] + " " + _strAry[7];
                    _dr[8] = _strAry[4];
                    _dr[9] = _strAry[5] + "-" + _strAry[6] + "-" + _strAry[7];
					_dr[10] = _strAry[8];
                    _strZipcode = RemoveDash(_strAry[9].Replace(" ", "")).Trim();

                    if (_strZipcode.Trim().Length == 5)
                    {
                        _dr[13] = _strZipcode;
                        _dr[14] = "1";
                    }
                    _dr[11] = _strZipcode;
                    _dr[12] = _strAry[10];					

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
