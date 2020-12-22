using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace _001_BC_BAN
{
	public class CONVERT
	{
		//기본 인코딩 설정
		private static string strEncoding = "ks_c_5601-1987";
		private static char chCSV = ',';
        private static string strCardTypeID = "101";
        private static string strCardTypeName = "비씨(은)";

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
            string _strReturn = "0014101";

            return _strReturn;
        }

		//등록 자료 생성
		//public static string ConvertRegister(string path, string xmlZipcodePath, string xmlZipcodeAreaPath, string xmlBankPath, string xmlPath) 
        public static string ConvertRegister(string path, string xmlZipcodePath, string xmlZipcodeAreaPath, string xmlZipcodePath_new, string xmlZipcodeAreaPath_new, string xmlBankPath, string xmlPath)
        {
			Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
			//FileInfo _fi = null;
			StreamReader _sr = null;																					//파일 읽기 스트림			
			StreamWriter _swError = null;																			//파일 쓰기 스트림
			DataSet _dsetZipcode = null, _dsetZipcdeArea = null, _dsetBank = null;
            DataSet _dsetZipcode_new = null, _dsetZipcdeArea_new = null;
            DataTable _dtable = null;																					//마스터 저장 테이블
			DataRow _dr = null;
			DataRow[] _drs = null;
            string _strReturn = "";
			string _strLine = "";
			string[] _strAry = null;
			string _strBankID = "", _strZipcode = "", _strAddress = "", _strAreaType = "", _strAreaGroup = "", _strBranch = "", _strBankToBankID = "";			
			int _iSeq = 1, _iErrorCount = 0;
            Branches _branches = new Branches();

            //2019.04.25 레이아웃 변경
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

                _dtable.Columns.Add("client_bank_request_no");              //dr[5]				
                _dtable.Columns.Add("client_send_date");                    
                _dtable.Columns.Add("client_send_number");
				_dtable.Columns.Add("customer_name");
                _dtable.Columns.Add("card_bank_ID");
                _dtable.Columns.Add("delivery_return_reason_last");         //dr[10]

                _dtable.Columns.Add("card_zipcode");                        //dr[11]
                _dtable.Columns.Add("card_address");
                _dtable.Columns.Add("customer_real_SSN");
                _dtable.Columns.Add("card_zipcode_new");                    //dr[14]
                _dtable.Columns.Add("card_zipcode_kind");                   //dr[15]
								
				//실제로는 은행 영업점 테이블
				
				_dsetBank = new DataSet();
                _dsetZipcode = new DataSet();
                _dsetZipcdeArea = new DataSet();
                _dsetZipcode.ReadXml(xmlZipcodePath);
                _dsetZipcode.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcode.Tables[0].Columns["zipcode"] };
                _dsetZipcdeArea.ReadXml(xmlZipcodeAreaPath);
                _dsetZipcdeArea.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcdeArea.Tables[0].Columns["zipcode"] };
                _dsetBank.ReadXml(xmlBankPath);

                _dsetZipcode_new = new DataSet();
                _dsetZipcdeArea_new = new DataSet();
                _dsetZipcode_new.ReadXml(xmlZipcodePath_new);
                _dsetZipcode_new.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcode_new.Tables[0].Columns["zipcode_new"] };
                _dsetZipcdeArea_new.ReadXml(xmlZipcodeAreaPath_new);
                _dsetZipcdeArea_new.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcdeArea_new.Tables[0].Columns["zipcode_new"] };

				//파일 읽기 Stream과 오류시 저장할 쓰기 Stream 생성
				_sr = new System.IO.StreamReader(path, _encoding);
				_swError = new System.IO.StreamWriter(path + ".Error", false, _encoding);

				while ((_strLine = _sr.ReadLine()) != null) {
					_dr = _dtable.NewRow();
					_dr[0] = _iSeq;
					_dr[1] = "012";
					_dr[2] = "012";
					_dr[3] = "A";
					_dr[4] = _iSeq;
					//CSV 분리
					_strAry = _strLine.Split(chCSV);
                    //_strAry[0];
                    //_strAry[1];
					_dr[5] = _strAry[2];
                    _dr[6] = _strAry[3];
                    _dr[7] = _strAry[4];
                    _dr[8] = _strAry[5];
					_strBankID = _strAry[6];
					_dr[9] = _strBankID;
                    //_strAry[7];
					_dr[10] = _strAry[8];

					if (_strBankID != "")
					{
						//지역 분류 선택
						_drs = _dsetBank.Tables[0].Select("bank_ID='" + _strBankID + "'");

						if (_drs.Length < 1) 
                        {
                            //_strZipcode = "";
                            //_strAddress = "";
                            _iErrorCount++;
                            _swError.WriteLine("1지점번호 오류 : " + _strLine);
						}
						else 
                        {
							_strZipcode = _drs[0]["bank_zipcode"].ToString();
							_strAddress = _drs[0]["bank_address"].ToString();
							_strBankToBankID = _drs[0]["bank_to_bank_ID"].ToString();
						}

						if (_strBankID != _strBankToBankID) 
                        {
							_drs = _dsetBank.Tables[0].Select("bank_ID='" + _strBankToBankID + "'");

							if (_drs.Length < 1) {
                                //_strZipcode = "";
                                //_strAddress = "";
                                _iErrorCount++;
                                _swError.WriteLine("2지점번호 오류 : " + _strLine);
							}
							else {
								_strZipcode = _drs[0]["bank_zipcode"].ToString();
								_strAddress = _drs[0]["bank_address"].ToString();								
							}
						}
                        
						_dr[11] = _strZipcode;

                        if (_strZipcode.Trim().Length == 5)
                        {
                            _dr[14] = _strZipcode.Trim();
                            _dr[15] = "1";
                        }

						_dr[12] = _strAddress;

						if (_strZipcode != "") 
                        {
							//지역 분류 선택

                            if (_strZipcode.Trim().Length == 5)
                            {
                                _drs = _dsetZipcdeArea_new.Tables[0].Select("zipcode_new = " + _strZipcode.Trim());    
                            }
                            else
                            {
                                _drs = _dsetZipcdeArea.Tables[0].Select("zipcode=" + _strZipcode.Trim());
                            }
							
							if (_drs.Length < 1) {
								_strAreaGroup = "012";
								_strBranch = "012";
							}
							else {
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
                                _drs = _dsetZipcode.Tables[0].Select("zipcode=" + _strZipcode.Trim());
                            }
							
							if (_drs.Length > 0) {
								_strAreaType = _drs[0]["area_type_code"].ToString();
							}
							else {
								_strAreaType = "";
							}

							//지점과 배송지점이 다르다면 본사로 등록시킨다.
							if (_strBankToBankID != _strBankID) {
								if (_strAreaType == "") {
									_strBranch = "000";
									_strAreaGroup = "000";
									_strAreaType = "A";
								}

								_dr["card_bank_ID"] = _strBankToBankID;
								_dr["customer_real_SSN"] = _strBankID;
							}
							else {
								_dr["customer_real_SSN"] = _strBankID;
							}

							//우편번호를 찾지 못 했다면 Error 파일에 기록
							if (_strAreaType.Equals("")) 
                            {
								_swError.WriteLine("우편번호 오류" + _strLine);
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
							_swError.WriteLine("우편번호오류 : " + _strLine);
							_iErrorCount++;
						}

                        //2012.08.14 태희철 수정[S]
                        //_strBankToBankID 를 초기화 하여 지점등록이 되어있지 않은건은 Error 처리
                        _strBankToBankID = "";
						_iSeq++;
					}
					else {
                        _strBankToBankID = "";
                        //2012.08.14 태희철 수정[E]
						_swError.WriteLine("은행지점번호 오류" + _strLine);
                        _iErrorCount++;
					}
				}

				//변환에 성공했다면
				if (_iErrorCount < 1) {					
					_sr.Close();					
					_dtable.WriteXml(xmlPath);					
                    _strReturn = string.Format("{0}건의 데이터 변환 성공", _iSeq-1);
				}
                else{
					_strReturn = string.Format("{0}건 변환, 영업점 조회 {1}건 실패", _iSeq - _iErrorCount - 1, _iErrorCount);
                }
			}
			catch (Exception ex) {
				if (_swError != null) {
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

        //마감 자료
        public static string ConvertResult(System.Data.DataTable dtable, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            StreamWriter _sw00 = null, _sw01 = null, _sw02 = null;																			//파일 쓰기 스트림
            int i = 0;
            string _strLine = "";
            string _strReturn = "", _strDeliveryStatus = "";
            try
            {
                _sw00 = new StreamWriter(fileName + ".00", true, _encoding);
                _sw01 = new StreamWriter(fileName + ".01", true, _encoding);
                _sw02 = new StreamWriter(fileName + ".02", true, _encoding);
                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    _strDeliveryStatus = dtable.Rows[i]["card_delivery_status"].ToString();

                    _strLine = GetStringAsLength(dtable.Rows[i]["client_number"].ToString(), 5, true, ' ');
                    _strLine += GetStringAsLength(dtable.Rows[i][""].ToString(),16,true,' ');
                    _strLine += GetStringAsLength(dtable.Rows[i][""].ToString(),7,true,' ');
                    _strLine += GetStringAsLength(dtable.Rows[i][""].ToString(),30,true,' ');
                    _strLine += GetStringAsLength(dtable.Rows[i][""].ToString(),30,true,' ');
                    _strLine += GetStringAsLength(dtable.Rows[i][""].ToString(),6,true,' ');
                    _strLine += GetStringAsLength(dtable.Rows[i][""].ToString(),1,true,' ');
                    _strLine += GetStringAsLength(dtable.Rows[i][""].ToString(),12,true,' ');
                    _strLine += GetStringAsLength(dtable.Rows[i][""].ToString(),13,true,' ');
                    _strLine += GetStringAsLength(dtable.Rows[i][""].ToString(),6,true,' ');
                    _strLine += GetStringAsLength(dtable.Rows[i][""].ToString(),110,true,' ');
                    _strLine += GetStringAsLength(dtable.Rows[i][""].ToString(),13,true,' ');
                    _strLine += GetStringAsLength(dtable.Rows[i][""].ToString(),1,true,' ');
                    _strLine += GetStringAsLength(dtable.Rows[i][""].ToString(),8,true,' ');
                   


                    if (_strDeliveryStatus == "1" || _strDeliveryStatus == "7")
                    {
                        _sw01.WriteLine(_strLine);
                    }
                    else if (_strDeliveryStatus == "2" || _strDeliveryStatus == "3")
                    {
                        _sw02.WriteLine(_strLine);
                    }
                    else if (_strDeliveryStatus == "4" || _strDeliveryStatus == "6")
                    {
                    }
                    else
                    {
                        _sw00.WriteLine(_strLine);
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


		private static string GetBCReturnDate(string value)
		{
			string _return = value;
			if (value.Length == 6)
			{
				_return += "01";
			}
			return _return;
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
