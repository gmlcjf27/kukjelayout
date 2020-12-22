using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace _005_하나_CONVERT
{
    public class CONVERT
    {
        //기본 인코딩 설정
        private static string strEncoding = "ks_c_5601-1987";
        private static string strCardTypeID = "005_CONV";
        private static string strCardTypeName = "하나_컨버터";

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

        //등록 자료 생성
        //2013.01.18 태희철 수정
        public static string ConvertRegister(string path, string xmlZipcodeAreaPath, string xmlZipcodePath)
        {
            System.Text.Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            StreamReader _sr = null;																					//파일 읽기 스트림
            StreamWriter _swError = null;
            StreamWriter _sw = null, _swTest = null;

            string _strBranch = "";          // 동의여부
            string _strAgree_Type = "";     // 일반동의, TM동의구분
            // 별지동의서구분, 공카드 구분
            string _strCodeType = "", _strGong = "", strDong_test = "";  

            string _strLine = "";
            byte[] _byteAry = null;
            string _strReturn = "";
            DataTable _dtable = null;
            DataSet _dsetZipcodeArea = null;

            try
            {
                _dtable = new DataTable("CONVERT");
                _dtable.Columns.Add("data_type");
                _dtable.Columns.Add("data");

                _dsetZipcodeArea = new DataSet();
                _dsetZipcodeArea.ReadXml(xmlZipcodeAreaPath);

                //파일 읽기 Stream과 오류시 저장할 쓰기 Stream 생성
                _sr = new StreamReader(path, _encoding);
                _swError = new StreamWriter(path + ".Error", false, _encoding);

                while ((_strLine = _sr.ReadLine()) != null)
                {
                    _byteAry = _encoding.GetBytes(_strLine);
                    //2012-05-07 태희철 수정 코드 "C" 중 동의서 분리
                    
                    _strGong = _encoding.GetString(_byteAry, 0, 1);

                    if (_strGong.ToUpper() == "D")
                    {
                        _strCodeType = _encoding.GetString(_byteAry, 267, 2);
                    }
                    else if (_strGong.ToUpper() == "H" || _strGong.ToUpper() == "T")
                    {
                        ;
                    }
                    else
                    {
                        _strBranch = _encoding.GetString(_byteAry, 0, 2); //일반 : 11, 동의 : 12, 지점 : 15, 지점동의서 : 16 
                        _strCodeType = _encoding.GetString(_byteAry, 1006, 2);
                        _strAgree_Type = _encoding.GetString(_byteAry, 713, 1); //Y,N,O : 일반동의서, T : TM동의서
                        //동의여부 Y or A 의 경우 동의서, I = 일반본인지정
                        strDong_test = _encoding.GetString(_byteAry, 714, 1);
                    }

                    //일반본인지정 = I
                    if (strDong_test == "I")
                    {
                        _sw = new StreamWriter(path + "_일반본인지정_1000", true, _encoding);
                        _sw.WriteLine(_strLine + "0051102");
                    }
                    //첫번째 자리가 "D" 일 경우 공카드
                    else if (_strGong.ToUpper() == "D")
                    {
                        if (_strCodeType == "01")
                        {
                            _sw = new StreamWriter(path + "_공카드_3000", true, _encoding);
                            _sw.WriteLine(_strLine + "0054201");
                        }
                        else
                        {
                            _sw = new StreamWriter(path + "_공카드_2000", true, _encoding);
                            _sw.WriteLine(_strLine + "0054301");
                        }
                    }
                    //헤더, 테일을 제외
                    else if (_strGong.ToUpper() == "H" || _strGong.ToUpper() == "T")
                    {
                        ;
                    }
                    //일반,동의서,지점 구분
                    else if (_strBranch == "11")
                    {
                        _sw = new StreamWriter(path + "_일반_100", true, _encoding);
                        _sw.WriteLine(_strLine + "0051101");
                    }
                    else if (_strBranch == "15")
                    {
                        _sw = new StreamWriter(path + "_지점_1000", true, _encoding);
                        _sw.WriteLine(_strLine + "0054101");
                    }
                    else if (_strBranch == "16")
                    {
                        _sw = new StreamWriter(path + "_지점_31000", true, _encoding);
                        _sw.WriteLine(_strLine + "0054102");
                    }
                    else
                    {
                        if (_strAgree_Type == "T")
                        {
                            if (_strCodeType == "01")
                            {
                                _sw = new StreamWriter(path + "_DONG_6500_T_CLUBSK", true, _encoding);
                                _sw.WriteLine(_strLine + "0052107");
                            }
                            else if (_strCodeType == "02")
                            {
                                _sw = new StreamWriter(path + "_DONG_7500_T_대한", true, _encoding);
                                _sw.WriteLine(_strLine + "0052108");
                            }
                            else if (_strCodeType == "03")
                            {
                                _sw = new StreamWriter(path + "_DONG_8500_T_아시아", true, _encoding);
                                _sw.WriteLine(_strLine + "0052109");
                            }
                            else if (_strCodeType == "04")
                            {
                                _sw = new StreamWriter(path + "_DONG_9500_T_하이플러스", true, _encoding);
                                _sw.WriteLine(_strLine + "0052110");
                            }
                            else
                            {
                                _sw = new StreamWriter(path + "_DONG_5500_T_일반", true, _encoding);
                                _sw.WriteLine(_strLine + "0052106");
                            }
                        }
                        else if (_strAgree_Type == "B")
                        {
                            _sw = new StreamWriter(path + "_DONG_10500_법인_일반", true, _encoding);
                                _sw.WriteLine(_strLine + "0052111");
                        }
                        //일반동의서
                        else
                        {
                            if (_strCodeType == "01")
                            {
                                _sw = new StreamWriter(path + "_DONG_1500_CLUB_SK", true, _encoding);
                                _sw.WriteLine(_strLine + "0052102");
                            }
                            else if (_strCodeType == "02")
                            {
                                _sw = new StreamWriter(path + "_DONG_2500_대한", true, _encoding);
                                _sw.WriteLine(_strLine + "0052103");
                            }
                            else if (_strCodeType == "03")
                            {
                                _sw = new StreamWriter(path + "_DONG_3500_아시아", true, _encoding);
                                _sw.WriteLine(_strLine + "0052104");
                            }
                            else if (_strCodeType == "04")
                            {
                                _sw = new StreamWriter(path + "_DONG_4500_하이플러스", true, _encoding);
                                _sw.WriteLine(_strLine + "0052105");
                            }
                            //2014.08.29 태희철 추가
                            else if (_strCodeType == "05")
                            {
                                _sw = new StreamWriter(path + "_DONG_11500_하이패스", true, _encoding);
                                _sw.WriteLine(_strLine + "0052112");
                            }
                            //2017.11.24 태희철 추가
                            else if (_strCodeType == "06")
                            {
                                _sw = new StreamWriter(path + "_DONG_14500_GS팝", true, _encoding);
                                _sw.WriteLine(_strLine + "0052115");
                            }
                            //2016.07.14 태희철 추가
                            else if (_strCodeType == "07")
                            {
                                _sw = new StreamWriter(path + "_DONG_12500_아모레퍼시픽", true, _encoding);
                                _sw.WriteLine(_strLine + "0052113");
                            }
                            //2017.04.20 태희철 추가
                            else if (_strCodeType == "08")
                            {
                                _sw = new StreamWriter(path + "_DONG_13500_더페이스샵", true, _encoding);
                                _sw.WriteLine(_strLine + "0052114");
                            }
                            //2018.05.02 태희철 추가
                            else if (_strCodeType == "09")
                            {
                                _sw = new StreamWriter(path + "_DONG_15500_청바G", true, _encoding);
                                _sw.WriteLine(_strLine + "0052116");
                            }
                            //2018.08.16 태희철 추가
                            else if (_strCodeType == "10")
                            {
                                _sw = new StreamWriter(path + "_DONG_16500_E1LPG", true, _encoding);
                                _sw.WriteLine(_strLine + "0052117");
                            }
                            //2018.11.13 태희철 추가
                            else if (_strCodeType == "11")
                            {
                                _sw = new StreamWriter(path + "_DONG_17500_등유나눔", true, _encoding);
                                _sw.WriteLine(_strLine + "0052118");
                            }
                            else
                            {
                                _sw = new StreamWriter(path + "_DONG_500_일반", true, _encoding);
                                _sw.WriteLine(_strLine + "0052101");
                            }
                        }
                    }

                    if (strDong_test == "A" && _strBranch != "12")
                    {
                        _swTest = new StreamWriter(path + "오류_동의서_확인필요", true, _encoding);
                        _swTest.WriteLine(_strLine + "0052101");
                    }

                    //_sw.WriteLine(_strLine);
                    if (_sw != null)
                    {
                        _sw.Close();    
                    }

                    if (_swTest != null) _swTest.Close();
                }
                _sr.Close();
                _strReturn = "성공";
            }
            catch (Exception ex)
            {
                _strReturn = "실패";
                if (_swError != null) _swError.WriteLine(ex.Message);
            }
            finally
            {
                if (_sr != null) _sr.Close();
                if (_sw != null) _sw.Close();
                if (_swTest != null) _swTest.Close();
                if (_swError != null) _swError.Close();
            }
            return _strReturn;
        }

        private static StreamWriter GetStreamWriter(string path)
        {
            StreamWriter _return = null;
            FileInfo _fi = new FileInfo(path);
            if (_fi.Exists)
            {
                _return = new StreamWriter(path, false, System.Text.Encoding.GetEncoding(949));
            }
            else
            {
                _return = new StreamWriter(path, true, System.Text.Encoding.GetEncoding(949));
            }
            return _return;
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
    }
}
