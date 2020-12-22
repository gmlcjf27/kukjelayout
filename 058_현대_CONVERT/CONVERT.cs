using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace _058_현대_CONVERT
{
    public class CONVERT
    {
        //기본 인코딩 설정
        private static string strEncoding = "ks_c_5601-1987";
        private static char chCSV = ';';
        private static string strCardTypeID = "058_CONV";
        private static string strCardTypeName = "현대_CONVERT";

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
            StreamWriter _sw = null;

            string _strDataCode = "";       // 사용데이터 구분
            string _strCodeType = "";       // 공카드구분
            string _strCustomer_ssn = "";       // 기프트구분
            string _strOwner_only = "";       // 기프트구분

            string _strLine = "";
            string _strReturn = "";

            string[] _strAry = null;

            DataTable _dtable = null;
            DataSet _dsetZipcodeArea = null;

            try
            {
                _dtable = new DataTable("CONVERT");

                _dsetZipcodeArea = new DataSet();
                _dsetZipcodeArea.ReadXml(xmlZipcodeAreaPath);

                //파일 읽기 Stream과 오류시 저장할 쓰기 Stream 생성
                _sr = new StreamReader(path, _encoding);
                _swError = new StreamWriter(path + ".Error", false, _encoding);

                while ((_strLine = _sr.ReadLine()) != null)
                {
                    //CSV  분리
                    _strAry = _strLine.Split(chCSV);
                    //발송구분코드 : 36,46 = 블랙, 37 = 스마트일반, 45 = 동의서, 47 = 스마트동의
                    _strCodeType = _strAry[5].Trim();      //구분코드
                    _strCustomer_ssn = _strAry[10].Trim();
                    //일반 중 본인만 수령 가능 코드 : 00 = 본인만, 01 = 3자수령가능
                    _strOwner_only = _strAry[74].Trim();  

                    //2014.09.02 태희철 기프트카드 구분
                    if (_strCustomer_ssn.Trim() == "")
                    {
                        _sw = new StreamWriter(path + "일반_기프트", true, _encoding);
                        _sw.WriteLine(_strLine + ";0581102");
                    }
                    //2013.01.14 태희철 수정 01:클럽, 02:대한, 03:아시아, 04:하이플러스
                    else if (_strCodeType == "36")
                    {
                        _sw = new StreamWriter(path + "블랙카드", true, _encoding);
                        _sw.WriteLine(_strLine + ";0583102");
                    }
                    else if (_strCodeType == "46")
                    {
                        _sw = new StreamWriter(path + "블랙_동의", true, _encoding);
                        _sw.WriteLine(_strLine + ";0582107");
                    }
                    else if (_strCodeType == "56")
                    {
                        _sw = new StreamWriter(path + "블랙_약식", true, _encoding);
                        _sw.WriteLine(_strLine + ";0582108");
                    }
                    else if (_strCodeType == "55")
                    {
                        _sw = new StreamWriter(path + "현대약식동의", true, _encoding);
                        _sw.WriteLine(_strLine + ";0582103");
                    }
                    else if (_strCodeType == "37")
                    {
                        if (_strOwner_only.Trim() == "00" || _strOwner_only.Trim() == "10" || _strOwner_only.Trim() == "20")
                        {
                            _sw = new StreamWriter(path + "스마트일반_본인", true, _encoding);
                            _sw.WriteLine(_strLine + ";0583103");
                        }
                        else
                        {
                            _sw = new StreamWriter(path + "스마트일반", true, _encoding);
                            _sw.WriteLine(_strLine + ";0583101");
                        }
                    }
                    else if (_strCodeType == "47")
                    {
                        _sw = new StreamWriter(path + "스마트동의", true, _encoding);
                        _sw.WriteLine(_strLine + ";0582102");
                    }
                    else if (_strCodeType == "45")
                    {
                        _sw = new StreamWriter(path + "현대동의", true, _encoding);
                        _sw.WriteLine(_strLine + ";0582101");
                    }
                    else if (_strCodeType == "57")
                    {
                        _sw = new StreamWriter(path + "현대스마트약식동의", true, _encoding);
                        _sw.WriteLine(_strLine + ";0582104");
                    }
                    else if (_strCodeType == "30" && (_strOwner_only.Trim() == "00" || _strOwner_only.Trim() == "10" || _strOwner_only.Trim() == "20"))
                    {
                        _sw = new StreamWriter(path + "일반_본인만", true, _encoding);
                        _sw.WriteLine(_strLine + ";0581103");
                    }
                    else if (_strCodeType == "30")
                    {
                        _sw = new StreamWriter(path + "일반", true, _encoding);
                        _sw.WriteLine(_strLine + ";0581101");
                    }
                    else
                    {
                        _sw = new StreamWriter(path + "_그외_오류", true, _encoding);
                        _sw.WriteLine(_strLine);
                    }
                    
                    _sw.Close();
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
