using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace _147BS렌탈_CONVERT
{
    public class CONVERT
    {
        //기본 인코딩 설정
        private static string strEncoding = "ks_c_5601-1987";
        private static char chCSV = ',';
        private static string strCardTypeID = "147_CONV";
        private static string strCardTypeName = "BS렌탈_CONVERT";

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

            string _strCodeType = "";       // 데이터구분


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

                    // 1=비에스렌탈, 2=비에스렌탈기타, 3=KB케피탈, 4=CJ헬로
                    _strCodeType = _strAry[21].Trim();      //구분코드

                    if (_strCodeType.Trim() == "1" || _strCodeType.Trim() == "7" || _strCodeType.Trim() == "9" || _strCodeType.Trim() == "11")
                    {
                        _sw = new StreamWriter(path + "비에스렌탈_500_비에스렌탈", true, _encoding);
                        _sw.WriteLine(_strLine + ",1473101");
                    }
                    else if (_strCodeType.Trim() == "2")
                    {
                        _sw = new StreamWriter(path + "비에스렌탈_1500_렌탈기타", true, _encoding);
                        _sw.WriteLine(_strLine + ",1473102");
                    }
                    else if (_strCodeType.Trim() == "3" || _strCodeType.Trim() == "8" || _strCodeType.Trim() == "10" || _strCodeType.Trim() == "12")
                    {
                        _sw = new StreamWriter(path + "비에스렌탈_3500_LNJ렌탈", true, _encoding);
                        _sw.WriteLine(_strLine + ",1473104");
                    }
                    else if (_strCodeType.Trim() == "4")
                    {
                        _sw = new StreamWriter(path + "비에스렌탈_2500_CJ헬로", true, _encoding);
                        _sw.WriteLine(_strLine + ",1473103");
                    }
                    else if (_strCodeType.Trim() == "5")
                    {
                        _sw = new StreamWriter(path + "비에스렌탈_4500_Tbroad", true, _encoding);
                        _sw.WriteLine(_strLine + ",1473105");
                    }
                    else if (_strCodeType.Trim() == "6")
                    {
                        _sw = new StreamWriter(path + "비에스렌탈_5500_현대HCN", true, _encoding);
                        _sw.WriteLine(_strLine + ",1473106");
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
