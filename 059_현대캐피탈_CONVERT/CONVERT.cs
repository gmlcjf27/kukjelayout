﻿using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace _060_현대캐피탈2_CONVERT
{
    public class CONVERT
    {
        //기본 인코딩 설정
        private static string strEncoding = "ks_c_5601-1987";
        private static char chCSV = '|';
        private static string strCardTypeID = "060_CONV";
        private static string strCardTypeName = "현대캐피탈2_CONVERT";

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
            string strFileName = "";

            string _strLine = "";
            string _strReturn = "";

            string[] _strAry = null;

            DataTable _dtable = null;
            DataSet _dsetZipcodeArea = null;

            FileInfo f = new FileInfo(path);

            try
            {
                _dtable = new DataTable("CONVERT");

                _dsetZipcodeArea = new DataSet();
                _dsetZipcodeArea.ReadXml(xmlZipcodeAreaPath);


                //파일 읽기 Stream과 오류시 저장할 쓰기 Stream 생성
                _sr = new StreamReader(path, _encoding);
                _swError = new StreamWriter(path + ".Error", false, _encoding);

                //파일명으로 구분 : LOP = 크로스, RLOP = 신용R(등기), XLOP = SME(MB), AUTO = 오토, WLOP = 워크, OREN = 전세대출
                strFileName = f.Name.Substring(0,7);

                while ((_strLine = _sr.ReadLine()) != null)
                {
                    _strAry = _strLine.Split(chCSV);
                    //_strCodeType = _strAry[_strAry.LongLength -2];

                    switch (strFileName)
                    {
                        case "NR_B001":
                            _sw = new StreamWriter(path + "_일반", true, _encoding);
                            _sw.WriteLine(_strLine + ";0591101");
                            break;
                        case "NR_B002":
                            _sw = new StreamWriter(path + "_약정서론카드", true, _encoding);
                            _sw.WriteLine(_strLine + ";0591102");
                            break;
                        case "TM_B001":
                            _sw = new StreamWriter(path + "_동의서", true, _encoding);
                            _sw.WriteLine(_strLine + ";0592101");
                            break;
                        default:
                            _sw = new StreamWriter("파일명오류_" + path + "_파일명오류", true, _encoding);
                            _sw.WriteLine(_strLine);
                            break;
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
