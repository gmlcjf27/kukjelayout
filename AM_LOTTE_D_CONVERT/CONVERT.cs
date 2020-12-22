using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Data;

namespace AM_LOTTE_D_CONVERT
{

    public class CONVERT
    {
        //기본 인코딩 설정
        private static string strEncoding = "ks_c_5601-1987";
        private static string strCardTypeID = "016_CONV";
        private static string strCardTypeName = "오전_롯데컨버터";


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
        public static string ConvertRegister(string path, string xmlZipcodeAreaPath, string xmlZipcodePath)
        {
            System.Text.Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            StreamReader _sr = null;																					//파일 읽기 스트림
            StreamWriter _swError = null;
            StreamWriter _sw = null;
            string _strLine = "", _strCode = "", _strCode2 = "", _strCode3 = "", _strCode4 = "", _strPlace = null, strMember = "", strOwner = "";
            string strCard_zipcode = "";
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
                    //2012.10.19 태희철 수정[S] 동의서 분리 방법 수정
                    //_strCode = _encoding.GetString(_byteAry, 1009, 5);
                    
                    _strCode = _encoding.GetString(_byteAry, 1045, 2);      //동의서식별코드 1->2byte, 09 : 특정제휴추가
                    _strCode2 = _encoding.GetString(_byteAry, 694, 1);      // 일반/동의 구분코드 Y : 일반, N : 동의서

                    ///카드배송유형코드
                    ///001 ~ 011 코드 : 일반 ~ 기타
                    /// 
                    _strCode3 = _encoding.GetString(_byteAry, 1048, 3);
                    _strCode4 = _encoding.GetString(_byteAry, 1044, 1);
                    //멤버스 유무 : 1 = 유, 0 = 무
                    strMember = _encoding.GetString(_byteAry, 1166, 1);
                    //본인만 유무 : 11 = 본인만
                    strOwner = _encoding.GetString(_byteAry, 1167, 2);

                    //if (strMember == "0" || strMember == "1")
                    //{
                    //    strMember == "1";
                    //}

                    //2011-11-22 태희철 수령지 지역 서울 / 지방 구분
                    // 수령지가 자택 / 직장 구분
                    //자택이 아니면 청구지를 직장으로 취급한다. 롯데카드 김지연 과장
                    if (_encoding.GetString(_byteAry, 698, 3) == "120")
                    {
                        _strPlace = _encoding.GetString(_byteAry, 147, 1);
                        strCard_zipcode = _encoding.GetString(_byteAry, 147, 6);
                    }
                    else
                    {
                        _strPlace = _encoding.GetString(_byteAry, 378, 1);
                        strCard_zipcode = _encoding.GetString(_byteAry, 378, 6);
                    }
                                        
                    switch (_strCode3.ToString())
                    {
                        // 패키지
                        case "002":
                            if (_strCode2.ToLower() == "y")
                            {
                                if (strMember == "1" || strMember == "2")
                                {
                                    if (strOwner == "11")
                                    {
                                        _sw = new StreamWriter(path + ".K_긴급_M_패키지_본인", true, _encoding);
                                        _sw.WriteLine(_strLine + "0863103");
                                        _sw.Close();
                                    }
                                    else
                                    {
                                        _sw = new StreamWriter(path + ".K_긴급_M_패키지", true, _encoding);
                                        _sw.WriteLine(_strLine + "0863103");
                                        _sw.Close();
                                    }
                                }
                                else
                                {
                                    if (strOwner == "11")
                                    {
                                        _sw = new StreamWriter(path + ".K_긴급_패키지_본인", true, _encoding);
                                        _sw.WriteLine(_strLine + "0863101");
                                        _sw.Close();
                                    }
                                    else
                                    {
                                        _sw = new StreamWriter(path + ".K_긴급_패키지", true, _encoding);
                                        _sw.WriteLine(_strLine + "0863101");
                                        _sw.Close();
                                    }
                                }
                            }
                            else
                            {
                                if (strMember == "1")
                                {
                                    switch (_strCode)
                                    {
                                        // [긴급-SK엔크린]
                                        case "10": _sw = new StreamWriter(path + ".M(별지)_패키지_긴급_동_SK엔크린_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862303");
                                            _sw.Close();
                                            break;
                                        // [긴급-웅진]
                                        case "11": _sw = new StreamWriter(path + ".M(별지)_패키지_긴급_동_웅진_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862303");
                                            _sw.Close();
                                            break;
                                        // [긴급-OK캐쉬백]
                                        case "12": _sw = new StreamWriter(path + ".M(별지)_패키지_긴급_동_OK캐쉬백_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862303");
                                            _sw.Close();
                                            break;
                                        case "13": _sw = new StreamWriter(path + ".M(별지)_패키지_긴급_동_아모레_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862303");
                                            _sw.Close();
                                            break;
                                        case "14": _sw = new StreamWriter(path + ".M(별지)_패키지_긴급_동_오포인트_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862303");
                                            _sw.Close();
                                            break;
                                        case "15": _sw = new StreamWriter(path + ".M(별지)_패키지_긴급_동_해피포인트_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862303");
                                            _sw.Close();
                                            break;
                                        case "16": _sw = new StreamWriter(path + ".M(별지)_패키지_긴급_동_교보문고_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862303");
                                            _sw.Close();
                                            break;
                                        case "17": _sw = new StreamWriter(path + ".M(별지)_패키지_긴급_동_뉴SOIL_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862303");
                                            _sw.Close();
                                            break;
                                        case "18": _sw = new StreamWriter(path + ".M(별지)_패키지_긴급_동_T멤버십_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862303");
                                            _sw.Close();
                                            break;
                                        case "19": _sw = new StreamWriter(path + ".M(별지)_패키지_긴급_동_아시아나_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862303");
                                            _sw.Close();
                                            break;
                                        case "20": _sw = new StreamWriter(path + ".M(별지)_패키지_긴급_동_대한항공_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862303");
                                            _sw.Close();
                                            break;
                                        case "21": _sw = new StreamWriter(path + ".M(별지)_패키지_긴급_동_E1LPG_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862303");
                                            _sw.Close();
                                            break;
                                        case "22": _sw = new StreamWriter(path + ".M(별지)_패키지_긴급_동_아이행복_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862303");
                                            _sw.Close();
                                            break;
                                        case "23": _sw = new StreamWriter(path + ".M(별지)_패키지_긴급_동_국민행복_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862303");
                                            _sw.Close();
                                            break;
                                        default: _sw = new StreamWriter(path + ".M(별지)_패키지_긴급_동_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862303");
                                            _sw.Close();
                                            break;
                                    }
                                }
                                //멤버스 별지 미징구
                                else if (strMember == "2")
                                {
                                    switch (_strCode)
                                    {
                                        // [긴급-SK엔크린]
                                        case "10": _sw = new StreamWriter(path + ".M_패키지_긴급_동_SK엔크린_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862309");
                                            _sw.Close();
                                            break;
                                        // [긴급-웅진]
                                        case "11": _sw = new StreamWriter(path + ".M_패키지_긴급_동_웅진_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862309");
                                            _sw.Close();
                                            break;
                                        // [긴급-OK캐쉬백]
                                        case "12": _sw = new StreamWriter(path + ".M_패키지_긴급_동_OK캐쉬백_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862309");
                                            _sw.Close();
                                            break;
                                        case "13": _sw = new StreamWriter(path + ".M_패키지_긴급_동_아모레_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862309");
                                            _sw.Close();
                                            break;
                                        case "14": _sw = new StreamWriter(path + ".M_패키지_긴급_동_오포인트_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862309");
                                            _sw.Close();
                                            break;
                                        case "15": _sw = new StreamWriter(path + ".M_패키지_긴급_동_해피포인트_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862309");
                                            _sw.Close();
                                            break;
                                        case "16": _sw = new StreamWriter(path + ".M_패키지_긴급_동_교보문고_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862309");
                                            _sw.Close();
                                            break;
                                        case "17": _sw = new StreamWriter(path + ".M_패키지_긴급_동_뉴SOIL_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862309");
                                            _sw.Close();
                                            break;
                                        case "18": _sw = new StreamWriter(path + ".M_패키지_긴급_동_T멤버십_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862309");
                                            _sw.Close();
                                            break;
                                        case "19": _sw = new StreamWriter(path + ".M_패키지_긴급_동_아시아나_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862309");
                                            _sw.Close();
                                            break;
                                        case "20": _sw = new StreamWriter(path + ".M_패키지_긴급_동_대한항공_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862309");
                                            _sw.Close();
                                            break;
                                        case "21": _sw = new StreamWriter(path + ".M_패키지_긴급_동_E1LPG_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862309");
                                            _sw.Close();
                                            break;
                                        case "22": _sw = new StreamWriter(path + ".M_패키지_긴급_동_아이행복_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862309");
                                            _sw.Close();
                                            break;
                                        case "23": _sw = new StreamWriter(path + ".M_패키지_긴급_동_국민행복_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862309");
                                            _sw.Close();
                                            break;
                                        default: _sw = new StreamWriter(path + ".M_패키지_긴급_동_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862309");
                                            _sw.Close();
                                            break;
                                    }
                                }
                                else
                                {
                                    switch (_strCode)
                                    {
                                        // [긴급-SK엔크린]
                                        case "10": _sw = new StreamWriter(path + ".패키지_긴급_동_SK엔크린_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862301");
                                            _sw.Close();
                                            break;
                                        // [긴급-웅진]
                                        case "11": _sw = new StreamWriter(path + ".패키지_긴급_동_웅진_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862301");
                                            _sw.Close();
                                            break;
                                        // [긴급-OK캐쉬백]
                                        case "12": _sw = new StreamWriter(path + ".패키지_긴급_동_OK캐쉬백_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862301");
                                            _sw.Close();
                                            break;
                                        case "13": _sw = new StreamWriter(path + ".패키지_긴급_동_아모레_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862301");
                                            _sw.Close();
                                            break;
                                        case "14": _sw = new StreamWriter(path + ".패키지_긴급_동_오포인트_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862301");
                                            _sw.Close();
                                            break;
                                        case "15": _sw = new StreamWriter(path + ".패키지_긴급_동_해피포인트_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862301");
                                            _sw.Close();
                                            break;
                                        case "16": _sw = new StreamWriter(path + ".패키지_긴급_동_교보문고_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862301");
                                            _sw.Close();
                                            break;
                                        case "17": _sw = new StreamWriter(path + ".패키지_긴급_동_뉴SOIL_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862301");
                                            _sw.Close();
                                            break;
                                        case "18": _sw = new StreamWriter(path + ".패키지_긴급_동_T멤버십_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862301");
                                            _sw.Close();
                                            break;
                                        case "19": _sw = new StreamWriter(path + ".패키지_긴급_동_아시아나_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862301");
                                            _sw.Close();
                                            break;
                                        case "20": _sw = new StreamWriter(path + ".패키지_긴급_동_대한항공_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862301");
                                            _sw.Close();
                                            break;
                                        case "21": _sw = new StreamWriter(path + ".패키지_긴급_동_E1LPG_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862301");
                                            _sw.Close();
                                            break;
                                        case "22": _sw = new StreamWriter(path + ".패키지_긴급_동_아이행복_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862301");
                                            _sw.Close();
                                            break;
                                        case "23": _sw = new StreamWriter(path + ".패키지_긴급_동_국민행복_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862301");
                                            _sw.Close();
                                            break;
                                        default: _sw = new StreamWriter(path + ".패키지_긴급_동_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862301");
                                            _sw.Close();
                                            break;
                                    }
                                }
                            }
                            break;
                        // 122롯데-긴급
                        case "003":
                            //긴급일반
                            if (_strCode2.ToLower() == "y")
                            {
                                if (strMember == "1" || strMember == "2")
                                {
                                    if (strOwner == "11")
                                    {
                                        _sw = new StreamWriter(path + ".K_M_긴급_오전_본인", true, _encoding);
                                        _sw.WriteLine(_strLine + "0863111");
                                        _sw.Close();
                                    }
                                    else
                                    {
                                        _sw = new StreamWriter(path + ".K_M_긴급_오전", true, _encoding);
                                        _sw.WriteLine(_strLine + "0863103");
                                        _sw.Close();
                                    }
                                }
                                else
                                {
                                    if (strOwner == "11")
                                    {
                                        _sw = new StreamWriter(path + ".K_긴급_오전_본인", true, _encoding);
                                        _sw.WriteLine(_strLine + "0863109");
                                        _sw.Close();
                                    }
                                    else
                                    {
                                        _sw = new StreamWriter(path + ".K_긴급_오전", true, _encoding);
                                        _sw.WriteLine(_strLine + "0863101");
                                        _sw.Close();
                                    }
                                }
                            }
                            //긴급동의
                            else
                            {
                                if (strMember == "1")
                                {
                                    switch (_strCode)
                                    {
                                        // [긴급-SK엔크린]
                                        case "10": _sw = new StreamWriter(path + ".M(별지)_긴급_동_SK엔크린_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862303");
                                            _sw.Close();
                                            break;
                                        // [긴급-웅진]
                                        case "11": _sw = new StreamWriter(path + ".M(별지)_긴급_동_웅진_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862303");
                                            _sw.Close();
                                            break;
                                        // [긴급-OK캐쉬백]
                                        case "12": _sw = new StreamWriter(path + ".M(별지)_긴급_동_OK캐쉬백_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862303");
                                            _sw.Close();
                                            break;
                                        case "13": _sw = new StreamWriter(path + ".M(별지)_긴급_동_아모레_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862303");
                                            _sw.Close();
                                            break;
                                        case "14": _sw = new StreamWriter(path + ".M(별지)_긴급_동_오포인트_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862303");
                                            _sw.Close();
                                            break;
                                        case "15": _sw = new StreamWriter(path + ".M(별지)_긴급_동_해피포인트_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862303");
                                            _sw.Close();
                                            break;
                                        case "16": _sw = new StreamWriter(path + ".M(별지)_긴급_동_교보문고_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862303");
                                            _sw.Close();
                                            break;
                                        case "17": _sw = new StreamWriter(path + ".M(별지)_긴급_동_뉴SOIL_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862303");
                                            _sw.Close();
                                            break;
                                        case "18": _sw = new StreamWriter(path + ".M(별지)_긴급_동_T멤버십_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862303");
                                            _sw.Close();
                                            break;
                                        case "19": _sw = new StreamWriter(path + ".M(별지)_긴급_동_아시아나_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862303");
                                            _sw.Close();
                                            break;
                                        case "20": _sw = new StreamWriter(path + ".M(별지)_긴급_동_대한항공_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862303");
                                            _sw.Close();
                                            break;
                                        case "21": _sw = new StreamWriter(path + ".M(별지)_긴급_동_E1LPG_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862303");
                                            _sw.Close();
                                            break;
                                        case "22": _sw = new StreamWriter(path + ".M(별지)_긴급_동_아이행복_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862303");
                                            _sw.Close();
                                            break;
                                        case "23": _sw = new StreamWriter(path + ".M(별지)_긴급_동_국민행복_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862303");
                                            _sw.Close();
                                            break;
                                        default: _sw = new StreamWriter(path + ".M(별지)_긴급_동_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862303");
                                            _sw.Close();
                                            break;
                                    }  
                                }
                                else if (strMember == "2")
                                {
                                    switch (_strCode)
                                    {
                                        // [긴급-SK엔크린]
                                        case "10": _sw = new StreamWriter(path + ".M_긴급_동_SK엔크린_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862309");
                                            _sw.Close();
                                            break;
                                        // [긴급-웅진]
                                        case "11": _sw = new StreamWriter(path + ".M_긴급_동_웅진_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862309");
                                            _sw.Close();
                                            break;
                                        // [긴급-OK캐쉬백]
                                        case "12": _sw = new StreamWriter(path + ".M_긴급_동_OK캐쉬백_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862309");
                                            _sw.Close();
                                            break;
                                        case "13": _sw = new StreamWriter(path + ".M_긴급_동_아모레_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862309");
                                            _sw.Close();
                                            break;
                                        case "14": _sw = new StreamWriter(path + ".M_긴급_동_오포인트_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862309");
                                            _sw.Close();
                                            break;
                                        case "15": _sw = new StreamWriter(path + ".M_긴급_동_해피포인트_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862309");
                                            _sw.Close();
                                            break;
                                        case "16": _sw = new StreamWriter(path + ".M_긴급_동_교보문고_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862309");
                                            _sw.Close();
                                            break;
                                        case "17": _sw = new StreamWriter(path + ".M_긴급_동_뉴SOIL_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862309");
                                            _sw.Close();
                                            break;
                                        case "18": _sw = new StreamWriter(path + ".M_긴급_동_T멤버십_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862309");
                                            _sw.Close();
                                            break;
                                        case "19": _sw = new StreamWriter(path + ".M_긴급_동_아시아나_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862309");
                                            _sw.Close();
                                            break;
                                        case "20": _sw = new StreamWriter(path + ".M_긴급_동_대한항공_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862309");
                                            _sw.Close();
                                            break;
                                        case "21": _sw = new StreamWriter(path + ".M_긴급_동_E1LPG_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862309");
                                            _sw.Close();
                                            break;
                                        case "22": _sw = new StreamWriter(path + ".M_긴급_동_아이행복_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862309");
                                            _sw.Close();
                                            break;
                                        case "23": _sw = new StreamWriter(path + ".M_긴급_동_국민행복_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862309");
                                            _sw.Close();
                                            break;
                                        default: _sw = new StreamWriter(path + ".M_긴급_동_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862309");
                                            _sw.Close();
                                            break;
                                    }
                                }
                                else
                                {
                                    switch (_strCode)
                                    {
                                        // [긴급-SK엔크린]
                                        case "10": _sw = new StreamWriter(path + ".긴급_동_SK엔크린_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862301");
                                            _sw.Close();
                                            break;
                                        // [긴급-웅진]
                                        case "11": _sw = new StreamWriter(path + ".긴급_동_웅진_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862301");
                                            _sw.Close();
                                            break;
                                        // [긴급-OK캐쉬백]
                                        case "12": _sw = new StreamWriter(path + ".긴급_동_OK캐쉬백_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862301");
                                            _sw.Close();
                                            break;
                                        case "13": _sw = new StreamWriter(path + ".긴급_동_아모레_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862301");
                                            _sw.Close();
                                            break;
                                        case "14": _sw = new StreamWriter(path + ".긴급_동_오포인트_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862301");
                                            _sw.Close();
                                            break;
                                        case "15": _sw = new StreamWriter(path + ".긴급_동_해피포인트_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862301");
                                            _sw.Close();
                                            break;
                                        case "16": _sw = new StreamWriter(path + ".긴급_동_교보문고_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862301");
                                            _sw.Close();
                                            break;
                                        case "17": _sw = new StreamWriter(path + ".긴급_동_뉴SOIL_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862301");
                                            _sw.Close();
                                            break;
                                        case "18": _sw = new StreamWriter(path + ".긴급_동_T멤버십_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862301");
                                            _sw.Close();
                                            break;
                                        case "19": _sw = new StreamWriter(path + ".긴급_동_아시아나_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862301");
                                            _sw.Close();
                                            break;
                                        case "20": _sw = new StreamWriter(path + ".긴급_동_대한항공_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862301");
                                            _sw.Close();
                                            break;
                                        case "21": _sw = new StreamWriter(path + ".긴급_동_E1LPG_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862301");
                                            _sw.Close();
                                            break;
                                        case "22": _sw = new StreamWriter(path + ".긴급_동_아이행복_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862301");
                                            _sw.Close();
                                            break;
                                        case "23": _sw = new StreamWriter(path + ".긴급_동_국민행복_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862301");
                                            _sw.Close();
                                            break;
                                        default: _sw = new StreamWriter(path + ".긴급_동_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862301");
                                            _sw.Close();
                                            break;
                                    }
                                }
                            }
                            break;
                        // 123롯데-vvip
                        case "004":
                            // 일반
                            if (_strCode2.ToLower()  == "y")
                            {
                                // "1" : 서울 , 신우편 "0" : 서울
                                if ((strCard_zipcode.Trim().Length == 6 && _strPlace == "1") || (strCard_zipcode.Trim().Length == 5 && _strPlace == "0"))
                                {
                                    if (strMember == "1" || strMember == "2")
                                    {
                                        if (strOwner == "11")
                                        {
                                            _sw = new StreamWriter(path + ".K_M_VVIP_오전(서울)_본인", true, _encoding);
                                            _sw.WriteLine(_strLine + "0863313");
                                            _sw.Close();
                                        }
                                        else
                                        {
                                            _sw = new StreamWriter(path + ".K_M_VVIP_오전(서울)", true, _encoding);
                                            _sw.WriteLine(_strLine + "0863305");
                                            _sw.Close();
                                        }
                                    }
                                    else
                                    {
                                        if (strOwner == "11")
                                        {
                                            _sw = new StreamWriter(path + ".K_VVIP_오전(서울)_본인", true, _encoding);
                                            _sw.WriteLine(_strLine + "0863309");
                                            _sw.Close();
                                        }
                                        else
                                        {
                                            _sw = new StreamWriter(path + ".K_VVIP_오전(서울)", true, _encoding);
                                            _sw.WriteLine(_strLine + "0863301");
                                            _sw.Close();
                                        }
                                    }
                                }
                                else
                                {
                                    if (strMember == "1" || strMember == "2")
                                    {
                                        if (strOwner == "11")
                                        {
                                            _sw = new StreamWriter(path + ".K_M_VVIP_오전(지방)_본인", true, _encoding);
                                            _sw.WriteLine(_strLine + "0863314");
                                            _sw.Close();
                                        }
                                        else
                                        {
                                            _sw = new StreamWriter(path + ".K_M_VVIP_오전(지방)", true, _encoding);
                                            _sw.WriteLine(_strLine + "0863306");
                                            _sw.Close();
                                        }
                                    }
                                    else
                                    {
                                        if (strOwner == "11")
                                        {
                                            _sw = new StreamWriter(path + ".K_VVIP_오전(지방)_본인", true, _encoding);
                                            _sw.WriteLine(_strLine + "0863310");
                                            _sw.Close();
                                        }
                                        else
                                        {
                                            _sw = new StreamWriter(path + ".K_VVIP_오전(지방)", true, _encoding);
                                            _sw.WriteLine(_strLine + "0863302");
                                            _sw.Close();
                                        }
                                    }
                                }
                            }
                            // 동의서
                            else
                            {
                                // "1" : 서울
                                if ((strCard_zipcode.Trim().Length == 6 && _strPlace == "1") || (strCard_zipcode.Trim().Length == 5 && _strPlace == "0"))
                                {
                                    if (strMember == "1")
                                    {
                                        _sw = new StreamWriter(path + ".M(별지)_VVIP_동_오전(서울)", true, _encoding);
                                        _sw.WriteLine(_strLine + "0862505");
                                        _sw.Close();
                                    }
                                    else if (strMember == "2")
                                    {
                                        _sw = new StreamWriter(path + ".M_VVIP_동_오전(서울)", true, _encoding);
                                        _sw.WriteLine(_strLine + "0862509");
                                        _sw.Close();
                                    }
                                    else
                                    {
                                        _sw = new StreamWriter(path + ".VVIP_동_오전(서울)", true, _encoding);
                                        _sw.WriteLine(_strLine + "0862501");
                                        _sw.Close();
                                    }
                                }
                                else
                                {
                                    if (strMember == "1")
                                    {
                                        _sw = new StreamWriter(path + ".M(별지)_VVIP_동_오전(지방)", true, _encoding);
                                        _sw.WriteLine(_strLine + "0862506");
                                        _sw.Close();
                                    }
                                    else if (strMember == "2")
                                    {
                                        _sw = new StreamWriter(path + ".M_VVIP_동_오전(지방)", true, _encoding);
                                        _sw.WriteLine(_strLine + "0862510");
                                        _sw.Close();
                                    }
                                    else
                                    {
                                        _sw = new StreamWriter(path + ".VVIP_동_오전(지방)", true, _encoding);
                                        _sw.WriteLine(_strLine + "0862502");
                                        _sw.Close();
                                    }
                                }
                            }
                            break;
                        // 087롯데-프리미어
                        case "005":
                            if (_strCode2.ToLower() == "y")
                            {
                                if (strMember == "1" || strMember == "2")
                                {
                                    if (strOwner == "11")
                                    {
                                        _sw = new StreamWriter(path + ".K_M_프리미어_오전_본인", true, _encoding);
                                        _sw.WriteLine(_strLine + "0863204");
                                        _sw.Close();
                                    }
                                    else
                                    {
                                        _sw = new StreamWriter(path + ".K_M_프리미어_오전", true, _encoding);
                                        _sw.WriteLine(_strLine + "0863202");
                                        _sw.Close();
                                    }
                                }
                                else
                                {
                                    if (strOwner == "11")
                                    {
                                        _sw = new StreamWriter(path + ".K_프리미어_오전_본인", true, _encoding);
                                        _sw.WriteLine(_strLine + "0863203");
                                        _sw.Close();
                                    }
                                    else
                                    {
                                        _sw = new StreamWriter(path + ".K_프리미어_오전", true, _encoding);
                                        _sw.WriteLine(_strLine + "0863201");
                                        _sw.Close();
                                    }
                                }
                            }
                            else
                            {
                                if (strMember == "1")
                                {
                                    switch (_strCode)
                                    {
                                        // [긴급-SK엔크린]
                                        case "10": _sw = new StreamWriter(path + ".M(별지)_프리미어_동_SK엔크린_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862402");
                                            _sw.Close();
                                            break;
                                        // [긴급-웅진]
                                        case "11": _sw = new StreamWriter(path + ".M(별지)_프리미어_동_웅진_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862402");
                                            _sw.Close();
                                            break;
                                        // [긴급-OK캐쉬백]
                                        case "12": _sw = new StreamWriter(path + ".M(별지)_프리미어_동_OK캐쉬백_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862402");
                                            _sw.Close();
                                            break;
                                        case "13": _sw = new StreamWriter(path + ".M(별지)_프리미어_동_아모레_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862402");
                                            _sw.Close();
                                            break;
                                        case "14": _sw = new StreamWriter(path + ".M(별지)_프리미어_동_오포인트_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862402");
                                            _sw.Close();
                                            break;
                                        case "15": _sw = new StreamWriter(path + ".M(별지)_프리미어_동_해피포인트_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862402");
                                            _sw.Close();
                                            break;
                                        case "16": _sw = new StreamWriter(path + ".M(별지)_프리미어_동_교보문고_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862402");
                                            _sw.Close();
                                            break;
                                        case "17": _sw = new StreamWriter(path + ".M(별지)_프리미어_동_뉴SOIL_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862402");
                                            _sw.Close();
                                            break;
                                        case "18": _sw = new StreamWriter(path + ".M(별지)_프리미어_동_T멤버십_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862402");
                                            _sw.Close();
                                            break;
                                        case "19": _sw = new StreamWriter(path + ".M(별지)_프리미어_동_아시아나_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862402");
                                            _sw.Close();
                                            break;
                                        case "20": _sw = new StreamWriter(path + ".M(별지)_프리미어_동_대한항공_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862402");
                                            _sw.Close();
                                            break;
                                        case "21": _sw = new StreamWriter(path + ".M(별지)_프리미어_동_E1LPG_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862402");
                                            _sw.Close();
                                            break;
                                        case "22": _sw = new StreamWriter(path + ".M(별지)_프리미어_동_아이행복_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862402");
                                            _sw.Close();
                                            break;
                                        case "23": _sw = new StreamWriter(path + ".M(별지)_프리미어_동_국민행복_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862402");
                                            _sw.Close();
                                            break;
                                        default: _sw = new StreamWriter(path + ".M(별지)_프리미어_동의서_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862402");
                                            _sw.Close();
                                            break;
                                    }
                                }
                                else if (strMember == "2")
                                {
                                    switch (_strCode)
                                    {
                                        // [긴급-SK엔크린]
                                        case "10": _sw = new StreamWriter(path + ".M_프리미어_동_SK엔크린_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862403");
                                            _sw.Close();
                                            break;
                                        // [긴급-웅진]
                                        case "11": _sw = new StreamWriter(path + ".M_프리미어_동_웅진_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862403");
                                            _sw.Close();
                                            break;
                                        // [긴급-OK캐쉬백]
                                        case "12": _sw = new StreamWriter(path + ".M_프리미어_동_OK캐쉬백_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862403");
                                            _sw.Close();
                                            break;
                                        case "13": _sw = new StreamWriter(path + ".M_프리미어_동_아모레_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862403");
                                            _sw.Close();
                                            break;
                                        case "14": _sw = new StreamWriter(path + ".M_프리미어_동_오포인트_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862403");
                                            _sw.Close();
                                            break;
                                        case "15": _sw = new StreamWriter(path + ".M_프리미어_동_해피포인트_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862403");
                                            _sw.Close();
                                            break;
                                        case "16": _sw = new StreamWriter(path + ".M_프리미어_동_교보문고_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862403");
                                            _sw.Close();
                                            break;
                                        case "17": _sw = new StreamWriter(path + ".M_프리미어_동_뉴SOIL_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862403");
                                            _sw.Close();
                                            break;
                                        case "18": _sw = new StreamWriter(path + ".M_프리미어_동_T멤버십_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862403");
                                            _sw.Close();
                                            break;
                                        case "19": _sw = new StreamWriter(path + ".M_프리미어_동_아시아나_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862403");
                                            _sw.Close();
                                            break;
                                        case "20": _sw = new StreamWriter(path + ".M_프리미어_동_대한항공_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862403");
                                            _sw.Close();
                                            break;
                                        case "21": _sw = new StreamWriter(path + ".M_프리미어_동_E1LPG_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862403");
                                            _sw.Close();
                                            break;
                                        case "22": _sw = new StreamWriter(path + ".M_프리미어_동_아이행복_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862403");
                                            _sw.Close();
                                            break;
                                        case "23": _sw = new StreamWriter(path + ".M_프리미어_동_국민행복_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862403");
                                            _sw.Close();
                                            break;
                                        default: _sw = new StreamWriter(path + ".M_프리미어_동의서_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862403");
                                            _sw.Close();
                                            break;
                                    }
                                }
                                else
                                {
                                    switch (_strCode)
                                    {
                                        // [긴급-SK엔크린]
                                        case "10": _sw = new StreamWriter(path + ".프리미어_동_SK엔크린_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862401");
                                            _sw.Close();
                                            break;
                                        // [긴급-웅진]
                                        case "11": _sw = new StreamWriter(path + ".프리미어_동_웅진_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862401");
                                            _sw.Close();
                                            break;
                                        // [긴급-OK캐쉬백]
                                        case "12": _sw = new StreamWriter(path + ".프리미어_동_OK캐쉬백_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862401");
                                            _sw.Close();
                                            break;
                                        case "13": _sw = new StreamWriter(path + ".프리미어_동_아모레_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862401");
                                            _sw.Close();
                                            break;
                                        case "14": _sw = new StreamWriter(path + ".프리미어_동_오포인트_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862401");
                                            _sw.Close();
                                            break;
                                        case "15": _sw = new StreamWriter(path + ".프리미어_동_해피포인트_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862401");
                                            _sw.Close();
                                            break;
                                        case "16": _sw = new StreamWriter(path + ".프리미어_동_교보문고_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862401");
                                            _sw.Close();
                                            break;
                                        case "17": _sw = new StreamWriter(path + ".프리미어_동_뉴SOIL_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862401");
                                            _sw.Close();
                                            break;
                                        case "18": _sw = new StreamWriter(path + ".프리미어_동_T멤버십_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862401");
                                            _sw.Close();
                                            break;
                                        case "19": _sw = new StreamWriter(path + ".프리미어_동_아시아나_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862401");
                                            _sw.Close();
                                            break;
                                        case "20": _sw = new StreamWriter(path + ".프리미어_동_대한항공_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862401");
                                            _sw.Close();
                                            break;
                                        case "21": _sw = new StreamWriter(path + ".프리미어_동_E1LPG_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862401");
                                            _sw.Close();
                                            break;
                                        case "22": _sw = new StreamWriter(path + ".프리미어_동_아이행복_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862401");
                                            _sw.Close();
                                            break;
                                        case "23": _sw = new StreamWriter(path + ".프리미어_동_국민행복_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862401");
                                            _sw.Close();
                                            break;
                                        default: _sw = new StreamWriter(path + ".프리미어_동의서_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862401");
                                            _sw.Close();
                                            break;
                                    }
                                }
                            }
                            break;
                        //친절(민원)
                        case "006":
                            //긴급일반
                            if (_strCode2.ToLower() == "y")
                            {
                                if (strMember == "1" || strMember == "2")
                                {
                                    if (strOwner == "11")
                                    {
                                        _sw = new StreamWriter(path + ".K_친절_M_긴급_오전_본인", true, _encoding);
                                        _sw.WriteLine(_strLine + "0863115");
                                        _sw.Close();
                                    }
                                    else
                                    {
                                        _sw = new StreamWriter(path + ".K_친절_M_긴급_오전", true, _encoding);
                                        _sw.WriteLine(_strLine + "0863107");
                                        _sw.Close();
                                    }
                                }
                                else
                                {
                                    if (strOwner == "11")
                                    {
                                        _sw = new StreamWriter(path + ".K_친절_긴급_오전_본인", true, _encoding);
                                        _sw.WriteLine(_strLine + "0863113");
                                        _sw.Close();
                                    }
                                    else
                                    {
                                        _sw = new StreamWriter(path + ".K_친절_긴급_오전", true, _encoding);
                                        _sw.WriteLine(_strLine + "0863105");
                                        _sw.Close();
                                    }
                                }
                            }
                            //긴급동의
                            else
                            {
                                if (strMember == "1")
                                {
                                    switch (_strCode)
                                    {
                                        // [긴급-SK엔크린]
                                        case "10": _sw = new StreamWriter(path + ".친절_M(별지)_긴급_동_SK엔크린_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862307");
                                            _sw.Close();
                                            break;
                                        // [긴급-웅진]
                                        case "11": _sw = new StreamWriter(path + ".친절_M(별지)_긴급_동_웅진_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862307");
                                            _sw.Close();
                                            break;
                                        // [긴급-OK캐쉬백]
                                        case "12": _sw = new StreamWriter(path + ".친절_M(별지)_긴급_동_OK캐쉬백_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862307");
                                            _sw.Close();
                                            break;
                                        case "13": _sw = new StreamWriter(path + ".친절_M(별지)_긴급_동_아모레_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862307");
                                            _sw.Close();
                                            break;
                                        case "14": _sw = new StreamWriter(path + ".친절_M(별지)_긴급_동_오포인트_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862307");
                                            _sw.Close();
                                            break;
                                        case "15": _sw = new StreamWriter(path + ".친절_M(별지)_긴급_동_해피포인트_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862307");
                                            _sw.Close();
                                            break;
                                        case "16": _sw = new StreamWriter(path + ".친절_M(별지)_긴급_동_교보문고_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862307");
                                            _sw.Close();
                                            break;
                                        case "17": _sw = new StreamWriter(path + ".친절_M(별지)_긴급_동_뉴SOIL_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862307");
                                            _sw.Close();
                                            break;
                                        case "18": _sw = new StreamWriter(path + ".친절_M(별지)_긴급_동_T멤버십_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862307");
                                            _sw.Close();
                                            break;
                                        case "19": _sw = new StreamWriter(path + ".친절_M(별지)_긴급_동_아시아나_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862307");
                                            _sw.Close();
                                            break;
                                        case "20": _sw = new StreamWriter(path + ".친절_M(별지)_긴급_동_대한항공_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862307");
                                            _sw.Close();
                                            break;
                                        case "21": _sw = new StreamWriter(path + ".친절_M(별지)_긴급_동_E1LPG_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862307");
                                            _sw.Close();
                                            break;
                                        case "22": _sw = new StreamWriter(path + ".친절_M(별지)_긴급_동_아이행복_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862307");
                                            _sw.Close();
                                            break;
                                        case "23": _sw = new StreamWriter(path + ".친절_M(별지)_긴급_동_국민행복_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862307");
                                            _sw.Close();
                                            break;
                                        default: _sw = new StreamWriter(path + ".친절_M(별지)_긴급_동_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862307");
                                            _sw.Close();
                                            break;
                                    }
                                }
                                else if (strMember == "2")
                                {
                                    switch (_strCode)
                                    {
                                        // [긴급-SK엔크린]
                                        case "10": _sw = new StreamWriter(path + ".친절_M_긴급_동_SK엔크린_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862311");
                                            _sw.Close();
                                            break;
                                        // [긴급-웅진]
                                        case "11": _sw = new StreamWriter(path + ".친절_M_긴급_동_웅진_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862311");
                                            _sw.Close();
                                            break;
                                        // [긴급-OK캐쉬백]
                                        case "12": _sw = new StreamWriter(path + ".친절_M_긴급_동_OK캐쉬백_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862311");
                                            _sw.Close();
                                            break;
                                        case "13": _sw = new StreamWriter(path + ".친절_M_긴급_동_아모레_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862311");
                                            _sw.Close();
                                            break;
                                        case "14": _sw = new StreamWriter(path + ".친절_M_긴급_동_오포인트_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862311");
                                            _sw.Close();
                                            break;
                                        case "15": _sw = new StreamWriter(path + ".친절_M_긴급_동_해피포인트_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862311");
                                            _sw.Close();
                                            break;
                                        case "16": _sw = new StreamWriter(path + ".친절_M_긴급_동_교보문고_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862311");
                                            _sw.Close();
                                            break;
                                        case "17": _sw = new StreamWriter(path + ".친절_M_긴급_동_뉴SOIL_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862311");
                                            _sw.Close();
                                            break;
                                        case "18": _sw = new StreamWriter(path + ".친절_M_긴급_동_T멤버십_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862311");
                                            _sw.Close();
                                            break;
                                        case "19": _sw = new StreamWriter(path + ".친절_M_긴급_동_아시아나_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862311");
                                            _sw.Close();
                                            break;
                                        case "20": _sw = new StreamWriter(path + ".친절_M_긴급_동_대한항공_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862311");
                                            _sw.Close();
                                            break;
                                        case "21": _sw = new StreamWriter(path + ".친절_M_긴급_동_E1LPG_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862311");
                                            _sw.Close();
                                            break;
                                        case "22": _sw = new StreamWriter(path + ".친절_M_긴급_동_아이행복_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862311");
                                            _sw.Close();
                                            break;
                                        case "23": _sw = new StreamWriter(path + ".친절_M_긴급_동_국민행복_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862311");
                                            _sw.Close();
                                            break;
                                        default: _sw = new StreamWriter(path + ".친절_M_긴급_동_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862311");
                                            _sw.Close();
                                            break;
                                    }
                                }
                                else
                                {
                                    switch (_strCode)
                                    {
                                        // [긴급-SK엔크린]
                                        case "10": _sw = new StreamWriter(path + ".친절_긴급_동_SK엔크린_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862305");
                                            _sw.Close();
                                            break;
                                        // [긴급-웅진]
                                        case "11": _sw = new StreamWriter(path + ".친절_긴급_동_웅진_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862305");
                                            _sw.Close();
                                            break;
                                        // [긴급-OK캐쉬백]
                                        case "12": _sw = new StreamWriter(path + ".친절_긴급_동_OK캐쉬백_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862305");
                                            _sw.Close();
                                            break;
                                        case "13": _sw = new StreamWriter(path + ".친절_긴급_동_아모레_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862305");
                                            _sw.Close();
                                            break;
                                        case "14": _sw = new StreamWriter(path + ".친절_긴급_동_오포인트_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862305");
                                            _sw.Close();
                                            break;
                                        case "15": _sw = new StreamWriter(path + ".친절_긴급_동_해피포인트_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862305");
                                            _sw.Close();
                                            break;
                                        case "16": _sw = new StreamWriter(path + ".친절_긴급_동_교보문고_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862305");
                                            _sw.Close();
                                            break;
                                        case "17": _sw = new StreamWriter(path + ".친절_긴급_동_뉴SOIL_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862305");
                                            _sw.Close();
                                            break;
                                        case "18": _sw = new StreamWriter(path + ".친절_긴급_동_T멤버십_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862305");
                                            _sw.Close();
                                            break;
                                        case "19": _sw = new StreamWriter(path + ".친절_긴급_동_아시아나_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862305");
                                            _sw.Close();
                                            break;
                                        case "20": _sw = new StreamWriter(path + ".친절_긴급_동_대한항공_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862305");
                                            _sw.Close();
                                            break;
                                        case "21": _sw = new StreamWriter(path + ".친절_긴급_동_E1LPG_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862305");
                                            _sw.Close();
                                            break;
                                        case "22": _sw = new StreamWriter(path + ".친절_긴급_동_아이행복_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862305");
                                            _sw.Close();
                                            break;
                                        case "23": _sw = new StreamWriter(path + ".친절_긴급_동_국민행복_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862305");
                                            _sw.Close();
                                            break;
                                        default: _sw = new StreamWriter(path + ".친절_긴급_동_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862305");
                                            _sw.Close();
                                            break;
                                    }
                                }
                            }
                            break;
                        // 광역
                        //2013.12.06 태희철 현재 데이터 인수 없음
                        case "008":
                            // 문자를 소문자로 변환 후 비교
                            if (_strCode2.ToLower() == "y")
                            {
                                if (strMember == "1" || strMember == "2")
                                {
                                    if (strOwner == "11")
                                    {
                                        _sw = new StreamWriter(path + ".K_M_광역_오전_본인", true, _encoding);
                                        _sw.WriteLine(_strLine + "       ");
                                        _sw.Close();
                                    }
                                    else
                                    {
                                        _sw = new StreamWriter(path + ".K_M_광역_오전", true, _encoding);
                                        _sw.WriteLine(_strLine + "       ");
                                        _sw.Close();
                                    }
                                }
                                else
                                {
                                    if (strOwner == "11")
                                    {
                                        _sw = new StreamWriter(path + ".K_광역_오전_본인", true, _encoding);
                                        _sw.WriteLine(_strLine + "       ");
                                        _sw.Close();
                                    }
                                    else
                                    {
                                        _sw = new StreamWriter(path + ".K_광역_오전", true, _encoding);
                                        _sw.WriteLine(_strLine + "       ");
                                        _sw.Close();
                                    }
                                }
                            }
                            else
                            {
                                if (strMember == "1")
                                {
                                    switch (_strCode)
                                    {
                                        // [긴급-SK엔크린]
                                        case "10": _sw = new StreamWriter(path + ".M(별지)_광역_동_SK엔크린_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        // [긴급-웅진]
                                        case "11": _sw = new StreamWriter(path + ".M(별지)_광역_동_SK웅진_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        // [긴급-OK캐쉬백]
                                        case "12": _sw = new StreamWriter(path + ".M(별지)_광역_동_OK캐쉬백_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        case "13": _sw = new StreamWriter(path + ".M(별지)_광역_동_아모레_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        case "14": _sw = new StreamWriter(path + ".M(별지)_광역_동_오포인트_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        case "15": _sw = new StreamWriter(path + ".M(별지)_광역_동_해피포인트_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        case "16": _sw = new StreamWriter(path + ".M(별지)_광역_동_교보문고_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        case "17": _sw = new StreamWriter(path + ".M(별지)_광역_동_뉴SOIL_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        case "18": _sw = new StreamWriter(path + ".M(별지)_광역_동_T멤버십_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        case "19": _sw = new StreamWriter(path + ".M(별지)_광역_동_아시아나_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        case "20": _sw = new StreamWriter(path + ".M(별지)_광역_동_대한항공_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        case "21": _sw = new StreamWriter(path + ".M(별지)_광역_동_E1LPG_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        case "22": _sw = new StreamWriter(path + ".M(별지)_광역_동_아이행복_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        case "23": _sw = new StreamWriter(path + ".M(별지)_광역_동_국민행복_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        default: _sw = new StreamWriter(path + ".M(별지)_광역_동_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                    }
                                }
                                else if (strMember == "2")
                                {
                                    switch (_strCode)
                                    {
                                        // [긴급-SK엔크린]
                                        case "10": _sw = new StreamWriter(path + ".M_광역_동_SK엔크린_긴급_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        // [긴급-웅진]
                                        case "11": _sw = new StreamWriter(path + ".M_광역_동_SK웅진_긴급_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        // [긴급-OK캐쉬백]
                                        case "12": _sw = new StreamWriter(path + ".M_광역_동_OK캐쉬백_긴급_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        case "13": _sw = new StreamWriter(path + ".M_광역_동_아모레_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        case "14": _sw = new StreamWriter(path + ".M_광역_동_오포인트_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        case "15": _sw = new StreamWriter(path + ".M_광역_동_해피포인트_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        case "16": _sw = new StreamWriter(path + ".M_광역_동_교보문고_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        case "17": _sw = new StreamWriter(path + ".M_광역_동_뉴SOIL_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        case "18": _sw = new StreamWriter(path + ".M_광역_동_T멤버십_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        case "19": _sw = new StreamWriter(path + ".M_광역_동_아시아나_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        case "20": _sw = new StreamWriter(path + ".M_광역_동_대한항공_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        case "21": _sw = new StreamWriter(path + ".M_광역_동_E1LPG_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        case "22": _sw = new StreamWriter(path + ".M_광역_동_아이행복_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        case "23": _sw = new StreamWriter(path + ".M_광역_동_국민행복_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        default: _sw = new StreamWriter(path + ".M_광역_동_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                    }
                                }
                                else
                                {
                                    switch (_strCode)
                                    {
                                        // [긴급-SK엔크린]
                                        case "10": _sw = new StreamWriter(path + ".광역_동_SK엔크린_긴급_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        // [긴급-웅진]
                                        case "11": _sw = new StreamWriter(path + ".광역_동_SK웅진_긴급_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        // [긴급-OK캐쉬백]
                                        case "12": _sw = new StreamWriter(path + ".광역_동_OK캐쉬백_긴급_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        case "13": _sw = new StreamWriter(path + ".광역_동_아모레_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        case "14": _sw = new StreamWriter(path + ".광역_동_오포인트_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        case "15": _sw = new StreamWriter(path + ".광역_동_해피포인트_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        case "16": _sw = new StreamWriter(path + ".광역_동_교보문고_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        case "17": _sw = new StreamWriter(path + ".광역_동_뉴SOIL_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        case "18": _sw = new StreamWriter(path + ".광역_동_T멤버십_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        case "19": _sw = new StreamWriter(path + ".광역_동_아시아나_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        case "20": _sw = new StreamWriter(path + ".광역_동_대한항공_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        case "21": _sw = new StreamWriter(path + ".광역_동_E1LPG_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        case "22": _sw = new StreamWriter(path + ".광역_동_아이행복_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        case "23": _sw = new StreamWriter(path + ".광역_동_국민행복_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                        default: _sw = new StreamWriter(path + ".광역_동_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "       ");
                                            _sw.Close();
                                            break;
                                    }
                                }
                            }
                            break;
                        default:
                            if (_strCode2.ToLower() == "y")
                            {
                                if (strMember == "1" || strMember == "2")
                                {
                                    if (strOwner == "11")
                                    {
                                        _sw = new StreamWriter(path + ".01_M_일반_오전_본인", true, _encoding);
                                        _sw.WriteLine(_strLine + "0861104");
                                        _sw.Close();
                                    }
                                    else
                                    {
                                        _sw = new StreamWriter(path + ".01_M_일반_오전", true, _encoding);
                                        _sw.WriteLine(_strLine + "0861102");
                                        _sw.Close();
                                    }
                                }
                                else
                                {
                                    if (strOwner == "11")
                                    {
                                        _sw = new StreamWriter(path + ".01_일반_오전_본인", true, _encoding);
                                        _sw.WriteLine(_strLine + "0861103");
                                        _sw.Close();
                                    }
                                    else
                                    {
                                        _sw = new StreamWriter(path + ".01_일반_오전", true, _encoding);
                                        _sw.WriteLine(_strLine + "0861101");
                                        _sw.Close();
                                    }
                                }
                            }
                            else
                            {
                                if (strMember == "1")
                                {
                                    switch (_strCode)
                                    {
                                        // [SK엔크린]
                                        case "10": _sw = new StreamWriter(path + ".M(별지)_" + _strCode + "_SK엔크린_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862603");
                                            _sw.Close();
                                            break;
                                        // [웅진]
                                        case "11": _sw = new StreamWriter(path + ".M(별지)_" + _strCode + "_웅진_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862602");
                                            _sw.Close();
                                            break;
                                        // [OK캐쉬백]
                                        case "12": _sw = new StreamWriter(path + ".M(별지)_" + _strCode + "_OK_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862604");
                                            _sw.Close();
                                            break;
                                        case "13": _sw = new StreamWriter(path + ".M(별지)_" + _strCode + "_아모레_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862605");
                                            _sw.Close();
                                            break;
                                        case "14": _sw = new StreamWriter(path + ".M(별지)_" + _strCode + "_오포인트_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862606");
                                            _sw.Close();
                                            break;
                                        case "15": _sw = new StreamWriter(path + ".M(별지)_" + _strCode + "_해피포인트_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862607");
                                            _sw.Close();
                                            break;
                                        case "16": _sw = new StreamWriter(path + ".M(별지)_" + _strCode + "_교보문고_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862608");
                                            _sw.Close();
                                            break;
                                        case "17": _sw = new StreamWriter(path + ".M(별지)_" + _strCode + "_뉴SOIL_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862609");
                                            _sw.Close();
                                            break;
                                        case "18": _sw = new StreamWriter(path + ".M(별지)_" + _strCode + "_T멤버십_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862610");
                                            _sw.Close();
                                            break;
                                        case "19": _sw = new StreamWriter(path + ".M(별지)_" + _strCode + "_아시아나_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862611");
                                            _sw.Close();
                                            break;
                                        case "20": _sw = new StreamWriter(path + ".M(별지)_" + _strCode + "_대한항공_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862612");
                                            _sw.Close();
                                            break;
                                        case "21": _sw = new StreamWriter(path + ".M(별지)_" + _strCode + "_E1LPG_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862613");
                                            _sw.Close();
                                            break;
                                        case "22": _sw = new StreamWriter(path + ".M(별지)_" + _strCode + "_아이행복_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862614");
                                            _sw.Close();
                                            break;
                                        case "23": _sw = new StreamWriter(path + ".M(별지)_" + _strCode + "_국민행복_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862615");
                                            _sw.Close();
                                            break;
                                        case "24": _sw = new StreamWriter(path + ".M(별지)_" + _strCode + "_이랜드디테일_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862616");
                                            _sw.Close();
                                            break;
                                        case "25": _sw = new StreamWriter(path + ".M(별지)_" + _strCode + "_E_POINT_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862617");
                                            _sw.Close();
                                            break;
                                        case "26": _sw = new StreamWriter(path + ".M(별지)_" + _strCode + "_E_LOLA_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862618");
                                            _sw.Close();
                                            break;
                                        default: _sw = new StreamWriter(path + ".M(별지)_동의서_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862601");
                                            _sw.Close();
                                            break;
                                    }
                                }
                                else if (strMember == "2")
                                {
                                    switch (_strCode)
                                    {
                                        // [SK엔크린]
                                        case "10": _sw = new StreamWriter(path + ".M_" + _strCode + "_SK엔크린_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862703");
                                            _sw.Close();
                                            break;
                                        // [웅진]
                                        case "11": _sw = new StreamWriter(path + ".M_" + _strCode + "_웅진_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862702");
                                            _sw.Close();
                                            break;
                                        // [OK캐쉬백]
                                        case "12": _sw = new StreamWriter(path + ".M_" + _strCode + "_OK_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862704");
                                            _sw.Close();
                                            break;
                                        case "13": _sw = new StreamWriter(path + ".M_" + _strCode + "_아모레_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862705");
                                            _sw.Close();
                                            break;
                                        case "14": _sw = new StreamWriter(path + ".M" + _strCode + "_오포인트_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862706");
                                            _sw.Close();
                                            break;
                                        case "15": _sw = new StreamWriter(path + ".M_" + _strCode + "_해피포인트_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862707");
                                            _sw.Close();
                                            break;
                                        case "16": _sw = new StreamWriter(path + ".M_" + _strCode + "_교보문고_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862708");
                                            _sw.Close();
                                            break;
                                        case "17": _sw = new StreamWriter(path + ".M_" + _strCode + "_뉴SOIL_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862709");
                                            _sw.Close();
                                            break;
                                        case "18": _sw = new StreamWriter(path + ".M_" + _strCode + "_T멤버십_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862710");
                                            _sw.Close();
                                            break;
                                        case "19": _sw = new StreamWriter(path + ".M_" + _strCode + "_아시아나_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862711");
                                            _sw.Close();
                                            break;
                                        case "20": _sw = new StreamWriter(path + ".M_" + _strCode + "_대한항공_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862712");
                                            _sw.Close();
                                            break;
                                        case "21": _sw = new StreamWriter(path + ".M_" + _strCode + "_E1LPG_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862713");
                                            _sw.Close();
                                            break;
                                        case "22": _sw = new StreamWriter(path + ".M_" + _strCode + "_아이행복_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862714");
                                            _sw.Close();
                                            break;
                                        case "23": _sw = new StreamWriter(path + ".M_" + _strCode + "_국민행복_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862715");
                                            _sw.Close();
                                            break;
                                        case "24": _sw = new StreamWriter(path + ".M_" + _strCode + "_이랜드리테일_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862716");
                                            _sw.Close();
                                            break;
                                        case "25": _sw = new StreamWriter(path + ".M_" + _strCode + "_E_POINT_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862717");
                                            _sw.Close();
                                            break;
                                        case "26": _sw = new StreamWriter(path + ".M_" + _strCode + "_LOLA_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862718");
                                            _sw.Close();
                                            break;
                                        default: _sw = new StreamWriter(path + ".M_동의서_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862701");
                                            _sw.Close();
                                            break;
                                    }
                                }
                                else
                                {
                                    switch (_strCode)
                                    {
                                        // [SK엔크린]
                                        case "10": _sw = new StreamWriter(path + ".동의_" + _strCode + "_SK엔크린_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862103");
                                            _sw.Close();
                                            break;
                                        // [웅진]
                                        case "11": _sw = new StreamWriter(path + ".동의_" + _strCode + "_웅진_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862102");
                                            _sw.Close();
                                            break;
                                        // [OK캐쉬백]
                                        case "12": _sw = new StreamWriter(path + ".동의_" + _strCode + "_OK_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862104");
                                            _sw.Close();
                                            break;
                                        case "13": _sw = new StreamWriter(path + ".동의_" + _strCode + "_아모레_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862106");
                                            _sw.Close();
                                            break;
                                        case "14": _sw = new StreamWriter(path + ".동의_" + _strCode + "_오포인트_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862107");
                                            _sw.Close();
                                            break;
                                        case "15": _sw = new StreamWriter(path + ".동의_" + _strCode + "_해피포인트_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862108");
                                            _sw.Close();
                                            break;
                                        case "16": _sw = new StreamWriter(path + ".동의_" + _strCode + "_교보문고_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862109");
                                            _sw.Close();
                                            break;
                                        case "17": _sw = new StreamWriter(path + ".동의_" + _strCode + "_뉴SOIL_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862110");
                                            _sw.Close();
                                            break;
                                        case "18": _sw = new StreamWriter(path + ".동의_" + _strCode + "_T멤버십_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862111");
                                            _sw.Close();
                                            break;
                                        case "19": _sw = new StreamWriter(path + ".동의_" + _strCode + "_아시아나_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862112");
                                            _sw.Close();
                                            break;
                                        case "20": _sw = new StreamWriter(path + ".동의_" + _strCode + "_대한항공_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862113");
                                            _sw.Close();
                                            break;
                                        case "21": _sw = new StreamWriter(path + ".동의_" + _strCode + "_E1LPG_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862114");
                                            _sw.Close();
                                            break;
                                        case "22": _sw = new StreamWriter(path + ".동의_" + _strCode + "_아이행복_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862115");
                                            _sw.Close();
                                            break;
                                        case "23": _sw = new StreamWriter(path + ".동의_" + _strCode + "_국민행복_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862116");
                                            _sw.Close();
                                            break;
                                        case "24": _sw = new StreamWriter(path + ".동의_" + _strCode + "_이랜드리테일_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862117");
                                            _sw.Close();
                                            break;
                                        case "25": _sw = new StreamWriter(path + ".동의_" + _strCode + "_E_POINT_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862118");
                                            _sw.Close();
                                            break;
                                        case "26": _sw = new StreamWriter(path + ".동의_" + _strCode + "_LOLA_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862119");
                                            _sw.Close();
                                            break;
                                        default: _sw = new StreamWriter(path + ".동의_동의서_오전", true, _encoding);
                                            _sw.WriteLine(_strLine + "0862101");
                                            _sw.Close();
                                            break;
                                    }
                                }
                            }
                            break;
                    }

                    //_sw.WriteLine(_strLine);
                    //_sw.Close();
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
