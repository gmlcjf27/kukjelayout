﻿using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace NPI_우체국_DOWN
{
	public class CONVERT
	{
		//기본 인코딩 설정
		private static string strEncoding = "ks_c_5601-1987";
        private static string strCardTypeID = "999";
        private static string strCardTypeName = "NPI_우체국_자료다운";

        //제휴사코드 반환
        public static string GetCardType(string path)
        {
            string _strReturn = "";

            return _strReturn;
        }

		//현 DLL의 카드 타입 코드 반환
		public static string GetCardTypeID() 
        {
            int _iReturn = 0;
            string _strReturn = "";
            FormSelectReceive _f = new FormSelectReceive();
            if (_f.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                _iReturn = _f.GetSelected;
            }
            //카드사 대분류 코드를 사용
            switch (_iReturn)
            {
                case 1:
                    strCardTypeID = "001";
                    break;
                case 2:
                    strCardTypeID = "002";
                    break;
                case 3:
                    strCardTypeID = "003";
                    break;
                case 4:
                    strCardTypeID = "005";
                    break;
                case 5:
                    strCardTypeID = "015";
                    break;
                case 6:
                    strCardTypeID = "016";
                    break;
                case 7:
                    strCardTypeID = "089";
                    break;
                case 8:
                    strCardTypeID = "004";
                    break;
                case 9:
                    strCardTypeID = "006";
                    break;
                case 10:
                    strCardTypeID = "018";
                    break;
                default:
                    strCardTypeID = "";
                    break;
            }
			return strCardTypeID;
		}

		//현 DLL의 카드 타입명 반환
		public static string GetCardTypeName() {
			return strCardTypeName;
		}

        //NPI자료생성
        public static string ConvertResult(DataTable dtable, string fileName)
        {
            string _strReturn = null;

            //카드사 대분류 코드를 사용
            switch (GetStringAsLength(strCardTypeID,3,true,' '))
            {
                case "001":
                    _strReturn = ConvertReceiveType1(dtable, fileName);
                    break;
                case "002":
                    _strReturn = ConvertReceiveType2(dtable, fileName);
                    break;
                case "003":
                    _strReturn = ConvertReceiveType3(dtable, fileName);
                    break;
                case "005":
                    _strReturn = ConvertReceiveType4(dtable, fileName);
                    break;
                case "015":
                    _strReturn = ConvertReceiveType5(dtable, fileName);
                    break;
                case "016":
                    _strReturn = ConvertReceiveType6(dtable, fileName);
                    break;
                case "089":
                    _strReturn = ConvertReceiveType7(dtable, fileName);
                    break;
                case "004":
                    _strReturn = ConvertReceiveType8(dtable, fileName);
                    break;
                case "006":
                    _strReturn = ConvertReceiveType9(dtable, fileName);
                    break;
                case "018":
                    _strReturn = ConvertReceiveType10(dtable, fileName);
                    break;
                default:
                    _strReturn = "";
                    break;
            }
            return _strReturn;
        }
        //일일마감자료
        public static string ConvertResultDay(System.Data.DataTable dtable, string fileName)
        {
            return ConvertResult(dtable, fileName);
        }
        //비씨NPI
        private static string ConvertReceiveType1(System.Data.DataTable dtable, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩
            StreamWriter _sw01 = null, _sw02 = null;
            StringBuilder _strLine = new StringBuilder("");
            string _strReturn = "", strBankID = "", strBranch = "", strCard_delivery_place = "", strDelivery_limit = "";
            string strCard_Kind = "", strZipcode = "", strCard_agree_code = "", strCard_vip_code = "", strBank_chk = "";
            string strZipcode_Kind = "", strZipcode_new = "", strCard_barcode_new = "";
            int i = 0;
            try
            {

                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    //영업점구분 2 or 3 or 4
                    strCard_delivery_place = dtable.Rows[i]["card_delivery_place_code"].ToString();
                    strBranch = dtable.Rows[i]["card_branch"].ToString();
                    strDelivery_limit = dtable.Rows[i]["delivery_limit_day"].ToString();
                    strBankID = dtable.Rows[i]["card_bank_ID"].ToString();
                    strCard_Kind = dtable.Rows[i]["card_kind"].ToString();
                    strZipcode = dtable.Rows[i]["card_zipcode"].ToString();
                    strCard_agree_code = dtable.Rows[i]["card_agree_code"].ToString();
                    strCard_vip_code = dtable.Rows[i]["card_vip_code"].ToString();

                    strZipcode_new = dtable.Rows[i]["card_zipcode_new"].ToString();
                    strZipcode_Kind = dtable.Rows[i]["card_zipcode_kind"].ToString();

                    strCard_barcode_new = dtable.Rows[i]["card_barcode_new"].ToString();

                    if (strBankID.Length > 2)
                    {
                        strBankID = strBankID.Substring(0, 2);
                    }

                    //데이터생성 시작
                    if (strBranch == "012")
                    {
                        //2018.09.06 태희철 수정
                        //KT링커스 업무로 인한 비씨동의서 등기건 구분 작업
                        _strLine = new StringBuilder(GetStringAsLength(strCard_barcode_new, 22, true, ' ') + ",");

                        if (dtable.Rows[i]["card_type_detail"].ToString().Substring(0,4) == "0012")
                        {
                            if (dtable.Rows[i]["card_type_detail"].ToString() == "0012101")
                            {
                                if (strZipcode_Kind == "1")
                                {
                                    _strLine = new StringBuilder(GetStringAsLength(strCard_barcode_new, 22, true, ' ') + "," + strZipcode_new);
                                    _sw02 = new StreamWriter(fileName + "bc_woori_DONG_new.txt", true, _encoding);
                                }
                                else
                                {
                                    _strLine = new StringBuilder(GetStringAsLength(strCard_barcode_new, 22, true, ' ') + "," + strZipcode.Substring(0, 5));
                                    _sw02 = new StreamWriter(fileName + "bc_woori_DONG_OLD.txt", true, _encoding);
                                    
                                }
                                _sw02.WriteLine(_strLine.ToString());
                                _sw02.Close();
                            }

                            if (strCard_barcode_new.Length > 27)
                            {
                                _strLine = new StringBuilder(GetStringAsLength(strCard_barcode_new, 28, true, ' ') + "," + strCard_barcode_new.Substring(23, 5));
                                _sw01 = new StreamWriter(fileName + "bc_012_DONG_bank.txt", true, _encoding);
                            }
                            else
                            {
                                _strLine = new StringBuilder(GetStringAsLength(strCard_barcode_new, 28, true, ' ') + ",");
                                _sw01 = new StreamWriter(fileName + "bc_012_DONG_bank.txt", true, _encoding);
                            }
                        }
                        else
                        {
                            if (strZipcode_Kind == "1")
                            {
                                _strLine = new StringBuilder(GetStringAsLength(strCard_barcode_new, 22, true, ' ') + "," + strZipcode_new);
                                _sw01 = new StreamWriter(fileName + "bc_012_DONG_new.txt", true, _encoding);
                            }
                            else
                            {
                                _strLine = new StringBuilder(GetStringAsLength(strCard_barcode_new, 22, true, ' ') + "," + strZipcode.Substring(0, 5));
                                _sw01 = new StreamWriter(fileName + "bc_012_DONG_OLD.txt", true, _encoding);
                            }
                        }
                    }
                    else
                    {
                        if (strZipcode_Kind == "1")
                        {
                            if (dtable.Rows[i]["card_type_detail"].ToString().Substring(0,4) == "0012" && (
                                strBranch == "100" || strBranch == "120" || strBranch == "140" || strBranch == "160" ||
                                strBranch == "180" || strBranch == "400" || strBranch == "401" || strBranch == "450")
                                )
                            {
                                _strLine = new StringBuilder(GetStringAsLength(strCard_barcode_new, 22, true, ' ') + "," + strZipcode_new);
                                _sw01 = new StreamWriter(fileName + "bc_new_서울수도권.txt", true, _encoding);
                            }
                            else
                            {
                                _strLine = new StringBuilder(GetStringAsLength(strCard_barcode_new, 22, true, ' ') + "," + strZipcode_new);
                                _sw01 = new StreamWriter(fileName + "bc_new.txt", true, _encoding);
                            }
                        }
                        else
                        {
                            if (dtable.Rows[i]["card_type_detail"].ToString().Substring(0, 4) == "0012" && (
                                strBranch == "100" || strBranch == "120" || strBranch == "140" || strBranch == "160" ||
                                strBranch == "180" || strBranch == "400" || strBranch == "401" || strBranch == "450")
                                )
                            {
                                _strLine = new StringBuilder(GetStringAsLength(strCard_barcode_new, 22, true, ' ') + "," + strZipcode.Substring(0, 5));
                                _sw01 = new StreamWriter(fileName + "bc_OLD_서울수도권.txt", true, _encoding);
                            }
                            else
                            {
                                _strLine = new StringBuilder(GetStringAsLength(strCard_barcode_new, 22, true, ' ') + "," + strZipcode.Substring(0, 5));
                                _sw01 = new StreamWriter(fileName + "bc_OLD.txt", true, _encoding);
                            }
                        }
                    }

                    _sw01.WriteLine(_strLine.ToString());
                    _sw01.Close();
                }

                _strReturn = string.Format("{0}건의 NPI데이타 다운 완료", i);
            }
            catch (Exception)
            {
                _strReturn = string.Format("{0}번째 데이터 인계 중 오류 발생", i + 1);
            }
            finally
            {
                if (_sw01 != null) _sw01.Close();
                if (_sw02 != null) _sw02.Close();
            }
            return _strReturn;
        }

        //국민NPI자료생성_back
        private static string ConvertReceiveType2_back(System.Data.DataTable dtable, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩
            StreamWriter _sw01 = null, _sw02 = null, _sw03 = null, _sw04 = null, _sw05 = null;
            StreamWriter _sw01_new = null, _sw02_new = null, _sw03_new = null, _sw04_new = null, _sw05_new = null;
            StringBuilder _strLine = new StringBuilder("");
            string _strReturn = "", strBankID = "", strBranch = "", strClient_register_type = "", strDelivery_limit = "";
            string strCard_Kind = "", strZipcode = "", strCard_agree_code = "", strCard_vip_code = "", strClient_insert_type ="";
            string strZipcode_Kind = "", strZipcode_new = "";
            int i = 0, icnt = 0;
            try
            {
                //StreamWriter 초기화
                _sw01 = new StreamWriter(fileName + "Kuk_OLD.txt", true, _encoding);
                _sw02 = new StreamWriter(fileName + "Kd_OLD.txt", true, _encoding);
                _sw03 = new StreamWriter(fileName + "Ke_OLD.txt", true, _encoding);
                _sw04 = new StreamWriter(fileName + "Kg_OLD.txt", true, _encoding);
                _sw05 = new StreamWriter(fileName + "Kg012_OLD.txt", true, _encoding);

                //StreamWriter 초기화
                _sw01_new = new StreamWriter(fileName + "Kuk_new.txt", true, _encoding);
                _sw02_new = new StreamWriter(fileName + "Kd_new.txt", true, _encoding);
                _sw03_new = new StreamWriter(fileName + "Ke_new.txt", true, _encoding);
                _sw04_new = new StreamWriter(fileName + "Kg_new.txt", true, _encoding);
                _sw05_new = new StreamWriter(fileName + "Kg012_new.txt", true, _encoding);

                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    //영업점구분 card_kind = B
                    strCard_Kind = dtable.Rows[i]["card_kind"].ToString();
                    strClient_register_type = dtable.Rows[i]["client_register_type"].ToString();
                    strClient_insert_type = dtable.Rows[i]["client_insert_type"].ToString();

                    strBranch = dtable.Rows[i]["card_branch"].ToString();
                    strDelivery_limit = dtable.Rows[i]["delivery_limit_day"].ToString();
                    strBankID = dtable.Rows[i]["card_bank_ID"].ToString();
                    
                    strZipcode = dtable.Rows[i]["card_zipcode"].ToString();
                    strCard_agree_code = dtable.Rows[i]["card_agree_code"].ToString();
                    strCard_vip_code = dtable.Rows[i]["card_vip_code"].ToString();

                    strZipcode_new = dtable.Rows[i]["card_zipcode_new"].ToString();
                    strZipcode_Kind = dtable.Rows[i]["card_zipcode_kind"].ToString();

                    //데이터생성 시작
                    _strLine = new StringBuilder(GetStringAsLength(dtable.Rows[i]["card_barcode_new"].ToString(), 22, true, ' ') + ",");

                    if (strZipcode_Kind == "0")
                    {
                        _strLine.Append(GetStringAsLength(strZipcode, 5, true, ' '));
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength(strZipcode_new, 5, true, ' '));
                    }

                    //_strLine.Append(GetStringAsLength(strZipcode, 5, true, ' '));

                    //긴급
                    if (strZipcode_Kind == "1")
                    {
                        if ((strClient_insert_type == "2" || strClient_insert_type == "3") && strClient_register_type == "I")
                        {
                            icnt++;
                            if (strBranch == "012")
                            {
                                _sw05_new.WriteLine(_strLine.ToString());
                            }
                            else
                            {
                                _sw04_new.WriteLine(_strLine.ToString());
                            }
                        }
                        else if (strClient_register_type == "Q" && strBranch != "012")
                        {
                            icnt++;
                            _sw03_new.WriteLine(_strLine.ToString());
                        }
                        //동의서(서울,수도권)
                        else if (((strBranch.Substring(0, 1) == "1" || strBranch.Substring(0, 1) == "4") && strClient_register_type == "D") && strBranch != "012")
                        {
                            icnt++;
                            _sw02_new.WriteLine(_strLine.ToString());
                        }
                        //일반 + 동의서(지방)
                        //strClient_register_type = 065국민긴급의 5000차(P), 21000차(G)
                        else if ((strClient_register_type != "P" && strClient_register_type != "G") && strBranch != "012")
                        {
                            icnt++;
                            _sw01_new.WriteLine(_strLine.ToString());
                        }
                    }
                    else
                    {
                        if ((strClient_insert_type == "2" || strClient_insert_type == "3") && strClient_register_type == "I")
                        {
                            icnt++;
                            if (strBranch == "012")
                            {
                                _sw05.WriteLine(_strLine.ToString());
                            }
                            else
                            {
                                _sw04.WriteLine(_strLine.ToString());
                            }
                        }
                        else if (strClient_register_type == "Q" && strBranch != "012")
                        {
                            icnt++;
                            _sw03.WriteLine(_strLine.ToString());
                        }
                        //동의서(서울)
                        else if (((strBranch.Substring(0, 1) == "1" || strBranch.Substring(0, 1) == "4") && strClient_register_type == "D") && strBranch != "012")
                        {
                            icnt++;
                            _sw02.WriteLine(_strLine.ToString());
                        }
                        //일반 + 동의서(지방)
                        //strClient_register_type = 065국민긴급의 5000차(P), 21000차(G)
                        else if ((strClient_register_type != "P" && strClient_register_type != "G") && strBranch != "012")
                        {
                            icnt++;
                            _sw01.WriteLine(_strLine.ToString());
                        }
                    }
                    
                }

                _strReturn = string.Format("{0}건의 NPI데이타 다운 완료", icnt);
            }
            catch (Exception)
            {
                _strReturn = string.Format("{0}번째 데이터 인계 중 오류 발생", i + 1);
            }
            finally
            {
                if (_sw01 != null) _sw01.Close();
                if (_sw02 != null) _sw02.Close();
                if (_sw03 != null) _sw03.Close();
                if (_sw04 != null) _sw04.Close();
                if (_sw05 != null) _sw05.Close();

                if (_sw01_new != null) _sw01_new.Close();
                if (_sw02_new != null) _sw02_new.Close();
                if (_sw03_new != null) _sw03_new.Close();
                if (_sw04_new != null) _sw04_new.Close();
                if (_sw05_new != null) _sw05_new.Close();
            }
            return _strReturn;
        }

        //국민NPI자료생성
        private static string ConvertReceiveType2(System.Data.DataTable dtable, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩
            StreamWriter _sw01 = null, _sw02 = null, _sw03 = null, _sw04 = null, _sw05 = null;
            StreamWriter _sw01_new = null, _sw02_new = null, _sw03_new = null, _sw04_new = null, _sw05_new = null;
            StringBuilder _strLine = new StringBuilder("");
            string _strReturn = "", strBankID = "", strBranch = "", strClient_register_type = "", strDelivery_limit = "";
            string strCard_Kind = "", strZipcode = "", strCard_agree_code = "", strCard_vip_code = "", strClient_insert_type = "";
            string strZipcode_Kind = "", strZipcode_new = "";
            int i = 0, icnt = 0;
            try
            {
                //StreamWriter 초기화
                _sw01 = new StreamWriter(fileName + "Kuk_OLD.txt", true, _encoding);
                _sw02 = new StreamWriter(fileName + "Kd_OLD.txt", true, _encoding);
                _sw03 = new StreamWriter(fileName + "Ke_OLD.txt", true, _encoding);
                _sw04 = new StreamWriter(fileName + "Kg_OLD.txt", true, _encoding);
                _sw05 = new StreamWriter(fileName + "Kg012_OLD.txt", true, _encoding);

                //StreamWriter 초기화
                _sw01_new = new StreamWriter(fileName + "Kuk_new.txt", true, _encoding);
                _sw02_new = new StreamWriter(fileName + "Kd_new.txt", true, _encoding);
                _sw03_new = new StreamWriter(fileName + "Ke_new.txt", true, _encoding);
                _sw04_new = new StreamWriter(fileName + "Kg_new.txt", true, _encoding);
                _sw05_new = new StreamWriter(fileName + "Kg012_new.txt", true, _encoding);

                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    //영업점구분 card_kind = B
                    strCard_Kind = dtable.Rows[i]["card_kind"].ToString();
                    strClient_register_type = dtable.Rows[i]["client_register_type"].ToString();
                    strClient_insert_type = dtable.Rows[i]["client_insert_type"].ToString();

                    strBranch = dtable.Rows[i]["card_branch"].ToString();
                    strDelivery_limit = dtable.Rows[i]["delivery_limit_day"].ToString();
                    strBankID = dtable.Rows[i]["card_bank_ID"].ToString();

                    strZipcode = dtable.Rows[i]["card_zipcode"].ToString();
                    strCard_agree_code = dtable.Rows[i]["card_agree_code"].ToString();
                    strCard_vip_code = dtable.Rows[i]["card_vip_code"].ToString();

                    strZipcode_new = dtable.Rows[i]["card_zipcode_new"].ToString();
                    strZipcode_Kind = dtable.Rows[i]["card_zipcode_kind"].ToString();

                    //데이터생성 시작
                    _strLine = new StringBuilder(GetStringAsLength(dtable.Rows[i]["card_barcode_new"].ToString(), 22, true, ' ') + ",");

                    if (strZipcode_Kind == "0")
                    {
                        _strLine.Append(GetStringAsLength(strZipcode, 5, true, ' '));
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength(strZipcode_new, 5, true, ' '));
                    }

                    //_strLine.Append(GetStringAsLength(strZipcode, 5, true, ' '));

                    //신우편
                    if (strZipcode_Kind == "1")
                    {   
                        //등기
                        if (strBranch == "012")
                        {
                            icnt++;
                            _sw05_new.WriteLine(_strLine.ToString());
                        }
                        //갱신
                        else if ((strClient_insert_type == "2" || strClient_insert_type == "3") && strClient_register_type == "I")
                        {
                            icnt++;
                            _sw04_new.WriteLine(_strLine.ToString());
                        }
                        //동의서(수도권)
                        else if (((strBranch.Substring(0, 1) == "1" || strBranch.Substring(0, 1) == "4") && strClient_register_type == "D") && strBranch != "012")
                        {
                            icnt++;
                            _sw02_new.WriteLine(_strLine.ToString());
                        }
                        else
                        {
                            icnt++;
                            _sw01_new.WriteLine(_strLine.ToString());
                        }
                    }
                    else
                    {
                        //등기
                        if (strBranch == "012")
                        {
                            icnt++;
                            _sw05.WriteLine(_strLine.ToString());
                        }
                        //갱신
                        else if ((strClient_insert_type == "2" || strClient_insert_type == "3") && strClient_register_type == "I")
                        {
                            icnt++;
                            _sw04.WriteLine(_strLine.ToString());
                        }
                        //동의서(수도권)
                        else if (((strBranch.Substring(0, 1) == "1" || strBranch.Substring(0, 1) == "4") && strClient_register_type == "D") && strBranch != "012")
                        {
                            icnt++;
                            _sw02.WriteLine(_strLine.ToString());
                        }
                        //일반 + 동의서(지방)
                        //strClient_register_type = 065국민긴급의 5000차(P), 21000차(G)
                        else
                        {
                            icnt++;
                            _sw01.WriteLine(_strLine.ToString());
                        }
                    }
                }

                _strReturn = string.Format("{0}건의 NPI데이타 다운 완료", icnt);
            }
            catch (Exception)
            {
                _strReturn = string.Format("{0}번째 데이터 인계 중 오류 발생", i + 1);
            }
            finally
            {
                if (_sw01 != null) _sw01.Close();
                if (_sw02 != null) _sw02.Close();
                if (_sw03 != null) _sw03.Close();
                if (_sw04 != null) _sw04.Close();
                if (_sw05 != null) _sw05.Close();

                if (_sw01_new != null) _sw01_new.Close();
                if (_sw02_new != null) _sw02_new.Close();
                if (_sw03_new != null) _sw03_new.Close();
                if (_sw04_new != null) _sw04_new.Close();
                if (_sw05_new != null) _sw05_new.Close();
            }
            return _strReturn;
        }


        //카카오뱅크NPI자료생성
        private static string ConvertReceiveType9(System.Data.DataTable dtable, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩
            StreamWriter _sw01 = null, _sw02 = null, _sw03 = null, _sw04 = null, _sw05 = null;
            StreamWriter _sw01_new = null, _sw02_new = null, _sw03_new = null, _sw04_new = null, _sw05_new = null;
            StringBuilder _strLine = new StringBuilder("");
            string _strReturn = "", strClient_register_type = "";
            string strZipcode = "", strClient_insert_type = "";
            string strZipcode_Kind = "", strZipcode_new = "";
            int i = 0, icnt = 0;
            try
            {
                //StreamWriter 초기화
                _sw01 = new StreamWriter(fileName + "kakao_OLD.txt", true, _encoding);

                //StreamWriter 초기화
                _sw01_new = new StreamWriter(fileName + "kakao_new.txt", true, _encoding);

                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    //영업점구분 card_kind = B
                    strClient_register_type = dtable.Rows[i]["client_register_type"].ToString();
                    strClient_insert_type = dtable.Rows[i]["client_insert_type"].ToString();

                    strZipcode_new = dtable.Rows[i]["card_zipcode_new"].ToString();
                    strZipcode_Kind = dtable.Rows[i]["card_zipcode_kind"].ToString();

                    //데이터생성 시작
                    _strLine = new StringBuilder(GetStringAsLength(dtable.Rows[i]["card_barcode_new"].ToString(), 22, true, ' ') + ",");

                    if (strZipcode_Kind == "0")
                    {
                        icnt++;
                        _strLine.Append(GetStringAsLength(strZipcode, 5, true, ' '));
                        _sw01.WriteLine(_strLine.ToString());
                    }
                    else
                    {
                        icnt++;
                        _strLine.Append(GetStringAsLength(strZipcode_new, 5, true, ' '));
                        _sw01_new.WriteLine(_strLine.ToString());
                    }
                }

                _strReturn = string.Format("{0}건의 NPI데이타 다운 완료", icnt);
            }
            catch (Exception)
            {
                _strReturn = string.Format("{0}번째 데이터 인계 중 오류 발생", i + 1);
            }
            finally
            {
                if (_sw01 != null) _sw01.Close();

                if (_sw01_new != null) _sw01_new.Close();
            }
            return _strReturn;
        }


        //신한NPI자료생성
        private static string ConvertReceiveType3(System.Data.DataTable dtable, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩
            StreamWriter _sw01 = null, _sw02 = null, _sw03 = null;
            StringBuilder _strLine = new StringBuilder("");
            string _strReturn = "", strBarcode_new = "";
            int i = 0;
            try
            {
                //StreamWriter 초기화
                _sw01 = new StreamWriter(fileName + "sh_일반.txt", true, _encoding);
                _sw02 = new StreamWriter(fileName + "sh_동의.txt", true, _encoding);
                _sw03 = new StreamWriter(fileName + "sh_기타.txt", true, _encoding);
                

                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    strBarcode_new = dtable.Rows[i]["card_barcode_new"].ToString();

                    //데이터생성 시작
                    if (strBarcode_new.Length != 0)
                    {
                        _strLine = new StringBuilder(GetStringAsLength(strBarcode_new, 13, true, ' ').Trim() + ",");

                        if (strBarcode_new.Length > 12)
                        {
                            _strLine.Append(GetStringAsLength(strBarcode_new.Substring(8,5), 5, true, ' '));
                        }

                        //일반
                        if (strBarcode_new.Substring(0, 6) == "407279")
                        {
                            _sw01.WriteLine(_strLine.ToString());
                        }
                        //동의
                        else if (strBarcode_new.Substring(0, 6) == "407418")
                        {
                            _sw02.WriteLine(_strLine.ToString());
                        }
                        else
                        {
                            _sw03.WriteLine(_strLine.ToString());
                        }
                    }
                }
                _strReturn = string.Format("{0}건의 NPI데이타 다운 완료", i);
            }
            catch (Exception)
            {
                _strReturn = string.Format("{0}번째 데이터 인계 중 오류 발생", i + 1);
            }
            finally
            {
                if (_sw01 != null) _sw01.Close();
                if (_sw02 != null) _sw02.Close();
                if (_sw03 != null) _sw03.Close();
            }
            return _strReturn;
        }

        //하나NPI자료생성
        private static string ConvertReceiveType4(System.Data.DataTable dtable, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩
            StreamWriter _sw01 = null, _sw01_new = null,_sw02 = null;
            StringBuilder _strLine = new StringBuilder("");
            string _strReturn = "", strclient_send_number = "", strClient_send_date = "";
            string strDegree_code = "", strZipcode = "";
            string strZipcode_Kind = "", strZipcode_new = "", strCard_type_detail = "";
            int i = 0, iCnt = 0;
            try
            {
                //StreamWriter 초기화
                _sw01 = new StreamWriter(fileName + "HANA_Post.txt", true, _encoding);
                //_sw01_new = new StreamWriter(fileName + "HANA_new.txt", true, _encoding);
                
                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    //일반동의구분 : Y = 동의, N = 일반
                    if (dtable.Rows[i]["client_send_date"].ToString().Length > 8)
                    {
                        strClient_send_date = dtable.Rows[i]["client_send_date"].ToString().Replace("-", "").Substring(0, 8);    
                    }
                    strclient_send_number = dtable.Rows[i]["client_send_number"].ToString();
                    strZipcode = dtable.Rows[i]["card_zipcode"].ToString();
                    strZipcode_new = dtable.Rows[i]["card_zipcode_new"].ToString();
                    strZipcode_Kind = dtable.Rows[i]["card_zipcode_kind"].ToString();
                    strCard_type_detail = dtable.Rows[i]["card_type_detail"].ToString();


                    //데이터생성 시작
                    _strLine = new StringBuilder(GetStringAsLength(strClient_send_date + strclient_send_number, 16, true, ' ').Trim() + ",");

                    if (strZipcode_Kind == "1")
                    {
                        _strLine.Append(GetStringAsLength(strZipcode_new, 5, true, ' '));
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength(strZipcode, 5, true, ' '));
                    }

                    iCnt++;
                    _sw01.WriteLine(_strLine.ToString());

                    /*
                    //동의서(서울지사)
                    if (strClient_enterprise_code != "Y" && (strBranch.Substring(0, 1) == "1" || strBranch.Substring(0, 1) == "4"))
                    {
                        if (strZipcode_Kind == "1")
                        {
                            _sw02 = new StreamWriter(fileName + strDegree_code + "_NEW.txt", true, _encoding);
                            _sw02.WriteLine(_strLine.ToString());
                            _sw02.Close();
                        }
                        else
                        {
                            _sw02 = new StreamWriter(fileName + strDegree_code + "_OLD.txt", true, _encoding);
                            _sw02.WriteLine(_strLine.ToString());
                            _sw02.Close();
                        }
                    }
                    //일반 + 동의서(서울지사 외)
                    else
                    {
                        if (strZipcode_Kind == "1")
                        {
                            iCnt++;
                            _sw01_new.WriteLine(_strLine.ToString());
                        }
                        else
                        {
                            iCnt++;
                            _sw01.WriteLine(_strLine.ToString());
                        }
                    }
                    */
                }
                _strReturn = string.Format("{0}건의 NPI데이타 다운 완료", iCnt);
            }
            catch (Exception)
            {
                _strReturn = string.Format("{0}번째 데이터 인계 중 오류 발생", i + 1);
            }
            finally
            {
                if (_sw01 != null) _sw01.Close();
            }
            return _strReturn;
        }

        //롯데NPI자료생성
        private static string ConvertReceiveType5(System.Data.DataTable dtable, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩
            StreamWriter _sw01 = null, _sw02 = null;
            StreamWriter _sw01_new = null, _sw02_new = null;
            StringBuilder _strLine = new StringBuilder("");
            string _strReturn = "", strBranch = "", strCard_agree_code = "", strZipcode = "", strZipcode_new = "", strZipcode_kind = "";
            
            int i = 0;
            try
            {
                //StreamWriter 초기화
                _sw01 = new StreamWriter(fileName + "LOTTE_OLD.txt", true, _encoding);
                _sw02 = new StreamWriter(fileName + "LOT-D_OLD.txt", true, _encoding);

                _sw01_new = new StreamWriter(fileName + "LOTTE_new.txt", true, _encoding);
                _sw02_new = new StreamWriter(fileName + "LOT-D_new.txt", true, _encoding);

                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    //카드구분 : 일반,동의
                    strCard_agree_code = dtable.Rows[i]["card_agree_code"].ToString();
                    strBranch = dtable.Rows[i]["card_branch"].ToString();
                    strZipcode = dtable.Rows[i]["card_zipcode"].ToString();

                    strZipcode_kind = dtable.Rows[i]["card_zipcode_kind"].ToString();
                    strZipcode_new = dtable.Rows[i]["card_zipcode_new"].ToString();
                    


                    //데이터생성 시작
                    _strLine = new StringBuilder(GetStringAsLength(dtable.Rows[i]["card_barcode_new"].ToString(), 22, true, ' ') + ",");
                    if (strZipcode_kind == "1")
                    {
                        _strLine.Append(GetStringAsLength(strZipcode_new, 5, true, ' '));

                        //일반, 동의 구분 : 일반 = Y, 동의서 = N
                        //일반 (서울지사)
                        if (strCard_agree_code == "Y" && (strBranch.Substring(0, 1) == "1" || strBranch.Substring(0, 1) == "4"))
                        {
                            _sw01_new.WriteLine(_strLine.ToString());
                        }
                        //일반 (서울외지사) + 동의서
                        else
                        {
                            _sw02_new.WriteLine(_strLine.ToString());
                        }
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength(strZipcode, 5, true, ' '));

                        //일반, 동의 구분 : 일반 = Y, 동의서 = N
                        //일반 (서울지사)
                        if (strCard_agree_code == "Y" && (strBranch.Substring(0, 1) == "1" || strBranch.Substring(0, 1) == "4"))
                        {
                            _sw01.WriteLine(_strLine.ToString());
                        }
                        //일반 (서울외지사) + 동의서
                        else
                        {
                            _sw02.WriteLine(_strLine.ToString());
                        }
                    }

                    ////일반, 동의 구분 : 일반 = Y, 동의서 = N
                    ////일반 (서울지사)
                    //if (strCard_agree_code == "Y" && strBranch.Substring(0,1) == "1")
                    //{
                    //    _sw01.WriteLine(_strLine.ToString());
                    //}
                    ////일반 (서울외지사) + 동의서
                    //else
                    //{
                    //    _sw02.WriteLine(_strLine.ToString());
                    //}
                }
                _strReturn = string.Format("{0}건의 NPI데이타 다운 완료", i);
            }
            catch (Exception)
            {
                _strReturn = string.Format("{0}번째 데이터 인계 중 오류 발생", i + 1);
            }
            finally
            {
                if (_sw01 != null) _sw01.Close();
                if (_sw02 != null) _sw02.Close();
                if (_sw01_new != null) _sw01_new.Close();
                if (_sw02_new != null) _sw02_new.Close();
            }
            return _strReturn;
        }

        //현대
        private static string ConvertReceiveType6(System.Data.DataTable dtable, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩
            StreamWriter _sw01 = null, _sw02 = null, _sw03 = null, _sw04 = null;
            StreamWriter _sw01_new = null, _sw02_new = null, _sw03_new = null, _sw04_new = null;
            StringBuilder _strLine = new StringBuilder("");
            string _strReturn = "", strBranch = "", strZipcode = "";
            string strZipcode_kind = "", strZipcode_new = "", strClient_send_number = "";

            int i = 0;
            try
            {
                //StreamWriter 초기화
                _sw01 = new StreamWriter(fileName + "HD_OLD.txt", true, _encoding);
                _sw02 = new StreamWriter(fileName + "HD_012_OLD.txt", true, _encoding);

                _sw01_new = new StreamWriter(fileName + "HD_new.txt", true, _encoding);
                _sw02_new = new StreamWriter(fileName + "HD_012_new.txt", true, _encoding);

                _sw03 = new StreamWriter(fileName + "HD_2D_OLD.txt", true, _encoding);
                _sw04 = new StreamWriter(fileName + "HD_2D_012_OLD.txt", true, _encoding);

                _sw03_new = new StreamWriter(fileName + "HD_2D_new.txt", true, _encoding);
                _sw04_new = new StreamWriter(fileName + "HD_2D_012_new.txt", true, _encoding);

                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    //카드구분 : 인편,등기
                    strBranch = dtable.Rows[i]["card_branch"].ToString();
                    strZipcode = dtable.Rows[i]["card_zipcode"].ToString();

                    strZipcode_kind = dtable.Rows[i]["card_zipcode_kind"].ToString();
                    strZipcode_new = dtable.Rows[i]["card_zipcode_new"].ToString();

                    strClient_send_number = dtable.Rows[i]["client_send_number"].ToString();

                    //데이터생성 시작
                    _strLine = new StringBuilder(dtable.Rows[i]["card_barcode_new"].ToString().Trim() + ",");

                    if (strZipcode_kind == "1")
                    {
                        _strLine.Append(GetStringAsLength(strZipcode_new, 5, true, ' '));

                        //등기구분
                        if (strBranch == "012")
                        {
                            _sw02_new.WriteLine(_strLine.ToString());
                        }
                        else
                        {
                            _sw01_new.WriteLine(_strLine.ToString());
                        }
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength(strZipcode, 5, true, ' '));

                        //등기구분
                        if (strBranch == "012")
                        {
                            _sw02.WriteLine(_strLine.ToString());
                        }
                        else
                        {
                            _sw01.WriteLine(_strLine.ToString());
                        }
                    }

                    

                    if (strZipcode_kind == "1")
                    {
                        _strLine = new StringBuilder("2" + strClient_send_number + "B002" + "30" + strZipcode_new + ",");
                        _strLine.Append(GetStringAsLength(strZipcode_new, 5, true, ' '));

                        //등기구분
                        if (strBranch == "012")
                        {
                            _sw04_new.WriteLine(_strLine.ToString());
                        }
                        else
                        {
                            _sw03_new.WriteLine(_strLine.ToString());
                        }
                    }
                    else
                    {
                        _strLine = new StringBuilder("2" + strClient_send_number + "B002" + "30" + strZipcode + ",");
                        _strLine.Append(GetStringAsLength(strZipcode, 5, true, ' '));

                        //등기구분
                        if (strBranch == "012")
                        {
                            _sw04.WriteLine(_strLine.ToString());
                        }
                        else
                        {
                            _sw03.WriteLine(_strLine.ToString());
                        }
                    }
                }
                _strReturn = string.Format("{0}건의 NPI데이타 다운 완료", i);
            }
            catch (Exception)
            {
                _strReturn = string.Format("{0}번째 데이터 인계 중 오류 발생", i + 1);
            }
            finally
            {
                if (_sw01 != null) _sw01.Close();
                if (_sw02 != null) _sw02.Close();

                if (_sw01_new != null) _sw01_new.Close();
                if (_sw02_new != null) _sw02_new.Close();

                if (_sw03 != null) _sw03.Close();
                if (_sw04 != null) _sw04.Close();

                if (_sw03_new != null) _sw03_new.Close();
                if (_sw04_new != null) _sw04_new.Close();
            }
            return _strReturn;
        }

        //농협
        private static string ConvertReceiveType7(System.Data.DataTable dtable, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩
            StreamWriter _sw01 = null, _sw02 = null;
            StringBuilder _strLine = new StringBuilder("");
            string _strReturn = "", strBranch = "", strZipcode = "";
            string strZipcode_kind = "", strZipcode_new = "";

            int i = 0;
            try
            {
                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    strZipcode_kind = dtable.Rows[i]["card_zipcode_kind"].ToString();
                    strZipcode_new = dtable.Rows[i]["card_zipcode_new"].ToString();
                    strBranch = dtable.Rows[i]["card_branch"].ToString();
                    //데이터생성 시작
                    _strLine = new StringBuilder(GetStringAsLength(dtable.Rows[i]["card_barcode_new"].ToString(), 22, true, ' ').Trim() + ",");

                    if (strZipcode_kind == "1")
                    {
                        if (strBranch == "012")
                        {
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_zipcode_new"].ToString(), 5, true, ' '));
                            _sw01 = new StreamWriter(fileName + "NH_012_new.txt", true, _encoding);
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_zipcode_new"].ToString(), 5, true, ' '));
                            _sw01 = new StreamWriter(fileName + "NH_new.txt", true, _encoding);
                        }
                    }
                    else
                    {
                        if (strBranch == "012")
                        {
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_zipcode"].ToString(), 5, true, ' '));
                            _sw01 = new StreamWriter(fileName + "NH_012_old.txt", true, _encoding);
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_zipcode"].ToString(), 5, true, ' '));
                            _sw01 = new StreamWriter(fileName + "NH_old.txt", true, _encoding);
                        }
                    }
                    _sw01.WriteLine(_strLine.ToString());
                    _sw01.Close();
                    
                }
                _strReturn = string.Format("{0}건의 NPI데이타 다운 완료", i);
            }
            catch (Exception)
            {
                _strReturn = string.Format("{0}번째 데이터 인계 중 오류 발생", i + 1);
            }
            finally
            {
                if (_sw01 != null) _sw01.Close();
            }
            return _strReturn;
        }

        //삼성
        private static string ConvertReceiveType8(System.Data.DataTable dtable, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩
            StreamWriter _sw01 = null, _sw02 = null;
            StringBuilder _strLine = new StringBuilder("");
            string _strReturn = "", strBranch = "", strZipcode = "";
            string strZipcode_kind = "", strZipcode_new = "";

            int i = 0;
            try
            {
                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    strZipcode_kind = dtable.Rows[i]["card_zipcode_kind"].ToString();
                    strZipcode_new = dtable.Rows[i]["card_zipcode_new"].ToString();
                    //데이터생성 시작
                    _strLine = new StringBuilder(dtable.Rows[i]["card_barcode_new"].ToString().Trim() + ",");

                    if (strZipcode_kind == "1")
                    {
                        _strLine.Append(GetStringAsLength(strZipcode_new, 5, true, ' '));
                        _sw01 = new StreamWriter(fileName + "SM_new.txt", true, _encoding);
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength(strZipcode, 5, true, ' '));
                        _sw01 = new StreamWriter(fileName + "SM_old.txt", true, _encoding);
                    }
                    _sw01.WriteLine(_strLine.ToString());
                    _sw01.Close();

                }
                _strReturn = string.Format("{0}건의 NPI데이타 다운 완료", i);
            }
            catch (Exception)
            {
                _strReturn = string.Format("{0}번째 데이터 인계 중 오류 발생", i + 1);
            }
            finally
            {
                if (_sw01 != null) _sw01.Close();
            }
            return _strReturn;
        }

        //인터파크
        private static string ConvertReceiveType10(System.Data.DataTable dtable, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩
            StreamWriter _sw01 = null, _sw02 = null;
            StringBuilder _strLine = new StringBuilder("");
            string _strReturn = "", strBranch = "", strZipcode = "";
            string strZipcode_kind = "", strZipcode_new = "";

            int i = 0;
            try
            {
                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    strZipcode_kind = dtable.Rows[i]["card_zipcode_kind"].ToString();
                    strZipcode_new = dtable.Rows[i]["card_zipcode_new"].ToString();
                    strBranch = dtable.Rows[i]["card_branch"].ToString();
                    //데이터생성 시작
                    if (strBranch == "012")
                    {
                        _strLine = new StringBuilder(dtable.Rows[i]["card_barcode_new"].ToString().Trim() + "," + strZipcode_new);
                        _sw01 = new StreamWriter(fileName + "INTER_NPI_012.txt", true, _encoding);
                    }
                    else
                    {
                        _strLine = new StringBuilder(dtable.Rows[i]["card_barcode_new"].ToString().Trim() + "," + strZipcode_new);
                        _sw01 = new StreamWriter(fileName + "INTER_NPI.txt", true, _encoding);
                    }

                    _sw01.WriteLine(_strLine.ToString());
                    _sw01.Close();

                }
                _strReturn = string.Format("{0}건의 NPI데이타 다운 완료", i);
            }
            catch (Exception)
            {
                _strReturn = string.Format("{0}번째 데이터 인계 중 오류 발생", i + 1);
            }
            finally
            {
                if (_sw01 != null) _sw01.Close();
            }
            return _strReturn;
        }




        #region 기타함수
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
        #endregion

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
