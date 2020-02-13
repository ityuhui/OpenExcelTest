// OpenExcelTest.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include "pch.h"
#include <iostream>
#include <string>

#import "C:\\Program Files\\Microsoft Office\\root\\VFS\\ProgramFilesCommonX64\\Microsoft Shared\\OFFICE16\\MSO.DLL" \
rename("RGB", "MsoRGB") \
rename("DocumentProperties", "MsoDocumentProperties") \
rename("SearchPath","MsoSearchPath")

#import "C:\\Program Files\\Microsoft Office\\root\\VFS\\ProgramFilesCommonX86\\Microsoft Shared\\VBA\\VBA6\\VBE6EXT.OLB"

#import "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE" \
rename( "DialogBox", "ExcelDialogBox" ) \
rename( "RGB", "ExcelRGB" ) \
rename( "CopyFile", "ExcelCopyFile" ) \
rename( "ReplaceText", "ExcelReplaceText" ) \
exclude( "IFont", "IPicture" ) no_dual_interfaces

Excel::_ApplicationPtr pExcelApp;
Excel::_WorkbookPtr pWorkbook;


std::string GetComErrorText(_com_error &e)
{
    std::string sRetString;
    std::string sDescr;
    std::string sErrMessage;
    char * pTmp = 0;

    _bstr_t bstrDescription = e.Description();
    try {
        if ((char *)bstrDescription == 0) {
            sDescr = "_com_error.Description() returned: <NULL>.";
        } else {
            pTmp = (char *)calloc(1, bstrDescription.length() + 128);
            sprintf_s(pTmp, bstrDescription.length() + 128, "%s", static_cast<LPTSTR> (bstrDescription));
            sDescr = std::string(pTmp);
        }
    }
    catch (...) {
        sDescr = " Failed to get _com_error.Description().";
    }

    free(pTmp);

    const TCHAR * errMsg = e.ErrorMessage();
    try {
        if ((char *)errMsg == 0) {
            sErrMessage = "_com_error.ErrorMessage() returned: <NULL>.";
        } else {
            pTmp = (char *)calloc(1, strlen(errMsg) + 128);
            sprintf_s(pTmp, bstrDescription.length() + 128, "%s", errMsg);
            sErrMessage = std::string(pTmp);
        }
    }
    catch (...) {
        sErrMessage = "Failed to get _com_error.ErrorMessage().";
    }
    free(pTmp);

    sRetString = sDescr + " " + sErrMessage;

    return sRetString;
}   // GetComErrorText

int main()
{

        CoInitializeEx(NULL, COINIT_MULTITHREADED);

        HRESULT hr = pExcelApp.CreateInstance(L"Excel.Application");

        pExcelApp->Visible = false;   // make Excel’s main window visible


        pWorkbook = pExcelApp->Workbooks->Open("C:\\20200211\\ConvertiblePricing_Complete_migrate_Symphony.xlsb");  // open excel file

        /*
        std::string sCaption = "hello";
        BSTR Caption = _bstr_t(sCaption.c_str()).copy();
        pExcelApp->PutCaption(Caption);
        

        std::string theFinalCaption = "Microsoft Excel - " + sCaption;
        HWND m_hdlExcelHandle = FindWindow("XLMAIN",NULL);

        unsigned long m_ExcelProcessId = 0;

        DWORD dwTheardId = GetWindowThreadProcessId(m_hdlExcelHandle, &m_ExcelProcessId);


        std::cout << m_ExcelProcessId << std::endl;
        */

        std::string sPath = "C:\\";
        std::string sMacroName = "MyMacro";
        std::string sParam = "2,2,2";
        BSTR RealPath = _bstr_t(sPath.c_str()).copy();
        BSTR Macro = _bstr_t(sMacroName.c_str()).copy();
        BSTR Param = _bstr_t(sParam.c_str()).copy();


        try {


            _variant_t qResult = pExcelApp->Run(Macro, Param, RealPath);
            std::string pResult = ((char*)_bstr_t(qResult.bstrVal));

            std::cout << pResult << std::endl;
        }
        catch (_com_error &ex) {
            std::cout << "cannot run macro " << GetComErrorText(ex) << std::endl;
            return -1;
        }

        sParam = "3,3,2";
        Param = _bstr_t(sParam.c_str()).copy();

        try {
            _variant_t qResult = pExcelApp->Run(Macro, Param, RealPath);
            std::string pResult = ((char*)_bstr_t(qResult.bstrVal));

            std::cout << pResult << std::endl;
        }
        catch (_com_error &ex) {
            std::cout << "cannot run macro 2 " << GetComErrorText(ex) << std::endl;
            return -1;
        }
        
        
        try {
            pWorkbook->Close(VARIANT_FALSE);  // save changes
        }
        catch (_com_error &ex ){
            std::cout << "cannot close the webook" << GetComErrorText(ex) << std::endl;
            return -1;
        }
        std::cout << "closed the webook" << std::endl;
        
        try {
            pWorkbook->Close(VARIANT_FALSE);  // save changes
        }
        catch (_com_error &ex) {
            std::cout << "cannot close the webook 1 " << GetComErrorText(ex) << std::endl;
            return -1;
        }
        std::cout << "closed the webook 1" << std::endl;

        try {
            pExcelApp->Quit();
        }
        catch (_com_error &ex) {
            std::cout << "cannot quit the excel" << std::endl;
            return -1;
        }


        std::cout << "Excel quited" << std::endl;
}

// Run program: Ctrl + F5 or Debug > Start Without Debugging menu
// Debug program: F5 or Debug > Start Debugging menu

// Tips for Getting Started: 
//   1. Use the Solution Explorer window to add/manage files
//   2. Use the Team Explorer window to connect to source control
//   3. Use the Output window to see build output and other messages
//   4. Use the Error List window to view errors
//   5. Go to Project > Add New Item to create new code files, or Project > Add Existing Item to add existing code files to the project
//   6. In the future, to open this project again, go to File > Open > Project and select the .sln file
