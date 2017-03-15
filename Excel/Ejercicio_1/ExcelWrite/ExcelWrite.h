#pragma once  //______________________________________ ExcelWrite.h  
#include "Resource.h"
#include "Excel.h"
class ExcelWrite: public Win::Dialog
{
public:
	ExcelWrite()
	{
		::CoInitialize(NULL);
	}
	~ExcelWrite()
	{
		::CoUninitialize();
	}
protected:
	//______ Wintempla GUI manager section begin: DO NOT EDIT AFTER THIS LINE
	void InitializeGui()
	{
		this->Text = L"ExcelWrite";
	}
	void Window_Open(Win::Event& e);
	void GetDialogTemplate(DLGTEMPLATE& dlgTemplate)
	{
		dlgTemplate.style = DS_CENTER | DS_MODALFRAME | WS_POPUP | WS_VISIBLE | WS_CAPTION | WS_SYSMENU;
	}
	bool EventHandler(Win::Event& e)
	{
		return false;
	}
};
