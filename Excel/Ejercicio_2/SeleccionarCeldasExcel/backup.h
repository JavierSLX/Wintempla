#pragma once  //______________________________________ SeleccionarCeldasExcel.h  
#include "Resource.h"
class SeleccionarCeldasExcel: public Win::Dialog
{
public:
	SeleccionarCeldasExcel()
	{
		::CoInitialize(NULL);
	}
	~SeleccionarCeldasExcel()
	{
		::CoUninitialize();
	}
protected:
	//______ Wintempla GUI manager section begin: DO NOT EDIT AFTER THIS LINE
protected:
	void GetDialogTemplate(DLGTEMPLATE& dlgTemplate)
	{
		dlgTemplate.style = WS_CAPTION | WS_POPUP | WS_SYSMENU | WS_VISIBLE | DS_CENTER | DS_MODALFRAME;
	}
	//_________________________________________________
	void InitializeGui()
	{
		this->Text = L"SeleccionarCeldasExcel";
	}
	//_________________________________________________
	void Window_Open(Win::Event& e);
	//_________________________________________________
	bool EventHandler(Win::Event& e)
	{
		return false;
	}
};
