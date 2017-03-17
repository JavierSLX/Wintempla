#pragma once  //______________________________________ PruebaExcel.h  
#include "Resource.h"
#include "Excel.h"
class PruebaExcel : public Win::Dialog
{
public:
	PruebaExcel()
	{
	}
	~PruebaExcel()
	{
	}
	//Objetos de Excel
	Com::Boot boot; //Internamente llama ::CoInitialize and ::CoUninitialize
	Excel::ApplicationX Aplicacion; //Crea el objeto para la aplicación
	Excel::Workbooks Libros;
	Excel::WorkbookX Libro;
	Excel::WorksheetX Hoja;
	Excel::Range Rango;
	Excel::Range Celdas;
	Excel::Range celda;
	Excel::Font fuente;
protected:
	//______ Wintempla GUI manager section begin: DO NOT EDIT AFTER THIS LINE
	Win::Toolbar toolbMenu;
	Win::Textbox tbx1;
protected:
	Win::Gdi::Font fontArial009A;
	void GetDialogTemplate(DLGTEMPLATE& dlgTemplate)
	{
		dlgTemplate.cx=Sys::Convert::CentimetersToDlgUnitX(9.23396);
		dlgTemplate.cy=Sys::Convert::CentimetersToDlgUnitY(3.17500);
		dlgTemplate.style = WS_CAPTION | WS_POPUP | WS_SYSMENU | WS_VISIBLE | DS_CENTER | DS_MODALFRAME;
	}
	//_________________________________________________
	void InitializeGui()
	{
		this->Text = L"PruebaExcel";
		toolbMenu.CreateX(NULL, NULL, WS_CHILD | WS_VISIBLE | CCS_NORESIZE | CCS_NOPARENTALIGN | CCS_ADJUSTABLE | CCS_NODIVIDER | TBSTYLE_FLAT | TBSTYLE_TOOLTIPS, 0.55563, 0.42333, 8.49313, 1.11125, hWnd, 1000);
		tbx1.CreateX(WS_EX_CLIENTEDGE, NULL, WS_CHILD | WS_TABSTOP | WS_VISIBLE | ES_AUTOHSCROLL | ES_LEFT | ES_WINNORMALCASE, 0.60854, 1.98438, 8.44021, 1.00542, hWnd, 1001);
		fontArial009A.CreateX(L"Arial", 0.317500, false, false, false, false);
		tbx1.Font = fontArial009A;
	}
	//_________________________________________________
	void Window_Open(Win::Event& e);
	void Cmd_Add(Win::Event& e);
	void Cmd_Msexcel(Win::Event& e);
	//_________________________________________________
	bool EventHandler(Win::Event& e)
	{
		if (this->IsEvent(e, IDM_ADD)) {Cmd_Add(e); return true;}
		if (this->IsEvent(e, IDM_MSEXCEL)) {Cmd_Msexcel(e); return true;}
		return false;
	}
};
