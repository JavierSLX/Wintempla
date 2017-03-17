#include "stdafx.h"  //________________________________________ PruebaExcel.cpp
#include "PruebaExcel.h"

int APIENTRY wWinMain(HINSTANCE hInstance, HINSTANCE, LPTSTR cmdLine, int cmdShow) {
	PruebaExcel app;
	return app.BeginDialog(IDI_PruebaExcel, hInstance);
}

void PruebaExcel::Window_Open(Win::Event& e)
{

	//________________________________________________________ toolbMenu
	TBBUTTON tbButton[9];//<< EDIT HERE THE NUMBER OF BUTTONS

	const int pixelsIconSize = Sys::Convert::CentimetersToScreenPixels(0.42333);
	const int pixelsButtonSize = pixelsIconSize + Sys::Convert::CentimetersToScreenPixels(0.2);
	toolbMenu.imageList.Create(pixelsIconSize, pixelsIconSize, 7);//<< EDIT HERE THE NUMBER OF IMAGES
	toolbMenu.imageList.AddIcon(this->hInstance, IDI_SAVE);
	toolbMenu.imageList.AddIcon(this->hInstance, IDI_COPY);
	toolbMenu.imageList.AddIcon(this->hInstance, IDI_ADD);
	toolbMenu.imageList.AddIcon(this->hInstance, IDI_EDIT);
	toolbMenu.imageList.AddIcon(this->hInstance, IDI_DELETE);
	toolbMenu.imageList.AddIcon(this->hInstance, IDI_PRINT);
	toolbMenu.imageList.AddIcon(this->hInstance, IDI_MSEXCEL);
	toolbMenu.SendMessage(TB_BUTTONSTRUCTSIZE, (WPARAM)(int)sizeof(TBBUTTON), 0);
	toolbMenu.SetImageList(toolbMenu.imageList);
	//_____________________________________
	tbButton[0].iBitmap = MAKELONG(0, 0); //<< IMAGE INDEX
	tbButton[0].idCommand = IDM_SAVE;
	tbButton[0].fsState = TBSTATE_ENABLED; // | TBSTATE_WRAP
	tbButton[0].fsStyle = BTNS_BUTTON;
	tbButton[0].dwData = 0L;
	tbButton[0].iString = (LONG_PTR)L"Save";
	//________________________ A separator
	tbButton[1].iBitmap = -1;
	tbButton[1].idCommand = 0;
	tbButton[1].fsState = TBSTATE_ENABLED; // | TBSTATE_WRAP
	tbButton[1].fsStyle = BTNS_SEP;
	tbButton[1].dwData = 0L;
	tbButton[1].iString = 0;
	//_____________________________________
	tbButton[2].iBitmap = MAKELONG(1, 0); //<< IMAGE INDEX
	tbButton[2].idCommand = IDM_COPY;
	tbButton[2].fsState = TBSTATE_ENABLED; // | TBSTATE_WRAP
	tbButton[2].fsStyle = BTNS_BUTTON;
	tbButton[2].dwData = 0L;
	tbButton[2].iString = (LONG_PTR)L"Copy";
	//_____________________________________
	tbButton[3].iBitmap = MAKELONG(2, 0); //<< IMAGE INDEX
	tbButton[3].idCommand = IDM_ADD;
	tbButton[3].fsState = TBSTATE_ENABLED; // | TBSTATE_WRAP
	tbButton[3].fsStyle = BTNS_BUTTON;
	tbButton[3].dwData = 0L;
	tbButton[3].iString = (LONG_PTR)L"Add";
	//_____________________________________
	tbButton[4].iBitmap = MAKELONG(3, 0); //<< IMAGE INDEX
	tbButton[4].idCommand = IDM_EDIT;
	tbButton[4].fsState = TBSTATE_ENABLED; // | TBSTATE_WRAP
	tbButton[4].fsStyle = BTNS_BUTTON;
	tbButton[4].dwData = 0L;
	tbButton[4].iString = (LONG_PTR)L"Edit";
	//_____________________________________
	tbButton[5].iBitmap = MAKELONG(4, 0); //<< IMAGE INDEX
	tbButton[5].idCommand = IDM_DELETE;
	tbButton[5].fsState = TBSTATE_ENABLED; // | TBSTATE_WRAP
	tbButton[5].fsStyle = BTNS_BUTTON;
	tbButton[5].dwData = 0L;
	tbButton[5].iString = (LONG_PTR)L"Delete";
	//________________________ A separator
	tbButton[6].iBitmap = -1;
	tbButton[6].idCommand = 0;
	tbButton[6].fsState = TBSTATE_ENABLED; // | TBSTATE_WRAP
	tbButton[6].fsStyle = BTNS_SEP;
	tbButton[6].dwData = 0L;
	tbButton[6].iString = 0;
	//_____________________________________
	tbButton[7].iBitmap = MAKELONG(5, 0); //<< IMAGE INDEX
	tbButton[7].idCommand = IDM_PRINT;
	tbButton[7].fsState = TBSTATE_ENABLED; // | TBSTATE_WRAP
	tbButton[7].fsStyle = BTNS_BUTTON;
	tbButton[7].dwData = 0L;
	tbButton[7].iString = (LONG_PTR)L"Print";
	//_____________________________________
	tbButton[8].iBitmap = MAKELONG(6, 0); //<< IMAGE INDEX
	tbButton[8].idCommand = IDM_MSEXCEL;
	tbButton[8].fsState = TBSTATE_ENABLED; // | TBSTATE_WRAP
	tbButton[8].fsStyle = BTNS_BUTTON;
	tbButton[8].dwData = 0L;
	tbButton[8].iString = (LONG_PTR)L"Export to Microsoft Excel";

	toolbMenu.SetBitmapSize(pixelsIconSize, pixelsIconSize);
	toolbMenu.SetButtonSize(pixelsButtonSize, pixelsButtonSize);
	toolbMenu.AddButtons(tbButton, 9);// << EDIT HERE THE NUMBER OF BUTTONS
	toolbMenu.SendMessage(TB_AUTOSIZE, 0, 0);
	toolbMenu.SetMaxTextRows(0);// EDIT HERE TO DISPLAY THE BUTTON TEXT
	toolbMenu.Show(SW_SHOWNORMAL);
	//	toolbMenu.ResizeToFit();

}

//Abre un Excel
void PruebaExcel::Cmd_Msexcel(Win::Event& e)
{
	wstring direccion;

	Win::FileDlg dlg;
	dlg.Clear();
	dlg.SetFilter(L"Excel files (*.xlsx)\0*.xlsx\0\0", 0, L"xlsx");

	if (dlg.BeginDialog(hWnd, L"Open", false) == TRUE)
	{
		direccion = dlg.GetFileNameFullPath();
	}

	try
	{
		//Crea la aplicación 
		Aplicacion.CreateInstance(L"Excel.Application", true);

		//Pone el excel a la vista
		Aplicacion.Visible = true;

		//Agrega un objeto Libros y Libro
		Libros = Aplicacion.WorkbooksX;

		//Abre un libro
		Libro = Libros.Open(direccion);

		//Cuenta cuantas hojas de trabajo tiene el Excel
		Libros = Aplicacion.get_Sheets();

		//Saca la hoja de los libros existentes y elige la seleccionada
		Com::Object sheet = Libros.get_Item(L"MESN");
		sheet.Method(L"Select");

		//Activa la hoja de trabajo para poder realizar los procesos
		Hoja = Aplicacion.ActiveSheet;

		//Pone un rango
		Rango = Hoja.get_Range(L"D4");
		Rango.Select();
		_variant_t valor = Rango.get_Value2();

		//Pone el valor en la caja de texto
		//tbxValor.Text = valor.bstrVal; //Cadenas string
		/*Sys::SqlTime fecha;
		Sys::Convert::SysTimeToSqlTime(Sys::Convert::VariantToTime(valor), fecha);
		wstring cadena;
		Sys::Format(cadena, L"%d/%d/%d", fecha.day, fecha.month, fecha.year);
		tbxValor.Text = cadena;
		tbxValor.Text = valor.bstrVal;*/
	}
	catch (Com::Exception excep)
	{
		wchar_t text[1024];
		excep.GetErrorText(text, 1024);
		this->MessageBoxW(text, L"Meses", MB_OK | MB_ICONERROR);
		Aplicacion.Quit();
	}
}

void PruebaExcel::Cmd_Add(Win::Event& e)
{
	try
	{
		//Crea la aplicación 
		Aplicacion.CreateInstance(L"Excel.Application", true);

		//Pone el excel a la vista
		Aplicacion.Visible = true;

		//Agrega un objeto Libros y Libro
		Libros = Aplicacion.WorkbooksX;
		Libro = Libros.Add(Excel::XlSheetType::xlWorksheet);

		//Pone nombre a la pestaña
		Hoja = Aplicacion.ActiveSheet;
		Hoja.NameX = L"Months of the Year";

		//Ponen el valor a la celda A1
		Rango = Hoja.get_Range(L"A1");
		Rango.Value2 = L"Enero";

		//Coloca en forma de celda
		Celdas = Hoja.Cells;
		Celdas.get_Item(1, 1);
		_variant_t elemento = Celdas.get_Item(2, 1);
		celda = elemento;
		celda.put_Value2(L"Febrero");

		//Cambia de color
		Rango = Hoja.get_Range(L"A3");
		Rango.Value2 = L"Marzo";
		fuente = Rango.FontX;
		fuente.ColorIndex = 3;
	}
	catch (Com::Exception excep)
	{
		wchar_t text[1024];
		excep.GetErrorText(text, 1024);
		this->MessageBoxW(text, L"Meses", MB_OK | MB_ICONERROR);
		Aplicacion.Quit();
	}
}

