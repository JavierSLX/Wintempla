#include "stdafx.h"  //________________________________________ ExcelWrite.cpp
#include "ExcelWrite.h"

int APIENTRY wWinMain(HINSTANCE hInstance, HINSTANCE , LPTSTR cmdLine, int cmdShow){
	ExcelWrite app;
	return app.BeginDialog(IDI_ExcelWrite, hInstance);
}

void ExcelWrite::Window_Open(Win::Event& e)
{
	//Crea un objeto tipo COM
	Com::Object aplicacion;

	try
	{
		//Crea el documento
		aplicacion.CreateInstance(L"Excel.Application", true);
		aplicacion.Put(L"Visible", false); //No se vea la creación de Excel

		//Obtiene los libros
		Com::Object libros;
		aplicacion.Get(L"Workbooks", libros);

		//Agrega un libro
		Com::Object libro;
		libros.Method(L"Add", (long)-4167, libro);

		//Activa la pestaña del libro
		Com::Object pestania;
		aplicacion.Get(L"ActiveSheet", pestania);

		//Da el nombre a la pestaña
		pestania.Put(L"Name", L"Accounting");

		//Saca el rango de la celda
		Com::Object rango;
		pestania.Get(L"Range", L"A1", rango);

		//Pone el valor a la celda A1
		rango.Put(L"Value2", L"Hola Mundo!");

		_variant_t resultado;
		libro.Method(L"SaveAs",
			L".\\info.xlsx",
			(short)51,
			L"",
			L"",
			false,
			false,
			true,
			resultado);

		//Termina la ejecución
		aplicacion.Method(L"Quit");
	}
	catch (Com::Exception excep)
	{
		excep.Display(hWnd, L"ExcelWrite");
		aplicacion.Method(L"Quit");
	}
}

