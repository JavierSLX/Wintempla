#include "stdafx.h"  //________________________________________ SeleccionarCeldasExcel.cpp
#include "SeleccionarCeldasExcel.h"

int APIENTRY wWinMain(HINSTANCE hInstance, HINSTANCE , LPTSTR cmdLine, int cmdShow){
	SeleccionarCeldasExcel app;
	return app.BeginDialog(IDI_SeleccionarCeldasExcel, hInstance);
}

void SeleccionarCeldasExcel::Window_Open(Win::Event& e)
{
	//Crea los objetos de COM de aplicación y rango
	Com::Object Aplicacion;
	Com::Object Rango;

	try
	{
		//Crea el Excel
		Aplicacion.CreateInstance(L"Excel.Application", true);
		Aplicacion.Put(L"Visible", true); //Lo hace visible

		//Crea los Libros
		Com::Object Libros;
		Aplicacion.Get(L"Workbooks", Libros);

		//Agrega un Libro
		Com::Object Libro;
		Libros.Method(L"Add", (long)-4167, Libro);

		//Activa una Pestaña de los Libros
		Com::Object Pestania;
		Aplicacion.Get(L"ActiveSheet", Pestania);

		//Coloca un valor a la celda A1
		Pestania.Get(L"Range", L"A1", Rango);
		Rango.Put(L"Value2", L"10");

		//Coloca un valor a la celda B1
		Pestania.Get(L"Range", L"B1", Rango);
		Rango.Put(L"Value2", L"20");

		//Coloca un valor a la celda C1
		Pestania.Get(L"Range", L"C1", Rango);
		Rango.Put(L"Value2", L"30");

		//Selecciona A1 a C1
		Pestania.Get(L"Range", L"A1", L"C1", Rango);
		Rango.Method(L"Select");
	}
	catch (Com::Exception excep)
	{
		excep.Display(hWnd, L"ExcelWrite");
		Aplicacion.Method(L"Quit");
	}
}

