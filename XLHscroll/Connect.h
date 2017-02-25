#pragma once
#include "resource.h"       // main symbols
//include order!!!
#include "mso.tlh"
using namespace Office;
#include "msaddndr.tlh"
//using namespace AddInDesignerObjects;
#include "vbe6ext.tlh"
using namespace VBIDE;
#include "excel.tlh"
//using namespace Excel;

/// <summary>The object for implementing an Add-in.</summary>
/// <seealso class='IDTExtensibility2' />
class ATL_NO_VTABLE CConnect : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CConnect, &CLSID_Connect>,
	public IDispatchImpl<AddInDesignerObjects::_IDTExtensibility2, &AddInDesignerObjects::IID__IDTExtensibility2,
							&AddInDesignerObjects::LIBID_AddInDesignerObjects, 1, 0>
	//public IDispatchImpl<CConnect, __uuidof(CConnect)/* &IID_IConnect*/, &LIBID_XLHScrollLib, 1, 0>
{
public:
	/// <summary>Implements the constructor for the Add-in object. Place your initialization code within this method.</summary>
	CConnect()
	{
	}

	DECLARE_REGISTRY_RESOURCEID(IDR_ADDIN)
	DECLARE_NOT_AGGREGATABLE(CConnect)

	BEGIN_COM_MAP(CConnect)
		COM_INTERFACE_ENTRY(IDispatch)
		COM_INTERFACE_ENTRY(AddInDesignerObjects::IDTExtensibility2)
	END_COM_MAP()

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct();
	void FinalRelease();

public:
//IDTExtensibility2 implementation:

	/// <summary>Implements the OnConnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being loaded.</summary>
	/// <param term='application'>Root object of the host application.</param>
	/// <param term='connectMode'>Describes how the Add-in is being loaded.</param>
	/// <param term='addInInst'>Object representing this Add-in.</param>
	/// <seealso class='IDTExtensibility2' />
	STDMETHOD(OnConnection)(IDispatch * Application, AddInDesignerObjects::ext_ConnectMode ConnectMode, IDispatch *AddInInst, SAFEARRAY **custom);

	/// <summary>Implements the OnDisconnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being unloaded.</summary>
	/// <param term='disconnectMode'>Describes how the Add-in is being unloaded.</param>
	/// <param term='custom'>Array of parameters that are host application specific.</param>
	/// <seealso class='IDTExtensibility2' />
	STDMETHOD(OnDisconnection)(AddInDesignerObjects::ext_DisconnectMode RemoveMode, SAFEARRAY **custom );

	/// <summary>Implements the OnAddInsUpdate method of the IDTExtensibility2 interface. Receives notification when the collection of Add-ins has changed.</summary>
	/// <param term='custom'>Array of parameters that are host application specific.</param>
	/// <seealso class='IDTExtensibility2' />	
	STDMETHOD(OnAddInsUpdate)(SAFEARRAY **custom )
	{
		UNREFERENCED_PARAMETER(custom);
		ATLTRACE("CConnect::OnAddInsUpdate\n");
		return S_OK;
	}

	/// <summary>Implements the OnStartupComplete method of the IDTExtensibility2 interface. Receives notification that the host application has completed loading.</summary>
	/// <param term='custom'>Array of parameters that are host application specific.</param>
	/// <seealso class='IDTExtensibility2' />
	STDMETHOD(OnStartupComplete)(SAFEARRAY **custom )
	{
		UNREFERENCED_PARAMETER(custom);
		ATLTRACE("CConnect::OnStartupComplete\n");
		return S_OK;
	}

	/// <summary>Implements the OnBeginShutdown method of the IDTExtensibility2 interface. Receives notification that the host application is being unloaded.</summary>
	/// <param term='custom'>Array of parameters that are host application specific.</param>
	/// <seealso class='IDTExtensibility2' />
	STDMETHOD(OnBeginShutdown)(SAFEARRAY **custom )
	{
		UNREFERENCED_PARAMETER(custom);
		ATLTRACE("CConnect::OnBeginShutdown\n");
		return S_OK;
	}

//?local var
private:
	//CComPtr<IDispatch> m_pExcelApp;
//	CComPtr<Excel::Window> m_pActiveWindow; //check katmouse hwnd classes! //EXCEL7   mso2010/13???
};

OBJECT_ENTRY_AUTO(__uuidof(Connect), CConnect)
