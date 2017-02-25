// Connect.cpp : Implementation of CConnect
#include "stdafx.h"
#include "AddIn.h"
#include "Connect.h"

extern CAddInModule _AtlModule;


HHOOK g_mouseHook;
Excel::_ApplicationPtr g_pExcelApp;

LRESULT CALLBACK MouseHookProc(int nCode, WPARAM wMsg, LPARAM lParam)
{
	if ((wMsg==WM_MOUSEWHEEL) && HIWORD(GetKeyState(VK_SHIFT)))
	{
//???stack unwind model (/EHsc ?)
try {
		short zDelta=GET_WHEEL_DELTA_WPARAM(((LPMOUSEHOOKSTRUCTEX)lParam)->mouseData);

		Excel::WindowPtr pActiveWindow=g_pExcelApp->GetActiveWindow();//throws if NULL
		ATLTRACE("pActiveWindow=%x\n",pActiveWindow.GetInterfacePtr());
		_variant_t vtl((long)1);
		if (zDelta < 0) pActiveWindow->SmallScroll(vtMissing, vtMissing, &vtl, vtMissing);
		else pActiveWindow->SmallScroll(vtMissing, vtMissing, vtMissing, &vtl);

} catch (...){ //(_com_error &err) {
		//ATLTRACE(L"MouseHookProc error.message: %s\n", (LPCWSTR)err.ErrorMessage());//?err hex syntax
        //ATLTRACE(L"MouseHookProc error.description: %s\n", (LPCWSTR)err.Description());
		ATLTRACE("MouseHookProc exception\n");
		}
		
		/*we must not pass this message to hook chain if there's no opened workbook,
		WM_MOUSEWHEEL+Shift crashes Excel
		Excel installs it's own application-level hooks (WH_MSGFILTER, WH_KEYBOARD, WH_CBT)
		This bug is present in Excel 97, 2000, 2002 and 2003 */
		return 1;
	}
return CallNextHookEx(g_mouseHook, nCode, wMsg, lParam);
}


// CConnect
HRESULT CConnect::FinalConstruct()
{
	g_mouseHook=NULL;
	return S_OK;
}

void CConnect::FinalRelease() 
{
	ATLASSERT(g_mouseHook==NULL);
	ATLASSERT(g_pExcelApp.GetInterfacePtr()==NULL);
}

STDMETHODIMP CConnect::OnConnection(IDispatch *pApplication, AddInDesignerObjects::ext_ConnectMode ConnectMode, IDispatch *pAddInInst, SAFEARRAY ** /*custom*/ )
{
	UNREFERENCED_PARAMETER(pAddInInst);
	UNREFERENCED_PARAMETER(ConnectMode);
	ATLTRACE("CConnect::OnConnection\n");
	HRESULT hr=S_OK;
try {
		g_pExcelApp=pApplication;
		ATLTRACE("g_pExcelApp=%x\n",g_pExcelApp.GetInterfacePtr());
} catch (_com_error &err) {
		UNREFERENCED_PARAMETER(err);
		ATLTRACE(L"CConnect::OnConnection error %x  %s\n", err.Error(), (LPCWSTR)err.ErrorMessage());
        ATLTRACE(L"Description: %s\n", (LPCWSTR)err.Description());
	}

	if SUCCEEDED(hr)
	{
		g_mouseHook=SetWindowsHookEx(WH_MOUSE, MouseHookProc, (HINSTANCE)&__ImageBase, GetCurrentThreadId());
		ATLTRACE("g_mouseHook=%x\n", g_mouseHook);	
	}
return hr;
}

STDMETHODIMP CConnect::OnDisconnection(AddInDesignerObjects::ext_DisconnectMode /*RemoveMode*/, SAFEARRAY ** /*custom*/ )
{
	ATLTRACE("CConnect::OnDisconnection\n");
	ATLASSERT(g_mouseHook);
	if (g_mouseHook)
	{
		UnhookWindowsHookEx(g_mouseHook);
		g_mouseHook=NULL;
		ATLTRACE("mouse unhooked\n");
	}
	ATLASSERT(g_pExcelApp.GetInterfacePtr());
	g_pExcelApp.Release();
return S_OK;
}
