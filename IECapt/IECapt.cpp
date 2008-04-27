////////////////////////////////////////////////////////////////////
//
// IECapt - A Internet Explorer Web Page Rendering Capture Utility
//
// Copyright (C) 2003-2008 Bjoern Hoehrmann <bjoern@hoehrmann.de>
//
// This program is free software; you can redistribute it and/or
// modify it under the terms of the GNU General Public License
// as published by the Free Software Foundation; either version 2
// of the License, or (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// $Id$
//
////////////////////////////////////////////////////////////////////

#define VC_EXTRALEAN
#include <stdlib.h>
#include <windows.h>
#include <mshtml.h>
#include <exdispid.h>
#include <atlbase.h>
#include <atlwin.h>
#include <atlcom.h>
#include <atlhost.h>
#include <atlimage.h>
#undef VC_EXTRALEAN

class CMain;
class CEventSink;

//////////////////////////////////////////////////////////////////
// CEventSink
//////////////////////////////////////////////////////////////////
class CEventSink :
    public CComObjectRootEx <CComSingleThreadModel>,
    public IDispatch
{
public:
    CEventSink() : m_pMain(NULL) {}

    BEGIN_COM_MAP(CEventSink)
        COM_INTERFACE_ENTRY(IDispatch)
        COM_INTERFACE_ENTRY_IID(DIID_DWebBrowserEvents2, IDispatch)
    END_COM_MAP()

    STDMETHOD(GetTypeInfoCount)(UINT* pctinfo);
    STDMETHOD(GetTypeInfo)(UINT itinfo, LCID lcid, ITypeInfo** pptinfo);
    STDMETHOD(GetIDsOfNames)(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgdispid);
    STDMETHOD(Invoke)(DISPID dispid, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pdispparams,
                      VARIANT* pvarResult, EXCEPINFO* pexcepinfo, UINT* puArgErr);

public:
    CMain* m_pMain;
};

//////////////////////////////////////////////////////////////////
// CMain
//////////////////////////////////////////////////////////////////
class CMain :
    public CWindowImpl <CMain>
{
public:

    CMain(LPTSTR uri, LPTSTR file, BOOL silent) :
      m_dwCookie(0), m_URI(uri), m_fileName(file), m_bSilent(silent) { }

    BEGIN_MSG_MAP(CMainWindow)
        MESSAGE_HANDLER(WM_CREATE,  OnCreate)
        MESSAGE_HANDLER(WM_SIZE,    OnSize)
        MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
    END_MSG_MAP()

    LRESULT OnCreate  (UINT nMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnSize    (UINT nMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnDestroy (UINT nMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);

    BOOL SaveSnapshot(IDispatch* pdisp, VARIANT* purl);

private:
    LPTSTR m_URI;
    LPTSTR m_fileName;
    BOOL   m_bSilent;

protected:
    CComPtr<IUnknown> m_pWebBrowserUnk;
    CComPtr<IWebBrowser2> m_pWebBrowser;
    CComObject<CEventSink>* m_pEventSink;
    HWND m_hwndWebBrowser;
    DWORD m_dwCookie;
};

//////////////////////////////////////////////////////////////////
// Implementation of CEventSink
//////////////////////////////////////////////////////////////////
STDMETHODIMP CEventSink::GetTypeInfoCount(UINT* pctinfo)
{
    return E_NOTIMPL;
}

STDMETHODIMP CEventSink::GetTypeInfo(UINT itinfo, LCID lcid,
                                     ITypeInfo** pptinfo)
{
    return E_NOTIMPL;
}

STDMETHODIMP CEventSink::GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames,
                                       UINT cNames, LCID lcid,
                                       DISPID* rgdispid)
{
    return E_NOTIMPL;
}

STDMETHODIMP CEventSink::Invoke(DISPID dispid, REFIID riid, LCID lcid,
                                WORD wFlags, DISPPARAMS* pdispparams,
                                VARIANT* pvarResult, EXCEPINFO* pexcepinfo,
                                UINT* puArgErr)
{
    if (dispid != DISPID_DOCUMENTCOMPLETE)
        return S_OK;

    if (pdispparams->cArgs != 2)
        return S_OK;

    if (pdispparams->rgvarg[0].vt != (VT_VARIANT | VT_BYREF))
        return S_OK;

    if (pdispparams->rgvarg[1].vt != VT_DISPATCH)
        return S_OK;

    if (m_pMain->SaveSnapshot(pdispparams->rgvarg[1].pdispVal,
                              pdispparams->rgvarg[0].pvarVal))
        m_pMain->PostMessage(WM_CLOSE);

    return S_OK;
}

//////////////////////////////////////////////////////////////////
// Implementation of CMain Messages
//////////////////////////////////////////////////////////////////
LRESULT CMain::OnCreate(UINT nMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
    HRESULT hr;
    RECT old;
    IUnknown * pUnk = NULL;
    GetClientRect(&old);

    m_hwndWebBrowser = ::CreateWindow(_T(ATLAXWIN_CLASS), m_URI,
        /*WS_POPUP|*/WS_CHILD|WS_DISABLED, old.top, old.left, old.right,
        old.bottom, m_hWnd, NULL, ::GetModuleHandle(NULL), NULL);

    hr = AtlAxGetControl(m_hwndWebBrowser, &m_pWebBrowserUnk);

    if (FAILED(hr))
        return 1;

    if (m_pWebBrowserUnk == NULL)
        return 1;

    hr = m_pWebBrowserUnk->QueryInterface(IID_IWebBrowser2, (void**)&m_pWebBrowser);

    if (FAILED(hr))
        return 1;

    // Set whether it should be silent
    m_pWebBrowser->put_Silent(m_bSilent ? VARIANT_TRUE : VARIANT_FALSE);

    hr = CComObject<CEventSink>::CreateInstance(&m_pEventSink);

    if (FAILED(hr))
        return 1;

    m_pEventSink->m_pMain = this;

    hr = AtlAdvise(m_pWebBrowserUnk, m_pEventSink->GetUnknown(),
                   DIID_DWebBrowserEvents2, &m_dwCookie);

    if (FAILED(hr))
        return 1;

    return 0;
}


LRESULT CMain::OnSize(UINT nMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
    if (m_hwndWebBrowser != NULL)
        ::MoveWindow(m_hwndWebBrowser, 0, 0, LOWORD(lParam), HIWORD(lParam), TRUE);

    return 0;
}

LRESULT CMain::OnDestroy(UINT nMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
    HRESULT hr;

    if (m_dwCookie != 0)
        hr = AtlUnadvise(m_pWebBrowserUnk, DIID_DWebBrowserEvents2, m_dwCookie);

    m_pWebBrowser.Release();
    m_pWebBrowserUnk.Release();

    PostQuitMessage(0);

    return 0;
}

//////////////////////////////////////////////////////////////////
// Implementation of CMain::SaveSnapshot
//////////////////////////////////////////////////////////////////
BOOL CMain::SaveSnapshot(IDispatch* pdisp, VARIANT* purl)
{
        long bodyHeight, bodyWidth, rootHeight, rootWidth, height, width;

        CComPtr<IDispatch> pDispatch;
        CComPtr<IDispatch> pWebBrowserDisp;

        HRESULT hr = m_pWebBrowser->get_Document(&pDispatch);

        if (FAILED(hr))
            return true;

        hr = m_pWebBrowserUnk->QueryInterface(IID_IDispatch, (void**)&pWebBrowserDisp);

        if (FAILED(hr))
            return true;

        // This is not the source we are looking for.
        if (pWebBrowserDisp != pdisp)
            return false;

        // TODO: if we fail past this point, we should not really try again
        // but abort the whole process to avoid being locked and halting.

        CComPtr<IHTMLDocument2> spDocument;
        hr = pDispatch->QueryInterface(IID_IHTMLDocument2, (void**)&spDocument);

        if (FAILED(hr))
            return true;

        CComPtr<IHTMLElement> spBody;
        hr = spDocument->get_body(&spBody);

        if (FAILED(hr))
            return true;

        CComPtr<IHTMLElement2> spBody2;
        hr = spBody->QueryInterface(IID_IHTMLElement2, (void**)&spBody2);

        if (FAILED(hr))
            return true;

        hr = spBody2->get_scrollHeight(&bodyHeight);

        if (FAILED(hr))
            return true;

        hr = spBody2->get_scrollWidth(&bodyWidth);

        if (FAILED(hr))
            return true;

        CComPtr<IHTMLDocument3> spDocument3;
        hr = pDispatch->QueryInterface(IID_IHTMLDocument2, (void**)&spDocument3);

        if (FAILED(hr))
            return true;

        // We also need to get the dimensions from the <html> due to quirks
        // and standards mode differences. Perhaps this should instead check
        // whether we are in quirks mode? How does it work with IE8?
        CComPtr<IHTMLElement> spHtml;
        hr = spDocument3->get_documentElement(&spHtml);

        if (FAILED(hr))
            return true;

        CComPtr<IHTMLElement2> spHtml2;
        hr = spHtml->QueryInterface(IID_IHTMLElement2, (void**)&spHtml2);

        if (FAILED(hr))
            return true;

        hr = spHtml2->get_scrollHeight(&rootHeight);

        if (FAILED(hr))
            return true;

        hr = spHtml2->get_scrollWidth(&rootWidth);

        if (FAILED(hr))
            return true;

        width = bodyWidth;
        height = rootHeight > bodyHeight ? rootHeight : bodyHeight;

        // TODO: What if width or height exceeds 32767? It seems Windows limits
        // the window size, and Internet Explorer does not draw what's not visible.
        ::MoveWindow(m_hwndWebBrowser, 0, 0, width, height, TRUE);

        CComPtr<IViewObject2> spViewObject;

        // This used to get the interface from the m_pWebBrowser but that seems
        // to be an undocumented feature, so we get it from the Document instead.
        hr = spDocument3->QueryInterface(IID_IViewObject2, (void**)&spViewObject);

        if (FAILED(hr))
            return true;

        RECTL rcBounds = { 0, 0, width, height };

        CImage image;

        // TODO: check return value;
        image.Create(width, height, 24);

        HDC imgDc = image.GetDC();
        hr = spViewObject->Draw(DVASPECT_CONTENT, -1, NULL, NULL, imgDc,
                                imgDc, &rcBounds, NULL, NULL, 0);
        image.ReleaseDC();

        if (SUCCEEDED(hr))
            hr = image.Save(m_fileName);

        return true;
}

static const GUID myGUID = { 0x445c10c2, 0xa6d4, 0x40a9, { 0x9c, 0x3f, 0x4e, 0x90, 0x42, 0x1d, 0x7e, 0x83 } };
static CComModule _Main;

void
IECaptHelp(void) {
    printf(" -----------------------------------------------------------------------------\n");
    printf(" Usage: iecapt --url=http://www.example.org/ --out=localfile.png\n");
    printf(" -----------------------------------------------------------------------------\n");
    printf("   --url=...        The URL to capture\n");
    printf("   --out=...        The target file (.png|jpeg|bmp|...)\n");
    printf("   --min-width=...  Minimal width for the image (default: 800)\n");
    printf("   --silent         Whether to surpress some dialogs\n");
    printf(" -----------------------------------------------------------------------------\n");
    printf(" http://iecapt.sf.net - (c) 2003-2008 Bjoern Hoehrmann - <bjoern@hoehrmann.de>\n");
}

int _tmain (int argc, _TCHAR* argv[])
{
    int ax;
    int argHelp = 0;
    int argSilent = 0;
    _TCHAR* argUrl = NULL;
    _TCHAR* argOut = NULL;
    unsigned int argMinWidth = 800;

    // Parse command line parameters
    for (ax = 1; ax < argc; ++ax) {
        size_t nlen;

        _TCHAR* s = argv[ax];
        _TCHAR* value;

        // boolean options
        if (_tcscmp(_T("--silent"), s) == 0) {
          argSilent = 1;
          continue;

        } else if (_tcscmp(_T("--help"), s) == 0) {
          argHelp = 1;
          break;
        } 

        value = _tcschr(s, '=');

        if (value == NULL) {
            if (argUrl == NULL) {
                argUrl = s;
                continue;
            } else if (argOut == NULL) {
                argOut = s;
                continue;
            }
            // error
            argHelp = 1;
            break;
        }

        nlen = value++ - s;

        // --name=value options
        if (_tcsncmp(_T("--url"), s, nlen) == 0) {
          argUrl = value;

        } else if (_tcsncmp(_T("--min-width"), s, nlen) == 0) {
            // TODO: add error checking here?
            argMinWidth = (unsigned int)_tstoi(value);

        } else if (_tcsncmp(_T("--out"), s, nlen) == 0) {
          argOut = value;

        } else {
          // TODO: error
          argHelp = 1;
        }
    }

    if (argUrl == NULL || argOut == NULL || argHelp) {
        IECaptHelp();
        return EXIT_FAILURE;
    }

    HRESULT hr = _Main.Init(NULL, ::GetModuleHandle(NULL), &myGUID);

    if (FAILED(hr))
        return EXIT_FAILURE;

    if (!AtlAxWinInit())
        return EXIT_FAILURE;

    CMain MainWnd(argUrl, argOut, argSilent);

    RECT rcMain = { 0, 0, argMinWidth, 600 };
    MainWnd.Create(NULL, rcMain, _T("IECapt"), WS_POPUP);

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    _Main.Term();

    return EXIT_SUCCESS;
}
