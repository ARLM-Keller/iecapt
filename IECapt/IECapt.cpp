////////////////////////////////////////////////////////////////////
//
// IECapt - A Internet Explorer Web Page Rendering Capture Utility
//
// Copyright (C) 2003-2006 Bjoern Hoehrmann <bjoern@hoehrmann.de>
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

public:

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

    CMain() : m_dwCookie(0) { }

public:

    BEGIN_MSG_MAP(CMainWindow)
        MESSAGE_HANDLER(WM_CREATE,  OnCreate)
        MESSAGE_HANDLER(WM_SIZE,    OnSize)
        MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
    END_MSG_MAP()

public:

    LRESULT OnCreate  (UINT nMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnSize    (UINT nMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnDestroy (UINT nMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);

public:

    BOOL SaveSnapshot(IDispatch* pdisp, VARIANT* purl);

public:
    LPTSTR m_URI;
    LPTSTR m_fileName;

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
        IDispatch* pDispatch = NULL;
        IDispatch* pWebBrowserDisp = NULL;

        HRESULT hr = m_pWebBrowser->get_Document(&pDispatch);

        if (FAILED(hr))
            return true;

        hr = m_pWebBrowserUnk->QueryInterface(IID_IDispatch, (void**)&pWebBrowserDisp);

        if (FAILED(hr))
            return true;

        if (pWebBrowserDisp != pdisp)
        {
            pWebBrowserDisp->Release();
            return false;
        }

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

        MoveWindow(0, 0, width, height, TRUE);      
        ::MoveWindow(m_hwndWebBrowser, 0, 0, width, height, TRUE);

        CComPtr<IViewObject2> spViewObject;
        hr = m_pWebBrowser->QueryInterface(IID_IViewObject2, (void**)&spViewObject);

        if (FAILED(hr))
            return true;

        BITMAPINFOHEADER bih;
        BITMAPINFO bi;
        RGBQUAD rgbquad;

        ZeroMemory(&bih, sizeof(BITMAPINFOHEADER));
        ZeroMemory(&rgbquad, sizeof(RGBQUAD));

        bih.biSize          = sizeof(BITMAPINFOHEADER);
        bih.biWidth         = width;
        bih.biHeight        = height;
        bih.biPlanes        = 1;
        bih.biBitCount      = 24;
        bih.biClrUsed       = 0;
        bih.biSizeImage     = 0;
        bih.biCompression   = BI_RGB;
        bih.biXPelsPerMeter = 0;
        bih.biYPelsPerMeter = 0;

        bi.bmiHeader = bih;
        bi.bmiColors[0] = rgbquad;

        HDC hdcMain = GetDC();

        if (!hdcMain)
            return true;

        HDC hdcMem = CreateCompatibleDC(hdcMain);

        if (!hdcMem)
            return true;

        char* bitmapData = NULL;
        HBITMAP hBitmap = CreateDIBSection(hdcMain, &bi, DIB_RGB_COLORS, (void**)&bitmapData, NULL, 0);

        if (!hBitmap) {
            // TODO: cleanup
            return true;
        }

        SelectObject(hdcMem, hBitmap);

        RECTL rcBounds = { 0, 0, width, height };
        hr = spViewObject->Draw(DVASPECT_CONTENT, -1, NULL, NULL, hdcMain,
                                hdcMem, &rcBounds, NULL, NULL, 0);

        if (SUCCEEDED(hr))
        {
            CImage image;
            image.Create(width, height, 24);
            CImageDC imageDC(image);
            ::BitBlt(imageDC, 0, 0, width, height, hdcMem, 0, 0, SRCCOPY);
            image.Save(m_fileName);
        }

        DeleteObject(hBitmap);
        DeleteDC(hdcMem);

        pWebBrowserDisp->Release();

        return true;
}

static const GUID myGUID = { 0x445c10c2, 0xa6d4, 0x40a9, { 0x9c, 0x3f, 0x4e, 0x90, 0x42, 0x1d, 0x7e, 0x83 } };
static CComModule _Main;

int _tmain (int argc, _TCHAR* argv[])
{
    if (argc != 3)
    {
        printf("Usage: %s http://www.example.org/ localfile.png\n", argv[0]);
        return EXIT_FAILURE;
    }
    
    HRESULT hr = _Main.Init(NULL, ::GetModuleHandle(NULL), &myGUID);

    if (FAILED(hr))
        return EXIT_FAILURE;

    if (!AtlAxWinInit())
        return EXIT_FAILURE;

    CMain MainWnd;

    MainWnd.m_URI = argv[1];
    MainWnd.m_fileName = argv[2];
    RECT rcMain = { 0, 0, 800, 600 };
    MainWnd.Create(NULL, rcMain, _T("IECapt"), WS_POPUP);

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) 
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    _Main.Term();

    return EXIT_SUCCESS;
}
