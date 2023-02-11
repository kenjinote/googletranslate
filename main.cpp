#pragma comment(linker,"/opt:nowin98")
#define UNICODE
#include<atlbase.h>
CComModule _Module;
#include<atlcom.h>
#include<atlhost.h>
#include<windows.h>

BEGIN_OBJECT_MAP(ObjectMap)
END_OBJECT_MAP()

TCHAR szClassName[]=TEXT("window");
CComQIPtr<IWebBrowser2> pWB;
CComPtr<IUnknown> p;

LRESULT CALLBACK WndProc(HWND hWnd,UINT msg,WPARAM wParam,LPARAM lParam)
{
	static HWND hBrowser,hEdit,hButton;
	switch(msg)
	{
	case WM_CREATE:
		{
			AtlAxWinInit();
			hBrowser=CreateWindowEx(
				WS_EX_CLIENTEDGE,
				TEXT("AtlAxWin"),
				TEXT("Shell.Explorer.2"),
				WS_CHILD|WS_VISIBLE,
				10,30*4+20,512,30*4,
				hWnd,
				NULL,
				((LPCREATESTRUCT)lParam)->hInstance,
				NULL);
			if(AtlAxGetControl(hBrowser,&p)==S_OK)
			{
				pWB=p;
			}
			pWB->put_Silent(TRUE);
			CComVariant vempty,vUrl(L"about:blank");
			pWB->Navigate2(&vUrl,&vempty,&vempty,&vempty,&vempty);
			hEdit=CreateWindowEx(
				WS_EX_CLIENTEDGE,
				TEXT("EDIT"),
				0,
				WS_VISIBLE|WS_CHILD|WS_VSCROLL|
				ES_MULTILINE|ES_AUTOHSCROLL|ES_AUTOVSCROLL,
				10,10,512,30*4,
				hWnd,
				0,
				((LPCREATESTRUCT)lParam)->hInstance,
				0);
			hButton=CreateWindow(
				TEXT("BUTTON"),
				TEXT("Google–|–ó"),
				WS_VISIBLE|WS_CHILD,
				512+20,10,128,30,
				hWnd,
				(HMENU)100,
				((LPCREATESTRUCT)lParam)->hInstance,
				0);
		}
		break;
	case WM_COMMAND:
		if(LOWORD(wParam)==100)
		{
			if(!GetWindowTextLength(hEdit))break;
			EnableWindow(hButton,0);
			CComPtr<IDispatch> pdisp;
			CComQIPtr<IHTMLDocument2> pDoc;
			while(1)
			{
				HRESULT hr=pWB->get_Document(&pdisp);
				if(SUCCEEDED(hr)&&pdisp!=NULL)
				{
					pDoc=pdisp;
					if(pDoc!=NULL)
					{
						break;
					}
				}
				Sleep(100);
			}
			VARIANT *param;
			SAFEARRAY *sfArray;
			DWORD szTextLength=GetWindowTextLength(hEdit);
			LPTSTR lpszText;
			lpszText=(LPTSTR)GlobalAlloc(
				GMEM_FIXED,
				sizeof(TCHAR)*(szTextLength+270+299+1));
			lstrcpy(
				lpszText,
				TEXT("<html>\r\n<head>\r\n<meta http-equiv=\"Content-Type\"")
				TEXT("content=\"text/html; charset=shift-jis\">\r\n<script ")
				TEXT("type=\"text/javascript\" src=\"http://www.google.com/jsapi")
				TEXT("\"></script>\r\n<script type=\"text/javascript\"><!--\r\n")
				TEXT("google.load(\"language\",\"1\");\r\nfunction textTranslate()")
				TEXT("{\r\nvar text=\""));
			GetWindowText(hEdit,lpszText+lstrlen(lpszText),szTextLength+1);
			lstrcat(
				lpszText,
				TEXT("\";\r\ngoogle.language.translate(text,\"ja\",\"en\",\r\n")
				TEXT("function(result) {\r\nif (result.translation) {\r\n")
				TEXT("document.getElementById(\"translation\").innerHTML=")
				TEXT("result.translation;\r\n}\r\n}\r\n);\r\n}\r\n")
				TEXT("google.setOnLoadCallback(textTranslate);\r\n// -->")
				TEXT("</script>\r\n</head>\r\n<body>\r\n<div id=\"translation")
				TEXT("\"></div>\r\n</body>\r\n</html>"));
			BSTR bstr=SysAllocString(lpszText);
			GlobalFree(lpszText);
			sfArray=SafeArrayCreateVector(VT_VARIANT,0,1);
			if(sfArray==NULL || pDoc==NULL)
			{
				goto cleanup;
			}
			SafeArrayAccessData(sfArray,(LPVOID*)&param);
			param->vt=VT_BSTR;
			param->bstrVal=bstr;
			SafeArrayUnaccessData(sfArray);
			pDoc->write(sfArray);
			pDoc->close();
cleanup:
			SysFreeString(bstr);
			if(sfArray!=NULL)
			{
				SafeArrayDestroy(sfArray);
			}
			pWB->Refresh();
			EnableWindow(hButton,1);
		}
		break;
	case WM_DESTROY:
		pWB.Release();
		PostQuitMessage(0);
		break;
	default:
		return(DefWindowProc(hWnd,msg,wParam,lParam));
	}
	return (0L);
}

int WINAPI WinMain(HINSTANCE hinst,HINSTANCE hPreInst,
				   LPSTR pCmdLine,int nCmdShow)
{
	HWND hWnd;
	MSG msg;
	WNDCLASS wndclass;
	_Module.Init(ObjectMap,hinst);
	if(!hPreInst)
	{
		wndclass.style=CS_HREDRAW|CS_VREDRAW;
		wndclass.lpfnWndProc=WndProc;
		wndclass.cbClsExtra=0;
		wndclass.cbWndExtra=0;
		wndclass.hInstance =hinst;
		wndclass.hIcon=NULL;
		wndclass.hCursor=LoadCursor(NULL,IDC_ARROW);
		wndclass.hbrBackground=(HBRUSH)(COLOR_WINDOW+1);
		wndclass.lpszMenuName=NULL;
		wndclass.lpszClassName=szClassName;
		if(!RegisterClass(&wndclass))
			return FALSE;
	}
	hWnd=CreateWindow(szClassName,
		TEXT("Google Translate ‚ðŽg‚Á‚½–|–ó"),
		WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT,
		0,
		CW_USEDEFAULT,
		0,
		0,
		0,
		hinst,
		0);
	ShowWindow(hWnd,nCmdShow);
	UpdateWindow(hWnd);
	while (GetMessage(&msg,NULL,0,0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	_Module.Term();
	return (msg.wParam);
}
