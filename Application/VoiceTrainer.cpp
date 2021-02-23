#pragma comment(linker,"\"/manifestdependency:type='win32'                                 \
                name='Microsoft.Windows.Common-Controls' version='6.0.0.0'                 \
                processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#include "windows.h"
#include "windowsx.h"
#include "richedit.h"
#include "tchar.h"
#include "stdio.h"
#include "ole2.h"
#include "atlbase.h"
#include "atlcom.h"
#include "atlctl.h"
#include "atlwin.h"
#include "VoiceTrainer.h"
#include "resource.h"
#undef NULL
#define NULL 0
#define APPCLASS "HmphVoiceTrainer"
#define APPVERSION "1.0"

static const GUID IID_IVoiceTrainer = {0x233ACB05, 0x6635, 0x4916, {0x9A, 0x4D, 0x70, 0x27, 0x71, 0x63, 0x05, 0xA5}};
static const GUID IID_IVoiceTrainerEvents = {0xE11C348B, 0xE78C, 0x491E, {0xB9, 0xA9, 0x7B, 0x7F, 0xA7, 0xD9, 0xF7, 0xEE}};

CComModule _module;

int CALLBACK WinMain(HINSTANCE instance, HINSTANCE prevInstance, LPSTR cmdLine, INT winOpts)
{
    HWND mainDlg;
    if (mainDlg = FindWindow(_T(APPCLASS), NULL))
    {
        ShowWindow(mainDlg, SW_SHOW);
        OpenIcon(mainDlg);
        SetForegroundWindow(mainDlg);
        return 0;
    }
    if (CoInitialize(NULL) != S_OK)
        return 1;
    LoadLibrary(_T("ThereKernel.dll"));
    LoadLibrary(_T("riched32.dll"));
    INITCOMMONCONTROLSEX initControls;
    initControls.dwSize = sizeof(initControls);
    initControls.dwICC = ICC_STANDARD_CLASSES | ICC_PROGRESS_CLASS;
    InitCommonControlsEx(&initControls);
    WNDCLASS mainClass = {0};
    mainClass.style = NULL;
    mainClass.lpfnWndProc = DefDlgProc;
    mainClass.cbClsExtra = 0;
    mainClass.cbWndExtra = DLGWINDOWEXTRA;
    mainClass.hInstance = instance;
    mainClass.hIcon = LoadIcon(NULL, IDI_APPLICATION);
    mainClass.hCursor = LoadCursor(NULL, IDC_ARROW);
    mainClass.hbrBackground = NULL;
    mainClass.lpszMenuName = NULL;
    mainClass.lpszClassName = _T(APPCLASS);
    RegisterClass(&mainClass);
    if (mainDlg = CreateDialog(instance, MAKEINTRESOURCE(ID_DIALOG_MAIN), NULL, (DLGPROC)MainDlgProc))
    {
        ShowWindow(mainDlg, winOpts);
        MSG msg;
        while (GetMessage(&msg, NULL, 0, 0))
        {
            if (!IsDialogMessage(mainDlg, &msg))
            {
                TranslateMessage(&msg);
                DispatchMessage(&msg);
            }
        }
    }
    UnregisterClass(mainClass.lpszClassName, mainClass.hInstance);
    CoUninitialize();
    return 0;
}

BOOL APIENTRY MainDlgProc(HWND dlg, UINT msg, WPARAM wPrm, LPARAM lPrm)
{
    VoiceTrainer *voiceTrainer = reinterpret_cast<VoiceTrainer*>(GetWindowLongPtr(dlg, DWLP_USER));
    switch (msg)
    {
        case WM_INITDIALOG:
        {
            voiceTrainer = new VoiceTrainer(dlg);
            if (voiceTrainer == NULL)
                return FALSE;
            SetWindowLongPtr(dlg, DWLP_USER, (LPARAM)voiceTrainer);
            return TRUE;
        }
        case WM_SYSCOMMAND:
        {
            if (wPrm == SC_CLOSE)
                DestroyWindow(dlg);
            break;
        }
        case WM_CLOSE:
        {
            DestroyWindow(dlg);
            break;
        }
        case WM_DESTROY:
        {
            if (voiceTrainer != NULL)
            {
                delete voiceTrainer;
                SetWindowLongPtr(dlg, DWLP_USER, NULL);
            }
            PostQuitMessage(0);
            break;
        }
        case WM_COMMAND:
        {
            switch (LOWORD(wPrm))
            {
                case ID_BUTTON_VOICE:
                {
                    if (voiceTrainer != NULL)
                    {
                        LONG state = voiceTrainer->GetState();
                        if (state > 0 && state < 3)
                            voiceTrainer->GetInterface()->Config(state + 1);
                    }
                    break;
                }
            }
            break;
        }
    }
    return FALSE;
}

VoiceTrainer::VoiceTrainer(HWND dlg):
    m_refCount(1),
    m_dlg(dlg),
    m_voiceTrainer(NULL),
    m_connectionPointID(0),
    m_state(-1)
{
    IUnknown *container = NULL;
    IUnknown *control = NULL;
    AtlAxCreateControlEx(_T("There.VoiceTrainer"), GetDlgItem(m_dlg, ID_BUTTON_TRAINER), NULL, &container, &control);
    if (control != NULL)
    {
        control->QueryInterface(IID_IVoiceTrainer, (void**)&m_voiceTrainer);
        if (m_voiceTrainer != NULL)
        {
            CComQIPtr<IConnectionPointContainer> voiceTrainer = m_voiceTrainer;
            voiceTrainer->FindConnectionPoint(IID_IVoiceTrainerEvents, &m_connectionPoint);
            CComPtr<IDispatch> recipient;
            recipient = this;
            m_connectionPoint->Advise(recipient, &m_connectionPointID);
            m_voiceTrainer->Init();
        }
        control->Release();
        container->Release();
    }
    CHARFORMAT cformat;
    cformat.cbSize      = sizeof(cformat);
    cformat.dwMask      = CFM_BOLD | CFM_COLOR;
    cformat.dwEffects   = CFE_BOLD;
    cformat.crTextColor = GetSysColor(COLOR_BTNTEXT);
    SendDlgItemMessage(m_dlg, ID_EDIT_TITLE, EM_SETCHARFORMAT, SCF_ALL, (LPARAM)&cformat);
    cformat.dwMask      = CFM_COLOR;
    cformat.dwEffects   = 0;
    cformat.crTextColor = GetSysColor(COLOR_BTNTEXT);
    SendDlgItemMessage(m_dlg, ID_EDIT_BLURB, EM_SETCHARFORMAT, SCF_ALL, (LPARAM)&cformat);
    if (m_voiceTrainer == NULL)
    {
        SetDlgItemText(m_dlg, ID_EDIT_TITLE, _T("Sorry!"));
        SetDlgItemText(m_dlg, ID_EDIT_BLURB, _T("This application must be installed in the There client directory to function properly."));
        SetDlgItemText(m_dlg, ID_EDIT_MESSAGE, _T(""));
        ShowWindow(GetDlgItem(m_dlg, ID_BUTTON_VOICE), SW_HIDE);
    }
}

VoiceTrainer::~VoiceTrainer()
{
    if (m_connectionPoint != NULL)
    {
        m_connectionPoint->Unadvise(m_connectionPointID);
        m_connectionPointID = 0;
        m_connectionPoint = NULL;
    }
    if (m_voiceTrainer != NULL)
    {
        m_voiceTrainer->Cancel();
        m_voiceTrainer->Release();
        m_voiceTrainer = NULL;
    }
}

IVoiceTrainer* VoiceTrainer::GetInterface()
{
    return m_voiceTrainer;
}

LONG VoiceTrainer::GetState()
{
    return m_state;
}

HRESULT STDMETHODCALLTYPE VoiceTrainer::QueryInterface(REFIID riid, _COM_Outptr_ void __RPC_FAR *__RPC_FAR *object)
{
    if (IsEqualIID(riid, IID_IVoiceTrainerEvents))
    {
        AddRef();
        *object = static_cast<IVoiceTrainerEvents*>(this);
        return S_OK;
    }
    *object = NULL;
    return E_NOINTERFACE;
}

ULONG STDMETHODCALLTYPE VoiceTrainer::AddRef()
{
    return ++m_refCount;
}

ULONG STDMETHODCALLTYPE VoiceTrainer::Release()
{
    return --m_refCount;
}

HRESULT STDMETHODCALLTYPE VoiceTrainer::GetTypeInfoCount(UINT *pctinfo)
{
    *pctinfo = 0;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE VoiceTrainer::GetTypeInfo(UINT itinfo, LCID lcid, ITypeInfo **tinfo)
{
    *tinfo = NULL;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE VoiceTrainer::GetIDsOfNames(REFIID riid, LPOLESTR *rgsznames, UINT cnames, LCID lcid, DISPID *rgdispid)
{
    return DispGetIDsOfNames(NULL, rgsznames, cnames, rgdispid);
}

HRESULT STDMETHODCALLTYPE VoiceTrainer::Invoke(DISPID dispidmember, REFIID riid, LCID lcid, WORD flags, DISPPARAMS *dispparams,
                                               VARIANT *varresult, EXCEPINFO *excepinfo, UINT *argerr)
{
    switch (dispidmember)
    {
        case 1:
            return OnBeginRecord();
        case 2:
            return OnEndRecord();
        case 3:
            return OnLevelChange();
        case 4:
            return OnConfigStateChange();
        case 5:
            return OnConfigMessage();
        case 6:
            return OnConfigError();
    }
    return E_FAIL;
}

HRESULT STDMETHODCALLTYPE VoiceTrainer::OnBeginRecord()
{
    ShowWindow(GetDlgItem(m_dlg, ID_PROGRESS_VOLUME), SW_SHOW);
    EnableWindow(GetDlgItem(m_dlg, ID_BUTTON_VOICE), FALSE);
    SetDlgItemText(m_dlg, ID_EDIT_MESSAGE, _T(""));
    return S_OK;
}

HRESULT STDMETHODCALLTYPE VoiceTrainer::OnEndRecord()
{
    ShowWindow(GetDlgItem(m_dlg, ID_PROGRESS_VOLUME), SW_HIDE);
    EnableWindow(GetDlgItem(m_dlg, ID_BUTTON_VOICE), TRUE);
    return S_OK;
}

HRESULT STDMETHODCALLTYPE VoiceTrainer::OnLevelChange()
{
    LONG level = 0;
    if (m_voiceTrainer != NULL && m_voiceTrainer->GetRecordLevel(&level) == S_OK)
    {
        SendDlgItemMessage(m_dlg, ID_PROGRESS_VOLUME, PBM_SETPOS, level, 0);
    }
    return S_OK;
}

HRESULT STDMETHODCALLTYPE VoiceTrainer::OnConfigStateChange()
{
    if (m_voiceTrainer != NULL && m_voiceTrainer->GetConfigState(&m_state) == S_OK)
    {
        switch (m_state)
        {
            case 1:
            {
                SetDlgItemText(m_dlg, ID_EDIT_TITLE, _T("Silence Test"));
                SetDlgItemText(m_dlg, ID_EDIT_BLURB, _T("This test will set the minimum volume levels required by your microphone. ")
                                                     _T("You must have your headset/microphone plugged in & activated for this test to work.\r\n\r\n")
                                                     _T("Click the button below to begin, and do not speak into your microphone.\r\n\r\n")
                                                     _T("The test will run approximately 15 seconds."));
                SetDlgItemText(m_dlg, ID_EDIT_MESSAGE, _T(""));
                ShowWindow(GetDlgItem(m_dlg, ID_BUTTON_VOICE), SW_SHOW);
                break;
            }
            case 2:
            {
                SetDlgItemText(m_dlg, ID_EDIT_TITLE, _T("Speaking Test"));
                SetDlgItemText(m_dlg, ID_EDIT_BLURB, _T("Click the button below to begin, then read the following paragraph out loud in a normal, conversational tone.\r\n\r\n")
                                                     _T("We are now gathering information about your normal speaking volume. ")
                                                     _T("Continue to read this paragraph until the test completes. ")
                                                     _T("When the test is done, a screen confirming the test completion will appear. ")
                                                     _T("If you get to the end of this paragraph before the end of the test, start reading it again from the beginning.\r\n\r\n")
                                                     _T("The entire test should take approximately 15 seconds."));
                SetDlgItemText(m_dlg, ID_EDIT_MESSAGE, _T(""));
                ShowWindow(GetDlgItem(m_dlg, ID_BUTTON_VOICE), SW_SHOW);
                break;
            }
            case 3:
            {
                SetDlgItemText(m_dlg, ID_EDIT_TITLE, _T("Setup Complete"));
                SetDlgItemText(m_dlg, ID_EDIT_BLURB, _T("Your headset/microphone is now adjusted and ready to use!"));
                SetDlgItemText(m_dlg, ID_EDIT_MESSAGE, _T(""));
                ShowWindow(GetDlgItem(m_dlg, ID_BUTTON_VOICE), SW_HIDE);
                break;
            }
        }
    }
    return S_OK;
}

HRESULT STDMETHODCALLTYPE VoiceTrainer::OnConfigMessage()
{
    BSTR str = NULL;
    if (m_voiceTrainer != NULL && m_voiceTrainer->GetConfigMessage(&str) == S_OK)
    {
        CHARFORMAT cformat;
        cformat.cbSize      = sizeof(cformat);
        cformat.dwMask      = CFM_COLOR;
        cformat.dwEffects   = 0;
        cformat.crTextColor = RGB(255, 242, 0);
        SendDlgItemMessage(m_dlg, ID_EDIT_MESSAGE, EM_SETCHARFORMAT, SCF_ALL, (LPARAM)&cformat);
        SetDlgItemText(m_dlg, ID_EDIT_MESSAGE, str);
    }
    return S_OK;
}

HRESULT STDMETHODCALLTYPE VoiceTrainer::OnConfigError()
{
    BSTR str = NULL;
    if (m_voiceTrainer != NULL && m_voiceTrainer->GetConfigError(&str) == S_OK)
    {
        CHARFORMAT cformat;
        cformat.cbSize      = sizeof(cformat);
        cformat.dwMask      = CFM_COLOR;
        cformat.dwEffects   = 0;
        cformat.crTextColor = RGB(237, 28, 36);
        SendDlgItemMessage(m_dlg, ID_EDIT_MESSAGE, EM_SETCHARFORMAT, SCF_ALL, (LPARAM)&cformat);
        SetDlgItemText(m_dlg, ID_EDIT_MESSAGE, str);
    }
    return S_OK;
}