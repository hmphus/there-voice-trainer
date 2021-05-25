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
#define APPCLASS L"HmphVoiceTrainer"
#define APPVERSION L"1.1"

static const GUID IID_IVoiceTrainer = {0x233ACB05, 0x6635, 0x4916, {0x9A, 0x4D, 0x70, 0x27, 0x71, 0x63, 0x05, 0xA5}};
static const GUID IID_IVoiceTrainerEvents = {0xE11C348B, 0xE78C, 0x491E, {0xB9, 0xA9, 0x7B, 0x7F, 0xA7, 0xD9, 0xF7, 0xEE}};

CComModule _module;

INT CALLBACK WinMain(HINSTANCE instance, HINSTANCE prevInstance, LPSTR cmdLine, INT winOpts)
{
    HWND mainDlg;
    if (mainDlg = FindWindow(APPCLASS, nullptr))
    {
        ShowWindow(mainDlg, SW_SHOW);
        OpenIcon(mainDlg);
        SetForegroundWindow(mainDlg);
        return 0;
    }

    if (CoInitialize(nullptr) != S_OK)
        return 1;

    LoadLibrary(L"ThereKernel.dll");
    LoadLibrary(L"riched32.dll");

    INITCOMMONCONTROLSEX initControls;
    initControls.dwSize = sizeof(initControls);
    initControls.dwICC = ICC_STANDARD_CLASSES | ICC_PROGRESS_CLASS;
    InitCommonControlsEx(&initControls);

    WNDCLASS mainClass = {0};
    mainClass.style = 0;
    mainClass.lpfnWndProc = DefDlgProc;
    mainClass.cbClsExtra = 0;
    mainClass.cbWndExtra = DLGWINDOWEXTRA;
    mainClass.hInstance = instance;
    mainClass.hIcon = LoadIcon(nullptr, IDI_APPLICATION);
    mainClass.hCursor = LoadCursor(nullptr, IDC_ARROW);
    mainClass.hbrBackground = nullptr;
    mainClass.lpszMenuName = nullptr;
    mainClass.lpszClassName = APPCLASS;
    RegisterClass(&mainClass);

    if (mainDlg = CreateDialog(instance, MAKEINTRESOURCE(ID_DIALOG_MAIN), nullptr, (DLGPROC)MainDlgProc))
    {
        ShowWindow(mainDlg, winOpts);
        MSG msg;
        while (GetMessage(&msg, nullptr, 0, 0))
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
            if (voiceTrainer == nullptr)
                return false;

            SetWindowLongPtr(dlg, DWLP_USER, (LPARAM)voiceTrainer);

            return true;
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
            if (voiceTrainer != nullptr)
            {
                delete voiceTrainer;
                SetWindowLongPtr(dlg, DWLP_USER, 0);
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
                    if (voiceTrainer != nullptr)
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
    return false;
}

VoiceTrainer::VoiceTrainer(HWND dlg):
    m_refCount(1),
    m_dlg(dlg),
    m_voiceTrainer(nullptr),
    m_connectionPointID(0),
    m_state(-1)
{
    IUnknown *container = nullptr;
    IUnknown *control = nullptr;
    AtlAxCreateControlEx(L"There.VoiceTrainer", GetDlgItem(m_dlg, ID_BUTTON_TRAINER), nullptr, &container, &control);
    if (control != nullptr)
    {
        control->QueryInterface(IID_IVoiceTrainer, (void**)&m_voiceTrainer);
        if (m_voiceTrainer != nullptr)
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

    if (m_voiceTrainer == nullptr)
    {
        SetDlgItemText(m_dlg, ID_EDIT_TITLE, L"Sorry!");
        SetDlgItemText(m_dlg, ID_EDIT_BLURB, L"This application must be installed in the There client directory to function properly.");
        SetDlgItemText(m_dlg, ID_EDIT_MESSAGE, L"");
        ShowWindow(GetDlgItem(m_dlg, ID_BUTTON_VOICE), SW_HIDE);
    }
}

VoiceTrainer::~VoiceTrainer()
{
    if (m_connectionPoint != nullptr)
    {
        m_connectionPoint->Unadvise(m_connectionPointID);
        m_connectionPointID = 0;
        m_connectionPoint = nullptr;
    }

    if (m_voiceTrainer != nullptr)
    {
        m_voiceTrainer->Cancel();
        m_voiceTrainer->Release();
        m_voiceTrainer = nullptr;
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

    *object = nullptr;
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
    *tinfo = nullptr;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE VoiceTrainer::GetIDsOfNames(REFIID riid, LPOLESTR *rgsznames, UINT cnames, LCID lcid, DISPID *rgdispid)
{
    return DispGetIDsOfNames(nullptr, rgsznames, cnames, rgdispid);
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
        default:
            return E_FAIL;
    }
}

HRESULT STDMETHODCALLTYPE VoiceTrainer::OnBeginRecord()
{
    ShowWindow(GetDlgItem(m_dlg, ID_PROGRESS_VOLUME), SW_SHOW);
    EnableWindow(GetDlgItem(m_dlg, ID_BUTTON_VOICE), false);
    SetDlgItemText(m_dlg, ID_EDIT_MESSAGE, L"");
    return S_OK;
}

HRESULT STDMETHODCALLTYPE VoiceTrainer::OnEndRecord()
{
    ShowWindow(GetDlgItem(m_dlg, ID_PROGRESS_VOLUME), SW_HIDE);
    EnableWindow(GetDlgItem(m_dlg, ID_BUTTON_VOICE), true);
    return S_OK;
}

HRESULT STDMETHODCALLTYPE VoiceTrainer::OnLevelChange()
{
    LONG level = 0;
    if (m_voiceTrainer != nullptr && SUCCEEDED(m_voiceTrainer->GetRecordLevel(&level)))
    {
        WPARAM state;
        if (level < 60)
            state = PBST_NORMAL;
        else if (level < 80)
            state = PBST_PAUSED;
        else
            state = PBST_ERROR;

        SendDlgItemMessage(m_dlg, ID_PROGRESS_VOLUME, PBM_SETPOS, level, 0);
        SendDlgItemMessage(m_dlg, ID_PROGRESS_VOLUME, PBM_SETSTATE, state, 0);
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE VoiceTrainer::OnConfigStateChange()
{
    if (m_voiceTrainer != nullptr && SUCCEEDED(m_voiceTrainer->GetConfigState(&m_state)))
    {
        switch (m_state)
        {
            case 1:
            {
                SetDlgItemText(m_dlg, ID_EDIT_TITLE, L"Silence Test");
                SetDlgItemText(m_dlg, ID_EDIT_BLURB, L"This test will set the minimum volume levels required by your microphone. "
                                                     L"You must have your headset/microphone plugged in & activated for this test to work.\r\n\r\n"
                                                     L"Click the button below to begin, and do not speak into your microphone.\r\n\r\n"
                                                     L"The test will run approximately 15 seconds.");
                SetDlgItemText(m_dlg, ID_EDIT_MESSAGE, L"");
                ShowWindow(GetDlgItem(m_dlg, ID_BUTTON_VOICE), SW_SHOW);
                break;
            }
            case 2:
            {
                SetDlgItemText(m_dlg, ID_EDIT_TITLE, L"Speaking Test");
                SetDlgItemText(m_dlg, ID_EDIT_BLURB, L"Click the button below to begin, then read the following paragraph out loud in a normal, conversational tone.\r\n\r\n"
                                                     L"We are now gathering information about your normal speaking volume. "
                                                     L"Continue to read this paragraph until the test completes. "
                                                     L"When the test is done, a screen confirming the test completion will appear. "
                                                     L"If you get to the end of this paragraph before the end of the test, start reading it again from the beginning.\r\n\r\n"
                                                     L"The entire test should take approximately 15 seconds.");
                SetDlgItemText(m_dlg, ID_EDIT_MESSAGE, L"");
                ShowWindow(GetDlgItem(m_dlg, ID_BUTTON_VOICE), SW_SHOW);
                break;
            }
            case 3:
            {
                SetDlgItemText(m_dlg, ID_EDIT_TITLE, L"Setup Complete");
                SetDlgItemText(m_dlg, ID_EDIT_BLURB, L"Your headset/microphone is now adjusted and ready to use!");
                SetDlgItemText(m_dlg, ID_EDIT_MESSAGE, L"");
                ShowWindow(GetDlgItem(m_dlg, ID_BUTTON_VOICE), SW_HIDE);
                break;
            }
        }
    }
    return S_OK;
}

HRESULT STDMETHODCALLTYPE VoiceTrainer::OnConfigMessage()
{
    BSTR str = nullptr;
    if (m_voiceTrainer != nullptr && SUCCEEDED(m_voiceTrainer->GetConfigMessage(&str)))
    {
        CHARFORMAT cformat;
        cformat.cbSize      = sizeof(cformat);
        cformat.dwMask      = CFM_COLOR;
        cformat.dwEffects   = 0;
        cformat.crTextColor = RGB(0, 182, 0);
        SendDlgItemMessage(m_dlg, ID_EDIT_MESSAGE, EM_SETCHARFORMAT, SCF_ALL, (LPARAM)&cformat);
        SetDlgItemText(m_dlg, ID_EDIT_MESSAGE, str);
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE VoiceTrainer::OnConfigError()
{
    BSTR str = nullptr;
    if (m_voiceTrainer != nullptr && SUCCEEDED(m_voiceTrainer->GetConfigError(&str)))
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