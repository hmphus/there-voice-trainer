#pragma once

BOOL APIENTRY MainDlgProc(HWND dlg, UINT msg, WPARAM wPrm, LPARAM lPrm);

class IVoiceTrainer: public IDispatch
{
public:
    virtual HRESULT STDMETHODCALLTYPE Init() = 0;
    virtual HRESULT STDMETHODCALLTYPE Config(LONG config) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetConfigState(LONG *pVal) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetRecordLevel(LONG *pVal) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetConfigMessage(BSTR *pVal) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetConfigError(BSTR *pVal) = 0;
    virtual HRESULT STDMETHODCALLTYPE Cancel() = 0;
    virtual HRESULT STDMETHODCALLTYPE Preprocess(LONG *pVal) = 0;
    virtual HRESULT STDMETHODCALLTYPE Preprocess(LONG newVal) = 0;
    virtual HRESULT STDMETHODCALLTYPE SupportsPreprocess(VARIANT_BOOL *pVal) = 0;
    virtual HRESULT STDMETHODCALLTYPE LaunchRecordingMixer() = 0;
};

class IVoiceTrainerEvents: public IDispatch
{
public:
    virtual HRESULT STDMETHODCALLTYPE OnBeginRecord() = 0;
    virtual HRESULT STDMETHODCALLTYPE OnEndRecord() = 0;
    virtual HRESULT STDMETHODCALLTYPE OnLevelChange() = 0;
    virtual HRESULT STDMETHODCALLTYPE OnConfigStateChange() = 0;
    virtual HRESULT STDMETHODCALLTYPE OnConfigMessage() = 0;
    virtual HRESULT STDMETHODCALLTYPE OnConfigError() = 0;
};

class VoiceTrainer: public IVoiceTrainerEvents
{
public:
    VoiceTrainer(HWND dlg);
    virtual ~VoiceTrainer();
    virtual IVoiceTrainer* GetInterface();
    virtual LONG GetState();

protected:
    virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, _COM_Outptr_ void __RPC_FAR *__RPC_FAR *object) override;
    virtual ULONG   STDMETHODCALLTYPE AddRef() override;
    virtual ULONG   STDMETHODCALLTYPE Release() override;

protected:
    virtual HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *pctinfo) override;
    virtual HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT itinfo, LCID lcid, ITypeInfo **tinfo) override;
    virtual HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, LPOLESTR *rgsznames, UINT cnames, LCID lcid, DISPID *rgdispid) override;
    virtual HRESULT STDMETHODCALLTYPE Invoke(DISPID dispidmember, REFIID riid, LCID lcid, WORD flags, DISPPARAMS *dispparams,
                                             VARIANT *varresult, EXCEPINFO *excepinfo, UINT *argerr) override;

protected:
    virtual HRESULT STDMETHODCALLTYPE OnBeginRecord() override;
    virtual HRESULT STDMETHODCALLTYPE OnEndRecord() override;
    virtual HRESULT STDMETHODCALLTYPE OnLevelChange() override;
    virtual HRESULT STDMETHODCALLTYPE OnConfigStateChange() override;
    virtual HRESULT STDMETHODCALLTYPE OnConfigMessage() override;
    virtual HRESULT STDMETHODCALLTYPE OnConfigError() override;

protected:
    ULONG m_refCount;
    HWND m_dlg;
    IVoiceTrainer *m_voiceTrainer;
    CComPtr<IConnectionPoint> m_connectionPoint;
    DWORD m_connectionPointID;
    LONG m_state;
};