#ifndef COMEVENTSINK_H
#define COMEVENTSINK_H

#include <napi.h>
#include <windows.h>
#include <comdef.h>
#include <OCIdl.h>
#include <map>
#include <vector>

struct EventData {
    DISPID dispId;
    std::vector<VARIANT> args;
};

class COMEventSink : public IDispatch {
public:
    COMEventSink(Napi::Env env, Napi::Function callback);
    ~COMEventSink();

    // IUnknown methods
    STDMETHOD(QueryInterface)(REFIID riid, void** ppvObject);
    STDMETHOD_(ULONG, AddRef)();
    STDMETHOD_(ULONG, Release)();

    // IDispatch methods
    STDMETHOD(GetTypeInfoCount)(UINT* pctinfo);
    STDMETHOD(GetTypeInfo)(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo);
    STDMETHOD(GetIDsOfNames)(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId);
    STDMETHOD(Invoke)(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr);

private:
    ULONG refCount;
    Napi::ThreadSafeFunction tsfn;
    
    static void CallJs(Napi::Env env, Napi::Function jsCallback, EventData* data);
};

#endif // COMEVENTSINK_H
