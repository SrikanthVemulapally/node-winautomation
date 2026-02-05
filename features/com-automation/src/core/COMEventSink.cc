#include "COMEventSink.h"
#include <comutil.h>

COMEventSink::COMEventSink(Napi::Env env, Napi::Function callback) 
    : refCount(1) {
    // Create thread-safe function with simple signature
    tsfn = Napi::ThreadSafeFunction::New(
        env,
        callback,
        "COMEventCallback",
        0,  // unlimited queue
        1   // initial thread count
    );
}

COMEventSink::~COMEventSink() {
    if (tsfn) {
        tsfn.Release();
    }
}

STDMETHODIMP COMEventSink::QueryInterface(REFIID riid, void** ppvObject) {
    if (riid == IID_IUnknown || riid == IID_IDispatch) {
        *ppvObject = static_cast<IDispatch*>(this);
        AddRef();
        return S_OK;
    }
    *ppvObject = nullptr;
    return E_NOINTERFACE;
}

STDMETHODIMP_(ULONG) COMEventSink::AddRef() {
    return InterlockedIncrement(&refCount);
}

STDMETHODIMP_(ULONG) COMEventSink::Release() {
    ULONG count = InterlockedDecrement(&refCount);
    if (count == 0) {
        delete this;
    }
    return count;
}

STDMETHODIMP COMEventSink::GetTypeInfoCount(UINT* pctinfo) {
    *pctinfo = 0;
    return S_OK;
}

STDMETHODIMP COMEventSink::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo) {
    return E_NOTIMPL;
}

STDMETHODIMP COMEventSink::GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId) {
    return E_NOTIMPL;
}

void COMEventSink::CallJs(Napi::Env env, Napi::Function jsCallback, EventData* data) {
    if (!data) return;
    
    // Create event object
    Napi::Object eventObj = Napi::Object::New(env);
    eventObj.Set("dispId", Napi::Number::New(env, data->dispId));
    
    // Convert parameters to JavaScript array
    Napi::Array args = Napi::Array::New(env, data->args.size());
    for (size_t i = 0; i < data->args.size(); i++) {
        VARIANT& var = data->args[i];
        
        Napi::Value jsValue;
        switch (var.vt) {
            case VT_I2:
                jsValue = Napi::Number::New(env, var.iVal);
                break;
            case VT_I4:
                jsValue = Napi::Number::New(env, var.lVal);
                break;
            case VT_R4:
                jsValue = Napi::Number::New(env, var.fltVal);
                break;
            case VT_R8:
                jsValue = Napi::Number::New(env, var.dblVal);
                break;
            case VT_BOOL:
                jsValue = Napi::Boolean::New(env, var.boolVal != VARIANT_FALSE);
                break;
            case VT_BSTR:
                if (var.bstrVal) {
                    int len = WideCharToMultiByte(CP_UTF8, 0, var.bstrVal, -1, NULL, 0, NULL, NULL);
                    std::string str(len - 1, 0);
                    WideCharToMultiByte(CP_UTF8, 0, var.bstrVal, -1, &str[0], len, NULL, NULL);
                    jsValue = Napi::String::New(env, str);
                } else {
                    jsValue = env.Null();
                }
                break;
            default:
                jsValue = env.Undefined();
                break;
        }
        
        args.Set(i, jsValue);
    }
    
    eventObj.Set("args", args);
    
    // Call JavaScript callback
    jsCallback.Call({eventObj});
}

STDMETHODIMP COMEventSink::Invoke(
    DISPID dispIdMember,
    REFIID riid,
    LCID lcid,
    WORD wFlags,
    DISPPARAMS* pDispParams,
    VARIANT* pVarResult,
    EXCEPINFO* pExcepInfo,
    UINT* puArgErr) {
    
    // Create event data
    EventData* data = new EventData();
    data->dispId = dispIdMember;
    
    // Copy arguments
    for (UINT i = 0; i < pDispParams->cArgs; i++) {
        VARIANT var;
        VariantInit(&var);
        VariantCopy(&var, &pDispParams->rgvarg[pDispParams->cArgs - 1 - i]); // Reverse order
        data->args.push_back(var);
    }
    
    // Call JavaScript callback on main thread via ThreadSafeFunction
    auto status = tsfn.BlockingCall(data, CallJs);
    
    if (status != napi_ok) {
        // Failed to queue callback
        for (auto& var : data->args) {
            VariantClear(&var);
        }
        delete data;
    }
    
    return S_OK;
}
