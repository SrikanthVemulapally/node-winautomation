#include "COMObject.h"
#include "COMEventSink.h"
#include <comutil.h>

#pragma comment(lib, "comsuppw.lib")

Napi::FunctionReference COMObject::constructor;

Napi::Object COMObject::Init(Napi::Env env, Napi::Object exports) {
    Napi::HandleScope scope(env);

    Napi::Function func = DefineClass(env, "COMObject", {
        InstanceMethod("getProperty", &COMObject::GetProperty),
        InstanceMethod("setProperty", &COMObject::SetProperty),
        InstanceMethod("invoke", &COMObject::Invoke),
        InstanceMethod("release", &COMObject::Release),
        InstanceMethod("adviseEvent", &COMObject::AdviseEvent),
        InstanceMethod("unadviseEvent", &COMObject::UnadviseEvent)
    });

    constructor = Napi::Persistent(func);
    constructor.SuppressDestruct();

    exports.Set("COMObject", func);
    return exports;
}

COMObject::COMObject(const Napi::CallbackInfo& info) 
    : Napi::ObjectWrap<COMObject>(info), pDispatch(nullptr) {
    
    Napi::Env env = info.Env();
    
    // Allow construction with null for internal use (wrapping existing IDispatch)
    if (info.Length() > 0 && info[0].IsNull()) {
        return;
    }
    
    if (info.Length() < 1 || !info[0].IsString()) {
        Napi::TypeError::New(env, "ProgID string expected").ThrowAsJavaScriptException();
        return;
    }

    std::string progId = info[0].As<Napi::String>().Utf8Value();
    std::wstring wProgId(progId.begin(), progId.end());

    CLSID clsid;
    HRESULT hr = CLSIDFromProgID(wProgId.c_str(), &clsid);
    if (FAILED(hr)) {
        Napi::Error::New(env, "Failed to get CLSID from ProgID").ThrowAsJavaScriptException();
        return;
    }

    hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)&pDispatch);
    if (FAILED(hr)) {
        Napi::Error::New(env, "Failed to create COM object").ThrowAsJavaScriptException();
        return;
    }
}

COMObject::~COMObject() {
    // Unadvise all events
    for (auto& pair : eventCookies) {
        IConnectionPointContainer* pCPC = nullptr;
        if (SUCCEEDED(pDispatch->QueryInterface(IID_IConnectionPointContainer, (void**)&pCPC))) {
            IConnectionPoint* pCP = nullptr;
            if (SUCCEEDED(pCPC->FindConnectionPoint(pair.first, &pCP))) {
                pCP->Unadvise(pair.second);
                pCP->Release();
            }
            pCPC->Release();
        }
    }
    
    // Release event sinks
    for (auto& pair : eventSinks) {
        if (pair.second) {
            pair.second->Release();
        }
    }
    
    if (pDispatch) {
        pDispatch->Release();
        pDispatch = nullptr;
    }
}

DISPID COMObject::GetDispId(const std::wstring& name) {
    if (!pDispatch) return DISPID_UNKNOWN;

    LPOLESTR memberName = const_cast<LPOLESTR>(name.c_str());
    DISPID dispId;
    HRESULT hr = pDispatch->GetIDsOfNames(IID_NULL, &memberName, 1, LOCALE_USER_DEFAULT, &dispId);
    
    return SUCCEEDED(hr) ? dispId : DISPID_UNKNOWN;
}

Napi::Value COMObject::VariantToValue(Napi::Env env, const VARIANT& var) {
    switch (var.vt) {
        case VT_EMPTY:
        case VT_NULL:
            return env.Null();
        
        case VT_I2:
            return Napi::Number::New(env, var.iVal);
        
        case VT_I4:
            return Napi::Number::New(env, var.lVal);
        
        case VT_R4:
            return Napi::Number::New(env, var.fltVal);
        
        case VT_R8:
            return Napi::Number::New(env, var.dblVal);
        
        case VT_BOOL:
            return Napi::Boolean::New(env, var.boolVal != VARIANT_FALSE);
        
        case VT_BSTR:
            if (var.bstrVal) {
                int len = WideCharToMultiByte(CP_UTF8, 0, var.bstrVal, -1, NULL, 0, NULL, NULL);
                std::string str(len - 1, 0);
                WideCharToMultiByte(CP_UTF8, 0, var.bstrVal, -1, &str[0], len, NULL, NULL);
                return Napi::String::New(env, str);
            }
            return env.Null();
        
        case VT_DISPATCH:
            if (var.pdispVal) {
                // Create a new COMObject instance manually without calling constructor
                Napi::Object obj = constructor.New({env.Null()});
                COMObject* comObj = Napi::ObjectWrap<COMObject>::Unwrap(obj);
                
                // Release any existing dispatch and set the new one
                if (comObj->pDispatch) {
                    comObj->pDispatch->Release();
                }
                comObj->pDispatch = var.pdispVal;
                comObj->pDispatch->AddRef();
                return obj;
            }
            return env.Null();
        
        default:
            return env.Undefined();
    }
}

VARIANT COMObject::ValueToVariant(const Napi::Value& value) {
    VARIANT var;
    VariantInit(&var);

    if (value.IsNull() || value.IsUndefined()) {
        var.vt = VT_NULL;
    }
    else if (value.IsBoolean()) {
        var.vt = VT_BOOL;
        var.boolVal = value.As<Napi::Boolean>().Value() ? VARIANT_TRUE : VARIANT_FALSE;
    }
    else if (value.IsNumber()) {
        double num = value.As<Napi::Number>().DoubleValue();
        if (num == (int)num) {
            var.vt = VT_I4;
            var.lVal = (int)num;
        } else {
            var.vt = VT_R8;
            var.dblVal = num;
        }
    }
    else if (value.IsString()) {
        std::string str = value.As<Napi::String>().Utf8Value();
        int len = MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, NULL, 0);
        wchar_t* wstr = new wchar_t[len];
        MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, wstr, len);
        var.vt = VT_BSTR;
        var.bstrVal = SysAllocString(wstr);
        delete[] wstr;
    }

    return var;
}

Napi::Value COMObject::GetProperty(const Napi::CallbackInfo& info) {
    Napi::Env env = info.Env();

    if (!pDispatch) {
        Napi::Error::New(env, "COM object not initialized").ThrowAsJavaScriptException();
        return env.Null();
    }

    if (info.Length() < 1 || !info[0].IsString()) {
        Napi::TypeError::New(env, "Property name expected").ThrowAsJavaScriptException();
        return env.Null();
    }

    std::string propName = info[0].As<Napi::String>().Utf8Value();
    std::wstring wPropName(propName.begin(), propName.end());

    DISPID dispId = GetDispId(wPropName);
    if (dispId == DISPID_UNKNOWN) {
        Napi::Error::New(env, "Property not found").ThrowAsJavaScriptException();
        return env.Null();
    }

    DISPPARAMS dispParams = { NULL, NULL, 0, 0 };
    VARIANT result;
    VariantInit(&result);

    HRESULT hr = pDispatch->Invoke(dispId, IID_NULL, LOCALE_USER_DEFAULT, 
                                    DISPATCH_PROPERTYGET, &dispParams, &result, NULL, NULL);

    if (FAILED(hr)) {
        Napi::Error::New(env, "Failed to get property").ThrowAsJavaScriptException();
        return env.Null();
    }

    Napi::Value retVal = VariantToValue(env, result);
    VariantClear(&result);
    return retVal;
}

Napi::Value COMObject::SetProperty(const Napi::CallbackInfo& info) {
    Napi::Env env = info.Env();

    if (!pDispatch) {
        Napi::Error::New(env, "COM object not initialized").ThrowAsJavaScriptException();
        return env.Undefined();
    }

    if (info.Length() < 2 || !info[0].IsString()) {
        Napi::TypeError::New(env, "Property name and value expected").ThrowAsJavaScriptException();
        return env.Undefined();
    }

    std::string propName = info[0].As<Napi::String>().Utf8Value();
    std::wstring wPropName(propName.begin(), propName.end());

    DISPID dispId = GetDispId(wPropName);
    if (dispId == DISPID_UNKNOWN) {
        Napi::Error::New(env, "Property not found").ThrowAsJavaScriptException();
        return env.Undefined();
    }

    VARIANT varValue = ValueToVariant(info[1]);
    
    DISPID dispidNamed = DISPID_PROPERTYPUT;
    DISPPARAMS dispParams;
    dispParams.rgvarg = &varValue;
    dispParams.rgdispidNamedArgs = &dispidNamed;
    dispParams.cArgs = 1;
    dispParams.cNamedArgs = 1;

    HRESULT hr = pDispatch->Invoke(dispId, IID_NULL, LOCALE_USER_DEFAULT,
                                    DISPATCH_PROPERTYPUT, &dispParams, NULL, NULL, NULL);

    VariantClear(&varValue);

    if (FAILED(hr)) {
        Napi::Error::New(env, "Failed to set property").ThrowAsJavaScriptException();
    }

    return env.Undefined();
}

Napi::Value COMObject::Invoke(const Napi::CallbackInfo& info) {
    Napi::Env env = info.Env();

    if (!pDispatch) {
        Napi::Error::New(env, "COM object not initialized").ThrowAsJavaScriptException();
        return env.Null();
    }

    if (info.Length() < 1 || !info[0].IsString()) {
        Napi::TypeError::New(env, "Method name expected").ThrowAsJavaScriptException();
        return env.Null();
    }

    std::string methodName = info[0].As<Napi::String>().Utf8Value();
    std::wstring wMethodName(methodName.begin(), methodName.end());

    DISPID dispId = GetDispId(wMethodName);
    if (dispId == DISPID_UNKNOWN) {
        Napi::Error::New(env, "Method not found").ThrowAsJavaScriptException();
        return env.Null();
    }

    // Convert arguments
    int argCount = info.Length() - 1;
    VARIANT* args = nullptr;
    if (argCount > 0) {
        args = new VARIANT[argCount];
        for (int i = 0; i < argCount; i++) {
            args[argCount - 1 - i] = ValueToVariant(info[i + 1]); // Reverse order for COM
        }
    }

    DISPPARAMS dispParams;
    dispParams.rgvarg = args;
    dispParams.rgdispidNamedArgs = NULL;
    dispParams.cArgs = argCount;
    dispParams.cNamedArgs = 0;

    VARIANT result;
    VariantInit(&result);

    HRESULT hr = pDispatch->Invoke(dispId, IID_NULL, LOCALE_USER_DEFAULT,
                                    DISPATCH_METHOD, &dispParams, &result, NULL, NULL);

    // Clean up arguments
    if (args) {
        for (int i = 0; i < argCount; i++) {
            VariantClear(&args[i]);
        }
        delete[] args;
    }

    if (FAILED(hr)) {
        Napi::Error::New(env, "Failed to invoke method").ThrowAsJavaScriptException();
        return env.Null();
    }

    Napi::Value retVal = VariantToValue(env, result);
    VariantClear(&result);
    return retVal;
}

Napi::Value COMObject::Release(const Napi::CallbackInfo& info) {
    if (pDispatch) {
        pDispatch->Release();
        pDispatch = nullptr;
    }
    return info.Env().Undefined();
}

Napi::Value COMObject::AdviseEvent(const Napi::CallbackInfo& info) {
    Napi::Env env = info.Env();

    if (!pDispatch) {
        Napi::Error::New(env, "COM object not initialized").ThrowAsJavaScriptException();
        return env.Undefined();
    }

    if (info.Length() < 2 || !info[0].IsString() || !info[1].IsFunction()) {
        Napi::TypeError::New(env, "IID string and callback function expected").ThrowAsJavaScriptException();
        return env.Undefined();
    }

    // Parse IID from string
    std::string iidStr = info[0].As<Napi::String>().Utf8Value();
    std::wstring wIidStr(iidStr.begin(), iidStr.end());
    
    IID iid;
    HRESULT hr = IIDFromString(wIidStr.c_str(), &iid);
    if (FAILED(hr)) {
        Napi::Error::New(env, "Invalid IID string").ThrowAsJavaScriptException();
        return env.Undefined();
    }

    // Get connection point container
    IConnectionPointContainer* pCPC = nullptr;
    hr = pDispatch->QueryInterface(IID_IConnectionPointContainer, (void**)&pCPC);
    if (FAILED(hr)) {
        Napi::Error::New(env, "Object does not support connection points").ThrowAsJavaScriptException();
        return env.Undefined();
    }

    // Find connection point
    IConnectionPoint* pCP = nullptr;
    hr = pCPC->FindConnectionPoint(iid, &pCP);
    pCPC->Release();
    
    if (FAILED(hr)) {
        Napi::Error::New(env, "Connection point not found").ThrowAsJavaScriptException();
        return env.Undefined();
    }

    // Create event sink
    Napi::Function callback = info[1].As<Napi::Function>();
    COMEventSink* pSink = new COMEventSink(env, callback);

    // Advise
    DWORD cookie;
    hr = pCP->Advise(pSink, &cookie);
    pCP->Release();

    if (FAILED(hr)) {
        pSink->Release();
        Napi::Error::New(env, "Failed to advise event sink").ThrowAsJavaScriptException();
        return env.Undefined();
    }

    // Store cookie and sink
    eventCookies[iid] = cookie;
    eventSinks[iid] = pSink;

    return Napi::Number::New(env, cookie);
}

Napi::Value COMObject::UnadviseEvent(const Napi::CallbackInfo& info) {
    Napi::Env env = info.Env();

    if (!pDispatch) {
        Napi::Error::New(env, "COM object not initialized").ThrowAsJavaScriptException();
        return env.Undefined();
    }

    if (info.Length() < 1 || !info[0].IsString()) {
        Napi::TypeError::New(env, "IID string expected").ThrowAsJavaScriptException();
        return env.Undefined();
    }

    // Parse IID from string
    std::string iidStr = info[0].As<Napi::String>().Utf8Value();
    std::wstring wIidStr(iidStr.begin(), iidStr.end());
    
    IID iid;
    HRESULT hr = IIDFromString(wIidStr.c_str(), &iid);
    if (FAILED(hr)) {
        Napi::Error::New(env, "Invalid IID string").ThrowAsJavaScriptException();
        return env.Undefined();
    }

    // Check if we have a cookie for this IID
    auto cookieIt = eventCookies.find(iid);
    if (cookieIt == eventCookies.end()) {
        Napi::Error::New(env, "No event connection found for this IID").ThrowAsJavaScriptException();
        return env.Undefined();
    }

    // Get connection point container
    IConnectionPointContainer* pCPC = nullptr;
    hr = pDispatch->QueryInterface(IID_IConnectionPointContainer, (void**)&pCPC);
    if (FAILED(hr)) {
        Napi::Error::New(env, "Object does not support connection points").ThrowAsJavaScriptException();
        return env.Undefined();
    }

    // Find connection point
    IConnectionPoint* pCP = nullptr;
    hr = pCPC->FindConnectionPoint(iid, &pCP);
    pCPC->Release();
    
    if (FAILED(hr)) {
        Napi::Error::New(env, "Connection point not found").ThrowAsJavaScriptException();
        return env.Undefined();
    }

    // Unadvise
    hr = pCP->Unadvise(cookieIt->second);
    pCP->Release();

    if (FAILED(hr)) {
        Napi::Error::New(env, "Failed to unadvise event sink").ThrowAsJavaScriptException();
        return env.Undefined();
    }

    // Release and remove sink
    auto sinkIt = eventSinks.find(iid);
    if (sinkIt != eventSinks.end()) {
        sinkIt->second->Release();
        eventSinks.erase(sinkIt);
    }
    
    eventCookies.erase(cookieIt);

    return env.Undefined();
}
