#ifndef COMOBJECT_H
#define COMOBJECT_H

#include <napi.h>
#include <windows.h>
#include <comdef.h>
#include <atlbase.h>
#include <OCIdl.h>
#include <map>

class COMEventSink;

// Comparator for IID/GUID to use in std::map
struct IIDComparator {
    bool operator()(const IID& a, const IID& b) const {
        return memcmp(&a, &b, sizeof(IID)) < 0;
    }
};

class COMObject : public Napi::ObjectWrap<COMObject> {
public:
    static Napi::Object Init(Napi::Env env, Napi::Object exports);
    COMObject(const Napi::CallbackInfo& info);
    ~COMObject();

private:
    static Napi::FunctionReference constructor;
    
    // Methods
    Napi::Value GetProperty(const Napi::CallbackInfo& info);
    Napi::Value SetProperty(const Napi::CallbackInfo& info);
    Napi::Value Invoke(const Napi::CallbackInfo& info);
    Napi::Value Release(const Napi::CallbackInfo& info);
    Napi::Value AdviseEvent(const Napi::CallbackInfo& info);
    Napi::Value UnadviseEvent(const Napi::CallbackInfo& info);

    // Helper methods
    DISPID GetDispId(const std::wstring& name);
    Napi::Value VariantToValue(Napi::Env env, const VARIANT& var);
    VARIANT ValueToVariant(const Napi::Value& value);

    // COM object
    IDispatch* pDispatch;
    
    // Event handling
    std::map<IID, DWORD, IIDComparator> eventCookies;
    std::map<IID, COMEventSink*, IIDComparator> eventSinks;
};

#endif // COMOBJECT_H
