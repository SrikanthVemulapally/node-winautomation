#include "COMVariant.h"
#include <comutil.h>

#pragma comment(lib, "comsuppw.lib")

// COMVariant implementation
COMVariant::COMVariant() {
    VariantInit(&variant);
}

COMVariant::~COMVariant() {
    Clear();
}

COMVariant::COMVariant(COMVariant&& other) noexcept {
    VariantInit(&variant);
    variant = other.variant;
    VariantInit(&other.variant);
}

COMVariant& COMVariant::operator=(COMVariant&& other) noexcept {
    if (this != &other) {
        Clear();
        variant = other.variant;
        VariantInit(&other.variant);
    }
    return *this;
}

VARIANT COMVariant::Detach() {
    VARIANT result = variant;
    VariantInit(&variant);
    return result;
}

void COMVariant::Clear() {
    VariantClear(&variant);
}

// COMVariantConverter implementation
namespace COMVariantConverter {

COMVariant ToVariant(const Napi::Value& value) {
    COMVariant result;
    VARIANT* var = result.Get();
    
    if (value.IsNull() || value.IsUndefined()) {
        var->vt = VT_NULL;
    }
    else if (value.IsBoolean()) {
        var->vt = VT_BOOL;
        var->boolVal = value.As<Napi::Boolean>().Value() ? VARIANT_TRUE : VARIANT_FALSE;
    }
    else if (value.IsNumber()) {
        double num = value.As<Napi::Number>().DoubleValue();
        // Use integer if possible for better compatibility
        if (num == static_cast<int32_t>(num)) {
            var->vt = VT_I4;
            var->lVal = static_cast<int32_t>(num);
        } else {
            var->vt = VT_R8;
            var->dblVal = num;
        }
    }
    else if (value.IsString()) {
        std::string str = value.As<Napi::String>().Utf8Value();
        var->vt = VT_BSTR;
        var->bstrVal = StringToBSTR(str);
    }
    else if (value.IsDate()) {
        // Convert JavaScript Date to VARIANT DATE
        Napi::Date date = value.As<Napi::Date>();
        double jsTime = date.ValueOf();
        // JavaScript time is milliseconds since 1970-01-01
        // COM DATE is days since 1899-12-30
        // Conversion: (jsTime / 86400000) + 25569
        var->vt = VT_DATE;
        var->date = (jsTime / 86400000.0) + 25569.0;
    }
    else {
        // For other types, store as VT_EMPTY
        var->vt = VT_EMPTY;
    }
    
    return result;
}

Napi::Value FromVariant(Napi::Env env, const VARIANT& variant) {
    switch (variant.vt) {
        case VT_EMPTY:
        case VT_NULL:
            return env.Null();
        
        case VT_I1:
            return Napi::Number::New(env, variant.cVal);
        case VT_I2:
            return Napi::Number::New(env, variant.iVal);
        case VT_I4:
            return Napi::Number::New(env, variant.lVal);
        case VT_I8:
            return Napi::Number::New(env, static_cast<double>(variant.llVal));
        
        case VT_UI1:
            return Napi::Number::New(env, variant.bVal);
        case VT_UI2:
            return Napi::Number::New(env, variant.uiVal);
        case VT_UI4:
            return Napi::Number::New(env, variant.ulVal);
        case VT_UI8:
            return Napi::Number::New(env, static_cast<double>(variant.ullVal));
        
        case VT_R4:
            return Napi::Number::New(env, variant.fltVal);
        case VT_R8:
            return Napi::Number::New(env, variant.dblVal);
        
        case VT_BOOL:
            return Napi::Boolean::New(env, variant.boolVal != VARIANT_FALSE);
        
        case VT_BSTR:
            if (variant.bstrVal) {
                return Napi::String::New(env, BSTRToString(variant.bstrVal));
            }
            return env.Null();
        
        case VT_DATE: {
            // Convert COM DATE to JavaScript Date
            // COM DATE is days since 1899-12-30
            // JavaScript time is milliseconds since 1970-01-01
            double jsTime = (variant.date - 25569.0) * 86400000.0;
            return Napi::Date::New(env, jsTime);
        }
        
        case VT_DISPATCH:
            // Return as COMObject - handled by caller
            return env.Undefined();
        
        case VT_ERROR:
            return Napi::Number::New(env, variant.scode);
        
        default:
            // Unsupported type
            return env.Undefined();
    }
}

std::vector<COMVariant> ToVariantArray(const std::vector<Napi::Value>& values) {
    std::vector<COMVariant> result;
    result.reserve(values.size());
    
    for (const auto& value : values) {
        result.push_back(ToVariant(value));
    }
    
    return result;
}

std::string BSTRToString(BSTR bstr) {
    if (!bstr) return "";
    
    int len = WideCharToMultiByte(CP_UTF8, 0, bstr, -1, NULL, 0, NULL, NULL);
    if (len <= 0) return "";
    
    std::string result(len - 1, 0);
    WideCharToMultiByte(CP_UTF8, 0, bstr, -1, &result[0], len, NULL, NULL);
    
    return result;
}

BSTR StringToBSTR(const std::string& str) {
    int len = MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, NULL, 0);
    if (len <= 0) return nullptr;
    
    wchar_t* wstr = new wchar_t[len];
    MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, wstr, len);
    
    BSTR bstr = SysAllocString(wstr);
    delete[] wstr;
    
    return bstr;
}

} // namespace COMVariantConverter
