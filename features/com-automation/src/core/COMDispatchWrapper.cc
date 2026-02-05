#include "COMDispatchWrapper.h"
#include <sstream>
#include <stdexcept>

COMDispatchWrapper::COMDispatchWrapper(IDispatch* dispatch) {
    if (!dispatch) {
        throw std::invalid_argument("IDispatch pointer cannot be null");
    }
    
    // CComPtr automatically calls AddRef
    this->dispatch = dispatch;
}

COMDispatchWrapper::~COMDispatchWrapper() {
    Release();
}

COMDispatchWrapper::COMDispatchWrapper(COMDispatchWrapper&& other) noexcept
    : dispatch(std::move(other.dispatch)), dispidCache(std::move(other.dispidCache)) {
    // CComPtr move constructor handles reference counting
}

COMDispatchWrapper& COMDispatchWrapper::operator=(COMDispatchWrapper&& other) noexcept {
    if (this != &other) {
        dispatch = std::move(other.dispatch);
        dispidCache = std::move(other.dispidCache);
    }
    return *this;
}

COMVariant COMDispatchWrapper::GetProperty(const std::string& name) {
    if (!IsValid()) {
        throw std::runtime_error("IDispatch is not valid");
    }
    
    DISPID dispId = GetDispID(name);
    if (dispId == DISPID_UNKNOWN) {
        throw std::runtime_error("Property not found: " + name);
    }
    
    DISPPARAMS dispParams = { NULL, NULL, 0, 0 };
    COMVariant result;
    EXCEPINFO excepInfo = { 0 };
    UINT argErr = 0;
    
    HRESULT hr = dispatch->Invoke(
        dispId,
        IID_NULL,
        LOCALE_USER_DEFAULT,
        DISPATCH_PROPERTYGET,
        &dispParams,
        result.Get(),
        &excepInfo,
        &argErr
    );
    
    if (FAILED(hr)) {
        if (hr == DISP_E_EXCEPTION) {
            throw std::runtime_error("Failed to get property '" + name + "': " + GetExceptionMessage(excepInfo));
        }
        throw std::runtime_error(FormatError(hr, "get property '" + name + "'"));
    }
    
    return result;
}

void COMDispatchWrapper::SetProperty(const std::string& name, const COMVariant& value) {
    if (!IsValid()) {
        throw std::runtime_error("IDispatch is not valid");
    }
    
    DISPID dispId = GetDispID(name);
    if (dispId == DISPID_UNKNOWN) {
        throw std::runtime_error("Property not found: " + name);
    }
    
    DISPID dispidNamed = DISPID_PROPERTYPUT;
    VARIANT varValue = *value.Get();
    
    DISPPARAMS dispParams;
    dispParams.rgvarg = &varValue;
    dispParams.rgdispidNamedArgs = &dispidNamed;
    dispParams.cArgs = 1;
    dispParams.cNamedArgs = 1;
    
    EXCEPINFO excepInfo = { 0 };
    UINT argErr = 0;
    
    HRESULT hr = dispatch->Invoke(
        dispId,
        IID_NULL,
        LOCALE_USER_DEFAULT,
        DISPATCH_PROPERTYPUT,
        &dispParams,
        NULL,
        &excepInfo,
        &argErr
    );
    
    if (FAILED(hr)) {
        if (hr == DISP_E_EXCEPTION) {
            throw std::runtime_error("Failed to set property '" + name + "': " + GetExceptionMessage(excepInfo));
        }
        throw std::runtime_error(FormatError(hr, "set property '" + name + "'"));
    }
}

COMVariant COMDispatchWrapper::Invoke(const std::string& name, const std::vector<COMVariant>& args) {
    if (!IsValid()) {
        throw std::runtime_error("IDispatch is not valid");
    }
    
    DISPID dispId = GetDispID(name);
    if (dispId == DISPID_UNKNOWN) {
        throw std::runtime_error("Method not found: " + name);
    }
    
    // Prepare arguments (COM expects reverse order)
    std::vector<VARIANT> varArgs;
    varArgs.reserve(args.size());
    for (auto it = args.rbegin(); it != args.rend(); ++it) {
        varArgs.push_back(*it->Get());
    }
    
    DISPPARAMS dispParams;
    dispParams.rgvarg = varArgs.empty() ? NULL : varArgs.data();
    dispParams.rgdispidNamedArgs = NULL;
    dispParams.cArgs = static_cast<UINT>(varArgs.size());
    dispParams.cNamedArgs = 0;
    
    COMVariant result;
    EXCEPINFO excepInfo = { 0 };
    UINT argErr = 0;
    
    HRESULT hr = dispatch->Invoke(
        dispId,
        IID_NULL,
        LOCALE_USER_DEFAULT,
        DISPATCH_METHOD,
        &dispParams,
        result.Get(),
        &excepInfo,
        &argErr
    );
    
    if (FAILED(hr)) {
        if (hr == DISP_E_EXCEPTION) {
            throw std::runtime_error("Failed to invoke method '" + name + "': " + GetExceptionMessage(excepInfo));
        }
        if (hr == DISP_E_BADPARAMCOUNT) {
            std::stringstream ss;
            ss << "Wrong number of arguments for method '" << name << "' (provided " << args.size() << ")";
            throw std::runtime_error(ss.str());
        }
        if (hr == DISP_E_TYPEMISMATCH) {
            std::stringstream ss;
            ss << "Type mismatch in argument " << argErr << " for method '" << name << "'";
            throw std::runtime_error(ss.str());
        }
        throw std::runtime_error(FormatError(hr, "invoke method '" + name + "'"));
    }
    
    return result;
}

void COMDispatchWrapper::Release() {
    if (dispatch) {
        dispatch.Release();  // CComPtr handles the actual Release call
        dispidCache.clear();
    }
}

DISPID COMDispatchWrapper::GetDispID(const std::string& name) {
    // Check cache first
    auto it = dispidCache.find(name);
    if (it != dispidCache.end()) {
        return it->second;
    }
    
    // Get DISPID from IDispatch
    std::wstring wName = StringToWide(name);
    LPOLESTR memberName = const_cast<LPOLESTR>(wName.c_str());
    DISPID dispId;
    
    HRESULT hr = dispatch->GetIDsOfNames(
        IID_NULL,
        &memberName,
        1,
        LOCALE_USER_DEFAULT,
        &dispId
    );
    
    if (SUCCEEDED(hr)) {
        // Cache the DISPID for future use
        dispidCache[name] = dispId;
        return dispId;
    }
    
    return DISPID_UNKNOWN;
}

std::wstring COMDispatchWrapper::StringToWide(const std::string& str) {
    if (str.empty()) return std::wstring();
    
    int len = MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, NULL, 0);
    if (len <= 0) return std::wstring();
    
    std::wstring result(len - 1, 0);
    MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, &result[0], len);
    
    return result;
}

std::string COMDispatchWrapper::GetExceptionMessage(const EXCEPINFO& excepInfo) {
    std::stringstream ss;
    
    if (excepInfo.bstrDescription) {
        int len = WideCharToMultiByte(CP_UTF8, 0, excepInfo.bstrDescription, -1, NULL, 0, NULL, NULL);
        if (len > 0) {
            std::string desc(len - 1, 0);
            WideCharToMultiByte(CP_UTF8, 0, excepInfo.bstrDescription, -1, &desc[0], len, NULL, NULL);
            ss << desc;
        }
    }
    
    if (excepInfo.bstrSource) {
        int len = WideCharToMultiByte(CP_UTF8, 0, excepInfo.bstrSource, -1, NULL, 0, NULL, NULL);
        if (len > 0) {
            std::string source(len - 1, 0);
            WideCharToMultiByte(CP_UTF8, 0, excepInfo.bstrSource, -1, &source[0], len, NULL, NULL);
            if (!ss.str().empty()) ss << " ";
            ss << "(Source: " << source << ")";
        }
    }
    
    if (ss.str().empty()) {
        ss << "Unknown COM exception (code: 0x" << std::hex << excepInfo.scode << ")";
    }
    
    return ss.str();
}

std::string COMDispatchWrapper::FormatError(HRESULT hr, const std::string& context) {
    _com_error error(hr);
    std::stringstream ss;
    ss << "Failed to " << context << ": " << error.ErrorMessage() << " (HRESULT: 0x" << std::hex << hr << ")";
    return ss.str();
}
