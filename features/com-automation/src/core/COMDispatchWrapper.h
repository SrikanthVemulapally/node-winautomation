#ifndef COM_DISPATCH_WRAPPER_H
#define COM_DISPATCH_WRAPPER_H

#include <windows.h>
#include <comdef.h>
#include <atlbase.h>
#include <string>
#include <unordered_map>
#include <memory>
#include "COMVariant.h"

/**
 * @class COMDispatchWrapper
 * @brief Thread-safe wrapper for IDispatch with automatic reference counting
 * 
 * This class provides a safe, RAII-compliant wrapper around IDispatch.
 * It uses ATL's CComPtr for automatic reference counting and provides
 * type-safe methods for property access and method invocation.
 * 
 * Memory Management:
 * - Uses CComPtr<IDispatch> for automatic AddRef/Release
 * - DISPID caching for performance
 * - Exception-safe operations
 * 
 * Thread Safety:
 * - Safe to use from single thread
 * - Not thread-safe for concurrent access (use mutex if needed)
 */
class COMDispatchWrapper {
public:
    /**
     * @brief Construct from IDispatch pointer
     * @param dispatch IDispatch pointer (AddRef is called automatically)
     * @throws std::invalid_argument if dispatch is nullptr
     */
    explicit COMDispatchWrapper(IDispatch* dispatch);
    
    /**
     * @brief Destructor - automatically releases IDispatch
     */
    ~COMDispatchWrapper();
    
    // Delete copy, allow move
    COMDispatchWrapper(const COMDispatchWrapper&) = delete;
    COMDispatchWrapper& operator=(const COMDispatchWrapper&) = delete;
    COMDispatchWrapper(COMDispatchWrapper&& other) noexcept;
    COMDispatchWrapper& operator=(COMDispatchWrapper&& other) noexcept;
    
    /**
     * @brief Get a property value
     * @param name Property name
     * @return COMVariant containing the property value
     * @throws std::runtime_error if property doesn't exist or access fails
     */
    COMVariant GetProperty(const std::string& name);
    
    /**
     * @brief Set a property value
     * @param name Property name
     * @param value Value to set
     * @throws std::runtime_error if property doesn't exist or is read-only
     */
    void SetProperty(const std::string& name, const COMVariant& value);
    
    /**
     * @brief Invoke a method
     * @param name Method name
     * @param args Method arguments (in correct order)
     * @return COMVariant containing the return value
     * @throws std::runtime_error if method doesn't exist or invocation fails
     */
    COMVariant Invoke(const std::string& name, const std::vector<COMVariant>& args);
    
    /**
     * @brief Get the underlying IDispatch pointer
     * @return Raw IDispatch pointer (do not Release)
     */
    IDispatch* GetDispatch() const { return dispatch.p; }
    
    /**
     * @brief Check if the wrapper is valid
     * @return true if IDispatch pointer is not null
     */
    bool IsValid() const { return dispatch != nullptr; }
    
    /**
     * @brief Release the IDispatch pointer
     * 
     * After calling this, the wrapper becomes invalid.
     * This is called automatically by the destructor.
     */
    void Release();

private:
    /**
     * @brief Get DISPID for a name (with caching)
     * @param name Member name
     * @return DISPID or DISPID_UNKNOWN if not found
     */
    DISPID GetDispID(const std::string& name);
    
    /**
     * @brief Convert std::string to wide string
     * @param str UTF-8 string
     * @return Wide string
     */
    std::wstring StringToWide(const std::string& str);
    
    /**
     * @brief Get detailed error message from EXCEPINFO
     * @param excepInfo Exception information
     * @return Error message string
     */
    std::string GetExceptionMessage(const EXCEPINFO& excepInfo);
    
    /**
     * @brief Format HRESULT as error message
     * @param hr HRESULT code
     * @param context Context description
     * @return Formatted error message
     */
    std::string FormatError(HRESULT hr, const std::string& context);

private:
    CComPtr<IDispatch> dispatch;                           // Smart pointer to IDispatch
    std::unordered_map<std::string, DISPID> dispidCache;  // DISPID cache for performance
};

#endif // COM_DISPATCH_WRAPPER_H
