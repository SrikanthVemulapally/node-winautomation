#ifndef COM_VARIANT_H
#define COM_VARIANT_H

#include <napi.h>
#include <windows.h>
#include <comdef.h>
#include <vector>
#include <string>

/**
 * @class COMVariant
 * @brief Safe wrapper for VARIANT with automatic cleanup
 * 
 * This class provides RAII management for VARIANT types,
 * ensuring VariantClear is always called.
 */
class COMVariant {
public:
    COMVariant();
    ~COMVariant();
    
    // Delete copy, allow move
    COMVariant(const COMVariant&) = delete;
    COMVariant& operator=(const COMVariant&) = delete;
    COMVariant(COMVariant&& other) noexcept;
    COMVariant& operator=(COMVariant&& other) noexcept;
    
    /**
     * @brief Get the underlying VARIANT
     * @return Pointer to VARIANT
     */
    VARIANT* Get() { return &variant; }
    const VARIANT* Get() const { return &variant; }
    
    /**
     * @brief Detach and return the VARIANT (caller takes ownership)
     * @return VARIANT value
     */
    VARIANT Detach();
    
    /**
     * @brief Clear the VARIANT
     */
    void Clear();
    
private:
    VARIANT variant;
};

/**
 * @namespace COMVariantConverter
 * @brief Utilities for converting between VARIANT and Napi::Value
 */
namespace COMVariantConverter {
    /**
     * @brief Convert Napi::Value to VARIANT
     * @param value JavaScript value
     * @return COMVariant containing the converted value
     * @throws Napi::TypeError if conversion fails
     */
    COMVariant ToVariant(const Napi::Value& value);
    
    /**
     * @brief Convert VARIANT to Napi::Value
     * @param env N-API environment
     * @param variant VARIANT to convert
     * @return Napi::Value containing the converted value
     * @throws Napi::Error if conversion fails
     */
    Napi::Value FromVariant(Napi::Env env, const VARIANT& variant);
    
    /**
     * @brief Convert array of Napi::Value to VARIANT array
     * @param values JavaScript values
     * @return Vector of COMVariant
     */
    std::vector<COMVariant> ToVariantArray(const std::vector<Napi::Value>& values);
    
    /**
     * @brief Convert BSTR to std::string (UTF-8)
     * @param bstr BSTR to convert
     * @return UTF-8 encoded string
     */
    std::string BSTRToString(BSTR bstr);
    
    /**
     * @brief Convert std::string (UTF-8) to BSTR
     * @param str UTF-8 encoded string
     * @return BSTR (caller must free with SysFreeString)
     */
    BSTR StringToBSTR(const std::string& str);
}

#endif // COM_VARIANT_H
