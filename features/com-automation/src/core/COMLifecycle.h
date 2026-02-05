#ifndef COM_LIFECYCLE_H
#define COM_LIFECYCLE_H

#include <windows.h>
#include <comdef.h>

/**
 * @class COMLifecycle
 * @brief Manages COM initialization and cleanup with RAII pattern
 * 
 * This class ensures proper COM initialization and cleanup.
 * It uses RAII (Resource Acquisition Is Initialization) to guarantee
 * that CoUninitialize is called even if exceptions occur.
 * 
 * Thread Safety: Each thread must have its own COMLifecycle instance
 * or use the thread-local GetInstance() method.
 */
class COMLifecycle {
public:
    /**
     * @brief Initialize COM for the current thread
     * @param apartmentType COINIT_APARTMENTTHREADED or COINIT_MULTITHREADED
     * @throws std::runtime_error if COM initialization fails
     */
    explicit COMLifecycle(DWORD apartmentType = COINIT_MULTITHREADED);
    
    /**
     * @brief Uninitialize COM for the current thread
     * 
     * This is called automatically when the object goes out of scope.
     * It's safe to call even if initialization failed.
     */
    ~COMLifecycle();
    
    // Delete copy constructor and assignment operator
    // COM initialization is per-thread and should not be copied
    COMLifecycle(const COMLifecycle&) = delete;
    COMLifecycle& operator=(const COMLifecycle&) = delete;
    
    /**
     * @brief Check if COM was successfully initialized
     * @return true if COM is initialized, false otherwise
     */
    bool IsInitialized() const { return initialized; }
    
    /**
     * @brief Get the HRESULT from initialization
     * @return HRESULT from CoInitializeEx call
     */
    HRESULT GetInitResult() const { return initResult; }
    
    /**
     * @brief Get thread-local COM lifecycle instance
     * @return Reference to thread-local COMLifecycle
     * 
     * This provides a convenient way to ensure COM is initialized
     * for the current thread without managing the lifecycle manually.
     */
    static COMLifecycle& GetInstance();

private:
    bool initialized;
    HRESULT initResult;
    DWORD threadId;
};

/**
 * @class COMScope
 * @brief RAII helper for scoped COM initialization
 * 
 * Usage:
 * @code
 * void MyFunction() {
 *     COMScope comScope;  // COM initialized here
 *     // ... use COM objects ...
 * }  // COM automatically uninitialized here
 * @endcode
 */
class COMScope {
public:
    COMScope() : lifecycle(COINIT_MULTITHREADED) {}
    explicit COMScope(DWORD apartmentType) : lifecycle(apartmentType) {}
    
    bool IsValid() const { return lifecycle.IsInitialized(); }
    
private:
    COMLifecycle lifecycle;
};

#endif // COM_LIFECYCLE_H
