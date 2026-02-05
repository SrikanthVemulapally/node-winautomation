#include "COMLifecycle.h"
#include <stdexcept>
#include <sstream>
#include <thread>
#include <mutex>
#include <unordered_map>

// Thread-local storage for COM lifecycle
thread_local std::unique_ptr<COMLifecycle> g_threadLocalCOM;

COMLifecycle::COMLifecycle(DWORD apartmentType) 
    : initialized(false), initResult(E_FAIL), threadId(GetCurrentThreadId()) {
    
    // Initialize COM for this thread
    initResult = CoInitializeEx(NULL, apartmentType);
    
    // S_OK means COM was initialized
    // S_FALSE means COM was already initialized (which is also success)
    // RPC_E_CHANGED_MODE means COM was initialized with a different apartment type
    if (initResult == S_OK || initResult == S_FALSE) {
        initialized = true;
    } else if (initResult == RPC_E_CHANGED_MODE) {
        // COM already initialized with different apartment type
        // This is acceptable - we'll use the existing initialization
        initialized = true;
        initResult = S_OK;
    } else {
        // Initialization failed
        _com_error error(initResult);
        std::stringstream ss;
        ss << "COM initialization failed: " << error.ErrorMessage();
        throw std::runtime_error(ss.str());
    }
}

COMLifecycle::~COMLifecycle() {
    // Only uninitialize if we successfully initialized
    // and we're on the same thread
    if (initialized && GetCurrentThreadId() == threadId) {
        CoUninitialize();
        initialized = false;
    }
}

COMLifecycle& COMLifecycle::GetInstance() {
    if (!g_threadLocalCOM) {
        g_threadLocalCOM = std::make_unique<COMLifecycle>();
    }
    return *g_threadLocalCOM;
}
