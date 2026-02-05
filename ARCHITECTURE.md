# COM Automation Architecture

## Overview

The COM automation system provides a robust, memory-safe way to interact with COM objects from Node.js.

## Architecture Layers

```
┌─────────────────────────────────────────────────────────┐
│                    JavaScript Layer                      │
│  (COMObject, OutlookConnector, ExcelConnector, etc.)   │
└─────────────────────────────────────────────────────────┘
                           ↓
┌─────────────────────────────────────────────────────────┐
│                   N-API Wrapper Layer                    │
│         (COMObject class - Napi::ObjectWrap)            │
└─────────────────────────────────────────────────────────┘
                           ↓
┌─────────────────────────────────────────────────────────┐
│                  COM Management Layer                    │
│  (COMDispatchWrapper, COMEventSink, COMLifecycle)       │
└─────────────────────────────────────────────────────────┘
                           ↓
┌─────────────────────────────────────────────────────────┐
│                   Windows COM Layer                      │
│        (IDispatch, IConnectionPoint, etc.)              │
└─────────────────────────────────────────────────────────┘
```

## Core Components

### 1. COMLifecycle
- Manages COM initialization/uninitialization
- Thread-local storage for apartment state
- RAII pattern for automatic cleanup

### 2. COMDispatchWrapper
- Wraps IDispatch interface
- Handles reference counting
- Provides type-safe method invocation
- Manages VARIANT conversions

### 3. COMEventSink
- Implements IDispatch for event callbacks
- Thread-safe event marshaling
- Automatic cleanup on destruction

### 4. COMObject (N-API)
- JavaScript-facing API
- Wraps COMDispatchWrapper
- Manages object lifetime
- Provides getProperty, setProperty, invoke, adviseEvent, unadviseEvent

## Memory Management Strategy

### Reference Counting
- All COM interfaces use AddRef/Release
- Smart pointers (CComPtr) for automatic management
- No manual Release calls in normal code paths

### Resource Cleanup
1. **Destructor chain**: ~COMObject → ~COMDispatchWrapper → IDispatch::Release
2. **Event cleanup**: Unadvise all connections before release
3. **ThreadSafeFunction**: Proper finalization callbacks
4. **Exception safety**: RAII ensures cleanup even on errors

### Leak Prevention
- No raw pointers stored without ownership
- All allocations paired with deallocations
- Finalizers for JavaScript objects
- Weak references where appropriate

## Thread Safety

### Apartment Threading
- COM initialized as MTA (COINIT_MULTITHREADED)
- Event callbacks marshaled to main thread via ThreadSafeFunction
- No direct V8 API calls from COM threads

### Synchronization
- ThreadSafeFunction for cross-thread callbacks
- Mutex protection for shared state
- Atomic reference counting

## Error Handling

### HRESULT Checking
- All COM calls check HRESULT
- Errors converted to JavaScript exceptions
- Detailed error messages with context

### Exception Safety
- Strong exception guarantee where possible
- RAII for resource cleanup
- No resource leaks on exception paths

## Project Structure

```
node-winautomation/
├── src/
│   ├── com/
│   │   ├── COMLifecycle.h/cc          # COM initialization
│   │   ├── COMDispatchWrapper.h/cc    # IDispatch wrapper
│   │   ├── COMEventSink.h/cc          # Event handling
│   │   ├── COMVariant.h/cc            # VARIANT conversions
│   │   └── COMObject.h/cc             # N-API wrapper
│   ├── connectors/
│   │   ├── OutlookConnector.h/cc      # Outlook-specific
│   │   ├── ExcelConnector.h/cc        # Excel-specific
│   │   └── WordConnector.h/cc         # Word-specific
│   └── automation/
│       └── (existing UI Automation code)
├── lib/
│   ├── com/
│   │   ├── COMObject.js               # Base class
│   │   ├── OutlookConnector.js        # Outlook wrapper
│   │   └── ExcelConnector.js          # Excel wrapper
│   └── automation/
│       └── (existing exports)
├── examples/
│   ├── com/
│   │   ├── outlook-basic.js
│   │   ├── outlook-events.js
│   │   └── excel-basic.js
│   └── automation/
│       └── (existing examples)
└── docs/
    ├── COM_API.md
    ├── CONNECTORS.md
    └── MEMORY_MANAGEMENT.md
```

## Design Principles

1. **Single Responsibility**: Each class has one clear purpose
2. **RAII**: Resource Acquisition Is Initialization
3. **Fail Fast**: Errors detected and reported immediately
4. **No Surprises**: Predictable behavior, clear documentation
5. **Performance**: Minimal overhead, efficient conversions
6. **Safety**: Memory-safe, thread-safe, exception-safe

## Testing Strategy

1. **Unit Tests**: Each component tested in isolation
2. **Integration Tests**: End-to-end scenarios
3. **Memory Tests**: Valgrind/ASAN for leak detection
4. **Stress Tests**: Long-running event subscriptions
5. **Error Tests**: Exception paths and error recovery
