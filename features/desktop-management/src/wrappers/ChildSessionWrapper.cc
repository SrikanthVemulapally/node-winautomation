#include "ChildSessionWrapper.h"
#include <TlHelp32.h>
#include <userenv.h>

#pragma comment(lib, "Wtsapi32.lib")
#pragma comment(lib, "Userenv.lib")

Napi::FunctionReference* ChildSessionWrapper::Initialize(Napi::Env env)
{
    Napi::Function func = DefineClass(env, "ChildSession", {
        InstanceMethod("getSessionId", &ChildSessionWrapper::GetSessionId),
        InstanceMethod("isActive", &ChildSessionWrapper::IsActive),
        InstanceMethod("terminate", &ChildSessionWrapper::Terminate),
        InstanceMethod("getProcesses", &ChildSessionWrapper::GetProcesses),
        InstanceMethod("launchProcess", &ChildSessionWrapper::LaunchProcess),
        InstanceMethod("getSessionInfo", &ChildSessionWrapper::GetSessionInfo),
    });

    Napi::FunctionReference* constructor = new Napi::FunctionReference();
    *constructor = Napi::Persistent(func);

    return constructor;
}

ChildSessionWrapper::ChildSessionWrapper(const Napi::CallbackInfo& info)
    : Napi::ObjectWrap<ChildSessionWrapper>(info), sessionId(0), isActive(false)
{
    Napi::Env env = info.Env();

    if (info.Length() > 0 && info[0].IsNumber())
    {
        sessionId = info[0].As<Napi::Number>().Uint32Value();
        isActive = true;
    }
    else
    {
        DWORD newSessionId = 0;
        BOOL result = WTSCreateChildSession(WTS_CURRENT_SERVER_HANDLE, &newSessionId);
        
        if (!result)
        {
            DWORD error = GetLastError();
            throw Napi::Error::New(env, "Failed to create child session. Error: " + std::to_string(error));
        }

        sessionId = newSessionId;
        isActive = true;
    }
}

ChildSessionWrapper::~ChildSessionWrapper()
{
    if (isActive)
    {
        WTSLogoffSession(WTS_CURRENT_SERVER_HANDLE, sessionId, FALSE);
    }
}

Napi::Value ChildSessionWrapper::GetSessionId(const Napi::CallbackInfo& info)
{
    Napi::Env env = info.Env();
    return Napi::Number::New(env, sessionId);
}

Napi::Value ChildSessionWrapper::IsActive(const Napi::CallbackInfo& info)
{
    Napi::Env env = info.Env();
    
    WTS_CONNECTSTATE_CLASS state;
    DWORD bytesReturned = 0;
    LPWSTR buffer = nullptr;

    if (WTSQuerySessionInformationW(WTS_CURRENT_SERVER_HANDLE, sessionId, 
        WTSConnectState, &buffer, &bytesReturned))
    {
        state = *reinterpret_cast<WTS_CONNECTSTATE_CLASS*>(buffer);
        WTSFreeMemory(buffer);
        
        isActive = (state == WTSActive || state == WTSConnected);
    }
    else
    {
        isActive = false;
    }

    return Napi::Boolean::New(env, isActive);
}

Napi::Value ChildSessionWrapper::Terminate(const Napi::CallbackInfo& info)
{
    Napi::Env env = info.Env();

    if (!isActive)
    {
        return Napi::Boolean::New(env, false);
    }

    BOOL result = WTSLogoffSession(WTS_CURRENT_SERVER_HANDLE, sessionId, FALSE);
    
    if (result)
    {
        isActive = false;
    }

    return Napi::Boolean::New(env, result == TRUE);
}

Napi::Value ChildSessionWrapper::GetProcesses(const Napi::CallbackInfo& info)
{
    Napi::Env env = info.Env();
    Napi::Array processArray = Napi::Array::New(env);

    PWTS_PROCESS_INFOW pProcessInfo = nullptr;
    DWORD processCount = 0;

    if (WTSEnumerateProcessesW(WTS_CURRENT_SERVER_HANDLE, 0, 1, &pProcessInfo, &processCount))
    {
        int index = 0;
        for (DWORD i = 0; i < processCount; i++)
        {
            if (pProcessInfo[i].SessionId == sessionId)
            {
                Napi::Object processObj = Napi::Object::New(env);
                processObj.Set("processId", Napi::Number::New(env, pProcessInfo[i].ProcessId));
                processObj.Set("processName", Napi::String::New(env, 
                    reinterpret_cast<const char16_t*>(pProcessInfo[i].pProcessName)));
                processObj.Set("sessionId", Napi::Number::New(env, pProcessInfo[i].SessionId));
                
                processArray[index++] = processObj;
            }
        }

        WTSFreeMemory(pProcessInfo);
    }

    return processArray;
}

Napi::Value ChildSessionWrapper::LaunchProcess(const Napi::CallbackInfo& info)
{
    Napi::Env env = info.Env();

    if (info.Length() < 1 || !info[0].IsString())
    {
        throw Napi::TypeError::New(env, "String expected for process path");
    }

    std::wstring processPath = info[0].As<Napi::String>().Utf16Value();
    std::wstring commandLine = processPath;
    
    if (info.Length() > 1 && info[1].IsString())
    {
        commandLine += L" " + info[1].As<Napi::String>().Utf16Value();
    }

    HANDLE hToken = nullptr;
    if (!WTSQueryUserToken(sessionId, &hToken))
    {
        throw Napi::Error::New(env, "Failed to get user token for session");
    }

    STARTUPINFOW si = { sizeof(si) };
    PROCESS_INFORMATION pi = { 0 };
    
    si.lpDesktop = const_cast<LPWSTR>(L"winsta0\\default");

    LPVOID pEnvironment = nullptr;
    CreateEnvironmentBlock(&pEnvironment, hToken, FALSE);

    BOOL result = CreateProcessAsUserW(
        hToken,
        nullptr,
        const_cast<LPWSTR>(commandLine.c_str()),
        nullptr,
        nullptr,
        FALSE,
        CREATE_UNICODE_ENVIRONMENT | CREATE_NEW_CONSOLE,
        pEnvironment,
        nullptr,
        &si,
        &pi
    );

    if (pEnvironment)
    {
        DestroyEnvironmentBlock(pEnvironment);
    }

    CloseHandle(hToken);

    if (!result)
    {
        DWORD error = GetLastError();
        throw Napi::Error::New(env, "Failed to launch process. Error: " + std::to_string(error));
    }

    Napi::Object processInfo = Napi::Object::New(env);
    processInfo.Set("processId", Napi::Number::New(env, pi.dwProcessId));
    processInfo.Set("threadId", Napi::Number::New(env, pi.dwThreadId));

    CloseHandle(pi.hProcess);
    CloseHandle(pi.hThread);

    return processInfo;
}

Napi::Value ChildSessionWrapper::GetSessionInfo(const Napi::CallbackInfo& info)
{
    Napi::Env env = info.Env();
    Napi::Object sessionInfo = Napi::Object::New(env);

    sessionInfo.Set("sessionId", Napi::Number::New(env, sessionId));

    LPWSTR buffer = nullptr;
    DWORD bytesReturned = 0;

    if (WTSQuerySessionInformationW(WTS_CURRENT_SERVER_HANDLE, sessionId, 
        WTSUserName, &buffer, &bytesReturned))
    {
        sessionInfo.Set("userName", Napi::String::New(env, reinterpret_cast<const char16_t*>(buffer)));
        WTSFreeMemory(buffer);
    }

    if (WTSQuerySessionInformationW(WTS_CURRENT_SERVER_HANDLE, sessionId, 
        WTSDomainName, &buffer, &bytesReturned))
    {
        sessionInfo.Set("domainName", Napi::String::New(env, reinterpret_cast<const char16_t*>(buffer)));
        WTSFreeMemory(buffer);
    }

    WTS_CONNECTSTATE_CLASS* pState = nullptr;
    if (WTSQuerySessionInformationW(WTS_CURRENT_SERVER_HANDLE, sessionId, 
        WTSConnectState, reinterpret_cast<LPWSTR*>(&pState), &bytesReturned))
    {
        const char* stateStr = "Unknown";
        switch (*pState)
        {
            case WTSActive: stateStr = "Active"; break;
            case WTSConnected: stateStr = "Connected"; break;
            case WTSConnectQuery: stateStr = "ConnectQuery"; break;
            case WTSShadow: stateStr = "Shadow"; break;
            case WTSDisconnected: stateStr = "Disconnected"; break;
            case WTSIdle: stateStr = "Idle"; break;
            case WTSListen: stateStr = "Listen"; break;
            case WTSReset: stateStr = "Reset"; break;
            case WTSDown: stateStr = "Down"; break;
            case WTSInit: stateStr = "Init"; break;
        }
        sessionInfo.Set("state", Napi::String::New(env, stateStr));
        WTSFreeMemory(pState);
    }

    return sessionInfo;
}
