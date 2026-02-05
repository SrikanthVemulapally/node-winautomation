#pragma once

#include "../../../ui-automation/src/shared.h"
#include <Windows.h>
#include <WtsApi32.h>

class ChildSessionWrapper : public Napi::ObjectWrap<ChildSessionWrapper>
{
private:
    DWORD sessionId;
    bool isActive;

public:
    static Napi::FunctionReference* Initialize(Napi::Env env);
    ChildSessionWrapper(const Napi::CallbackInfo& info);
    ~ChildSessionWrapper();

    Napi::Value GetSessionId(const Napi::CallbackInfo& info);
    Napi::Value IsActive(const Napi::CallbackInfo& info);
    Napi::Value Terminate(const Napi::CallbackInfo& info);
    Napi::Value GetProcesses(const Napi::CallbackInfo& info);
    Napi::Value LaunchProcess(const Napi::CallbackInfo& info);
    Napi::Value GetSessionInfo(const Napi::CallbackInfo& info);
};
