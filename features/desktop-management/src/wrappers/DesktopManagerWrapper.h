#pragma once

#include "../../../ui-automation/src/shared.h"
#include "ChildSessionWrapper.h"
#include "RDPClientWrapper.h"
#include <Windows.h>

class DesktopManagerWrapper : public Napi::ObjectWrap<DesktopManagerWrapper>
{
private:
    Napi::Reference<Napi::Object> childSessionRef;
    Napi::Reference<Napi::Object> rdpClientRef;
    DWORD sessionId;
    bool isManaged;

public:
    static Napi::FunctionReference* Initialize(Napi::Env env);
    DesktopManagerWrapper(const Napi::CallbackInfo& info);
    ~DesktopManagerWrapper();

    Napi::Value CreateDesktopInWindow(const Napi::CallbackInfo& info);
    Napi::Value GetChildSession(const Napi::CallbackInfo& info);
    Napi::Value GetRDPClient(const Napi::CallbackInfo& info);
    Napi::Value LaunchApplication(const Napi::CallbackInfo& info);
    Napi::Value ResizeDesktop(const Napi::CallbackInfo& info);
    Napi::Value Cleanup(const Napi::CallbackInfo& info);
    Napi::Value GetSessionId(const Napi::CallbackInfo& info);
    Napi::Value IsConnected(const Napi::CallbackInfo& info);
};
