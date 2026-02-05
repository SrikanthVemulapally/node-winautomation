#pragma once

#include "../../../ui-automation/src/shared.h"
#include <Windows.h>
#include <MsTscAx.h>
#include <atlbase.h>
#include <atlcom.h>

class RDPClientWrapper : public Napi::ObjectWrap<RDPClientWrapper>
{
private:
    CComPtr<IMsRdpClient10> rdpClient;
    HWND containerWindow;
    HWND rdpWindow;
    bool isConnected;
    bool isFullscreen;

public:
    static Napi::FunctionReference* Initialize(Napi::Env env);
    RDPClientWrapper(const Napi::CallbackInfo& info);
    ~RDPClientWrapper();

    Napi::Value Connect(const Napi::CallbackInfo& info);
    Napi::Value Disconnect(const Napi::CallbackInfo& info);
    Napi::Value SetServer(const Napi::CallbackInfo& info);
    Napi::Value SetPort(const Napi::CallbackInfo& info);
    Napi::Value SetDesktopSize(const Napi::CallbackInfo& info);
    Napi::Value SetFullscreen(const Napi::CallbackInfo& info);
    Napi::Value GetConnectionState(const Napi::CallbackInfo& info);
    Napi::Value EmbedInWindow(const Napi::CallbackInfo& info);
    Napi::Value SendKeys(const Napi::CallbackInfo& info);
    Napi::Value GetWindowHandle(const Napi::CallbackInfo& info);
    Napi::Value SetColorDepth(const Napi::CallbackInfo& info);
    Napi::Value SetAuthenticationLevel(const Napi::CallbackInfo& info);
    Napi::Value EnableCredentialSaving(const Napi::CallbackInfo& info);
    Napi::Value SetUsername(const Napi::CallbackInfo& info);
    Napi::Value SetDomain(const Napi::CallbackInfo& info);
    Napi::Value Reconnect(const Napi::CallbackInfo& info);
};
