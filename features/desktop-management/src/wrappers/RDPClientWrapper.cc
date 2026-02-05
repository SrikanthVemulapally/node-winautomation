#include "RDPClientWrapper.h"

#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")

Napi::FunctionReference* RDPClientWrapper::Initialize(Napi::Env env)
{
    Napi::Function func = DefineClass(env, "RDPClient", {
        InstanceMethod("connect", &RDPClientWrapper::Connect),
        InstanceMethod("disconnect", &RDPClientWrapper::Disconnect),
        InstanceMethod("setServer", &RDPClientWrapper::SetServer),
        InstanceMethod("setPort", &RDPClientWrapper::SetPort),
        InstanceMethod("setDesktopSize", &RDPClientWrapper::SetDesktopSize),
        InstanceMethod("setFullscreen", &RDPClientWrapper::SetFullscreen),
        InstanceMethod("getConnectionState", &RDPClientWrapper::GetConnectionState),
        InstanceMethod("embedInWindow", &RDPClientWrapper::EmbedInWindow),
        InstanceMethod("sendKeys", &RDPClientWrapper::SendKeys),
        InstanceMethod("getWindowHandle", &RDPClientWrapper::GetWindowHandle),
        InstanceMethod("setColorDepth", &RDPClientWrapper::SetColorDepth),
        InstanceMethod("setAuthenticationLevel", &RDPClientWrapper::SetAuthenticationLevel),
        InstanceMethod("enableCredentialSaving", &RDPClientWrapper::EnableCredentialSaving),
        InstanceMethod("setUsername", &RDPClientWrapper::SetUsername),
        InstanceMethod("setDomain", &RDPClientWrapper::SetDomain),
        InstanceMethod("reconnect", &RDPClientWrapper::Reconnect),
    });

    Napi::FunctionReference* constructor = new Napi::FunctionReference();
    *constructor = Napi::Persistent(func);

    return constructor;
}

RDPClientWrapper::RDPClientWrapper(const Napi::CallbackInfo& info)
    : Napi::ObjectWrap<RDPClientWrapper>(info), containerWindow(nullptr), 
      rdpWindow(nullptr), isConnected(false), isFullscreen(false)
{
    Napi::Env env = info.Env();

    HRESULT hr = rdpClient.CoCreateInstance(__uuidof(MsRdpClient10));
    
    if (FAILED(hr))
    {
        throw Napi::Error::New(env, "Failed to create RDP client instance. Error: 0x" + 
            std::to_string(hr));
    }

    CComPtr<IMsRdpClientAdvancedSettings8> advSettings;
    hr = rdpClient->get_AdvancedSettings9(reinterpret_cast<IMsRdpClientAdvancedSettings8**>(&advSettings));
    
    if (SUCCEEDED(hr))
    {
        advSettings->put_EnableCredSspSupport(VARIANT_TRUE);
        advSettings->put_AuthenticationLevel(0);
        advSettings->put_EnableAutoReconnect(VARIANT_TRUE);
        advSettings->put_MaxReconnectAttempts(5);
    }
}

RDPClientWrapper::~RDPClientWrapper()
{
    if (isConnected && rdpClient)
    {
        rdpClient->Disconnect();
    }
}

Napi::Value RDPClientWrapper::Connect(const Napi::CallbackInfo& info)
{
    Napi::Env env = info.Env();

    if (!rdpClient)
    {
        throw Napi::Error::New(env, "RDP client not initialized");
    }

    HRESULT hr = rdpClient->Connect();
    
    if (FAILED(hr))
    {
        throw Napi::Error::New(env, "Failed to connect. Error: 0x" + std::to_string(hr));
    }

    isConnected = true;
    return Napi::Boolean::New(env, true);
}

Napi::Value RDPClientWrapper::Disconnect(const Napi::CallbackInfo& info)
{
    Napi::Env env = info.Env();

    if (!rdpClient)
    {
        return Napi::Boolean::New(env, false);
    }

    rdpClient->Disconnect();
    isConnected = false;

    return Napi::Boolean::New(env, true);
}

Napi::Value RDPClientWrapper::SetServer(const Napi::CallbackInfo& info)
{
    Napi::Env env = info.Env();

    if (info.Length() < 1 || !info[0].IsString())
    {
        throw Napi::TypeError::New(env, "String expected for server address");
    }

    std::wstring server = info[0].As<Napi::String>().Utf16Value();
    
    HRESULT hr = rdpClient->put_Server(CComBSTR(server.c_str()));
    
    if (FAILED(hr))
    {
        throw Napi::Error::New(env, "Failed to set server");
    }

    return Napi::Boolean::New(env, true);
}

Napi::Value RDPClientWrapper::SetPort(const Napi::CallbackInfo& info)
{
    Napi::Env env = info.Env();

    if (info.Length() < 1 || !info[0].IsNumber())
    {
        throw Napi::TypeError::New(env, "Number expected for port");
    }

    long port = info[0].As<Napi::Number>().Int32Value();
    
    CComPtr<IMsRdpClientAdvancedSettings8> advSettings;
    HRESULT hr = rdpClient->get_AdvancedSettings9(reinterpret_cast<IMsRdpClientAdvancedSettings8**>(&advSettings));
    
    if (SUCCEEDED(hr))
    {
        advSettings->put_RDPPort(port);
    }

    return Napi::Boolean::New(env, SUCCEEDED(hr));
}

Napi::Value RDPClientWrapper::SetDesktopSize(const Napi::CallbackInfo& info)
{
    Napi::Env env = info.Env();

    if (info.Length() < 2 || !info[0].IsNumber() || !info[1].IsNumber())
    {
        throw Napi::TypeError::New(env, "Two numbers expected (width, height)");
    }

    long width = info[0].As<Napi::Number>().Int32Value();
    long height = info[1].As<Napi::Number>().Int32Value();

    HRESULT hr1 = rdpClient->put_DesktopWidth(width);
    HRESULT hr2 = rdpClient->put_DesktopHeight(height);

    return Napi::Boolean::New(env, SUCCEEDED(hr1) && SUCCEEDED(hr2));
}

Napi::Value RDPClientWrapper::SetFullscreen(const Napi::CallbackInfo& info)
{
    Napi::Env env = info.Env();

    if (info.Length() < 1 || !info[0].IsBoolean())
    {
        throw Napi::TypeError::New(env, "Boolean expected");
    }

    bool fullscreen = info[0].As<Napi::Boolean>().Value();
    isFullscreen = fullscreen;

    HRESULT hr = rdpClient->put_FullScreen(fullscreen ? VARIANT_TRUE : VARIANT_FALSE);

    return Napi::Boolean::New(env, SUCCEEDED(hr));
}

Napi::Value RDPClientWrapper::GetConnectionState(const Napi::CallbackInfo& info)
{
    Napi::Env env = info.Env();

    if (!rdpClient)
    {
        return Napi::Number::New(env, -1);
    }

    short state = 0;
    rdpClient->get_Connected(&state);

    return Napi::Number::New(env, state);
}

Napi::Value RDPClientWrapper::EmbedInWindow(const Napi::CallbackInfo& info)
{
    Napi::Env env = info.Env();

    if (info.Length() < 1)
    {
        throw Napi::TypeError::New(env, "Window handle expected");
    }

    HWND parentHwnd = nullptr;
    
    if (info[0].IsNumber())
    {
        parentHwnd = reinterpret_cast<HWND>(info[0].As<Napi::Number>().Int64Value());
    }
    else if (info[0].IsBuffer())
    {
        Napi::Buffer<void*> buffer = info[0].As<Napi::Buffer<void*>>();
        parentHwnd = *reinterpret_cast<HWND*>(buffer.Data());
    }
    else
    {
        throw Napi::TypeError::New(env, "Invalid window handle type");
    }

    if (!IsWindow(parentHwnd))
    {
        throw Napi::Error::New(env, "Invalid window handle");
    }

    containerWindow = parentHwnd;

    CComPtr<IOleObject> oleObject;
    HRESULT hr = rdpClient->QueryInterface(&oleObject);
    
    if (FAILED(hr))
    {
        throw Napi::Error::New(env, "Failed to get OLE object interface");
    }

    CComPtr<IOleWindow> oleWindow;
    hr = rdpClient->QueryInterface(&oleWindow);
    
    if (SUCCEEDED(hr))
    {
        oleWindow->GetWindow(&rdpWindow);
        
        if (rdpWindow)
        {
            SetParent(rdpWindow, parentHwnd);
            
            RECT rect;
            GetClientRect(parentHwnd, &rect);
            SetWindowPos(rdpWindow, nullptr, 0, 0, 
                rect.right - rect.left, rect.bottom - rect.top,
                SWP_NOZORDER | SWP_SHOWWINDOW);
        }
    }

    return Napi::Boolean::New(env, rdpWindow != nullptr);
}

Napi::Value RDPClientWrapper::SendKeys(const Napi::CallbackInfo& info)
{
    Napi::Env env = info.Env();

    if (info.Length() < 1 || !info[0].IsString())
    {
        throw Napi::TypeError::New(env, "String expected for keys");
    }

    std::wstring keys = info[0].As<Napi::String>().Utf16Value();

    CComPtr<IMsRdpClientNonScriptable5> nonScriptable;
    HRESULT hr = rdpClient->QueryInterface(&nonScriptable);
    
    if (FAILED(hr))
    {
        throw Napi::Error::New(env, "Failed to get non-scriptable interface");
    }

    for (wchar_t ch : keys)
    {
        nonScriptable->SendKeys(1, VARIANT_FALSE, ch);
    }

    return Napi::Boolean::New(env, true);
}

Napi::Value RDPClientWrapper::GetWindowHandle(const Napi::CallbackInfo& info)
{
    Napi::Env env = info.Env();

    if (!rdpWindow)
    {
        return env.Null();
    }

    return Napi::Number::New(env, reinterpret_cast<int64_t>(rdpWindow));
}

Napi::Value RDPClientWrapper::SetColorDepth(const Napi::CallbackInfo& info)
{
    Napi::Env env = info.Env();

    if (info.Length() < 1 || !info[0].IsNumber())
    {
        throw Napi::TypeError::New(env, "Number expected for color depth");
    }

    long colorDepth = info[0].As<Napi::Number>().Int32Value();
    
    HRESULT hr = rdpClient->put_ColorDepth(colorDepth);

    return Napi::Boolean::New(env, SUCCEEDED(hr));
}

Napi::Value RDPClientWrapper::SetAuthenticationLevel(const Napi::CallbackInfo& info)
{
    Napi::Env env = info.Env();

    if (info.Length() < 1 || !info[0].IsNumber())
    {
        throw Napi::TypeError::New(env, "Number expected for authentication level");
    }

    long level = info[0].As<Napi::Number>().Int32Value();
    
    CComPtr<IMsRdpClientAdvancedSettings8> advSettings;
    HRESULT hr = rdpClient->get_AdvancedSettings9(reinterpret_cast<IMsRdpClientAdvancedSettings8**>(&advSettings));
    
    if (SUCCEEDED(hr))
    {
        advSettings->put_AuthenticationLevel(level);
    }

    return Napi::Boolean::New(env, SUCCEEDED(hr));
}

Napi::Value RDPClientWrapper::EnableCredentialSaving(const Napi::CallbackInfo& info)
{
    Napi::Env env = info.Env();

    if (info.Length() < 1 || !info[0].IsBoolean())
    {
        throw Napi::TypeError::New(env, "Boolean expected");
    }

    bool enable = info[0].As<Napi::Boolean>().Value();
    
    CComPtr<IMsRdpClientAdvancedSettings8> advSettings;
    HRESULT hr = rdpClient->get_AdvancedSettings9(reinterpret_cast<IMsRdpClientAdvancedSettings8**>(&advSettings));
    
    if (SUCCEEDED(hr))
    {
        advSettings->put_EnableCredSspSupport(enable ? VARIANT_TRUE : VARIANT_FALSE);
    }

    return Napi::Boolean::New(env, SUCCEEDED(hr));
}

Napi::Value RDPClientWrapper::SetUsername(const Napi::CallbackInfo& info)
{
    Napi::Env env = info.Env();

    if (info.Length() < 1 || !info[0].IsString())
    {
        throw Napi::TypeError::New(env, "String expected for username");
    }

    std::wstring username = info[0].As<Napi::String>().Utf16Value();
    
    HRESULT hr = rdpClient->put_UserName(CComBSTR(username.c_str()));

    return Napi::Boolean::New(env, SUCCEEDED(hr));
}

Napi::Value RDPClientWrapper::SetDomain(const Napi::CallbackInfo& info)
{
    Napi::Env env = info.Env();

    if (info.Length() < 1 || !info[0].IsString())
    {
        throw Napi::TypeError::New(env, "String expected for domain");
    }

    std::wstring domain = info[0].As<Napi::String>().Utf16Value();
    
    HRESULT hr = rdpClient->put_Domain(CComBSTR(domain.c_str()));

    return Napi::Boolean::New(env, SUCCEEDED(hr));
}

Napi::Value RDPClientWrapper::Reconnect(const Napi::CallbackInfo& info)
{
    Napi::Env env = info.Env();

    if (!rdpClient)
    {
        throw Napi::Error::New(env, "RDP client not initialized");
    }

    long width = 0, height = 0;
    rdpClient->get_DesktopWidth(&width);
    rdpClient->get_DesktopHeight(&height);

    HRESULT hr = rdpClient->Reconnect(width, height);

    return Napi::Boolean::New(env, SUCCEEDED(hr));
}
