#pragma once

#include "../Shared.h"

class RectWrapper : public Napi::ObjectWrap<RectWrapper>
{
public:
  static Napi::FunctionReference *Initialize(Napi::Env env);
  static Napi::Value New(Napi::Env env, RECT *rect);

  RECT m_pRECT;

  RectWrapper(const Napi::CallbackInfo &info);

  Napi::Value GetBottom(const Napi::CallbackInfo &info);
  Napi::Value GetLeft(const Napi::CallbackInfo &info);
  Napi::Value GetRight(const Napi::CallbackInfo &info);
  Napi::Value GetTop(const Napi::CallbackInfo &info);
};
