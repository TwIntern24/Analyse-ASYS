/*
* Copyright 2022 Axel Waggershauser
*/
// SPDX-License-Identifier: Apache-2.0

#pragma once

#include <string>
#include <iterator>

namespace ZXing {

std::string ToUtf8(std::wstring str);
std::wstring FromUtf8(std::string utf8);
#if __cplusplus > 201703L
std::wstring FromUtf8(std::u8string utf8);
#endif

std::wstring EscapeNonGraphical(std::wstring str);
std::string EscapeNonGraphical(std::string utf8);

} // namespace ZXing
