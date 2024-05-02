/*
* Copyright 2016 Huy Cuong Nguyen
* Copyright 2006-2007 Jeremias Maerki.
*/
// SPDX-License-Identifier: Apache-2.0

#pragma once

#include <DataMatrix_ZXing/DMSymbolShape.h>
#include <string>

namespace ZXing {

class ByteArray;

namespace DataMatrix {

enum class SymbolShape;

//↓ DataMatrix ECC 200 data encoder following the algorithm described in ISO/IEC 16022:200(E) in annex S.
ByteArray Encode(const std::wstring& msg);
ByteArray Encode(const std::wstring& msg, ZXing::DataMatrix::SymbolShape shape, int minWidth, int minHeight, int maxWidth, int maxHeight);

} // DataMatrix
} // ZXing
