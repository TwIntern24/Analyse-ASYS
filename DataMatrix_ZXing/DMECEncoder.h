/*
* Copyright 2016 Huy Cuong Nguyen
* Copyright 2006 Jeremias Maerki.
*/
// SPDX-License-Identifier: Apache-2.0

#pragma once

namespace ZXing {

class ByteArray;

namespace DataMatrix {

class SymbolInfo;


//↓ Creates and interleaves the ECC200 error correction for an encoded message. Return the codewords with interleaved error correction.
void EncodeECC200(ByteArray& codewords, const SymbolInfo& symbolInfo);  //codewords: the codewords, symbolInfo: information about the symbol to be encoded

} // DataMatrix
} // ZXing
