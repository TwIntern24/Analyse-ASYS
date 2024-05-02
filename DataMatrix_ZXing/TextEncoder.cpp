/*
* Copyright 2016 Nu-book Inc.
*/
// SPDX-License-Identifier: Apache-2.0

#include "DataMatrix_ZXing/TextEncoder.h"
#include "DataMatrix_ZXing/CharacterSet.h"
#include "DataMatrix_ZXing/ECI.h"
#include "DataMatrix_ZXing/Utf.h"
#include "DataMatrix_ZXing/ZXAlgorithms.h"
#include "DataMatrix_ZXing/libzueci/zueci.h"

#include <stdexcept>
#include <algorithm>

namespace ZXing {

void TextEncoder::GetBytes(const std::string& str, CharacterSet charset, std::string& bytes)
{
	int eci = ToInt(ToECI(charset));
	const int str_len = narrow_cast<int>(str.length());
	int eci_len;

	if (eci == -1)
		eci = 899; // Binary

	bytes.clear();

    int error_number = zueci_dest_len_eci(eci, reinterpret_cast<const unsigned char *>(str.data()), str_len, &eci_len);

    if (error_number >= ZUECI_ERROR) // Shouldn't happen
        throw std::logic_error("Internal error `zueci_dest_len_eci()`");

    bytes.resize(eci_len); // Sufficient but approximate length

    std::copy(str.begin(), str.end(), bytes.begin());
    error_number = zueci_utf8_to_eci(eci,
                                     reinterpret_cast<const unsigned char*>(str.data()),
                                     str_len,
                                     reinterpret_cast<unsigned char*>(&bytes[0]),
                                     &eci_len);

    if (error_number >= ZUECI_ERROR) {
        bytes.clear();
        throw std::invalid_argument("Unexpected charcode");
    }

    bytes.resize(eci_len); // Actual length
}

void TextEncoder::GetBytes(const std::wstring& str, CharacterSet charset, std::string& bytes)
{
    GetBytes(ToUtf8(str), charset, bytes);
}

} // ZXing
