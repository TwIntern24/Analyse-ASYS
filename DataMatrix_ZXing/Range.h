/*
* Copyright 2022 Axel Waggershauser
*/
// SPDX-License-Identifier: Apache-2.0

#pragma once

#include "DataMatrix_ZXing/ZXAlgorithms.h"

#include <iterator>

namespace ZXing {

template <typename Iterator>
struct StrideIter
{
	Iterator pos;
	int stride;

	using iterator_category = std::random_access_iterator_tag;
	using difference_type   = typename std::iterator_traits<Iterator>::difference_type;
	using value_type        = typename std::iterator_traits<Iterator>::value_type;
	using pointer           = Iterator;
	using reference         = typename std::iterator_traits<Iterator>::reference;

        auto operator*() const -> decltype(*pos) { return *pos; }
        auto operator[](int i) const -> decltype(*(pos + i * stride)) { return *(pos + i * stride); }
	StrideIter<Iterator>& operator++() { return pos += stride, *this; }
	StrideIter<Iterator> operator++(int) { auto temp = *this; ++*this; return temp; }
	bool operator!=(const StrideIter<Iterator>& rhs) const { return pos != rhs.pos; }
	StrideIter<Iterator> operator+(int i) const { return {pos + i * stride, stride}; }
	StrideIter<Iterator> operator-(int i) const { return {pos - i * stride, stride}; }
	int operator-(const StrideIter<Iterator>& rhs) const { return narrow_cast<int>((pos - rhs.pos) / stride); }
        StrideIter(const Iterator& iter, int stride) : pos(iter), stride(stride) {}
};

template <typename Iterator>
StrideIter<Iterator> make_StrideIter(const Iterator& iter, int stride)  //StrideIter(const Iterator&, int) -> StrideIter<Iterator>;
{
    return StrideIter<Iterator>(iter, stride);
}


template <typename Iterator>
struct Range
{
	Iterator _begin, _end;

	Range(Iterator b, Iterator e) : _begin(b), _end(e) {}

	template <typename C>
	Range(const C& c) : _begin(std::begin(c)), _end(std::end(c)) {}

	Iterator begin() const noexcept { return _begin; }
	Iterator end() const noexcept { return _end; }
	explicit operator bool() const { return begin() < end(); }
	int size() const { return narrow_cast<int>(end() - begin()); }
};

template <typename C>
Range<typename C::const_iterator> make_Range(const C& c)        //Range(const C&) -> Range<typename C::const_iterator>;
{
    return Range<typename C::const_iterator>(std::begin(c), std::end(c));
}

} // namespace ZXing
