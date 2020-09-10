//  Copyright 2020 John Buonagurio
//
//  Distributed under the Boost Software License, Version 1.0.
//
//  See accompanying file LICENSE_1_0.txt or copy at
//  http://www.boost.org/LICENSE_1_0.txt

#pragma once

#include <xll/config.hpp>

#include <xll/constants.hpp>
#include <xll/callback.hpp>
#include <xll/detail/assert.hpp>

#include <type_traits>
#include <utility>

#include <boost/mp11/integral.hpp>
#include <boost/mp11/list.hpp>
#include <boost/mp11/utility.hpp>
#include <boost/mp11/algorithm.hpp>

namespace xll {
namespace detail {

namespace mp11 = boost::mp11;

// Overload resolution based on P0608R3: "A sane variant converting constructor"
// https://wg21.link/p0608

template <class T> using remove_cvref_t = std::remove_cv_t<std::remove_reference_t<T>>;

struct narrowing_check_impl
{
    template<class T>
    static auto check(T (&&)[1]) -> mp11::mp_identity<T>;

    template<class T, class U>
    using fn = decltype(check<T>({ std::declval<U>() }));
};

template<class T, class U> using narrowing_check = mp11::mp_if<std::is_arithmetic<T>, mp11::mp_identity<T>, mp11::mp_invoke_q<narrowing_check_impl, T, U>>;

template<class... T> struct overload;

template<> struct overload<>
{
    void operator()() const;
};

template<class T1, class... T> struct overload<T1, T...> : overload<T...>
{
    using overload<T...>::operator();
    template<class U> auto operator()(T1, U&&) const -> narrowing_check<T1, U>;
};

template<class... T> struct overload<bool, T...> : overload<T...>
{
    using overload<T...>::operator();
    template<class U> auto operator()(bool, U&&) const -> mp11::mp_if<std::is_same<remove_cvref_t<U>, bool>, mp11::mp_identity<bool>>;
};

template<class... T> struct overload<const bool, T...> : overload<T...>
{
    using overload<T...>::operator();
    template<class U> auto operator()(bool, U&&) const -> mp11::mp_if<std::is_same<remove_cvref_t<U>, bool>, mp11::mp_identity<const bool>>;
};

template<class... T> struct overload<volatile bool, T...> : overload<T...>
{
    using overload<T...>::operator();
    template<class U> auto operator()(bool, U&&) const -> mp11::mp_if<std::is_same<remove_cvref_t<U>, bool>, mp11::mp_identity<volatile bool>>;
};

template<class... T> struct overload<const volatile bool, T...> : overload<T...>
{
    using overload<T...>::operator();
    template<class U> auto operator()(bool, U&&) const -> mp11::mp_if<std::is_same<remove_cvref_t<U>, bool>, mp11::mp_identity<const volatile bool>>;
};

template<class U, class... T> using resolve_overload_type = typename decltype(overload<T...>()(std::declval<U>(), std::declval<U>()))::type;
template<class U, class... T> using resolve_overload_index = mp11::mp_find<mp11::mp_list<T...>, resolve_overload_type<U, T...>>;


// Variant implementation based on Boost.Variant2:
// Copyright 2017-2019 Peter Dimov.
// Distributed under the Boost Software License, Version 1.0.
// https://github.com/boostorg/variant2

// +==================================+=======+=======+
// | query                            |   x86 | amd64 |
// +==================================+=======+=======+
// | sizeof(XLOPER12::val.num)        |     8 |     8 |
// | sizeof(XLOPER12::val.str)        |     4 |     8 |
// | sizeof(XLOPER12::val.xbool)      |     4 |     4 |
// | sizeof(XLOPER12::val.err)        |     4 |     4 |
// | sizeof(XLOPER12::val.w)          |     4 |     4 |
// | sizeof(XLOPER12::val.sref)       |    20 |    20 |
// | sizeof(XLOPER12::val.mref)       |     8 |    16 |
// | sizeof(XLOPER12::val.array)      |    12 |    16 |
// | sizeof(XLOPER12::val.flow)       |    16 |    24 |
// | sizeof(XLOPER12::val.bigdata)    |     8 |    16 |
// +----------------------------------+-------+-------+
// | sizeof(XLOPER12::val)            |    24 |    24 |
// | alignof(decltype(XLOPER12::val)) |     8 |     8 |
// +==================================+=======+=======+

template <class... T> union variant_storage_impl;

template <class... T> using variant_storage = variant_storage_impl<T...>;

template<> union variant_storage_impl<> {};

template <class T1, class... T> union variant_storage_impl<T1, T...>
{
    T1 first_;
    variant_storage<T...> rest_;
    unsigned char padding_[24];

    template <class... Args>
    constexpr explicit variant_storage_impl(mp11::mp_size_t<0>, Args&&... args)
        : first_(std::forward<Args>(args)...) {}

    template <std::size_t I, class... Args>
    constexpr explicit variant_storage_impl(mp11::mp_size_t<I>, Args&&... args)
        : rest_(mp11::mp_size_t<I-1>(), std::forward<Args>(args)...) {}

    ~variant_storage_impl() {}

    template <class... Args>
    void emplace(mp11::mp_size_t<0>, Args&&... args)
        { ::new(&first_) T1(std::forward<Args>(args)...); }

    template <std::size_t I, class... Args>
    void emplace(mp11::mp_size_t<I>, Args&&... args)
        { rest_.emplace(mp11::mp_size_t<I-1>(), std::forward<Args>(args)...); }

    constexpr T1& get(mp11::mp_size_t<0>) noexcept { return first_; }
    constexpr const T1& get(mp11::mp_size_t<0>) const noexcept { return first_; }

    template <std::size_t I>
    constexpr mp11::mp_at_c<mp11::mp_list<T...>, I-1>& get(mp11::mp_size_t<I>) noexcept
        { return rest_.get(mp11::mp_size_t<I-1>()); }

    template <std::size_t I>
    constexpr const mp11::mp_at_c<mp11::mp_list<T...>, I-1>& get(mp11::mp_size_t<I>) const noexcept
        { return rest_.get(mp11::mp_size_t<I-1>()); }
};

} // namespace detail

template <class T> struct in_place_type_t {};
template <class T> constexpr in_place_type_t<T> in_place_type{};

namespace detail {

template <class T> struct is_in_place_type: std::false_type {};
template <class T> struct is_in_place_type<in_place_type_t<T>>: std::true_type {};

struct variant_common_type {};

template <class... Ts>
struct variant_base_impl : variant_common_type
{
protected:
    variant_storage<Ts...> st_;
    uint32_t xltype_; // DWORD

    // XLOPER12-compatible aligned storage must be 24 bytes.
    static_assert(sizeof(st_) == 24U, "invalid variant size");

    struct copy_construct_impl
    {
        variant_base_impl *self_;
        const variant_base_impl& other_;
        template <class I> constexpr void operator()(I)
        {
            ::new(&self_->st_) variant_storage<Ts...>(I(), other_.st_.get(I()));
            self_->xltype_ = other_.xltype_;
        }
    };

    struct move_construct_impl
    {
        variant_base_impl *self_;
        variant_base_impl& other_;
        template <class I> constexpr void operator()(I) noexcept
        {
            ::new(&self_->st_) variant_storage<Ts...>(I(), std::move(other_.st_.get(I())));
            self_->xltype_ = std::exchange(other_.xltype_, XLTYPE::xltypeMissing);
        }
    };

    struct destroy_impl
    {
        variant_base_impl *self_;
        template <class I> void operator()(I) const noexcept
        {
            using T = mp11::mp_at<mp11::mp_list<Ts...>, I>;
            switch (T::xltype::value) {
            case XLTYPE::xltypeStr:
            case XLTYPE::xltypeRef:
            case XLTYPE::xltypeMulti:
                if (self_->flags() & xlbitXLFree) {
                    //Excel12(xlFree, nullptr, self_); // Excel allocated memory
                }
                else {
                    self_->st_.get(I()).~T();
                    self_->xltype_ = xltypeMissing;
                }
                break;
            default:
                break;
            }
        }
    };
    
    constexpr std::size_t index_impl(mp11::mp_list<>, uint32_t) const noexcept
        { return 0; }

    template <class U1, class... U>
    constexpr std::size_t index_impl(mp11::mp_list<U1, U...>, uint32_t xltype) const noexcept
        { return (xltype == U1::xltype::value) ? 0 : index_impl(mp11::mp_list<U...>(), xltype) + 1; }

    constexpr std::size_t index_impl(uint32_t xltype) const noexcept
        { return index_impl(mp11::mp_list<Ts...>(), xltype); }

    // Constructs value directly in storage. Variant must be empty.
    template <class U, class... Args> void replace(Args&&... args)
    {
        static_assert(mp11::mp_contains<mp11::mp_list<Ts...>, U>::value, "invalid variant alternative type");
        ::new(&st_) variant_storage<Ts...>(mp11::mp_find<mp11::mp_list<Ts...>, U>(), std::forward<Args>(args)...);
        xltype_ = U::xltype::value;
    }

public:
    constexpr variant_base_impl() noexcept
        : xltype_(XLTYPE::xltypeMissing), st_(mp11::mp_find<mp11::mp_list<Ts...>, xlmissing>()) {}

    // Constructs the variant with the alternative U.
    template <class U, class... Args>
    constexpr explicit variant_base_impl(in_place_type_t<U>, Args&&... args)
        : xltype_(U::xltype::value), st_(mp11::mp_find<mp11::mp_list<Ts...>, U>(), std::forward<Args>(args)...) {}

    void destroy() noexcept
        { boost::mp11::mp_with_index<sizeof...(Ts)>(index(), destroy_impl{this}); }

    void release() noexcept
        { destroy(); }
    
    ~variant_base_impl() noexcept
        { destroy(); }

    // Returns index of current stored type.
    constexpr std::size_t index() const noexcept
        { return index_impl(mp11::mp_list<Ts...>(), xltype()); }
    
    // Construct value in-place.
    template <class U, class... Args> void emplace(Args&&... args)
    {
        static_assert(mp11::mp_contains<mp11::mp_list<Ts...>, U>::value, "invalid variant alternative type");
        U temp(std::forward<Args>(args)...);
        destroy();
        st_.emplace(mp11::mp_find<mp11::mp_list<Ts...>, U>(), std::move(temp));
        xltype_ = U::xltype::value;
    }

    template <class U> constexpr U& get() noexcept
    {
        static_assert(mp11::mp_contains<mp11::mp_list<Ts...>, U>::value, "invalid variant alternative type");
        XLL_ASSERT(xltype() == U::xltype::value);
        return st_.get(mp11::mp_find<mp11::mp_list<Ts...>, U>());
    }

    template <class U> constexpr const U& get() const noexcept
    {
        static_assert(mp11::mp_contains<mp11::mp_list<Ts...>, U>::value, "invalid variant alternative type");
        //XLL_ASSERT(xltype() == U::xltype::value);
        return st_.get(mp11::mp_find<mp11::mp_list<Ts...>, U>());
    }

    // Extensions for variant containing a single type.
    template <class L = mp11::mp_list<Ts...>,
        class E = std::enable_if_t<std::is_same_v<mp11::mp_size<L>, mp11::mp_size_t<1>>>,
        class U = boost::mp11::mp_front<L>>
    constexpr U& value() noexcept
        { return st_.get(mp11::mp_find<L, U>()); }

    template <class L = mp11::mp_list<Ts...>,
        class E = std::enable_if_t<std::is_same_v<mp11::mp_size<L>, mp11::mp_size_t<1>>>,
        class U = boost::mp11::mp_front<L>>
    constexpr const U& value() const noexcept
        { return st_.get(mp11::mp_find<L, U>()); }
    
    // Returns xltype without xlbit flags (xlbitXLFree, xlbitDLLFree).
    constexpr uint32_t xltype() const noexcept
        { return (xltype_ & 0x0FFF); }

    // Returns xlbit flags (xlbitXLFree, xlbitDLLFree).
    constexpr uint32_t flags() const noexcept
        { return (xltype_ & 0xF000); }

    constexpr void set_flags(uint32_t flags) noexcept
        { xltype_ |= (flags & 0xF000); }

    constexpr void clear_flags(uint32_t flags) noexcept
        { xltype_ &= ~(flags & 0xF000); }
};

template <class... Ts>
struct variant_base : variant_base_impl<Ts...>
{
    using variant_base_impl = variant_base_impl<Ts...>;

    variant_base() = default;
    
    constexpr variant_base(const variant_base& other) noexcept(mp11::mp_all<std::is_nothrow_copy_constructible<Ts>...>::value)
        { mp11::mp_with_index<sizeof...(Ts)>(other.index(), copy_construct_impl{this, other}); }

    constexpr variant_base(variant_base&& other) noexcept
        { mp11::mp_with_index<sizeof...(Ts)>(other.index(), move_construct_impl{this, other}); }

    constexpr variant_base& operator=(const variant_base& rhs) noexcept(mp11::mp_all<std::is_nothrow_copy_constructible<Ts>...>::value)
    {
        destroy();
        mp11::mp_with_index<sizeof...(Ts)>(rhs.index(), copy_construct_impl{*this, rhs});
        return *this;
    }

    constexpr variant_base& operator=(variant_base&& rhs) noexcept
    {
        destroy();
        mp11::mp_with_index<sizeof...(Ts)>(rhs.index(), move_construct_impl{*this, rhs});
        return *this;
    }

    // Converting move constructor for stored value types.
    template <class U,
        class Ud = std::decay_t<U>,
        class E1 = std::enable_if_t<!std::is_same_v<Ud, variant_base> && !is_in_place_type<Ud>::value>,
        class V = resolve_overload_type<U&&, Ts...>,
        class E2 = std::enable_if_t<std::is_constructible_v<V, U&&>>>
    constexpr variant_base(U&& val) noexcept
        : variant_base_impl(in_place_type_t<V>(), std::forward<U>(val)) {}

    // Converting move assignment operator for stored value types.
    template<class U,
        class Ud = std::decay_t<U>,
        class E1 = std::enable_if_t<!std::is_same_v<Ud, variant_base>>,
        class V = resolve_overload_type<U, Ts...>,
        class E2 = std::enable_if_t<std::is_assignable_v<V&, U&&> && std::is_constructible_v<V, U&&>>>
    constexpr variant_base& operator=(U&& rhs) noexcept(std::is_nothrow_constructible_v<V, U&&>)
    {
        emplace<V>(std::forward<U>(rhs));
        return *this;
    }
};

struct placeholder {
    using xltype = std::integral_constant<XLTYPE, xltypeMissing>;
};

using variant_opaque_ptr = detail::variant_base<detail::placeholder> *;

} // namespace detail
} // namespace xll
