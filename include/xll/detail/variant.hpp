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
#include <xll/detail/memory.hpp>
#include <xll/detail/tagged_union.hpp>

#include <boost/mp11/list.hpp>
#include <boost/mp11/algorithm.hpp>
#include <boost/mp11/integer_sequence.hpp>

#include <algorithm>
#include <array>
#include <cstdint>
#include <limits>
#include <new>
#include <string>
#include <type_traits>
#include <utility>

namespace xll {
namespace detail {

template <class... Ts>
struct variant : public tagged_union
{
    template <class T>
    using xltype_t = typename T::xltype;
    
    using types = boost::mp11::mp_list<Ts...>;
    using xltypes = boost::mp11::mp_transform<xltype_t, types>;

    template <class Tag>
    struct map {
        using index = boost::mp11::mp_find<xltypes, xltype_t<Tag>>;
        using value = boost::mp11::mp_at<types, index>;
    };

    using is_monostate =
      std::is_same<
        boost::mp11::mp_size<types>,
        boost::mp11::mp_size_t<1>>;

    using default_xltype =
      std::conditional_t<
        is_monostate::value,
        typename boost::mp11::mp_front<types>::xltype,
        std::integral_constant<uint32_t, xltypeMissing>>;

protected:
    struct destroy_impl
    {
        variant& self;

        template <class I>
        void operator()(I)
        {
            using T = boost::mp11::mp_at_c<types, I::value>;
            if (self.xltype_ & xlbitXLFree) {
                Excel12(xlFree, nullptr, &self);
            }
            else {
                launder_cast<T*>(&self.val_)->~T();
                self.xltype_ = xltypeMissing;
            }
        }
    };

    template <class... Us>
    struct copy_construct_impl
    {
        variant& self;
        const variant<Us...>& other;

        template <class I>
        constexpr void operator()(I)
        {
            using T = boost::mp11::mp_at_c<types, I::value>;
            ::new(&self.val_) T(
                *launder_cast<const T*>(&other.val_));
            self.set_xltype(other.xltype_);
        }
    };

    template <class... Us>
    struct move_construct_impl
    {
        variant& self;
        variant<Us...>& other;

        template <class I>
        constexpr void operator()(I)
        {
            using T = boost::mp11::mp_at_c<types, I::value>;
            ::new(&self.val_) T(std::move(
                *launder_cast<T*>(&other.val_)));
            self.set_xltype(other.xltype_);
        }
    };
    
    struct equals
    {
        const variant& self;
        const variant& other;

        template<class I>
        constexpr bool operator()(I)
        {
            using T = boost::mp11::mp_at_c<types, I::value>;
            return *launder_cast<const T*>(&self.val_) ==
                   *launder_cast<const T*>(&other.val_);
        }
    };
    
    void destroy()
    {
        boost::mp11::mp_with_index<sizeof...(Ts)>(index(),
            destroy_impl{*this});
    }

    template <class... Us>
    constexpr void copy_construct(const variant<Us...>& rhs)
    {
        boost::mp11::mp_with_index<sizeof...(Us)>(rhs.index(),
            copy_construct_impl<Us...>{*this, rhs});
    }

    template <class... Us>
    constexpr void move_construct(variant<Us...>& rhs)
    {
        boost::mp11::mp_with_index<sizeof...(Us)>(rhs.index(),
            move_construct_impl<Us...>{*this, rhs});
    }

public:
    constexpr variant()
        : tagged_union(default_xltype::value)
    {}

    ~variant()
    {
        if (xltype_ & xlbitDLLFree)
            return; // destroy in xlAutoFree12
        destroy();
    }

    void release()
    {
        destroy();
    }
    
    template <class... Us,
        typename = std::enable_if_t<boost::mp11::mp_all<boost::mp11::mp_contains<types, Us>...>::value>>
    constexpr variant(const variant<Us...>& rhs)
        noexcept(boost::mp11::mp_all<std::is_nothrow_copy_constructible<Us>...>::value)
    {
        copy_construct(rhs);
    }

    template <class... Us,
        typename = std::enable_if_t<boost::mp11::mp_all<boost::mp11::mp_contains<types, Us>...>::value>>
    constexpr variant(variant<Us...>&& rhs)
        noexcept(boost::mp11::mp_all<std::is_nothrow_move_constructible<Us>...>::value)
    {
        move_construct(rhs);
    }

    template <class... Us,
        typename = std::enable_if_t<boost::mp11::mp_all<boost::mp11::mp_contains<types, Us>...>::value>>
    variant& operator=(const variant<Us...>& rhs)
    {
        destroy();
        copy_construct(rhs);
        return *this;
    }

    template <class... Us,
        typename = std::enable_if_t<boost::mp11::mp_all<boost::mp11::mp_contains<types, Us>...>::value>>
    variant& operator=(variant<Us...>&& rhs)
    {
        destroy();
        move_construct(rhs);
        return *this;
    }

    constexpr bool operator==(const variant& rhs) const
    {
        if (xltype_ != rhs.xltype_)
            return false;
        return boost::mp11::mp_with_index<sizeof...(Ts)>(rhs.index(),
            equals{*this, rhs});       
    }

    // Converting constructors for stored value types.
    template <class T,
        std::enable_if_t<boost::mp11::mp_contains<types, T>::value>* = nullptr>
    constexpr variant(const T& rhs)
    {
        ::new(&val_) T(rhs);
        set_xltype(T::xltype::value);
    }

    template <class T,
        std::enable_if_t<boost::mp11::mp_contains<types, T>::value>* = nullptr>
    constexpr variant(T&& rhs)
    {
        ::new(&val_) T(std::forward<T>(rhs));
        set_xltype(T::xltype::value);
    }
    
    // Monostate converting constructors for arguments.
    template <class... Args,
        std::enable_if_t<is_monostate::value && !boost::mp11::mp_or<boost::mp11::mp_contains<types, Args>...>::value>* = nullptr>
    constexpr variant(const Args&... args)
    {
        using T = boost::mp11::mp_front<types>;
        ::new(&val_) T(args...);
        set_xltype(T::xltype::value);
    }

    template <class... Args,
        std::enable_if_t<is_monostate::value && !boost::mp11::mp_or<boost::mp11::mp_contains<types, Args>...>::value>* = nullptr>
    constexpr variant(Args&&... args)
    {
        using T = boost::mp11::mp_front<types>;
        ::new(&val_) T(std::forward<Args>(args)...);
        set_xltype(T::xltype::value);
    }
    
    // Construct value in-place.
    template <class Tag, class... Args>
    void emplace(Args&&... args) noexcept
    {
        destroy();
        using T = typename map<Tag>::value;
        ::new(&val_) T(std::forward<Args>(args)...);
        set_xltype(T::xltype::value);
    }

    // Returns index of current stored type.
    constexpr std::size_t index() const noexcept
    {
        using N = boost::mp11::mp_size<types>;
        using Is = boost::mp11::mp_iota<N>;
        
        const uint32_t current = xltype();
        std::size_t result = N::value;
        boost::mp11::mp_for_each<Is>([&](auto I) {
            if (boost::mp11::mp_at_c<types, I>::xltype::value == current)
                result = I;
        });

        return result;
    }

    // Tag-based value accessor.
    template <class Tag>
    constexpr auto& get() noexcept
    {
        XLL_ASSERT(xltype() == Tag::xltype::value);
        using T = typename map<Tag>::value;
        return *launder_cast<T*>(&val_);
    }

    template <class Tag>
    constexpr auto& get() const noexcept
    {
        XLL_ASSERT(xltype() == Tag::xltype::value);
        using T = typename map<Tag>::value;
        return *launder_cast<const T*>(&val_);
    }

    // Monostate value accessor.
    template <class V = boost::mp11::mp_front<types>>
    constexpr std::enable_if_t<is_monostate::value, V>& value() noexcept
    {
        return *launder_cast<V*>(&val_);
    }

    template <class V = boost::mp11::mp_front<types>>
    constexpr const std::enable_if_t<is_monostate::value, V>& value() const noexcept
    {
        return *launder_cast<const V*>(&val_);
    }
};

} // namespace detail
} // namespace xll
