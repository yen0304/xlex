# Unsafe Rust Guide

## When to Use Unsafe

Use `unsafe` only when:
1. Interfacing with C/FFI
2. Implementing low-level data structures
3. Performance-critical code where safe alternatives are too slow
4. Accessing hardware or OS primitives

## Unsafe Operations

```rust
unsafe {
    // 1. Dereference raw pointers
    let ptr: *const i32 = &x;
    let val = *ptr;
    
    // 2. Call unsafe functions
    libc::malloc(size);
    
    // 3. Access mutable statics
    static mut COUNTER: i32 = 0;
    COUNTER += 1;
    
    // 4. Implement unsafe traits
    // 5. Access union fields
}
```

## Raw Pointers

```rust
let mut x = 42;

// Creating raw pointers (safe)
let ptr: *const i32 = &x;
let mut_ptr: *mut i32 = &mut x;

// Dereferencing (unsafe)
unsafe {
    println!("{}", *ptr);
    *mut_ptr = 100;
}

// Pointer arithmetic
unsafe {
    let arr = [1, 2, 3];
    let ptr = arr.as_ptr();
    let second = *ptr.add(1);  // arr[1]
}
```

## Unsafe Functions

```rust
/// # Safety
/// - `ptr` must be valid and properly aligned
/// - `ptr` must point to initialized memory
/// - No other references to this memory may exist
unsafe fn dangerous(ptr: *mut i32) {
    *ptr = 42;
}

// Safe wrapper pattern
pub fn safe_wrapper(data: &mut i32) {
    // SAFETY: We have exclusive access via &mut
    unsafe { dangerous(data) }
}
```

## Unsafe Traits

```rust
/// # Safety
/// Implementors must ensure memory is properly aligned
unsafe trait Aligned {
    fn alignment() -> usize;
}

unsafe impl Aligned for u64 {
    fn alignment() -> usize { 8 }
}
```

## Common Patterns

### Transmute (Type Punning)
```rust
// Prefer: from_ne_bytes, to_ne_bytes
let bytes: [u8; 4] = unsafe { std::mem::transmute(42i32) };

// Better alternative
let bytes = 42i32.to_ne_bytes();
```

### ManuallyDrop
```rust
use std::mem::ManuallyDrop;

let mut data = ManuallyDrop::new(String::from("hello"));
// data won't be dropped automatically

// Manual cleanup when needed
unsafe { ManuallyDrop::drop(&mut data); }
```

### MaybeUninit
```rust
use std::mem::MaybeUninit;

let mut arr: [MaybeUninit<i32>; 10] = unsafe {
    MaybeUninit::uninit().assume_init()
};

for (i, elem) in arr.iter_mut().enumerate() {
    elem.write(i as i32);
}

// Now safe to use
let arr: [i32; 10] = unsafe {
    std::mem::transmute(arr)
};
```

## Safety Documentation

Always document safety requirements:

```rust
/// Copies `count` elements from `src` to `dst`.
///
/// # Safety
///
/// - `src` must be valid for reads of `count * size_of::<T>()` bytes
/// - `dst` must be valid for writes of `count * size_of::<T>()` bytes
/// - Both `src` and `dst` must be properly aligned
/// - The regions must not overlap
pub unsafe fn copy<T>(src: *const T, dst: *mut T, count: usize) {
    std::ptr::copy_nonoverlapping(src, dst, count);
}
```

## Miri for Validation

```bash
# Install Miri
rustup +nightly component add miri

# Run tests with Miri
cargo +nightly miri test

# Check for undefined behavior
cargo +nightly miri run
```

Miri detects:
- Use after free
- Out-of-bounds access
- Invalid alignment
- Data races
- Memory leaks
